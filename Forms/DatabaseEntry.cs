using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MarsAddIn.HelperClasses;
using json = MarsAddIn.HelperClasses.MarsSettings;
using db = MarsAddIn.HelperClasses.DatabaseClass;
using helper = MarsAddIn.HelperClasses.HelperClass;
using Excel = Microsoft.Office.Interop.Excel;
// ReSharper disable CoVariantArrayConversion

namespace MarsAddIn.Forms
{
    public partial class DatabaseEntry : Form
    {
        private int _comboBoxCount;
        public string CustomQuery;

        private List<string> _excelHeaderList = new List<string>();
        private List<DatabaseObject> _dbColumnList = new List<DatabaseObject>();
        private List<DatabaseObject> _dbColumnListAll = new List<DatabaseObject>();

        public bool showForm = true;

        public DatabaseEntry(string intType = "Write")
        {
            var isRead = intType != "Write";

            InitializeComponent();

            if (!File.Exists(helper.SettingsPath))
            {
                showForm = false;
                return;
            }

            var json = JsonConvert.DeserializeObject<json>(File.ReadAllText(helper.SettingsPath));
            var databaseServers = json.DatabaseServers;

            if (databaseServers == null)
            {
                MessageBox.Show(@"The settings array, 'DatabaseServers' was not found; enter and retry", caption: @"Missing Setting");
                showForm = false;
                return;
            }

            comboServer.Items.AddRange(databaseServers.ToArray());
            _comboBoxCount = 0;

            if (!isRead)
            {
                Excel.Worksheet wks = Globals.ThisAddIn.Application.ActiveSheet;

                foreach (Excel.Name nameValue in wks.Names)
                {
                    var nameValueShort = nameValue.Name.Replace(wks.Name, "").Replace("!", "").Replace("'", "");
                    var val = nameValue.Value.Replace("=", "").Replace("\"", "");

                    switch (nameValueShort)
                    {
                        case "DBServer":
                            comboServer.SelectedItem = val;
                            PopulateDatabaseNames();
                            break;
                        case "DBName":
                            comboDatabase.SelectedItem = val;
                            PopulateTablesNames();
                            break;
                        case "TableName":
                            comboTable.SelectedItem = val;
                            break;
                        case "AutoTrunc":
                            cBoxTrunc.Checked = val == "1" || val == "yes" || val == "true";
                            break;
                    }
                }

                return;
            }

            var widthMinus = 75;
            Text = @"Read from Database";
            btnDoAction.Text = @"Read Data";
            btnDoAction.Left = btnDoAction.Left - widthMinus / 2;

            lblTarget.Visible = false;

            cBoxTrunc.Visible = false;
            cboxHeaderRow.Visible = false;

            lblReturnRows.Visible = true;
            txtNumRows.Visible = true;

            btnPlusPlus.Visible = false;

            Width = Width - widthMinus;
            MaximumSize = new Size(MaximumSize.Width - widthMinus, MaximumSize.Height);
            MinimumSize = new Size(MinimumSize.Width - widthMinus, MinimumSize.Height);

            Name = @"DatabaseRead";
        }

        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (Name == "DatabaseEntry")
                _excelHeaderList = db.GetExcelColumns(_excelHeaderList, cboxHeaderRow.Checked);

            if (_dbColumnList == null || _dbColumnList.Count == 0)
                UpdateDbColumnList();

            var maxColumns = Name == "DatabaseEntry" ? _dbColumnListAll.Count : _dbColumnList.Count;
            if (maxColumns == _comboBoxCount || maxColumns == 0) return;

            helper.OptimizationOn(true);
            AddColumnRow();
            helper.OptimizationOff(true);
        }

        private void AddColumnRow()
        {
            Height += 25;
            btnDoAction.Top += 25;
            _comboBoxCount++;

            var columnComboBoxLeft = new ComboBox
            {
                Top = (btnDoAction.Top - 40),
                Left = comboTable.Left,
                Height = 18,
                Width = 148,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Name = $"ComboBoxLeft_{_comboBoxCount}",
            };
            Controls.Add(columnComboBoxLeft);

            if (Name == "DatabaseRead")
            {
                var columnsR = _dbColumnList.Select(x => x.ColumnName).ToArray();
                columnComboBoxLeft.Items.AddRange(columnsR);

                if (_comboBoxCount <= _dbColumnList.Count)
                    columnComboBoxLeft.SelectedIndex = _comboBoxCount - 1;

                return;
            }

            var columnComboBoxRight = new ComboBox
            {
                Top = columnComboBoxLeft.Top,
                Left = ((comboTable.Left + comboTable.Width) - columnComboBoxLeft.Width),
                Height = columnComboBoxLeft.Height,
                Width = columnComboBoxLeft.Width,
                DropDownStyle = columnComboBoxLeft.DropDownStyle,
                Name = $"ComboBoxRight_{_comboBoxCount}"
            };
            Controls.Add(columnComboBoxRight);
            columnComboBoxRight.SelectionChangeCommitted += SelectionChanged;

            var lblText = "";
            if (_comboBoxCount <= _dbColumnListAll.Count)
                lblText = _dbColumnListAll[_comboBoxCount - 1].DataType;

            var labelDbColumn = new Label
            {
                Top = columnComboBoxRight.Top + 2,
                Left = columnComboBoxRight.Left + columnComboBoxRight.Width + 10,
                Height = 18,
                Width = 85,
                Name = $"lblDB_{_comboBoxCount}",
                Text = lblText,
                TextAlign = ContentAlignment.MiddleLeft
            };
            Controls.Add(labelDbColumn);

            columnComboBoxLeft.Items.AddRange(_excelHeaderList.ToArray());

            var columnsW0 = new List<string> { "" };
            columnsW0.AddRange(_dbColumnListAll.Select(x => x.ColumnName));

            columnComboBoxRight.Items.AddRange(columnsW0.ToArray());

            if (_comboBoxCount <= _excelHeaderList.Count)
                columnComboBoxLeft.SelectedIndex = _comboBoxCount - 1;

            if (columnComboBoxLeft.SelectedItem != null)
            {
                var matchIndex = columnsW0.FindIndex(x => string.Equals(x, columnComboBoxLeft.SelectedItem.ToString(), StringComparison.OrdinalIgnoreCase));
                if (matchIndex >= 0)
                    columnComboBoxRight.SelectedIndex = matchIndex;
                else if (_comboBoxCount <= _dbColumnListAll.Count)
                    columnComboBoxRight.SelectedIndex = _comboBoxCount;
            }
            else if (_comboBoxCount <= _dbColumnListAll.Count)
                columnComboBoxRight.SelectedIndex = _comboBoxCount;
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (_comboBoxCount < 1) return;

            helper.OptimizationOn(true);

            var searchArray = new List<string> { $"ComboBoxLeft_{_comboBoxCount}", $"ComboBoxRight_{_comboBoxCount}", $"lblDB_{_comboBoxCount}" };
            for (var i = Controls.Count - 1; i != -1; i--)
            {
                var indComboBox = Controls[i] as ComboBox;
                if (indComboBox == null) continue;

                if (searchArray.Contains(indComboBox.Name))
                    indComboBox.Dispose();
            }

            for (var i = Controls.Count - 1; i != -1; i--)
            {
                var indLblBox = Controls[i] as Label;
                if (indLblBox == null) continue;

                if (searchArray.Contains(indLblBox.Name))
                    indLblBox.Dispose();
            }

            Height -= 25;
            btnDoAction.Top -= 25;
            _comboBoxCount--;

            helper.OptimizationOff(true);
        }

        private void btnDoAction_Click(object sender, EventArgs e)
        {
            helper.OptimizationOn();

            string query;

            if (Name == "DatabaseRead")

            {
                query = CustomQuery.Length > 0 ? CustomQuery : ConstructReadQuery();

                Excel.Range startRange = Globals.ThisAddIn.Application.Selection;
                startRange = startRange.Offset[1, 0];

                db.ReadData(comboServer.SelectedItem, query, startRange, CustomQuery.Length > 0);

                Excel.Worksheet wks = Globals.ThisAddIn.Application.ActiveSheet;
                wks.Names.Add(Name: "DBServer", RefersTo: comboServer.SelectedItem.ToString());
                wks.Names.Add(Name: "DBName", RefersTo: comboDatabase.SelectedItem.ToString());
                wks.Names.Add(Name: "TableName", RefersTo: comboTable.SelectedItem.ToString());

                helper.OptimizationOff();
            }
            else
            {
                var startTime = DateTime.Now;

                query = ConstructWriteQuery();
                var selectedItemsSource = new List<string>();
                var selectedItemsTarget = new List<DatabaseObject>();

                for (var i = 0; i < Controls.Count; i++)
                {
                    var indComboBox = Controls[i] as ComboBox;
                    if (indComboBox == null) continue;

                    if (indComboBox.Name.Contains("ComboBoxLeft") && indComboBox.SelectedItem != null)
                        selectedItemsSource.Add(indComboBox.SelectedItem.ToString());


                    if (indComboBox.Name.Contains("ComboBoxRight") && indComboBox.SelectedItem != null && (string)indComboBox.SelectedItem != "")
                        selectedItemsTarget.Add(_dbColumnListAll.First(db =>
                            db.ColumnName == indComboBox.SelectedItem.ToString()));


                }

                Excel.Range rangeFinal = null;

                Excel.Worksheet wks = Globals.ThisAddIn.Application.ActiveSheet;
                var lastColumnInt = wks.Cells[1, 255].End(Excel.XlDirection.xlToLeft).Column();
                Excel.Range headerRange = wks.Range[wks.Cells[1, 1], wks.Cells[1, lastColumnInt]];

                var maxRow = 0;
                foreach (Excel.Range indCell in headerRange)
                {
                    if (!selectedItemsSource.Contains(indCell.Value)) continue;

                    var lastRowInt = Globals.ThisAddIn.Application.ActiveSheet.Cells[1048576, indCell.Column].End(Excel.XlDirection.xlUp).Row();
                    if (lastRowInt > maxRow)
                        maxRow = lastRowInt;
                }

                foreach (Excel.Range indCell in headerRange)
                {
                    if (!selectedItemsSource.Contains(indCell.Value)) continue;

                    Excel.Range colRange = wks.Range[wks.Cells[1, indCell.Column], wks.Cells[maxRow, indCell.Column]];
                    rangeFinal = rangeFinal == null ? colRange : Globals.ThisAddIn.Application.Union(rangeFinal, colRange);
                }

                if (rangeFinal == null)
                {
                    helper.OptimizationOff();
                    return;
                }

                var writeRslt = db.WriteData(comboServer.SelectedItem, comboDatabase.SelectedItem,
                    comboTable.SelectedItem, query, rangeFinal, selectedItemsTarget);


                for (var i = Controls.Count - 1; i != 0; i--)
                {
                    var indLblBox = Controls[i] as Label;
                    if (indLblBox == null) continue;

                    if (indLblBox.Name.Contains("lblDB"))
                    {
                        indLblBox.BackColor = BackColor;
                        indLblBox.ForeColor = Color.Black;
                    }

                    foreach (var truncColumn in writeRslt.TruncColumns)
                        if (indLblBox.Name == $"lblDB_{truncColumn + 1}")
                            indLblBox.BackColor = Color.LightGray;


                    if (writeRslt.BadTypeColumns.Count <= 0) continue;

                    foreach (var bdtColumn in writeRslt.BadTypeColumns)
                        if (indLblBox.Name == $"lblDB_{bdtColumn + 1}")
                            indLblBox.ForeColor = Color.Red;
                }

                helper.OptimizationOff();


                var message = $"Rows added: {writeRslt.WriteRows}";
                if (writeRslt.TruncColumns.Count > 0)
                    message += $"{Environment.NewLine}Columns truncated: {writeRslt.TruncColumns.Count}";

                if (writeRslt.BadTypeColumns.Count > 0)
                    message += $"{Environment.NewLine}Columns with type mismatch: {writeRslt.BadTypeColumns.Count}";

                var timeToLoad = DateTime.Now - startTime;
                message += $"{Environment.NewLine}Time to load: {timeToLoad.Seconds} second";
                message += timeToLoad.Seconds == 1 ? "" : "s";

                MessageBox.Show(message, caption: @"Table Load");

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }



        private void comboServerChanged_Committed(object sender, EventArgs e)
        {
            ClearControls();
            QueryEditorManipulations();

            comboServer.DroppedDown = false;
            BeginInvoke(new Action(() => PopulateDatabaseNames()));
        }

        private void PopulateDatabaseNames()
        {
            var server = comboServer.SelectedItem.ToString();

            comboDatabase.Items.Clear();
            comboTable.Items.Clear();

            var dbs = db.GetDatabases(db.GetDatabaseObjects(server));

            if (dbs.Length > 0)
                comboDatabase.Items.AddRange(dbs);
        }

        private void comboDatabaseChanged_Committed(object sender, EventArgs e)
        {
            ClearControls();
            QueryEditorManipulations();

            PopulateTablesNames();
        }

        private void PopulateTablesNames()
        {
            var server = comboServer.SelectedItem.ToString();
            var database = comboDatabase.SelectedItem.ToString();

            comboTable.Items.Clear();

            var tbls = db.GetTables(db.GetDatabaseObjects(server, database, includeView: Name == "DatabaseRead"));

            if (tbls.Length > 0)
                comboTable.Items.AddRange(tbls);
        }

        private void comboTableChanged_Committed(object sender, EventArgs e)
        {
            ClearControls();
            QueryEditorManipulations();

            UpdateDbColumnList();
        }



        private void lblSource_Click(object sender, EventArgs e)
        {
            string[] itemArray;

            if (Name == "DatabaseEntry")
            {
                _excelHeaderList.Clear();
                _excelHeaderList = db.GetExcelColumns(_excelHeaderList, cboxHeaderRow.Checked);
                itemArray = _excelHeaderList.ToArray();
            }
            else
            {
                UpdateDbColumnList();
                var columns = _dbColumnList.Select(x => x.ColumnName).ToArray();
                itemArray = columns;
            }

            for (var i = 0; i < Controls.Count; i++)
            {
                var indComboBox = Controls[i] as ComboBox;
                if (indComboBox == null) continue;

                if (!indComboBox.Name.Contains("ComboBoxLeft")) continue;

                indComboBox.Items.Clear();
                indComboBox.Items.AddRange(itemArray);
            }

        }

        private void lblSource_DblClick(object sender, EventArgs e)
        {
            UpdateSelected();
        }

        private void lblTarget_Click(object sender, EventArgs e)
        {
            var server = comboServer.SelectedItem.ToString();
            var database = comboDatabase.SelectedItem.ToString();
            var table = comboTable.SelectedItem.ToString();

            _dbColumnList = db.GetDatabaseObjects(server, database, table);

            if (_dbColumnList.Count <= 0) return;

            for (var i = 0; i < Controls.Count; i++)
            {
                var indComboBox = Controls[i] as ComboBox;
                if (indComboBox == null) continue;

                if (!indComboBox.Name.Contains("ComboBoxRight")) continue;

                indComboBox.Items.Clear();
                var columns = _dbColumnList.Select(x => x.ColumnName).ToArray();
                indComboBox.Items.AddRange(columns);
            }
        }

        private void lblTarget_DblClick(object sender, EventArgs e)
        {
            UpdateSelected("ComboBoxRight");
        }


        private void UpdateDbColumnList(bool includeId = true)
        {
            if (Name == "DatabaseEntry")
                includeId = false;

            if (comboServer.SelectedIndex == -1 || comboDatabase.SelectedIndex == -1 || comboTable.SelectedIndex == -1) return;

            var server = comboServer.SelectedItem.ToString();
            var database = comboDatabase.SelectedItem.ToString();
            var table = comboTable.SelectedItem.ToString();

            _dbColumnList = db.GetDatabaseObjects(server, database, table);
            _dbColumnListAll = new List<DatabaseObject>(_dbColumnList);

            if (!includeId)
            {
                var dbColumnListFiltered = (from indItem in _dbColumnList
                                            where !indItem.IsIdentity
                                            select new DatabaseObject
                                            {
                                                ColumnName = indItem.ColumnName,
                                                DataType = indItem.DataType,
                                                ColumnLength = indItem.ColumnLength,
                                                DatabaseName = indItem.DatabaseName,
                                                TableName = indItem.TableName
                                            }).ToList();
                _dbColumnList = dbColumnListFiltered;
            }

            btnDoAction.Enabled = true;
        }

        private void UpdateSelected(string combBoxName = "ComboBoxLeft")
        {
            var j = 0;
            for (var i = 0; i < Controls.Count; i++)
            {
                var indComboBox = Controls[i] as ComboBox;
                if (indComboBox == null) continue;

                if (!indComboBox.Name.Contains(combBoxName)) continue;

                if (j < indComboBox.Items.Count)
                    indComboBox.SelectedIndex = j;
                j++;
            }
        }



        private string ConstructReadQuery()
        {
            string columns = null;
            var numRows = txtNumRows.Text;
            var query = "SELECT " + (numRows.Length > 0 ? $" TOP {numRows} " : " ");

            Excel.Range startRange = Globals.ThisAddIn.Application.Selection;

            var j = 0;
            for (var i = 0; i < Controls.Count; i++)
            {
                var indComboBox = Controls[i] as ComboBox;
                if (indComboBox == null) continue;

                if (!indComboBox.Name.Contains("ComboBoxLeft")) continue;

                if (indComboBox.SelectedItem == null) continue;

                columns += $"[{indComboBox.SelectedItem}], ";
                startRange.Offset[0, j].Value = indComboBox.SelectedItem;
                j++;
            }

            return columns == null ? null : $"{query} {columns.Substring(0, columns.Length - 2)} FROM [{comboDatabase.SelectedItem}].{comboTable.SelectedItem}";
        }

        private string ConstructWriteQuery()
        {
            var query = "";

            if (cBoxTrunc.Checked)
                query = string.Format("TRUNCATE TABLE [{0}].{1}", comboDatabase.SelectedItem, comboTable.SelectedItem);

            return query;
        }

        private void ClearControls()
        {
            _dbColumnList.Clear();
            _dbColumnListAll.Clear();

            var clrCnt = 0;
            var searchArray = new List<string> { "ComboBoxLeft", "ComboBoxRight", "lblDB" };
            for (var i = Controls.Count - 1; i != 0; i--)
            {
                var indComboBox = Controls[i] as ComboBox;
                if (indComboBox == null) continue;

                if (searchArray.Count(ar => indComboBox.Name.Contains(ar)) == 0) continue;
                indComboBox.Dispose();

                if (!indComboBox.Name.Contains(searchArray.First())) continue;
                clrCnt++;
            }

            for (var i = Controls.Count - 1; i != 0; i--)
            {
                var indLblBox = Controls[i] as Label;
                if (indLblBox == null) continue;

                if (searchArray.Count(ar => indLblBox.Name.Contains(ar)) == 0) continue;
                indLblBox.Dispose();
            }

            Height = Height - (clrCnt * 25);
            btnDoAction.Top = btnDoAction.Top - (clrCnt * 25);
            _comboBoxCount = 0;
        }

        private void SelectionChanged(object sender, EventArgs e)
        {
            var ctrl = (ComboBox)sender;
            var cntrlNum = int.Parse(ctrl.Name.Substring(ctrl.Name.LastIndexOf('_') + 1));

            for (var i = Controls.Count - 1; i != 0; i--)
            {
                var indLblBox = Controls[i] as Label;
                if (indLblBox == null) continue;

                if (!indLblBox.Name.Contains("lblDB")) continue;

                var lblNum = int.Parse(indLblBox.Name.Substring(indLblBox.Name.LastIndexOf('_') + 1));
                if (lblNum != cntrlNum) continue;

                foreach (var dbobj in _dbColumnListAll)
                    if (dbobj.ColumnName == (string)ctrl.SelectedItem)
                    {
                        indLblBox.Text = dbobj.DataType;
                    }

                indLblBox.ForeColor = Color.Black;
                indLblBox.BackColor = BackColor;
                if ((string)ctrl.SelectedItem == "")
                    indLblBox.Text = "";

            }
        }

        private void lblTable_Click(object sender, EventArgs e)
        {
            lblTable.Enabled = false;

            if (comboServer.SelectedIndex == -1 || comboDatabase.SelectedIndex == -1) return;

            var columnList = "";
            if (_dbColumnList.Count > 0)
            {
                _dbColumnList.ForEach(clist => columnList += Environment.NewLine + "\t" + clist.ColumnName + ",");
                columnList = columnList.Substring(0, columnList.Length - 1);
            }
            else
                columnList = $"{Environment.NewLine}\t*";



            var baseQuery = $"SELECT{columnList}{Environment.NewLine}FROM{Environment.NewLine}\t[{comboDatabase.SelectedItem}]";
            if (comboTable.SelectedIndex > -1)
                baseQuery += $".{comboTable.SelectedItem}";


            var qe = new QueryEdit(CustomQuery == "" ? baseQuery : CustomQuery) { StartPosition = FormStartPosition.Manual };
            qe.Location = helper.GetLocation(-(qe.Width + 465), leftMost: false);
            qe.TopMost = true;

            qe.Show(this);
        }

        public void QueryEditorManipulations(string query = "")
        {
            if (query.Length > 0)
                ClearControls();

            txtNumRows.Enabled = query.Length == 0;
            btnMinus.Enabled = query.Length == 0;
            btnPlus.Enabled = query.Length == 0;
            CustomQuery = query;

            lblTable.Enabled = true;
        }

        private void btnPlusPlus_Click(object sender, EventArgs e)
        {
            if (Name == "DatabaseEntry")
                _excelHeaderList = db.GetExcelColumns(_excelHeaderList, cboxHeaderRow.Checked);

            if (_dbColumnList == null || _dbColumnList.Count == 0)
                UpdateDbColumnList();

            var maxColumns = Name == "DatabaseEntry" ? _dbColumnListAll.Count : _dbColumnList.Count;
            if (maxColumns == _comboBoxCount || maxColumns == 0) return;

            helper.OptimizationOn(true);

            var remaining = maxColumns - _comboBoxCount;
            for (var i = 0; i < remaining; i++)
                AddColumnRow();

            helper.OptimizationOff(true);
        }
    }
}
