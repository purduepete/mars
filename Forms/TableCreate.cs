using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MarsAddIn.HelperClasses;
using json = MarsAddIn.HelperClasses.MarsSettings;
using db = MarsAddIn.HelperClasses.DatabaseClass;
using helper = MarsAddIn.HelperClasses.HelperClass;

// ReSharper disable CoVariantArrayConversion

namespace MarsAddIn.Forms
{
    public partial class TableCreate : Form
    {
        private int _comboBoxCount;

        public bool showForm = true;
        public TableCreate()
        {
            InitializeComponent();
            var settingsLocation = Environment.GetEnvironmentVariable("LOCALAPPDATA") + @"\MarsAddIn\settings.json";

            if (!File.Exists(settingsLocation))
            {
                showForm = false;
                return;
            }

            var json = JsonConvert.DeserializeObject<json>(File.ReadAllText(settingsLocation));
            var databaseServers = json.DatabaseServers;

            if (databaseServers == null)
            {
                MessageBox.Show(@"The settings array, 'DatabaseServers' was not found; enter and retry", caption: @"Missing Setting");
                showForm = false;
                return;
            }

            btnDoAction.Left = (Width - btnDoAction.Width) / 2;

            lblColumName.Width = 100;
            lblColumName.Left = tboxTableName.Left;
            lblColumName.AutoSize = false;
            lblColumName.Top = btnPlus.Top + 2;

            lblDataType.Width = 80;
            lblDataType.Left = lblColumName.Left + lblColumName.Width + 10;
            lblDataType.AutoSize = false;
            lblDataType.Top = btnPlus.Top + 2;

            lblColumnSize.Width = 40;
            lblColumnSize.Left = lblDataType.Left + lblDataType.Width + 10;
            lblColumnSize.AutoSize = false;
            lblColumnSize.Top = btnPlus.Top + 2;

            lblId.Width = 30;
            lblId.Left = lblColumnSize.Left + lblColumnSize.Width + 10;
            lblId.AutoSize = false;
            lblId.Top = btnPlus.Top + 2;

            lblPK.Width = 30;
            lblPK.Left = lblId.Left + lblId.Width + 7;
            lblPK.AutoSize = false;
            lblPK.Top = btnPlus.Top + 2;

            lblNull.Width = 52;
            lblNull.Left = lblPK.Left + lblPK.Width + 5;
            lblNull.AutoSize = false;
            lblNull.Top = btnPlus.Top + 2;

            comboServer.Items.AddRange(databaseServers.ToArray());
            _comboBoxCount = 0;
        }

        private void btnPlus_Click(object sender, EventArgs e)
        {
            helper.OptimizationOn(true);

            /* Get date for source/target */

            Height += 25;
            btnDoAction.Top += 25;
            _comboBoxCount++;

            var top = (_comboBoxCount - 0) * 25;

            var columnColumnName = new TextBox
            {
                Top = btnPlus.Top + top,
                Left = lblColumName.Left,
                Height = 18,
                Width = 100,
                Name = $"ColumnName_{_comboBoxCount}"
            };
            Controls.Add(columnColumnName);


            var columnDataType = new ComboBox
            {
                Top = btnPlus.Top + top,
                Left = lblDataType.Left,
                Height = 18,
                Width = 80,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Name = $"DataType_{_comboBoxCount}"
            };
            columnDataType.Items.AddRange(new[] { "varchar", "int", "datetime" });
            columnDataType.SelectedIndex = 0;

            columnDataType.SelectionChangeCommitted += dataType_Commited;
            Controls.Add(columnDataType);


            var columnColumnSize = new TextBox
            {
                Top = btnPlus.Top + top,
                Left = lblColumnSize.Left,
                Height = 18,
                Width = 40,
                Name = $"ColumnSize_{_comboBoxCount}",
                Text = @"255"
            };
            Controls.Add(columnColumnSize);


            var columnIsId = new CheckBox()
            {
                Top = btnPlus.Top + top + 2,
                Left = lblId.Left + 1,
                Width = 15,
                Height = 15,
                Name = $"ColumnIsId_{_comboBoxCount}",
                Enabled = false
            };
            columnIsId.Click += chkBox_Click;
            Controls.Add(columnIsId);


            var columnIsPk = new CheckBox()
            {
                Top = btnPlus.Top + top + 2,
                Left = lblPK.Left + 1,
                Width = 15,
                Height = 15,
                Name = $"ColumnIsPK_{_comboBoxCount}"
            };
            columnIsPk.Click += chkBox_Click;
            Controls.Add(columnIsPk);

            var columnNullable = new CheckBox()
            {
                Top = btnPlus.Top + top + 2,
                Left = lblNull.Left + (lblNull.Width - 18) / 2,
                Width = 15,
                Height = 15,
                Name = $"ColumnNullable_{_comboBoxCount}",
                Checked = true

            };
            Controls.Add(columnNullable);

            helper.OptimizationOff(true);
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (_comboBoxCount < 1) return;

            helper.OptimizationOn(true);

            var searchArray = new List<string> { $"ColumnName_{_comboBoxCount}", $"DataType_{_comboBoxCount}", $"ColumnSize_{_comboBoxCount}", $"ColumnIsId_{_comboBoxCount}", $"ColumnIsPK_{_comboBoxCount}", $"ColumnNullable_{_comboBoxCount}" };

            for (var i = Controls.Count - 1; i != -1; i--)
            {
                switch (Controls[i])
                {
                    case ComboBox indComboBox:
                        if (searchArray.Contains(indComboBox.Name))
                            indComboBox.Dispose();
                        break;
                    case TextBox indTextBox:
                        if (searchArray.Contains(indTextBox.Name))
                            indTextBox.Dispose();
                        break;

                    case CheckBox indCheckBox:
                        if (searchArray.Contains(indCheckBox.Name))
                            indCheckBox.Dispose();
                        break;
                }
            }

            Height -= 25;
            btnDoAction.Top -= 25;
            _comboBoxCount--;

            helper.OptimizationOff(true);
        }

        private void btnDoAction_Click(object sender, EventArgs e)
        {
            helper.OptimizationOn();

            var tblObject = new TableObject { ServerName = comboServer.SelectedItem.ToString(), DatabaseName = comboDatabase.SelectedItem.ToString(), TableName = tboxTableName.Text, Recreate = cboxRecreate.Checked };

            for (var j = 1; j <= _comboBoxCount; j++)
            {
                var searchArray = new List<string> { $"ColumnName_{j}", $"DataType_{j}", $"ColumnSize_{j}", $"ColumnIsId_{j}", $"ColumnIsPK_{j}", $"ColumnNullable_{j}" };
                for (var i = 0; i < Controls.Count; i++)
                {
                    var control = Controls[i];
                    if (searchArray.All(ar => control.Name != ar)) continue;

                    switch (control.Name.Replace("_", "").Replace(j.ToString(), ""))
                    {
                        case "ColumnName":
                            tblObject.ColumnName.Add(control.Text);
                            break;

                        case "DataType":
                            var dt = (((ComboBox)control).SelectedItem).ToString();
                            tblObject.DataType.Add(dt.ToUpper());
                            break;

                        case "ColumnSize":
                            tblObject.ColumnSize.Add(int.TryParse(control.Text, out var dataLength)
                                ? dataLength
                                : 0);
                            break;

                        case "ColumnIsId":
                            tblObject.IsId.Add(((CheckBox)control).Checked);
                            break;

                        case "ColumnIsPK":
                            tblObject.IsPrimaryKey.Add(((CheckBox)control).Checked);
                            break;

                        case "ColumnNullable":
                            tblObject.Nullable.Add(((CheckBox)control).Checked);
                            break;

                    }
                }
            }

            if (tblObject.ColumnName.Count == 0)
            {
                MessageBox.Show(@"ERROR: Columns needed for table creation, not found", caption: @"Info");
                return;
            }

            db.CreateTable(tblObject);

            helper.OptimizationOff();
        }


        private void ClearControls()
        {
            //return;

            if (_comboBoxCount == 0) return;

            var searchArray = new List<string> { "ColumnName_", "DataType_", "ColumnSize_", "ColumnIsId_", "ColumnIsPK_", "ColumnNullable_" };
            for (var i = Controls.Count - 1; i != -1; i--)
            {
                switch (Controls[i])
                {
                    case ComboBox indComboBox:
                        if (searchArray.Any(ar => indComboBox.Name.Contains(ar)))
                            indComboBox.Dispose();
                        break;
                    case TextBox indTextBox:
                        if (searchArray.Any(ar => indTextBox.Name.Contains(ar)))
                            indTextBox.Dispose();
                        break;

                    case CheckBox indCheckBox:
                        if (searchArray.Any(ar => indCheckBox.Name.Contains(ar)))
                            indCheckBox.Dispose();
                        break;
                }
            }

            Height -= _comboBoxCount * 25;
            btnDoAction.Top -= _comboBoxCount * 25;
            _comboBoxCount = 0;
        }

        private void ToggleActionButton()
        {
            btnDoAction.Enabled = (comboServer.SelectedIndex != 1 && comboDatabase.SelectedIndex != -1 && tboxTableName != null && tboxTableName.Text.Length > 0);
        }

        private void UnCheckOthers(string controlName)
        {
            var searchString = controlName.Substring(0, controlName.LastIndexOf("_", StringComparison.Ordinal));
            for (var i = Controls.Count - 1; i != -1; i--)
            {
                var indCheckBox = Controls[i] as CheckBox;
                if (indCheckBox == null) continue;

                if (indCheckBox.Name.Contains(searchString) && indCheckBox.Name != controlName)
                    indCheckBox.Checked = false;
            }
        }

        private void DataTypeChange(ComboBox control)
        {
            var controlName = control.Name;
            if (!int.TryParse(controlName.Substring(controlName.LastIndexOf("_", StringComparison.Ordinal) + 1), out var controlNum)) return;

            for (var i = Controls.Count - 1; i != -1; i--)
            {

                switch (Controls[i])
                {
                    case TextBox indTextBox:
                        if (indTextBox.Name == $"ColumnSize_{controlNum}")
                            if ((string)control.SelectedItem != "varchar")
                            {
                                indTextBox.Text = "";
                                indTextBox.Enabled = false;
                            }
                            else
                            {
                                indTextBox.Text = @"255";
                                indTextBox.Enabled = true;
                            }
                        break;

                    case CheckBox indCheckBox:
                        if (indCheckBox.Name == $"ColumnIsId_{controlNum}")
                            if ((string)control.SelectedItem != "int")
                            {
                                indCheckBox.Enabled = false;
                                indCheckBox.Checked = false;
                            }
                            else
                                indCheckBox.Enabled = true;
                        break;
                }
            }
        }


        private void dataType_Commited(object sender, EventArgs e)
        {
            var comboBox = sender as ComboBox;
            DataTypeChange(comboBox);
        }

        private void chkBox_Click(object sender, EventArgs e)
        {
            var chkBox = sender as CheckBox;
            UnCheckOthers(chkBox?.Name);
        }

        private void tboxTableName_TextChanged(object sender, EventArgs e)
        {
            ToggleActionButton();
        }

        private void comboServerChanged_Committed(object sender, EventArgs e)
        {
            ToggleActionButton();

            var server = comboServer.SelectedItem.ToString();

            comboDatabase.Items.Clear();

            var dbs = db.GetDatabases(db.GetDatabaseObjects(server));

            if (dbs.Length > 0)
                comboDatabase.Items.AddRange(dbs);
        }

        private void comboDatabaseChanged_Committed(object sender, EventArgs e)
        {
            ToggleActionButton();
        }

    }
}
