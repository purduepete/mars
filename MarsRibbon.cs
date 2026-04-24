using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using MarsAddIn.Forms;
using helper = MarsAddIn.HelperClasses.HelperClass;
using group = MarsAddIn.HelperClasses.GroupClass;
using json = MarsAddIn.HelperClasses.MarsSettings;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using XlHAlign = Microsoft.Office.Interop.Excel.XlHAlign;
using XlVAlign = Microsoft.Office.Interop.Excel.XlVAlign;

namespace MarsAddIn
{
    public partial class MarsRibbon
    {
        private void MarsRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            var ribbonSettings = helper.LoadRibbonItemSettings();
            if (ribbonSettings.Count == 0)
            {
                helper.SaveRibbonItemSettings((OfficeRibbon)sender);
                return;
            }

            var ribbonItems = (OfficeRibbon)sender;
            var ribbonTab = ribbonItems.Tabs[1];

            foreach (var ribbonGroup in ribbonTab.Groups)
                foreach (var ribbonGroupItem in ribbonGroup.Items)
                    try
                    {
                        if (ribbonGroupItem.Tag != null && ribbonSettings.ContainsKey(ribbonGroupItem.Tag.ToString()))
                            ribbonGroupItem.Visible = ribbonSettings[ribbonGroupItem.Tag.ToString()];
                    }
                    catch (Exception)
                    {
                        // ignored
                    }
        }

        private void rowShadingBtn_Click(object sender, RibbonControlEventArgs e)
        {

            helper.OptimizationOn();

            try
            {
                Excel.Range selRange = Globals.ThisAddIn.Application.Selection;
                Excel.FormatCondition fc = selRange.FormatConditions.Add(Type: Excel.XlFormatConditionType.xlExpression,
                    Formula1: (selRange.Row % 2 == 1 ? "=MOD(ROW(), 2)" : "=NOT(MOD(ROW(), 2))"));

                fc.Font.Color = Color.Black;
                fc.Interior.Color = Color.FromArgb(1, 217, 217, 217);

                Excel.FormatCondition fc1 = selRange.FormatConditions.Add(Type: Excel.XlFormatConditionType.xlExpression,
                    Formula1: (selRange.Row % 2 == 1 ? "=NOT(MOD(ROW(), 2))" : "=MOD(ROW(), 2)"));

                fc1.Font.Color = Color.Black;
                fc1.Interior.Color = Color.White;
            }
            catch (Exception)
            {
                // ignored
            }
            finally
            {
                helper.OptimizationOff();
            }
        }

        private void btnCreateStructure_Click(object sender, RibbonControlEventArgs e)
        {
            var tblCreateForm = new TableCreate()
            {
                StartPosition = FormStartPosition.Manual
            };

            var frmWidth = tblCreateForm.Width;
            var location = helper.GetLocation(-(frmWidth + 15), leftMost: false);
            tblCreateForm.Location = location;
            tblCreateForm.TopMost = true;

            if (tblCreateForm.showForm)
                tblCreateForm.Show();
        }

        private void addRowsButton_Click(object sender, RibbonControlEventArgs e)
        {
            var addRowsForm = new AddToExcel(addType: "row")
            {
                StartPosition = FormStartPosition.Manual
            };

            var frmWidth = addRowsForm.Width;
            var location = helper.GetLocation(-(frmWidth + 15));
            addRowsForm.Location = location;
            addRowsForm.TopMost = true;

            addRowsForm.Show();
        }

        private void addColumnsButton_Click(object sender, RibbonControlEventArgs e)
        {
            var addColumnsForm = new AddToExcel
            {
                StartPosition = FormStartPosition.Manual
            };

            var frmWidth = addColumnsForm.Width;
            var location = helper.GetLocation(-(frmWidth + 15));
            addColumnsForm.Location = location;
            addColumnsForm.TopMost = true;

            addColumnsForm.Show();
        }

        private void dfltPivot_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range rng = Globals.ThisAddIn.Application.ActiveCell;
            try
            {
                Excel.PivotTable pt = rng.PivotTable;
                pt.HasAutoFormat = false;
                pt.RowGrand = true;
                pt.ColumnGrand = true;
                pt.InGridDropZones = true;
                pt.RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow);
            }
            catch (Exception)
            {
                MessageBox.Show(text: @"A pivot table must be selected for this functionality to work",
                    caption: @"No Pivot Table", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void wksSheetList_Click(object sender, RibbonControlEventArgs e)
        {

            var worksheetListForm = new WorksheetListForm
            {
                StartPosition = FormStartPosition.Manual
            };

            var frmWidth = worksheetListForm.Width;
            var location = helper.GetLocation(-(frmWidth + 15));
            worksheetListForm.Location = location;
            worksheetListForm.TopMost = true;

            worksheetListForm.Show();
        }

        private void wksShowAll_Click(object sender, RibbonControlEventArgs e)
        {
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            foreach (Excel.Worksheet wks in wb.Worksheets)
            {
                wks.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }
        }

        private void wksResize_Click(object sender, RibbonControlEventArgs e)
        {
            var iRow = 0;
            var iColumn = 0;
            Excel.Worksheet wks = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Window aWindow = Globals.ThisAddIn.Application.ActiveWindow;

            foreach (Excel.Range rng in wks.Range["A1:BA500"])
            {
                try
                {
                    if (rng.Value == null) continue;

                    string rngValue = rng.Value.ToString();
                    var rngValueLength = rngValue.Length;

                    if (rngValueLength == 0) continue;
                    iRow = rng.Row > iRow ? rng.Row : iRow;
                    iColumn = rng.Column > iColumn ? rng.Column : iColumn;
                }
                catch (Exception)
                {
                    // ignored
                }
            }

            if (iRow <= 0 || iColumn <= 0) return;

            helper.OptimizationOn();

            try
            {
                Excel.Range rszRange = wks.Range[wks.Range["A1"], wks.Cells[iRow, iColumn]];
                rszRange.Select();
                aWindow.Zoom = true;
                aWindow.Zoom = aWindow.Zoom > 120 ? 120 : aWindow.Zoom;
                wks.Range["A1"].Select();
            }
            catch (Exception)
            {
                // ignored
            }
            finally
            {
                helper.OptimizationOff();
            }
        }

        private void dupPivotBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range rng = Globals.ThisAddIn.Application.ActiveCell;
            try
            {
                Excel.PivotTable pt = rng.PivotTable;

                const string message = "Select a target range";
                const string caption = "Target Range";

                var rngTarget = Globals.ThisAddIn.Application.InputBox(Prompt: message, Title: caption,
                    Default: "", Top: 10, Left: 10, Type: 8);

                pt.TableRange2.Copy(Destination: rngTarget);
            }
            catch (Exception)
            {
                MessageBox.Show(text: @"A pivot table must be selected for this functionality to work",
                    caption: @"No Pivot Table", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void headerFormatBtn_Click(object sender, RibbonControlEventArgs e)
        {
            helper.OptimizationOn();

            try
            {
                Excel.Range selRange = Globals.ThisAddIn.Application.Selection;
                foreach (Excel.Range rng in selRange)
                {
                    rng.Font.Size = 10;
                    rng.Font.Bold = true;
                    rng.Font.Color = Color.WhiteSmoke;
                    rng.Interior.Color = Color.FromArgb(1, 49, 84, 150);
                    rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    rng.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    rng.WrapText = true;
                }
            }
            catch (Exception)
            {
                //ignored
            }
            finally
            {
                helper.OptimizationOff();
            }
        }

        private void fmtNumberBtn_Click(object sender, RibbonControlEventArgs e)
        {
            helper.OptimizationOn();

            Excel.Range selRange = Globals.ThisAddIn.Application.Selection;
            foreach (Excel.Range rng in selRange)
            {
                rng.NumberFormat = "#,##0;-#,##0";
            }

            helper.OptimizationOff();
        }

        private void fmtDollarsBtn_Click(object sender, RibbonControlEventArgs e)
        {
            helper.OptimizationOn();

            Excel.Range selRange = Globals.ThisAddIn.Application.Selection;
            foreach (Excel.Range rng in selRange)
            {
                rng.NumberFormat = "$#,##0;[Red]($#,##0)";
            }

            helper.OptimizationOff();
        }

        private void validationBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("DataValidation");
        }

        private void pasteSpecialButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("PasteSpecialDialog");
        }

        private void myFontBtn_Click(object sender, RibbonControlEventArgs e)
        {
            helper.OptimizationOn();

            try
            {
                Excel.Worksheet wks = Globals.ThisAddIn.Application.ActiveSheet;
                wks.Cells.Font.Size = 11;
                wks.Cells.Font.Name = "Aptos Narrow";
            }
            catch (Exception)
            {
                // ignored
            }
            finally
            {
                helper.OptimizationOff();
            }
        }

        private void gridlinesBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Window wnd = Globals.ThisAddIn.Application.ActiveWindow;
            wnd.DisplayGridlines = !wnd.DisplayGridlines;
        }

        private void fixedRangeBtn_Click(object sender, RibbonControlEventArgs e)
        {
            var fixedRangeFrm = new FixedRangeForm()
            {
                StartPosition = FormStartPosition.Manual,
            };

            var frmWidth = fixedRangeFrm.Width;
            var location = helper.GetLocation(-(frmWidth + 15));
            fixedRangeFrm.Location = location;
            fixedRangeFrm.TopMost = true;

            fixedRangeFrm.Show();
        }

        private void dynamicNameRangeBtn_Click(object sender, RibbonControlEventArgs e)
        {
            var dynamicRangeFrm = new DynamicRangeForm()
            {
                StartPosition = FormStartPosition.Manual,
            };

            var frmWidth = dynamicRangeFrm.Width;
            var location = helper.GetLocation(-(frmWidth + 15));
            dynamicRangeFrm.Location = location;
            dynamicRangeFrm.TopMost = true;

            dynamicRangeFrm.Show();
        }

        private void aboutBtn_Click(object sender, RibbonControlEventArgs e)
        {
            var vrs = helper.GetRunningVersion();
            var msg = "Version " + vrs.Major.ToString() + "." + vrs.Minor.ToString() + "." + vrs.Build.ToString() +
                      "." +
                      vrs.Revision.ToString();
            msg = msg + Environment.NewLine + Environment.NewLine + "This application and the material contained " +
                  Environment.NewLine +
                  "within it is Microsoft confidential.";
            MessageBox.Show(text: msg, caption: @"About");
        }

        private void addinSettingBtn_Click(object sender, RibbonControlEventArgs e)
        {
            const string addInDir = @"\MarsAddIn\";

            var dirLocation = Environment.GetEnvironmentVariable("LOCALAPPDATA") + addInDir;
            if (!Directory.Exists(dirLocation))
            {
                Directory.CreateDirectory(dirLocation);
            }

            var fileLocation = Environment.GetEnvironmentVariable("LOCALAPPDATA") + addInDir + "settings.json";
            if (File.Exists(fileLocation))
            {
                Process.Start(fileLocation);
            }
            else
            {
                var marsetting = new json();
                using (var file = File.CreateText(fileLocation))
                {
                    var serializer = new JsonSerializer();
                    serializer.Serialize(file, marsetting);
                }

                Process.Start(fileLocation);
            }
        }

        private void rangeGroup_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("NameManager");
        }

        private void grpUngrpBtn_Click(object sender, RibbonControlEventArgs e)
        {
            var shtCnt = Globals.ThisAddIn.Application.ActiveWindow.SelectedSheets.Count;
            var wkSheet = Globals.ThisAddIn.Application.ActiveSheet;

            if (shtCnt == 1 && (wkSheet.Name.Substring(0, 1) == "-" || wkSheet.Name.Substring(0, 1) == "+"))
                group.UngroupSheets(wkSheet);
            else
            {
                group.GroupSheets();

                if (Globals.ThisAddIn.Application.ActiveWindow.SelectedSheets.Count > 1)
                {
                    // ReSharper disable once UnusedVariable
                    var a = Globals.ThisAddIn.Application.ActiveWindow.SelectedSheets[1].Select;
                }
            }

            if (group.HasGrouping())
            {
                Globals.Ribbons.MarsRibbon.showHideAllBtn.Enabled = true;
                Globals.Ribbons.MarsRibbon.showHideGrpBtn.Enabled = true;
            }
            else
            {
                Globals.Ribbons.MarsRibbon.showHideAllBtn.Enabled = false;
                Globals.Ribbons.MarsRibbon.showHideGrpBtn.Enabled = false;
            }
        }

        private void showHideGrpBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            if (Globals.ThisAddIn.Application.ActiveWindow.SelectedSheets.Count > 1)
            {
                // ReSharper disable once UnusedVariable
                var a = Globals.ThisAddIn.Application.ActiveWindow.SelectedSheets[1].Select;
            }

            group.ShowHideGroupSheets();

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void showHideAllBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            if (Globals.ThisAddIn.Application.ActiveWindow.SelectedSheets.Count > 1)
            {
                // ReSharper disable once UnusedVariable
                var a = Globals.ThisAddIn.Application.ActiveWindow.SelectedSheets[1].Select;
            }

            helper.ShowHideAllSheets();

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void tglFieldListBtn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("PivotFieldListShowHide");
            }
            catch (Exception)
            {
                MessageBox.Show(text: @"A pivot table must be selected for this functionality to work",
                    caption: @"No Pivot Table", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void tglFormulaBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("PivotTableGenerateGetPivotData");
        }

        private void btnConcat_Click(object sender, RibbonControlEventArgs e)
        {
            var hasString = false;
            var rngList = new List<string>();

            Excel.Range arng = Globals.ThisAddIn.Application.Selection;

            foreach (Excel.Range rng in arng)
            {
                if (rng.Value == null) continue;

                var val = rng.Value.ToString();

                if (val.Length == 0) continue;

                if (!int.TryParse(val, out int value) && !hasString)
                    hasString = true;

                rngList.Add(val);
            }

            if (rngList.Count == 0) return;

            var finalString = string.Join(",", hasString ? rngList.Select(x => "'" + x + "'").ToList() : rngList);
            Clipboard.SetText(finalString);

            finalString = finalString.Substring(0, 1) == "'" ? "'" + finalString : finalString;
            var saveCell = Globals.ThisAddIn.Application.ActiveSheet.Cells(arng.Row + arng.Count, arng.Column);
            saveCell.Value = finalString;
        }

        private void btnReadData_Click(object sender, RibbonControlEventArgs e)
        {
            var dbEntryFrm = new DatabaseEntry(intType: "Read")
            {
                StartPosition = FormStartPosition.Manual,
            };

            var frmWidth = dbEntryFrm.Width;
            var location = helper.GetLocation(-(frmWidth + 15), leftMost: false);
            dbEntryFrm.Location = location;
            dbEntryFrm.TopMost = true;


            if (dbEntryFrm.showForm)
                dbEntryFrm.Show();
        }

        private void btnWriteData_Click(object sender, RibbonControlEventArgs e)
        {
            var dbEntryFrm = new DatabaseEntry()
            {
                StartPosition = FormStartPosition.Manual
            };

            var frmWidth = dbEntryFrm.Width;
            var location = helper.GetLocation(-(frmWidth + 15), leftMost: false);
            dbEntryFrm.Location = location;
            dbEntryFrm.TopMost = true;

            if (dbEntryFrm.showForm)
                dbEntryFrm.Show();
        }


        private void wksGroup_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            if (group.HasGrouping())
            {
                Globals.Ribbons.MarsRibbon.showHideAllBtn.Enabled = true;
                Globals.Ribbons.MarsRibbon.showHideGrpBtn.Enabled = true;
            }
            else
            {
                Globals.Ribbons.MarsRibbon.showHideAllBtn.Enabled = false;
                Globals.Ribbons.MarsRibbon.showHideGrpBtn.Enabled = false;
            }
        }

        private void adminGroup_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            helper.OptimizationOff();
        }

        private void addinConfigItemesBtn_Click(object sender, RibbonControlEventArgs e)
        {
            var flist = new frmFeatureList()
            {
                StartPosition = FormStartPosition.Manual
            };

            var frmWidth = flist.Width;
            var location = helper.GetLocation(-(frmWidth + 15));
            flist.Location = location;
            flist.TopMost = true;

            flist.Show();
        }
    }
}
