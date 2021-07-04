using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace MarsAddIn.Forms
{
    public partial class DynamicRangeForm : Form
    {
        public DynamicRangeForm()
        {
            InitializeComponent();
        }

        private void dynamicRangeOKBtn_Click(object sender, EventArgs e)
        {
            Workbook wkb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet wks = Globals.ThisAddIn.Application.ActiveSheet;
            Range rng = Globals.ThisAddIn.Application.ActiveCell;

            var rangeName = rngNameTextBox.Text;

            if (!string.IsNullOrEmpty(rangeName))
            {
                var rowToExclude = rowExcludeTextBox.Text.Substring(0, 1).ToLower() != "r"
                    ? int.Parse(rowExcludeTextBox.Text)
                    : 0;
                var rowOffset = rowExcludeTextBox.Text.Substring(0, 1).ToLower() == "r"
                    ? int.Parse(rowExcludeTextBox.Text.Substring(1))
                    : 0;

                var columnToExclude = columnExcludeTextBox.Text.Substring(0, 1).ToLower() != "c"
                    ? int.Parse(columnExcludeTextBox.Text)
                    : 0;
                var columnOffset = columnExcludeTextBox.Text.Substring(0, 1).ToLower() == "c"
                    ? int.Parse(columnExcludeTextBox.Text.Substring(1))
                    : 0;

                var rowHeight = int.Parse(rowHeightTextBox.Text);
                var columnWidth = int.Parse(columnWidthTextBox.Text);

                foreach (Name nm in wkb.Names)
                    if (nm.Name == rangeName)
                        nm.Delete();

                foreach (Name nm in wks.Names)
                {
                    var currentRangeName = nm.Name.Replace(wks.Name + "!", "");
                    if (currentRangeName == rangeName)
                        nm.Delete();
                }

                var c1 = wks.Rows.Count;
                var c2 = rng.Offset[0, columnOffset].Column;
                var cf = wks.Cells[c1, c2].Address;
                var rowVal = rowHeight == -1
                    ? $"COUNTA({rng.Offset[0, columnOffset].Address}:{cf}) - {rowToExclude}"
                    : rowHeight.ToString();

                c1 = rng.Offset[rowOffset, 0].Row;
                c2 = wks.Columns.Count;
                cf = wks.Cells[c1, c2].Address;
                var colVal = columnWidth == -1
                    ? $"COUNTA({rng.Offset[rowOffset, 0].Address}:{cf}) - {columnToExclude}"
                    : columnWidth.ToString();

                if (globalCheckBox.CheckState == CheckState.Checked)
                    wkb.Names.Add(Name = rangeName, RefersTo: $"=OFFSET({rng.Address}, 0, 0, {rowVal}, {colVal})");

                else
                    wks.Names.Add(Name = rangeName, RefersTo: $"=OFFSET({rng.Address}, 0, 0, {rowVal}, {colVal})");
                Hide();
            }
            else
            {
                MessageBox.Show(text: "The name range entered, is invalid; please try again", caption: "Invalid Range Name", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }
    }
}
