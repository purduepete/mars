using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using helper = MarsAddIn.HelperClasses.HelperClass;

namespace MarsAddIn.Forms
{
    public partial class AddToExcel : Form
    {
        private string AddType { get; set; }

        public AddToExcel(string addType = "column")
        {
            InitializeComponent();
            AddType = addType;

            if (addType != "row") return;

            lblAddToExcel.Text = @"How many rows do you wish to add?";
            Text = @"Add Rows";
        }

        private void btnAddToExcelCancel_Click(object sender, EventArgs e)
        {
            Hide();
        }

        private void btnAddToExcelOK_Click(object sender, EventArgs e)
        {
            helper.OptimizationOn();

            if (AddType == @"column")
            {
                if (int.TryParse(tboxAddToExcel.Text, out var answ))
                {

                    Excel.Worksheet wks = Globals.ThisAddIn.Application.ActiveSheet;
                    var currentColumn = Globals.ThisAddIn.Application.ActiveCell.Column;

                    string cls = wks.Cells[1, currentColumn].Address;
                    cls = cls.Substring(0, cls.LastIndexOf("$", StringComparison.Ordinal)).Replace("$", "");

                    string cle = wks.Cells[1, currentColumn + answ - 1].Address;
                    cle = cle.Substring(0, cle.LastIndexOf("$", StringComparison.Ordinal)).Replace("$", "");

                    var columnRange = $"{cls}:{cle}";
                    wks.Columns[columnRange].Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
                }
            }
            else
            {
                if (int.TryParse(tboxAddToExcel.Text, out var answ))
                {
                    Excel.Worksheet wks = Globals.ThisAddIn.Application.ActiveSheet;
                    var currentRow = Globals.ThisAddIn.Application.ActiveCell.Row;
                    var rowRange = $"{currentRow}:{currentRow + answ - 1}";
                    wks.Rows[rowRange].Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown);
                }
            }

            Hide();

            helper.OptimizationOff();
        }
    }
}
