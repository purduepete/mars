using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace MarsAddIn.Forms
{
    public partial class FixedRangeForm : Form
    {
        public FixedRangeForm()
        {
            InitializeComponent();
        }

        private void fixedRangeOKBtn_Click(object sender, EventArgs e)
        {
            Workbook wkb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet wks = Globals.ThisAddIn.Application.ActiveSheet;
            Range rng = Globals.ThisAddIn.Application.Selection;

            var rangeName = rngNameTextBox.Text;
            var myValue = valueTextBox.Text;

            if (!string.IsNullOrEmpty(rangeName))
            {
                foreach (Name nm in wkb.Names)
                    if (nm.Name == rangeName)
                        nm.Delete();

                foreach (Name nm in wks.Names)
                {
                    var currentRangeName = nm.Name.Replace(wks.Name + "!", "");
                    if (currentRangeName == rangeName)
                        nm.Delete();
                }

                if (globalCheckBox.CheckState == CheckState.Checked)
                    wkb.Names.Add(Name = rangeName,
                        RefersTo:
                            (string.IsNullOrEmpty(myValue)
                                ? $"='{wks.Name}'!{rng.Address}"
                                : $"={valueTextBox.Text}"));
                else
                    wks.Names.Add(Name = rangeName,
                        RefersTo:
                            (string.IsNullOrEmpty(myValue)
                                ? $"='{wks.Name}'!{rng.Address}"
                                : $"={valueTextBox.Text}"));

                Hide();
            }
            else
            {
                MessageBox.Show(text: "The name range entered, is invalid; please try again",
                    caption: "Invalid Range Name", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }
    }
}
