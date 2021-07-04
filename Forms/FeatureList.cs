using System;
using System.Collections.Generic;
using System.Windows.Forms;
using helper = MarsAddIn.HelperClasses.HelperClass;

namespace MarsAddIn.Forms
{
    public partial class frmFeatureList : Form
    {
        private bool _lbRange;
        private bool _lbPivotTable;
        private bool _lbWorksheet;
        private bool _lbFormatting;
        private bool _lbCopyPaste;
        private bool _lbDatabase;
        private bool _lbFile;
        private bool _lbAdmin;

        public frmFeatureList()
        {
            InitializeComponent();

            _lbRange = true;
            _lbPivotTable = true;
            _lbWorksheet = true;
            _lbFormatting = true;
            _lbCopyPaste = true;
            _lbDatabase = true;
            _lbFile = true;
            _lbAdmin = true;

            var ribbonSettings = helper.LoadRibbonItemSettings();

            ListBoxWorkerWrite(checkedListBoxRange, ribbonSettings);
            ListBoxWorkerWrite(checkedListBoxPivotTable, ribbonSettings);
            ListBoxWorkerWrite(checkedListBoxWorksheet, ribbonSettings);
            ListBoxWorkerWrite(checkedListBoxFormatting, ribbonSettings);
            ListBoxWorkerWrite(checkedListBoxCopyPaste, ribbonSettings);
            ListBoxWorkerWrite(checkedListBoxDatabase, ribbonSettings);
            ListBoxWorkerWrite(checkedListBoxFile, ribbonSettings);
            ListBoxWorkerWrite(checkedListBoxAdmin, ribbonSettings);
        }

        private void btnFLOK_Click(object sender, EventArgs e)
        {
            var ribbonSettings = new Dictionary<string, bool>();

            ribbonSettings = (ListBoxWorkerRead(checkedListBoxRange, ribbonSettings));
            ribbonSettings = ListBoxWorkerRead(checkedListBoxPivotTable, ribbonSettings);
            ribbonSettings = ListBoxWorkerRead(checkedListBoxWorksheet, ribbonSettings);
            ribbonSettings = ListBoxWorkerRead(checkedListBoxFormatting, ribbonSettings);
            ribbonSettings = ListBoxWorkerRead(checkedListBoxCopyPaste, ribbonSettings);
            ribbonSettings = ListBoxWorkerRead(checkedListBoxDatabase, ribbonSettings);
            ribbonSettings = ListBoxWorkerRead(checkedListBoxFile, ribbonSettings);
            ribbonSettings = ListBoxWorkerRead(checkedListBoxAdmin, ribbonSettings);

            helper.SaveRibbonItemSettings(ribbonSettings);
            Hide();
        }

        private static void ListBoxWorkerWrite(CheckedListBox lbox, Dictionary<string, bool> ribbonSettings)
        {
            for (var i = 0; i < lbox.Items.Count; i++)
                try
                {
                    lbox.SetItemCheckState(i, ribbonSettings[lbox.Items[i].ToString()] ? CheckState.Checked : CheckState.Unchecked);
                }
                catch (Exception)
                {
                    // ignored
                }
        }

        private static Dictionary<string, bool> ListBoxWorkerRead(CheckedListBox lbox, Dictionary<string, bool> ribbonSettings)
        {
            for (var i = 0; i < lbox.Items.Count; i++)
                ribbonSettings.Add(lbox.Items[i].ToString(), lbox.GetItemChecked(i));

            return ribbonSettings;
        }

        private void lblRange_Click(object sender, EventArgs e)
        {
            ListBoxWorkerToggleAll(checkedListBoxRange, _lbRange);
            _lbRange = !_lbRange;
        }

        private void lblPivotTable_Click(object sender, EventArgs e)
        {
            ListBoxWorkerToggleAll(checkedListBoxPivotTable, _lbPivotTable);
            _lbPivotTable = !_lbPivotTable;
        }

        private void lblWorksheet_Click(object sender, EventArgs e)
        {
            ListBoxWorkerToggleAll(checkedListBoxWorksheet, _lbWorksheet);
            _lbWorksheet = !_lbWorksheet;
        }

        private void lblFormatting_Click(object sender, EventArgs e)
        {
            ListBoxWorkerToggleAll(checkedListBoxFormatting, _lbFormatting);
            _lbFormatting = !_lbFormatting;
        }

        private void lblCopyPaste_Click(object sender, EventArgs e)
        {
            ListBoxWorkerToggleAll(checkedListBoxCopyPaste, _lbCopyPaste);
            _lbCopyPaste = !_lbCopyPaste;
        }

        private void lblDatabase_Click(object sender, EventArgs e)
        {
            ListBoxWorkerToggleAll(checkedListBoxDatabase, _lbDatabase);
            _lbDatabase = !_lbDatabase;
        }

        private void lblFile_Click(object sender, EventArgs e)
        {
            ListBoxWorkerToggleAll(checkedListBoxFile, _lbFile);
            _lbFile = !_lbFile;
        }

        private void lblAdministration_Click(object sender, EventArgs e)
        {
            ListBoxWorkerToggleAll(checkedListBoxAdmin, _lbAdmin);
            _lbAdmin = !_lbAdmin;
        }

        private static void ListBoxWorkerToggleAll(CheckedListBox lbox, bool lbitem)
        {
            for (var i = 0; i < lbox.Items.Count; i++)
                lbox.SetItemCheckState(i, lbitem ? CheckState.Checked : CheckState.Unchecked);
        }
    }
}
