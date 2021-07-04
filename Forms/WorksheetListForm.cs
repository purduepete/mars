using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using ListBox = System.Windows.Forms.ListBox;
using helper = MarsAddIn.HelperClasses.HelperClass;

namespace MarsAddIn.Forms
{
    public partial class WorksheetListForm : Form
    {
        public WorksheetListForm()
        {
            InitializeComponent();
            FillList();
        }


        private void FillList()
        {
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            const int col1Width = 235;
            worksheetListFormListView.Columns[0].Width = col1Width;
            worksheetListFormListView.Columns[1].Width = worksheetListFormListView.Width - col1Width - 25;
            worksheetListFormListView.View = View.Details;
            worksheetListFormListView.FullRowSelect = true;

            worksheetListFormListView.Items.Clear();

            foreach (Worksheet wks in wb.Worksheets)
            {
                string visString;
                switch (wks.Visible)
                {
                    case XlSheetVisibility.xlSheetHidden:
                        visString = "Hidden";
                        break;
                    case XlSheetVisibility.xlSheetVeryHidden:
                        visString = "Very Hidden";
                        break;
                    case XlSheetVisibility.xlSheetVisible:
                        visString = "Visible";
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                var newList = new ListViewItem(wks.Name);
                newList.SubItems.Add(visString);
                worksheetListFormListView.Items.Add(newList);
            }
        }

        private void SetVisibleAttribute(XlSheetVisibility visibility = XlSheetVisibility.xlSheetVisible)
        {
            helper.OptimizationOn();

            var visType = "";
            switch (visibility)
            {
                case XlSheetVisibility.xlSheetVisible:
                    visType = "visible";
                    break;
                case XlSheetVisibility.xlSheetHidden:
                    visType = "hidden";
                    break;
                case XlSheetVisibility.xlSheetVeryHidden:
                    visType = "very hidden";
                    break;
                default:
                    break;
            }

            var si = worksheetListFormListView.SelectedItems;
            foreach (ListViewItem sii in si)
            {
                var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                Worksheet wks = wb.Worksheets[sii.Text];

                if (wks.Visible == visibility) continue;
                try
                {
                    wks.Visible = visibility;
                    sii.Selected = false;
                    FillList();
                }
                catch (Exception)
                {
                    helper.OptimizationOff();
                    MessageBox.Show(
                        text: $@"An error was encountered while trying to make the worksheet '{ wks.Name }', {visType}.  Reminder, at least one worksheet needs to be visible",
                        caption: @"Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
                }
                {
                }
            }
            helper.OptimizationOff();
        }


        private void worksheetListFormHiddenButton_Click(object sender, EventArgs e)
        {
            SetVisibleAttribute(XlSheetVisibility.xlSheetHidden);
        }

        private void worksheetListFormVHiddenButton_Click(object sender, EventArgs e)
        {
            SetVisibleAttribute(XlSheetVisibility.xlSheetVeryHidden);
        }

        private void worksheetListFormVisibleButton_Click(object sender, EventArgs e)
        {
            SetVisibleAttribute();
        }


        private void worksheetListFormListView_DoubleClick(object sender, EventArgs e)
        {
            helper.OptimizationOn();

            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            var lv = (ListView)sender;
            var wksName = lv.SelectedItems[0].Text;

            Worksheet wks = wb.Worksheets[wksName];
            string userAction = null;
            try
            {
                switch (wks.Visible)
                {
                    case XlSheetVisibility.xlSheetHidden:
                        wks.Visible = XlSheetVisibility.xlSheetVisible;
                        break;
                    case XlSheetVisibility.xlSheetVeryHidden:
                        wks.Visible = XlSheetVisibility.xlSheetVisible;
                        break;
                    case XlSheetVisibility.xlSheetVisible:
                        userAction = "hidden";
                        wks.Visible = XlSheetVisibility.xlSheetHidden;
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
                FillList();

            }
            catch (Exception)
            {
                MessageBox.Show(
                    text:
                    $@"An error was encountered while trying to make the worksheet '{wksName}', {
                            userAction
                        }  Reminder, at least one worksheet needs to be visible",
                    caption: @"Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
            finally
            {
                helper.OptimizationOff();
            }
        }

        private void worksheetListFormListView_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
