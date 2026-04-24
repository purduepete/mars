using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using group = MarsAddIn.HelperClasses.GroupClass;

namespace MarsAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            Application.WorkbookOpen += Application_WorkbookChanged;
            Application.WorkbookActivate += Application_WorkbookChanged;
            Application.SheetActivate += Application_SheetActivate;

            if (Application.ActiveWorkbook != null)
                UpdateGroupButtons();
        }

        private static void Application_WorkbookChanged(Excel.Workbook wb)
        {
            UpdateGroupButtons();
        }

        private static void Application_SheetActivate(object sh)
        {
            UpdateGroupButtons();
        }

        internal static void UpdateGroupButtons()
        {
            var hasGrouping = group.HasGrouping();
            Globals.Ribbons.MarsRibbon.showHideAllBtn.Enabled = hasGrouping;
            Globals.Ribbons.MarsRibbon.showHideGrpBtn.Enabled = hasGrouping;
        }

        private static void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
