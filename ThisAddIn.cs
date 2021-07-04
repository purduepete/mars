using Excel = Microsoft.Office.Interop.Excel;
using group = MarsAddIn.HelperClasses.GroupClass;

namespace MarsAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.WorkbookOpen += Application_ActiveWorkbookOpen;
        }

        private static void Application_ActiveWorkbookOpen(Excel.Workbook wb)
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
