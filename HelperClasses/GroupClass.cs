using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace MarsAddIn.HelperClasses
{
    internal static class GroupClass
    {
        public static string GetGid(Excel.Worksheet wkSheet)
        {
            string gid = null;

            foreach (Excel.CustomProperty cust in wkSheet.CustomProperties)
                if (cust.Name == "SheetId")
                    gid = cust.Value;

            if (gid != null) return gid;

            gid = Guid.NewGuid().ToString();
            wkSheet.CustomProperties.Add("SheetId", gid);
            return gid;
        }

        public static void UngroupSheets(Excel.Worksheet wkSheet)
        {
            var wkb = Globals.ThisAddIn.Application.ActiveWorkbook;
            var wksList = new List<string>();

            if (wkSheet.Name.Substring(0, 1) == "-" || wkSheet.Name.Substring(0, 1) == "+")
                wkSheet.Name = wkSheet.Name.Substring(1).Trim();

            foreach (Excel.CustomProperty cust in wkSheet.CustomProperties)
                if (cust.Name == "ChildrenIds")
                {
                    string s = cust.Value;
                    wksList = s.Split(',').ToList();
                    cust.Delete();
                }

            if (wksList.Count == 0) return;

            foreach (Excel.Worksheet wks in wkb.Worksheets)
            {
                var gid = GetGid(wks);
                if (wksList.Contains(gid))
                    wks.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }
        }

        public static void GroupSheets()
        {
            var wksList = new List<string>();
            string gid;
            var gidm = "";


            var i = 0;
            var wkSheets = Globals.ThisAddIn.Application.ActiveWindow.SelectedSheets;

            // Get GID of all selected worksheets
            foreach (Excel.Worksheet wkSheet in wkSheets)
            {
                gid = GetGid(wkSheet);

                if (i == 0)
                {
                    gidm = gid;
                    if (wkSheet.Name.Substring(0, 1) != "-" && wkSheet.Name.Substring(0, 1) != "+")
                        wkSheet.Name = $"- {wkSheet.Name}";
                }
                else
                    wksList.Add(gid);

                i++;
            }

            if (wksList.Count == 0) return;

            // Define children GID's
            foreach (Excel.Worksheet wkSheet in wkSheets)
            {
                gid = GetGid(wkSheet);
                if (gid != gidm) continue;

                foreach (Excel.CustomProperty cust in wkSheet.CustomProperties)
                    if (cust.Name == "ChildrenIds")
                        cust.Delete();

                wkSheet.CustomProperties.Add("ChildrenIds", string.Join(",", wksList));
            }
        }

        public static void ShowHideGroupSheets()
        {
            var wkSheet = Globals.ThisAddIn.Application.ActiveSheet;
            var wkb = Globals.ThisAddIn.Application.ActiveWorkbook;

            var showHideList = new List<string>();
            var gidp = GetGid(wkSheet);
            string gid;
            var hidden = false;

            if (gidp == null) return;

            foreach (Excel.CustomProperty cust in wkSheet.CustomProperties)
                if (cust.Name == "ChildrenIds")
                {
                    string s = cust.Value;
                    showHideList = s.Split(',').ToList();
                }

            if (showHideList.Count == 0) return;

            foreach (Excel.Worksheet wks in wkb.Worksheets)
            {
                gid = GetGid(wks);
                if (!showHideList.Contains(gid)) continue;

                if (wks.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                    hidden = true;
            }

            foreach (Excel.Worksheet wks in wkb.Worksheets)
            {
                gid = GetGid(wks);
                if (gid == gidp)
                    HelperClass.UpdateWorksheetName(hidden, wks);

                if (!showHideList.Contains(gid)) continue;
                wks.Visible = hidden ? Excel.XlSheetVisibility.xlSheetVisible : Excel.XlSheetVisibility.xlSheetHidden;
            }
        }

        public static bool HasGrouping()
        {
            var wkb = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (wkb == null) return false;

            foreach (Excel.Worksheet wks in wkb.Worksheets)
            {
                foreach (Excel.CustomProperty cust in wks.CustomProperties)
                {
                    if (cust.Name == "ChildrenIds")
                    {
                        return true;
                    }

                }
            }
            return false;
        }
    }
}
