using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using group = MarsAddIn.HelperClasses.GroupClass;
using json = MarsAddIn.HelperClasses.MarsSettings;

namespace MarsAddIn.HelperClasses
{
    internal static class HelperClass
    {
        public static void OptimizationOn(bool lite = false)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.EnableEvents = false;
            if (lite) return;

            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
        }

        public static void OptimizationOff(bool lite = false)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.EnableEvents = true;

            if (lite) return;
            Globals.ThisAddIn.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        }

        public static void ShowInputDialog(ref string input)
        {

            var size = new Size(750, 70);
            var inputBox = new Form
            {
                StartPosition = FormStartPosition.Manual,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                ClientSize = size,
                Text = "Workbook Path"
            };

            var frmWidth = inputBox.Width;
            var location = GetLocation(-(frmWidth + 15));
            inputBox.Location = location;

            var textBox = new TextBox
            {
                Size = new Size(size.Width - 10, 23),
                Location = new Point(5, 5),
                Text = input,
            };

            inputBox.Controls.Add(textBox);


            var okButton = new Button
            {
                DialogResult = DialogResult.OK,
                Name = "okButton",
                Size = new Size(75, 23),
                Text = "&OK",
                Location = new Point(size.Width - 80, 39)
            };

            inputBox.Controls.Add(okButton);
            inputBox.AcceptButton = okButton;

            var result = inputBox.ShowDialog();
            input = textBox.Text;
        }

        public static Version GetRunningVersion()
        {
            try
            {
                return System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
            }
            catch
            {
                return Assembly.GetExecutingAssembly().GetName().Version;
            }
        }

        public static void ShowHideAllSheets()
        {
            var wkb = Globals.ThisAddIn.Application.ActiveWorkbook;
            var showHideList = new List<string>();
            var parentList = new List<string>();

            var hidden = false;
            string gid;


            foreach (Excel.Worksheet wks in wkb.Worksheets)
            {
                gid = group.GetGid(wks);

                foreach (Excel.CustomProperty cust in wks.CustomProperties)
                    if (cust.Name == "ChildrenIds")
                    {
                        string s = cust.Value;
                        showHideList.AddRange(s.Split(','));
                        parentList.Add(gid);
                    }
            }

            foreach (Excel.Worksheet wks in wkb.Worksheets)
            {
                gid = group.GetGid(wks);
                if (!showHideList.Contains(gid)) continue;

                if (wks.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                    hidden = true;
            }

            foreach (Excel.Worksheet wks in wkb.Worksheets)
            {
                gid = group.GetGid(wks);

                if (parentList.Contains(gid))
                    UpdateWorksheetName(hidden, wks);

                if (!showHideList.Contains(gid)) continue;
                wks.Visible = hidden ? Excel.XlSheetVisibility.xlSheetVisible : Excel.XlSheetVisibility.xlSheetHidden;
            }
        }

        public static void UpdateWorksheetName(bool hidden, Excel.Worksheet wks)
        {
            if (hidden)
            {
                if (wks.Name.Substring(0, 1) == "+")
                    wks.Name = "- " + wks.Name.Substring(1).Trim();
            }
            else
            {
                if (wks.Name.Substring(0, 1) == "-")
                    wks.Name = "+ " + wks.Name.Substring(1).Trim();
            }
        }

        public static Point GetLocation(int offsetX, int offsetY = 15, bool leftMost = true)
        {
            var appLocation = Globals.ThisAddIn.Application.Left;
            var screen = Screen.AllScreens.First(s => s.Bounds.Left < (appLocation + 15));

            var pnt = new Point
            {
                Y = screen.Bounds.Top + offsetY,
                X = screen.Bounds.X + 15
            };

            if (!leftMost)
                pnt.X = screen.Bounds.X + screen.Bounds.Width + offsetX;

            return pnt;
        }

        public static Dictionary<string, bool> LoadRibbonItemSettings()
        {
            var settingsLocation = Environment.GetEnvironmentVariable("LOCALAPPDATA") + @"\MarsAddIn\settings.json";

            var settingsDict = new Dictionary<string, bool>();

            if (File.Exists(settingsLocation))
            {
                var jsonSettings = JsonConvert.DeserializeObject<json>(File.ReadAllText(settingsLocation));
                if (jsonSettings.RibbonSettings == null) return settingsDict;


                foreach (var ribbonSetting in jsonSettings.RibbonSettings)
                    settingsDict[ribbonSetting.Key] = ribbonSetting.Value;

                return settingsDict;
            }

            CreateSettingsFile(settingsLocation);

            return settingsDict;
        }

        public static void SaveRibbonItemSettings(OfficeRibbon officeRibbon)
        {
            var settingsLocation = Environment.GetEnvironmentVariable("LOCALAPPDATA") + @"\MarsAddIn\settings.json";

            if (File.Exists(settingsLocation))
            {
                var json = JsonConvert.DeserializeObject<json>(File.ReadAllText(settingsLocation));
                json.RibbonSettings = new Dictionary<string, bool>();

                var ribbonTabs = officeRibbon.Tabs[1];
                foreach (var tabGroup in ribbonTabs.Groups)
                    foreach (var ribbonItem in tabGroup.Items)
                        if (!new[] { "menuAddInSettings" }.Contains(ribbonItem.Name) && ribbonItem.Tag != null)
                            try
                            {
                                json.RibbonSettings.Add(ribbonItem.Tag.ToString(), ribbonItem.Visible);
                            }
                            catch (Exception)
                            {
                                // ignored
                            }

                File.WriteAllText(settingsLocation, JsonConvert.SerializeObject(json));
            }
            else
            {
                CreateSettingsFile(settingsLocation);
                var json = JsonConvert.DeserializeObject<json>(File.ReadAllText(settingsLocation));
                json.RibbonSettings = new Dictionary<string, bool>();

                var ribbonTabs = officeRibbon.Tabs[1];
                foreach (var tabGroup in ribbonTabs.Groups)
                    foreach (var ribbonItem in tabGroup.Items)
                        if (!new[] { "menuAddInSettings" }.Contains(ribbonItem.Name) && ribbonItem.Tag != null)
                            try
                            {
                                json.RibbonSettings.Add(ribbonItem.Tag.ToString(), ribbonItem.Visible);
                            }
                            catch (Exception)
                            {
                                // ignored
                            }

                File.WriteAllText(settingsLocation, JsonConvert.SerializeObject(json));
            }

        }

        public static void SaveRibbonItemSettings(Dictionary<string, bool> dictItems)
        {
            var settingsLocation = Environment.GetEnvironmentVariable("LOCALAPPDATA") + @"\MarsAddIn\settings.json";

            if (File.Exists(settingsLocation))
            {
                var json = JsonConvert.DeserializeObject<json>(File.ReadAllText(settingsLocation));
                json.RibbonSettings = dictItems;

                File.WriteAllText(settingsLocation, JsonConvert.SerializeObject(json));
            }
            else
            {
                CreateSettingsFile(settingsLocation);
                var json = JsonConvert.DeserializeObject<json>(File.ReadAllText(settingsLocation));
                json.RibbonSettings = dictItems;

                File.WriteAllText(settingsLocation, JsonConvert.SerializeObject(json));
            }

        }

        private static void CreateSettingsFile(string settingsLocation)
        {
            var dirLocation = Path.GetDirectoryName(settingsLocation);
            if (!Directory.Exists(dirLocation))
            {
                Directory.CreateDirectory(dirLocation);
            }

            var marsetting = new json();
            using (var file = File.CreateText(settingsLocation))
            {
                var serializer = new JsonSerializer();
                serializer.Serialize(file, marsetting);
            }
        }

        public static void ReleaseObject(object o)
        {
            try
            {
                if (o == null) return;
                var retryCount = 0;
                while (Marshal.ReleaseComObject(o) > 0 && retryCount < 10)
                {
                    Thread.Sleep(1000);
                    retryCount++;
                }
            }
            catch
            {
                // ignored
            }
            finally
            {
                // ReSharper disable once RedundantAssignment
                o = null;
            }
        }
    }
}

