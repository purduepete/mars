using Microsoft.Office.Core;

namespace MarsAddIn

{
    partial class MarsRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MarsRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl2 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.marsTab = this.Factory.CreateRibbonTab();
            this.rngGroup = this.Factory.CreateRibbonGroup();
            this.addRowsBtn = this.Factory.CreateRibbonButton();
            this.addColumnsBtn = this.Factory.CreateRibbonButton();
            this.rngSeparator = this.Factory.CreateRibbonSeparator();
            this.nameMenu = this.Factory.CreateRibbonMenu();
            this.fixedRangeBtn = this.Factory.CreateRibbonButton();
            this.dynamicNameRangeBtn = this.Factory.CreateRibbonButton();
            this.validationBtn = this.Factory.CreateRibbonButton();
            this.pivotGroup = this.Factory.CreateRibbonGroup();
            this.dfltPivot = this.Factory.CreateRibbonButton();
            this.dupPivotBtn = this.Factory.CreateRibbonButton();
            this.ptSeparator = this.Factory.CreateRibbonSeparator();
            this.tglFormulaBtn = this.Factory.CreateRibbonButton();
            this.tglFieldListBtn = this.Factory.CreateRibbonButton();
            this.wksGroup = this.Factory.CreateRibbonGroup();
            this.wksSheetListBtn = this.Factory.CreateRibbonButton();
            this.wksShowAllBtn = this.Factory.CreateRibbonButton();
            this.wksResizeBtn = this.Factory.CreateRibbonButton();
            this.wksSeparator = this.Factory.CreateRibbonSeparator();
            this.grpUngrpBtn = this.Factory.CreateRibbonButton();
            this.showHideGrpBtn = this.Factory.CreateRibbonButton();
            this.showHideAllBtn = this.Factory.CreateRibbonButton();
            this.fmtGroup = this.Factory.CreateRibbonGroup();
            this.headerFormatBtn = this.Factory.CreateRibbonButton();
            this.rowShadingBtn = this.Factory.CreateRibbonButton();
            this.defaultMenu = this.Factory.CreateRibbonMenu();
            this.myFontBtn = this.Factory.CreateRibbonButton();
            this.gridlinesBtn = this.Factory.CreateRibbonButton();
            this.fmtSeparator = this.Factory.CreateRibbonSeparator();
            this.fmtNumberBtn = this.Factory.CreateRibbonButton();
            this.fmtDollarsBtn = this.Factory.CreateRibbonButton();
            this.copyPasteGroup = this.Factory.CreateRibbonGroup();
            this.pasteSpecialButton = this.Factory.CreateRibbonButton();
            this.copyPasteFormulaBtn = this.Factory.CreateRibbonButton();
            this.dbGroup = this.Factory.CreateRibbonGroup();
            this.btnConcatStrings = this.Factory.CreateRibbonButton();
            this.btnCreateStructure = this.Factory.CreateRibbonButton();
            this.dbSeperator = this.Factory.CreateRibbonSeparator();
            this.btnReadData = this.Factory.CreateRibbonButton();
            this.btnWriteData = this.Factory.CreateRibbonButton();
            this.adminGroup = this.Factory.CreateRibbonGroup();
            this.aboutBtn = this.Factory.CreateRibbonButton();
            this.menuAddInSettings = this.Factory.CreateRibbonMenu();
            this.addinSettingBtn = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.settingsBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.marsTab.SuspendLayout();
            this.rngGroup.SuspendLayout();
            this.pivotGroup.SuspendLayout();
            this.wksGroup.SuspendLayout();
            this.fmtGroup.SuspendLayout();
            this.copyPasteGroup.SuspendLayout();
            this.dbGroup.SuspendLayout();
            this.adminGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            this.tab1.Visible = false;
            // 
            // marsTab
            // 
            this.marsTab.Groups.Add(this.rngGroup);
            this.marsTab.Groups.Add(this.pivotGroup);
            this.marsTab.Groups.Add(this.wksGroup);
            this.marsTab.Groups.Add(this.fmtGroup);
            this.marsTab.Groups.Add(this.copyPasteGroup);
            this.marsTab.Groups.Add(this.dbGroup);
            this.marsTab.Groups.Add(this.adminGroup);
            this.marsTab.Label = "MARS";
            this.marsTab.Name = "marsTab";
            // 
            // rngGroup
            // 
            ribbonDialogLauncherImpl1.ScreenTip = "Add new name";
            this.rngGroup.DialogLauncher = ribbonDialogLauncherImpl1;
            this.rngGroup.Items.Add(this.addRowsBtn);
            this.rngGroup.Items.Add(this.addColumnsBtn);
            this.rngGroup.Items.Add(this.rngSeparator);
            this.rngGroup.Items.Add(this.nameMenu);
            this.rngGroup.Items.Add(this.validationBtn);
            this.rngGroup.Label = "Range";
            this.rngGroup.Name = "rngGroup";
            this.rngGroup.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rangeGroup_DialogLauncherClick);
            // 
            // addRowsBtn
            // 
            this.addRowsBtn.Label = "Add Rows";
            this.addRowsBtn.Name = "addRowsBtn";
            this.addRowsBtn.OfficeImageId = "InsertRows";
            this.addRowsBtn.ScreenTip = "Add a defined numbers of rows, starting at the selected cell";
            this.addRowsBtn.ShowImage = true;
            this.addRowsBtn.Tag = "Add Rows";
            this.addRowsBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addRowsButton_Click);
            // 
            // addColumnsBtn
            // 
            this.addColumnsBtn.Label = "Add Columns";
            this.addColumnsBtn.Name = "addColumnsBtn";
            this.addColumnsBtn.OfficeImageId = "InsertColumns";
            this.addColumnsBtn.ScreenTip = "Add a defined numbers of columns, starting at the selected cell";
            this.addColumnsBtn.ShowImage = true;
            this.addColumnsBtn.Tag = "Add Columns";
            this.addColumnsBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addColumnsButton_Click);
            // 
            // rngSeparator
            // 
            this.rngSeparator.Name = "rngSeparator";
            // 
            // nameMenu
            // 
            this.nameMenu.Items.Add(this.fixedRangeBtn);
            this.nameMenu.Items.Add(this.dynamicNameRangeBtn);
            this.nameMenu.Label = "Names";
            this.nameMenu.Name = "nameMenu";
            this.nameMenu.OfficeImageId = "NameManager";
            this.nameMenu.ShowImage = true;
            this.nameMenu.Tag = "Names";
            // 
            // fixedRangeBtn
            // 
            this.fixedRangeBtn.Label = "Fixed Range";
            this.fixedRangeBtn.Name = "fixedRangeBtn";
            this.fixedRangeBtn.OfficeImageId = "NameDefine";
            this.fixedRangeBtn.ShowImage = true;
            this.fixedRangeBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fixedRangeBtn_Click);
            // 
            // dynamicNameRangeBtn
            // 
            this.dynamicNameRangeBtn.Label = "Dynamic Range";
            this.dynamicNameRangeBtn.Name = "dynamicNameRangeBtn";
            this.dynamicNameRangeBtn.OfficeImageId = "NameCreateFromSelection";
            this.dynamicNameRangeBtn.ShowImage = true;
            this.dynamicNameRangeBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dynamicNameRangeBtn_Click);
            // 
            // validationBtn
            // 
            this.validationBtn.Label = "Validation";
            this.validationBtn.Name = "validationBtn";
            this.validationBtn.OfficeImageId = "DataValidation";
            this.validationBtn.ShowImage = true;
            this.validationBtn.Tag = "Validation";
            this.validationBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.validationBtn_Click);
            // 
            // pivotGroup
            // 
            this.pivotGroup.Items.Add(this.dfltPivot);
            this.pivotGroup.Items.Add(this.dupPivotBtn);
            this.pivotGroup.Items.Add(this.ptSeparator);
            this.pivotGroup.Items.Add(this.tglFormulaBtn);
            this.pivotGroup.Items.Add(this.tglFieldListBtn);
            this.pivotGroup.Label = "Pivot Table";
            this.pivotGroup.Name = "pivotGroup";
            // 
            // dfltPivot
            // 
            this.dfltPivot.Label = "Default Pivot";
            this.dfltPivot.Name = "dfltPivot";
            this.dfltPivot.OfficeImageId = "DefaultPivotTable";
            this.dfltPivot.ScreenTip = "For the selected pivot, apply personal settings";
            this.dfltPivot.ShowImage = true;
            this.dfltPivot.Tag = "Default Pivot";
            this.dfltPivot.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dfltPivot_Click);
            // 
            // dupPivotBtn
            // 
            this.dupPivotBtn.Label = "Duplicate Pivot";
            this.dupPivotBtn.Name = "dupPivotBtn";
            this.dupPivotBtn.OfficeImageId = "DuplicateView";
            this.dupPivotBtn.ShowImage = true;
            this.dupPivotBtn.Tag = "Duplicate Pivot";
            this.dupPivotBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dupPivotBtn_Click);
            // 
            // ptSeparator
            // 
            this.ptSeparator.Name = "ptSeparator";
            // 
            // tglFormulaBtn
            // 
            this.tglFormulaBtn.Label = "Toggle Formulas";
            this.tglFormulaBtn.Name = "tglFormulaBtn";
            this.tglFormulaBtn.OfficeImageId = "ShapeSheetShowFormulas";
            this.tglFormulaBtn.ShowImage = true;
            this.tglFormulaBtn.Tag = "Toggle Formulas";
            this.tglFormulaBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tglFormulaBtn_Click);
            // 
            // tglFieldListBtn
            // 
            this.tglFieldListBtn.Label = "Toggle Field List";
            this.tglFieldListBtn.Name = "tglFieldListBtn";
            this.tglFieldListBtn.OfficeImageId = "PivotFieldListShowHide";
            this.tglFieldListBtn.ShowImage = true;
            this.tglFieldListBtn.Tag = "Toggle Field List";
            this.tglFieldListBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tglFieldListBtn_Click);
            // 
            // wksGroup
            // 
            ribbonDialogLauncherImpl2.ScreenTip = "Assess workbook for grouping";
            this.wksGroup.DialogLauncher = ribbonDialogLauncherImpl2;
            this.wksGroup.Items.Add(this.wksSheetListBtn);
            this.wksGroup.Items.Add(this.wksShowAllBtn);
            this.wksGroup.Items.Add(this.wksResizeBtn);
            this.wksGroup.Items.Add(this.wksSeparator);
            this.wksGroup.Items.Add(this.grpUngrpBtn);
            this.wksGroup.Items.Add(this.showHideGrpBtn);
            this.wksGroup.Items.Add(this.showHideAllBtn);
            this.wksGroup.Label = "Worksheet";
            this.wksGroup.Name = "wksGroup";
            this.wksGroup.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.wksGroup_DialogLauncherClick);
            // 
            // wksSheetListBtn
            // 
            this.wksSheetListBtn.Label = "Worksheet List";
            this.wksSheetListBtn.Name = "wksSheetListBtn";
            this.wksSheetListBtn.OfficeImageId = "ListView";
            this.wksSheetListBtn.ShowImage = true;
            this.wksSheetListBtn.Tag = "Worksheet List";
            this.wksSheetListBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.wksSheetList_Click);
            // 
            // wksShowAllBtn
            // 
            this.wksShowAllBtn.Label = "Show All Worksheets";
            this.wksShowAllBtn.Name = "wksShowAllBtn";
            this.wksShowAllBtn.OfficeImageId = "PageVersionsShow";
            this.wksShowAllBtn.ShowImage = true;
            this.wksShowAllBtn.Tag = "Show All Worksheets";
            this.wksShowAllBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.wksShowAll_Click);
            // 
            // wksResizeBtn
            // 
            this.wksResizeBtn.Label = "Resize Worksheet";
            this.wksResizeBtn.Name = "wksResizeBtn";
            this.wksResizeBtn.OfficeImageId = "PageSize";
            this.wksResizeBtn.ShowImage = true;
            this.wksResizeBtn.Tag = "Resize Worksheet";
            this.wksResizeBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.wksResize_Click);
            // 
            // wksSeparator
            // 
            this.wksSeparator.Name = "wksSeparator";
            // 
            // grpUngrpBtn
            // 
            this.grpUngrpBtn.Label = "Group/Ungroup";
            this.grpUngrpBtn.Name = "grpUngrpBtn";
            this.grpUngrpBtn.OfficeImageId = "GroupAdpQueryType";
            this.grpUngrpBtn.ShowImage = true;
            this.grpUngrpBtn.Tag = "Group/Ungroup";
            this.grpUngrpBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.grpUngrpBtn_Click);
            // 
            // showHideGrpBtn
            // 
            this.showHideGrpBtn.Enabled = false;
            this.showHideGrpBtn.Label = "Show/Hide Group";
            this.showHideGrpBtn.Name = "showHideGrpBtn";
            this.showHideGrpBtn.OfficeImageId = "PivotPlusMinusButtonsShowHide";
            this.showHideGrpBtn.ShowImage = true;
            this.showHideGrpBtn.Tag = "Show/Hide Group";
            this.showHideGrpBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.showHideGrpBtn_Click);
            // 
            // showHideAllBtn
            // 
            this.showHideAllBtn.Enabled = false;
            this.showHideAllBtn.Label = "Show/Hide All";
            this.showHideAllBtn.Name = "showHideAllBtn";
            this.showHideAllBtn.OfficeImageId = "InsertActionGallery";
            this.showHideAllBtn.ShowImage = true;
            this.showHideAllBtn.Tag = "Show/Hide All";
            this.showHideAllBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.showHideAllBtn_Click);
            // 
            // fmtGroup
            // 
            this.fmtGroup.Items.Add(this.headerFormatBtn);
            this.fmtGroup.Items.Add(this.rowShadingBtn);
            this.fmtGroup.Items.Add(this.defaultMenu);
            this.fmtGroup.Items.Add(this.fmtSeparator);
            this.fmtGroup.Items.Add(this.fmtNumberBtn);
            this.fmtGroup.Items.Add(this.fmtDollarsBtn);
            this.fmtGroup.Label = "Formatting";
            this.fmtGroup.Name = "fmtGroup";
            // 
            // headerFormatBtn
            // 
            this.headerFormatBtn.Label = "Header Format";
            this.headerFormatBtn.Name = "headerFormatBtn";
            this.headerFormatBtn.OfficeImageId = "HeaderCell";
            this.headerFormatBtn.ShowImage = true;
            this.headerFormatBtn.Tag = "Header Format";
            this.headerFormatBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.headerFormatBtn_Click);
            // 
            // rowShadingBtn
            // 
            this.rowShadingBtn.Label = "Row Shading";
            this.rowShadingBtn.Name = "rowShadingBtn";
            this.rowShadingBtn.OfficeImageId = "HeaderSourceEdit";
            this.rowShadingBtn.ShowImage = true;
            this.rowShadingBtn.Tag = "Row Shading";
            this.rowShadingBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rowShadingBtn_Click);
            // 
            // defaultMenu
            // 
            this.defaultMenu.Items.Add(this.myFontBtn);
            this.defaultMenu.Items.Add(this.gridlinesBtn);
            this.defaultMenu.Label = "Preferred Settings";
            this.defaultMenu.Name = "defaultMenu";
            this.defaultMenu.OfficeImageId = "OutlookGears";
            this.defaultMenu.ShowImage = true;
            this.defaultMenu.Tag = "Preferred Settings";
            // 
            // myFontBtn
            // 
            this.myFontBtn.Label = "Default Font";
            this.myFontBtn.Name = "myFontBtn";
            this.myFontBtn.OfficeImageId = "FontColorPicker";
            this.myFontBtn.ShowImage = true;
            this.myFontBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.myFontBtn_Click);
            // 
            // gridlinesBtn
            // 
            this.gridlinesBtn.Label = "Toggle Gridlines";
            this.gridlinesBtn.Name = "gridlinesBtn";
            this.gridlinesBtn.OfficeImageId = "GridSettings";
            this.gridlinesBtn.ShowImage = true;
            this.gridlinesBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gridlinesBtn_Click);
            // 
            // fmtSeparator
            // 
            this.fmtSeparator.Name = "fmtSeparator";
            // 
            // fmtNumberBtn
            // 
            this.fmtNumberBtn.Label = "Format Number";
            this.fmtNumberBtn.Name = "fmtNumberBtn";
            this.fmtNumberBtn.OfficeImageId = "FormattingIncreaseDecimals";
            this.fmtNumberBtn.ShowImage = true;
            this.fmtNumberBtn.Tag = "Format Number";
            this.fmtNumberBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fmtNumberBtn_Click);
            // 
            // fmtDollarsBtn
            // 
            this.fmtDollarsBtn.Label = "Format Dollar";
            this.fmtDollarsBtn.Name = "fmtDollarsBtn";
            this.fmtDollarsBtn.OfficeImageId = "AccountingFormat";
            this.fmtDollarsBtn.ShowImage = true;
            this.fmtDollarsBtn.Tag = "Format Dollar";
            this.fmtDollarsBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fmtDollarsBtn_Click);
            // 
            // copyPasteGroup
            // 
            this.copyPasteGroup.Items.Add(this.pasteSpecialButton);
            this.copyPasteGroup.Items.Add(this.copyPasteFormulaBtn);
            this.copyPasteGroup.Label = "Copy/Paste";
            this.copyPasteGroup.Name = "copyPasteGroup";
            // 
            // pasteSpecialButton
            // 
            this.pasteSpecialButton.Label = "Paste Special";
            this.pasteSpecialButton.Name = "pasteSpecialButton";
            this.pasteSpecialButton.OfficeImageId = "Paste";
            this.pasteSpecialButton.ShowImage = true;
            this.pasteSpecialButton.Tag = "Paste Special";
            this.pasteSpecialButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pasteSpecialButton_Click);
            // 
            // copyPasteFormulaBtn
            // 
            this.copyPasteFormulaBtn.Label = "Copy/Paste Formula";
            this.copyPasteFormulaBtn.Name = "copyPasteFormulaBtn";
            this.copyPasteFormulaBtn.Tag = "Copy/Paste Formula";
            this.copyPasteFormulaBtn.Visible = false;
            // 
            // dbGroup
            // 
            this.dbGroup.Items.Add(this.btnConcatStrings);
            this.dbGroup.Items.Add(this.btnCreateStructure);
            this.dbGroup.Items.Add(this.dbSeperator);
            this.dbGroup.Items.Add(this.btnReadData);
            this.dbGroup.Items.Add(this.btnWriteData);
            this.dbGroup.Label = "Database";
            this.dbGroup.Name = "dbGroup";
            // 
            // btnConcatStrings
            // 
            this.btnConcatStrings.Label = "Concat Selection";
            this.btnConcatStrings.Name = "btnConcatStrings";
            this.btnConcatStrings.OfficeImageId = "ManageWebPartConnections";
            this.btnConcatStrings.ShowImage = true;
            this.btnConcatStrings.Tag = "Concat Selection";
            this.btnConcatStrings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConcat_Click);
            // 
            // btnCreateStructure
            // 
            this.btnCreateStructure.Label = "Create Table";
            this.btnCreateStructure.Name = "btnCreateStructure";
            this.btnCreateStructure.OfficeImageId = "CreateTableInDesignView";
            this.btnCreateStructure.ShowImage = true;
            this.btnCreateStructure.Tag = "Create Table";
            this.btnCreateStructure.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateStructure_Click);
            // 
            // dbSeperator
            // 
            this.dbSeperator.Name = "dbSeperator";
            // 
            // btnReadData
            // 
            this.btnReadData.Label = "Read Data";
            this.btnReadData.Name = "btnReadData";
            this.btnReadData.OfficeImageId = "GroupDatabaseModelingModel";
            this.btnReadData.ShowImage = true;
            this.btnReadData.Tag = "Read Data";
            this.btnReadData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReadData_Click);
            // 
            // btnWriteData
            // 
            this.btnWriteData.Label = "Write Data";
            this.btnWriteData.Name = "btnWriteData";
            this.btnWriteData.OfficeImageId = "DatabaseSqlServer";
            this.btnWriteData.ShowImage = true;
            this.btnWriteData.Tag = "Write Data";
            this.btnWriteData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWriteData_Click);
            // 
            // adminGroup
            // 
            this.adminGroup.Items.Add(this.aboutBtn);
            this.adminGroup.Items.Add(this.menuAddInSettings);
            this.adminGroup.Items.Add(this.settingsBtn);
            this.adminGroup.Label = "Administration";
            this.adminGroup.Name = "adminGroup";
            this.adminGroup.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.adminGroup_DialogLauncherClick);
            // 
            // aboutBtn
            // 
            this.aboutBtn.Label = "About";
            this.aboutBtn.Name = "aboutBtn";
            this.aboutBtn.OfficeImageId = "About";
            this.aboutBtn.ShowImage = true;
            this.aboutBtn.Tag = "About";
            this.aboutBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.aboutBtn_Click);
            // 
            // menuAddInSettings
            // 
            this.menuAddInSettings.Items.Add(this.addinSettingBtn);
            this.menuAddInSettings.Items.Add(this.button1);
            this.menuAddInSettings.Label = "Settings";
            this.menuAddInSettings.Name = "menuAddInSettings";
            this.menuAddInSettings.OfficeImageId = "OutlookGears";
            this.menuAddInSettings.ShowImage = true;
            this.menuAddInSettings.Tag = "Settings";
            // 
            // addinSettingBtn
            // 
            this.addinSettingBtn.Label = "Add-In Settings";
            this.addinSettingBtn.Name = "addinSettingBtn";
            this.addinSettingBtn.OfficeImageId = "OutlookGears";
            this.addinSettingBtn.ShowImage = true;
            this.addinSettingBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addinSettingBtn_Click);
            // 
            // button1
            // 
            this.button1.Label = "Add-In Config Items";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "OutlineSettings";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addinConfigItemesBtn_Click);
            // 
            // settingsBtn
            // 
            this.settingsBtn.Label = "";
            this.settingsBtn.Name = "settingsBtn";
            // 
            // MarsRibbon
            // 
            this.Name = "MarsRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.marsTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MarsRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.marsTab.ResumeLayout(false);
            this.marsTab.PerformLayout();
            this.rngGroup.ResumeLayout(false);
            this.rngGroup.PerformLayout();
            this.pivotGroup.ResumeLayout(false);
            this.pivotGroup.PerformLayout();
            this.wksGroup.ResumeLayout(false);
            this.wksGroup.PerformLayout();
            this.fmtGroup.ResumeLayout(false);
            this.fmtGroup.PerformLayout();
            this.copyPasteGroup.ResumeLayout(false);
            this.copyPasteGroup.PerformLayout();
            this.dbGroup.ResumeLayout(false);
            this.dbGroup.PerformLayout();
            this.adminGroup.ResumeLayout(false);
            this.adminGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab marsTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rngGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addRowsBtn;

        internal Microsoft.Office.Tools.Ribbon.RibbonButton addColumnsBtn;

        internal Microsoft.Office.Tools.Ribbon.RibbonButton dfltPivot;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton wksSheetListBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup wksGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton wksShowAllBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton wksResizeBtn;

        internal Microsoft.Office.Tools.Ribbon.RibbonButton dupPivotBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup adminGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton aboutBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup fmtGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton headerFormatBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rowShadingBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fmtNumberBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fmtDollarsBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup pivotGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton validationBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pasteSpecialButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton settingsBtn;

        internal Microsoft.Office.Tools.Ribbon.RibbonButton myFontBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup copyPasteGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup dbGroup;

        internal Microsoft.Office.Tools.Ribbon.RibbonButton dynamicNameRangeBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyPasteFormulaBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu defaultMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton gridlinesBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu nameMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fixedRangeBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addinSettingBtn;

        internal Microsoft.Office.Tools.Ribbon.RibbonButton grpUngrpBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showHideGrpBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tglFieldListBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tglFormulaBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showHideAllBtn;

        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuAddInSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConcatStrings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateStructure;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator dbSeperator;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReadData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWriteData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator rngSeparator;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator ptSeparator;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator wksSeparator;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator fmtSeparator;
    }

    partial class ThisRibbonCollection
    {
        internal MarsRibbon MarsRibbon
        {
            get { return this.GetRibbon<MarsRibbon>(); }
        }
    }
}
