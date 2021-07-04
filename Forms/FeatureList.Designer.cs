namespace MarsAddIn.Forms
{
    partial class frmFeatureList
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.checkedListBoxRange = new System.Windows.Forms.CheckedListBox();
            this.checkedListBoxPivotTable = new System.Windows.Forms.CheckedListBox();
            this.lblPivotTable = new System.Windows.Forms.Label();
            this.lblWorksheet = new System.Windows.Forms.Label();
            this.checkedListBoxWorksheet = new System.Windows.Forms.CheckedListBox();
            this.btnFLOK = new System.Windows.Forms.Button();
            this.lblFormatting = new System.Windows.Forms.Label();
            this.checkedListBoxFormatting = new System.Windows.Forms.CheckedListBox();
            this.lblCopyPaste = new System.Windows.Forms.Label();
            this.checkedListBoxCopyPaste = new System.Windows.Forms.CheckedListBox();
            this.lblDatabase = new System.Windows.Forms.Label();
            this.checkedListBoxDatabase = new System.Windows.Forms.CheckedListBox();
            this.lblFile = new System.Windows.Forms.Label();
            this.checkedListBoxFile = new System.Windows.Forms.CheckedListBox();
            this.lblAdministration = new System.Windows.Forms.Label();
            this.checkedListBoxAdmin = new System.Windows.Forms.CheckedListBox();
            this.lblRange = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // checkedListBoxRange
            // 
            this.checkedListBoxRange.AllowDrop = true;
            this.checkedListBoxRange.CheckOnClick = true;
            this.checkedListBoxRange.FormattingEnabled = true;
            this.checkedListBoxRange.Items.AddRange(new object[] {
            "Add Rows",
            "Add Columns",
            "Validation",
            "Names"});
            this.checkedListBoxRange.Location = new System.Drawing.Point(12, 29);
            this.checkedListBoxRange.Name = "checkedListBoxRange";
            this.checkedListBoxRange.Size = new System.Drawing.Size(225, 64);
            this.checkedListBoxRange.TabIndex = 0;
            this.checkedListBoxRange.ThreeDCheckBoxes = true;
            // 
            // checkedListBoxPivotTable
            // 
            this.checkedListBoxPivotTable.CheckOnClick = true;
            this.checkedListBoxPivotTable.FormattingEnabled = true;
            this.checkedListBoxPivotTable.Items.AddRange(new object[] {
            "Default Pivot",
            "Duplicate Pivot",
            "Toggle Field List",
            "Toggle Formulas"});
            this.checkedListBoxPivotTable.Location = new System.Drawing.Point(12, 118);
            this.checkedListBoxPivotTable.Name = "checkedListBoxPivotTable";
            this.checkedListBoxPivotTable.Size = new System.Drawing.Size(225, 64);
            this.checkedListBoxPivotTable.TabIndex = 1;
            // 
            // lblPivotTable
            // 
            this.lblPivotTable.AutoSize = true;
            this.lblPivotTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPivotTable.Location = new System.Drawing.Point(9, 101);
            this.lblPivotTable.Name = "lblPivotTable";
            this.lblPivotTable.Size = new System.Drawing.Size(72, 13);
            this.lblPivotTable.TabIndex = 3;
            this.lblPivotTable.Text = "Pivot Table";
            this.lblPivotTable.Click += new System.EventHandler(this.lblPivotTable_Click);
            // 
            // lblWorksheet
            // 
            this.lblWorksheet.AutoSize = true;
            this.lblWorksheet.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWorksheet.Location = new System.Drawing.Point(12, 190);
            this.lblWorksheet.Name = "lblWorksheet";
            this.lblWorksheet.Size = new System.Drawing.Size(68, 13);
            this.lblWorksheet.TabIndex = 4;
            this.lblWorksheet.Text = "Worksheet";
            this.lblWorksheet.Click += new System.EventHandler(this.lblWorksheet_Click);
            // 
            // checkedListBoxWorksheet
            // 
            this.checkedListBoxWorksheet.CheckOnClick = true;
            this.checkedListBoxWorksheet.FormattingEnabled = true;
            this.checkedListBoxWorksheet.Items.AddRange(new object[] {
            "Worksheet List",
            "Show All Worksheets",
            "Resize Worksheet",
            "Group/Ungroup",
            "Show/Hide Group",
            "Show/Hide All"});
            this.checkedListBoxWorksheet.Location = new System.Drawing.Point(12, 207);
            this.checkedListBoxWorksheet.Name = "checkedListBoxWorksheet";
            this.checkedListBoxWorksheet.Size = new System.Drawing.Size(225, 94);
            this.checkedListBoxWorksheet.TabIndex = 5;
            // 
            // btnFLOK
            // 
            this.btnFLOK.Location = new System.Drawing.Point(86, 679);
            this.btnFLOK.Name = "btnFLOK";
            this.btnFLOK.Size = new System.Drawing.Size(75, 23);
            this.btnFLOK.TabIndex = 6;
            this.btnFLOK.Text = "OK";
            this.btnFLOK.UseVisualStyleBackColor = true;
            this.btnFLOK.Click += new System.EventHandler(this.btnFLOK_Click);
            // 
            // lblFormatting
            // 
            this.lblFormatting.AutoSize = true;
            this.lblFormatting.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFormatting.Location = new System.Drawing.Point(12, 309);
            this.lblFormatting.Name = "lblFormatting";
            this.lblFormatting.Size = new System.Drawing.Size(66, 13);
            this.lblFormatting.TabIndex = 7;
            this.lblFormatting.Text = "Formatting";
            this.lblFormatting.Click += new System.EventHandler(this.lblFormatting_Click);
            // 
            // checkedListBoxFormatting
            // 
            this.checkedListBoxFormatting.CheckOnClick = true;
            this.checkedListBoxFormatting.FormattingEnabled = true;
            this.checkedListBoxFormatting.Items.AddRange(new object[] {
            "Header Format",
            "Row Shading",
            "Preferred Settings",
            "Format Number",
            "Format Number (min)",
            "Format Dollar"});
            this.checkedListBoxFormatting.Location = new System.Drawing.Point(12, 326);
            this.checkedListBoxFormatting.Name = "checkedListBoxFormatting";
            this.checkedListBoxFormatting.Size = new System.Drawing.Size(225, 94);
            this.checkedListBoxFormatting.TabIndex = 8;
            // 
            // lblCopyPaste
            // 
            this.lblCopyPaste.AutoSize = true;
            this.lblCopyPaste.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCopyPaste.Location = new System.Drawing.Point(12, 428);
            this.lblCopyPaste.Name = "lblCopyPaste";
            this.lblCopyPaste.Size = new System.Drawing.Size(73, 13);
            this.lblCopyPaste.TabIndex = 9;
            this.lblCopyPaste.Text = "Copy/Paste";
            this.lblCopyPaste.Click += new System.EventHandler(this.lblCopyPaste_Click);
            // 
            // checkedListBoxCopyPaste
            // 
            this.checkedListBoxCopyPaste.CheckOnClick = true;
            this.checkedListBoxCopyPaste.FormattingEnabled = true;
            this.checkedListBoxCopyPaste.Items.AddRange(new object[] {
            "Paste Special"});
            this.checkedListBoxCopyPaste.Location = new System.Drawing.Point(12, 445);
            this.checkedListBoxCopyPaste.Name = "checkedListBoxCopyPaste";
            this.checkedListBoxCopyPaste.Size = new System.Drawing.Size(225, 19);
            this.checkedListBoxCopyPaste.TabIndex = 10;
            // 
            // lblDatabase
            // 
            this.lblDatabase.AutoSize = true;
            this.lblDatabase.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDatabase.Location = new System.Drawing.Point(12, 472);
            this.lblDatabase.Name = "lblDatabase";
            this.lblDatabase.Size = new System.Drawing.Size(61, 13);
            this.lblDatabase.TabIndex = 11;
            this.lblDatabase.Text = "Database";
            this.lblDatabase.Click += new System.EventHandler(this.lblDatabase_Click);
            // 
            // checkedListBoxDatabase
            // 
            this.checkedListBoxDatabase.CheckOnClick = true;
            this.checkedListBoxDatabase.FormattingEnabled = true;
            this.checkedListBoxDatabase.Items.AddRange(new object[] {
            "Concat Selection",
            "Create Table",
            "Read Data",
            "Write Data"});
            this.checkedListBoxDatabase.Location = new System.Drawing.Point(12, 489);
            this.checkedListBoxDatabase.Name = "checkedListBoxDatabase";
            this.checkedListBoxDatabase.Size = new System.Drawing.Size(225, 64);
            this.checkedListBoxDatabase.TabIndex = 12;
            // 
            // lblFile
            // 
            this.lblFile.AutoSize = true;
            this.lblFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFile.Location = new System.Drawing.Point(12, 562);
            this.lblFile.Name = "lblFile";
            this.lblFile.Size = new System.Drawing.Size(27, 13);
            this.lblFile.TabIndex = 13;
            this.lblFile.Text = "File";
            this.lblFile.Click += new System.EventHandler(this.lblFile_Click);
            // 
            // checkedListBoxFile
            // 
            this.checkedListBoxFile.CheckOnClick = true;
            this.checkedListBoxFile.FormattingEnabled = true;
            this.checkedListBoxFile.Items.AddRange(new object[] {
            "File Path",
            "My Files"});
            this.checkedListBoxFile.Location = new System.Drawing.Point(12, 579);
            this.checkedListBoxFile.Name = "checkedListBoxFile";
            this.checkedListBoxFile.Size = new System.Drawing.Size(225, 34);
            this.checkedListBoxFile.TabIndex = 14;
            // 
            // lblAdministration
            // 
            this.lblAdministration.AutoSize = true;
            this.lblAdministration.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAdministration.Location = new System.Drawing.Point(12, 621);
            this.lblAdministration.Name = "lblAdministration";
            this.lblAdministration.Size = new System.Drawing.Size(86, 13);
            this.lblAdministration.TabIndex = 15;
            this.lblAdministration.Text = "Administration";
            this.lblAdministration.Click += new System.EventHandler(this.lblAdministration_Click);
            // 
            // checkedListBoxAdmin
            // 
            this.checkedListBoxAdmin.CheckOnClick = true;
            this.checkedListBoxAdmin.FormattingEnabled = true;
            this.checkedListBoxAdmin.Items.AddRange(new object[] {
            "About",
            "Enable Add-Ins"});
            this.checkedListBoxAdmin.Location = new System.Drawing.Point(12, 638);
            this.checkedListBoxAdmin.Name = "checkedListBoxAdmin";
            this.checkedListBoxAdmin.Size = new System.Drawing.Size(225, 34);
            this.checkedListBoxAdmin.TabIndex = 16;
            // 
            // lblRange
            // 
            this.lblRange.AutoSize = true;
            this.lblRange.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRange.Location = new System.Drawing.Point(12, 12);
            this.lblRange.Name = "lblRange";
            this.lblRange.Size = new System.Drawing.Size(44, 13);
            this.lblRange.TabIndex = 17;
            this.lblRange.Text = "Range";
            this.lblRange.Click += new System.EventHandler(this.lblRange_Click);
            // 
            // frmFeatureList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(249, 714);
            this.Controls.Add(this.lblRange);
            this.Controls.Add(this.checkedListBoxAdmin);
            this.Controls.Add(this.lblAdministration);
            this.Controls.Add(this.checkedListBoxFile);
            this.Controls.Add(this.lblFile);
            this.Controls.Add(this.checkedListBoxDatabase);
            this.Controls.Add(this.lblDatabase);
            this.Controls.Add(this.checkedListBoxCopyPaste);
            this.Controls.Add(this.lblCopyPaste);
            this.Controls.Add(this.checkedListBoxFormatting);
            this.Controls.Add(this.lblFormatting);
            this.Controls.Add(this.btnFLOK);
            this.Controls.Add(this.checkedListBoxWorksheet);
            this.Controls.Add(this.lblWorksheet);
            this.Controls.Add(this.lblPivotTable);
            this.Controls.Add(this.checkedListBoxPivotTable);
            this.Controls.Add(this.checkedListBoxRange);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(265, 800);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(265, 39);
            this.Name = "frmFeatureList";
            this.ShowIcon = false;
            this.Text = "Feature List";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox checkedListBoxRange;
        private System.Windows.Forms.CheckedListBox checkedListBoxPivotTable;
        private System.Windows.Forms.Label lblPivotTable;
        private System.Windows.Forms.Label lblWorksheet;
        private System.Windows.Forms.CheckedListBox checkedListBoxWorksheet;
        private System.Windows.Forms.Button btnFLOK;
        private System.Windows.Forms.Label lblFormatting;
        private System.Windows.Forms.CheckedListBox checkedListBoxFormatting;
        private System.Windows.Forms.Label lblCopyPaste;
        private System.Windows.Forms.CheckedListBox checkedListBoxCopyPaste;
        private System.Windows.Forms.Label lblDatabase;
        private System.Windows.Forms.CheckedListBox checkedListBoxDatabase;
        private System.Windows.Forms.Label lblFile;
        private System.Windows.Forms.CheckedListBox checkedListBoxFile;
        private System.Windows.Forms.Label lblAdministration;
        private System.Windows.Forms.CheckedListBox checkedListBoxAdmin;
        private System.Windows.Forms.Label lblRange;
    }
}