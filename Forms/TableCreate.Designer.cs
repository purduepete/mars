namespace MarsAddIn.Forms
{
    partial class TableCreate
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
            this.comboServer = new System.Windows.Forms.ComboBox();
            this.lblServer = new System.Windows.Forms.Label();
            this.lblDatabase = new System.Windows.Forms.Label();
            this.comboDatabase = new System.Windows.Forms.ComboBox();
            this.lblTable = new System.Windows.Forms.Label();
            this.btnPlus = new System.Windows.Forms.Button();
            this.btnMinus = new System.Windows.Forms.Button();
            this.btnDoAction = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tboxTableName = new System.Windows.Forms.TextBox();
            this.lblColumName = new System.Windows.Forms.Label();
            this.lblDataType = new System.Windows.Forms.Label();
            this.lblColumnSize = new System.Windows.Forms.Label();
            this.lblId = new System.Windows.Forms.Label();
            this.lblPK = new System.Windows.Forms.Label();
            this.cboxRecreate = new System.Windows.Forms.CheckBox();
            this.lblNull = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // comboServer
            // 
            this.comboServer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboServer.FormattingEnabled = true;
            this.comboServer.Location = new System.Drawing.Point(86, 6);
            this.comboServer.Name = "comboServer";
            this.comboServer.Size = new System.Drawing.Size(377, 21);
            this.comboServer.TabIndex = 0;
            this.comboServer.SelectionChangeCommitted += new System.EventHandler(this.comboServerChanged_Committed);
            // 
            // lblServer
            // 
            this.lblServer.AutoSize = true;
            this.lblServer.Location = new System.Drawing.Point(12, 9);
            this.lblServer.Name = "lblServer";
            this.lblServer.Size = new System.Drawing.Size(38, 13);
            this.lblServer.TabIndex = 1;
            this.lblServer.Text = "Server";
            // 
            // lblDatabase
            // 
            this.lblDatabase.AutoSize = true;
            this.lblDatabase.Location = new System.Drawing.Point(12, 36);
            this.lblDatabase.Name = "lblDatabase";
            this.lblDatabase.Size = new System.Drawing.Size(53, 13);
            this.lblDatabase.TabIndex = 2;
            this.lblDatabase.Text = "Database";
            // 
            // comboDatabase
            // 
            this.comboDatabase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboDatabase.FormattingEnabled = true;
            this.comboDatabase.Location = new System.Drawing.Point(86, 33);
            this.comboDatabase.Name = "comboDatabase";
            this.comboDatabase.Size = new System.Drawing.Size(377, 21);
            this.comboDatabase.TabIndex = 3;
            this.comboDatabase.SelectionChangeCommitted += new System.EventHandler(this.comboDatabaseChanged_Committed);
            // 
            // lblTable
            // 
            this.lblTable.AutoSize = true;
            this.lblTable.Location = new System.Drawing.Point(12, 63);
            this.lblTable.Name = "lblTable";
            this.lblTable.Size = new System.Drawing.Size(34, 13);
            this.lblTable.TabIndex = 4;
            this.lblTable.Text = "Table";
            // 
            // btnPlus
            // 
            this.btnPlus.Location = new System.Drawing.Point(15, 98);
            this.btnPlus.Name = "btnPlus";
            this.btnPlus.Size = new System.Drawing.Size(20, 20);
            this.btnPlus.TabIndex = 6;
            this.btnPlus.Text = "+";
            this.btnPlus.UseVisualStyleBackColor = true;
            this.btnPlus.Click += new System.EventHandler(this.btnPlus_Click);
            // 
            // btnMinus
            // 
            this.btnMinus.Location = new System.Drawing.Point(39, 98);
            this.btnMinus.Name = "btnMinus";
            this.btnMinus.Size = new System.Drawing.Size(20, 20);
            this.btnMinus.TabIndex = 7;
            this.btnMinus.Text = "-";
            this.btnMinus.UseVisualStyleBackColor = true;
            this.btnMinus.Click += new System.EventHandler(this.btnMinus_Click);
            // 
            // btnDoAction
            // 
            this.btnDoAction.Enabled = false;
            this.btnDoAction.Location = new System.Drawing.Point(167, 132);
            this.btnDoAction.Name = "btnDoAction";
            this.btnDoAction.Size = new System.Drawing.Size(84, 23);
            this.btnDoAction.TabIndex = 8;
            this.btnDoAction.Text = "Create Table";
            this.btnDoAction.UseVisualStyleBackColor = true;
            this.btnDoAction.Click += new System.EventHandler(this.btnDoAction_Click);
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Location = new System.Drawing.Point(12, 89);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(455, 2);
            this.label1.TabIndex = 14;
            // 
            // tboxTableName
            // 
            this.tboxTableName.Location = new System.Drawing.Point(86, 60);
            this.tboxTableName.Name = "tboxTableName";
            this.tboxTableName.Size = new System.Drawing.Size(285, 20);
            this.tboxTableName.TabIndex = 15;
            this.tboxTableName.TextChanged += new System.EventHandler(this.tboxTableName_TextChanged);
            // 
            // lblColumName
            // 
            this.lblColumName.AutoSize = true;
            this.lblColumName.Location = new System.Drawing.Point(83, 98);
            this.lblColumName.Name = "lblColumName";
            this.lblColumName.Size = new System.Drawing.Size(73, 13);
            this.lblColumName.TabIndex = 16;
            this.lblColumName.Text = "Column Name";
            this.lblColumName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblDataType
            // 
            this.lblDataType.AutoSize = true;
            this.lblDataType.Location = new System.Drawing.Point(175, 98);
            this.lblDataType.Name = "lblDataType";
            this.lblDataType.Size = new System.Drawing.Size(57, 13);
            this.lblDataType.TabIndex = 17;
            this.lblDataType.Text = "Data Type";
            this.lblDataType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblColumnSize
            // 
            this.lblColumnSize.AutoSize = true;
            this.lblColumnSize.Location = new System.Drawing.Point(256, 98);
            this.lblColumnSize.Name = "lblColumnSize";
            this.lblColumnSize.Size = new System.Drawing.Size(40, 13);
            this.lblColumnSize.TabIndex = 18;
            this.lblColumnSize.Text = "Length";
            this.lblColumnSize.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblId
            // 
            this.lblId.AutoSize = true;
            this.lblId.Location = new System.Drawing.Point(311, 98);
            this.lblId.Name = "lblId";
            this.lblId.Size = new System.Drawing.Size(24, 13);
            this.lblId.TabIndex = 19;
            this.lblId.Text = "ID?";
            this.lblId.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblPK
            // 
            this.lblPK.AutoSize = true;
            this.lblPK.Location = new System.Drawing.Point(351, 98);
            this.lblPK.Name = "lblPK";
            this.lblPK.Size = new System.Drawing.Size(27, 13);
            this.lblPK.TabIndex = 20;
            this.lblPK.Text = "PK?";
            this.lblPK.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cboxRecreate
            // 
            this.cboxRecreate.AutoSize = true;
            this.cboxRecreate.Location = new System.Drawing.Point(387, 62);
            this.cboxRecreate.Name = "cboxRecreate";
            this.cboxRecreate.Size = new System.Drawing.Size(76, 17);
            this.cboxRecreate.TabIndex = 21;
            this.cboxRecreate.Text = "Recreate?";
            this.cboxRecreate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.cboxRecreate.UseVisualStyleBackColor = true;
            // 
            // lblNull
            // 
            this.lblNull.AutoSize = true;
            this.lblNull.Location = new System.Drawing.Point(384, 98);
            this.lblNull.Name = "lblNull";
            this.lblNull.Size = new System.Drawing.Size(51, 13);
            this.lblNull.TabIndex = 22;
            this.lblNull.Text = "Nullable?";
            this.lblNull.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // TableCreate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 167);
            this.Controls.Add(this.lblNull);
            this.Controls.Add(this.cboxRecreate);
            this.Controls.Add(this.lblPK);
            this.Controls.Add(this.lblId);
            this.Controls.Add(this.lblColumnSize);
            this.Controls.Add(this.lblDataType);
            this.Controls.Add(this.lblColumName);
            this.Controls.Add(this.tboxTableName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnDoAction);
            this.Controls.Add(this.btnMinus);
            this.Controls.Add(this.btnPlus);
            this.Controls.Add(this.lblTable);
            this.Controls.Add(this.comboDatabase);
            this.Controls.Add(this.lblDatabase);
            this.Controls.Add(this.lblServer);
            this.Controls.Add(this.comboServer);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(500, 1000);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(500, 206);
            this.Name = "TableCreate";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Create Table";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboServer;
        private System.Windows.Forms.Label lblServer;
        private System.Windows.Forms.Label lblDatabase;
        private System.Windows.Forms.ComboBox comboDatabase;
        private System.Windows.Forms.Label lblTable;
        private System.Windows.Forms.Button btnPlus;
        private System.Windows.Forms.Button btnMinus;
        private System.Windows.Forms.Button btnDoAction;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tboxTableName;
        private System.Windows.Forms.Label lblColumName;
        private System.Windows.Forms.Label lblDataType;
        private System.Windows.Forms.Label lblColumnSize;
        private System.Windows.Forms.Label lblId;
        private System.Windows.Forms.Label lblPK;
        private System.Windows.Forms.CheckBox cboxRecreate;
        private System.Windows.Forms.Label lblNull;
    }
}