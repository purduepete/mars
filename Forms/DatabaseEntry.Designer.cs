namespace MarsAddIn.Forms
{
    partial class DatabaseEntry
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
            this.comboTable = new System.Windows.Forms.ComboBox();
            this.btnPlus = new System.Windows.Forms.Button();
            this.btnMinus = new System.Windows.Forms.Button();
            this.btnDoAction = new System.Windows.Forms.Button();
            this.lblSource = new System.Windows.Forms.Label();
            this.lblTarget = new System.Windows.Forms.Label();
            this.cboxHeaderRow = new System.Windows.Forms.CheckBox();
            this.lblReturnRows = new System.Windows.Forms.Label();
            this.txtNumRows = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cBoxTrunc = new System.Windows.Forms.CheckBox();
            this.btnPlusPlus = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // comboServer
            // 
            this.comboServer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboServer.FormattingEnabled = true;
            this.comboServer.Location = new System.Drawing.Point(86, 6);
            this.comboServer.Name = "comboServer";
            this.comboServer.Size = new System.Drawing.Size(302, 21);
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
            this.comboDatabase.Size = new System.Drawing.Size(302, 21);
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
            this.lblTable.Click += new System.EventHandler(this.lblTable_Click);
            // 
            // comboTable
            // 
            this.comboTable.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboTable.FormattingEnabled = true;
            this.comboTable.Location = new System.Drawing.Point(86, 60);
            this.comboTable.Name = "comboTable";
            this.comboTable.Size = new System.Drawing.Size(302, 21);
            this.comboTable.TabIndex = 5;
            this.comboTable.SelectionChangeCommitted += new System.EventHandler(this.comboTableChanged_Committed);
            // 
            // btnPlus
            // 
            this.btnPlus.Location = new System.Drawing.Point(15, 125);
            this.btnPlus.Name = "btnPlus";
            this.btnPlus.Size = new System.Drawing.Size(20, 20);
            this.btnPlus.TabIndex = 6;
            this.btnPlus.Text = "+";
            this.btnPlus.UseVisualStyleBackColor = true;
            this.btnPlus.Click += new System.EventHandler(this.btnPlus_Click);
            // 
            // btnMinus
            // 
            this.btnMinus.Location = new System.Drawing.Point(39, 125);
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
            this.btnDoAction.Location = new System.Drawing.Point(211, 166);
            this.btnDoAction.Name = "btnDoAction";
            this.btnDoAction.Size = new System.Drawing.Size(75, 23);
            this.btnDoAction.TabIndex = 8;
            this.btnDoAction.Text = "Write Data";
            this.btnDoAction.UseVisualStyleBackColor = true;
            this.btnDoAction.Click += new System.EventHandler(this.btnDoAction_Click);
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.Location = new System.Drawing.Point(83, 129);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(41, 13);
            this.lblSource.TabIndex = 9;
            this.lblSource.Text = "Source";
            this.lblSource.Click += new System.EventHandler(this.lblSource_Click);
            this.lblSource.DoubleClick += new System.EventHandler(this.lblSource_DblClick);
            // 
            // lblTarget
            // 
            this.lblTarget.AutoSize = true;
            this.lblTarget.Location = new System.Drawing.Point(238, 129);
            this.lblTarget.Name = "lblTarget";
            this.lblTarget.Size = new System.Drawing.Size(38, 13);
            this.lblTarget.TabIndex = 10;
            this.lblTarget.Text = "Target";
            this.lblTarget.Click += new System.EventHandler(this.lblTarget_Click);
            this.lblTarget.DoubleClick += new System.EventHandler(this.lblTarget_DblClick);
            // 
            // cboxHeaderRow
            // 
            this.cboxHeaderRow.AutoSize = true;
            this.cboxHeaderRow.Checked = true;
            this.cboxHeaderRow.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cboxHeaderRow.Location = new System.Drawing.Point(86, 88);
            this.cboxHeaderRow.Name = "cboxHeaderRow";
            this.cboxHeaderRow.Size = new System.Drawing.Size(114, 17);
            this.cboxHeaderRow.TabIndex = 11;
            this.cboxHeaderRow.Text = "Has Header Row?";
            this.cboxHeaderRow.UseVisualStyleBackColor = true;
            // 
            // lblReturnRows
            // 
            this.lblReturnRows.AutoSize = true;
            this.lblReturnRows.Location = new System.Drawing.Point(12, 90);
            this.lblReturnRows.Name = "lblReturnRows";
            this.lblReturnRows.Size = new System.Drawing.Size(69, 13);
            this.lblReturnRows.TabIndex = 12;
            this.lblReturnRows.Text = "Return Rows";
            this.lblReturnRows.Visible = false;
            // 
            // txtNumRows
            // 
            this.txtNumRows.Location = new System.Drawing.Point(86, 87);
            this.txtNumRows.Name = "txtNumRows";
            this.txtNumRows.Size = new System.Drawing.Size(302, 20);
            this.txtNumRows.TabIndex = 13;
            this.txtNumRows.Visible = false;
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Location = new System.Drawing.Point(12, 116);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(385, 2);
            this.label1.TabIndex = 14;
            // 
            // cBoxTrunc
            // 
            this.cBoxTrunc.AutoSize = true;
            this.cBoxTrunc.Location = new System.Drawing.Point(279, 89);
            this.cBoxTrunc.Name = "cBoxTrunc";
            this.cBoxTrunc.Size = new System.Drawing.Size(109, 17);
            this.cBoxTrunc.TabIndex = 15;
            this.cBoxTrunc.Text = "Truncate Target?";
            this.cBoxTrunc.UseVisualStyleBackColor = true;
            // 
            // btnPlusPlus
            // 
            this.btnPlusPlus.Location = new System.Drawing.Point(63, 125);
            this.btnPlusPlus.Name = "btnPlusPlus";
            this.btnPlusPlus.Size = new System.Drawing.Size(20, 20);
            this.btnPlusPlus.TabIndex = 16;
            this.btnPlusPlus.Text = "|";
            this.btnPlusPlus.UseVisualStyleBackColor = true;
            this.btnPlusPlus.Click += new System.EventHandler(this.btnPlusPlus_Click);
            // 
            // DatabaseEntry
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(494, 197);
            this.Controls.Add(this.btnPlusPlus);
            this.Controls.Add(this.cBoxTrunc);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblReturnRows);
            this.Controls.Add(this.cboxHeaderRow);
            this.Controls.Add(this.lblTarget);
            this.Controls.Add(this.lblSource);
            this.Controls.Add(this.btnDoAction);
            this.Controls.Add(this.btnMinus);
            this.Controls.Add(this.btnPlus);
            this.Controls.Add(this.comboTable);
            this.Controls.Add(this.lblTable);
            this.Controls.Add(this.comboDatabase);
            this.Controls.Add(this.lblDatabase);
            this.Controls.Add(this.lblServer);
            this.Controls.Add(this.comboServer);
            this.Controls.Add(this.txtNumRows);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(510, 1000);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(510, 236);
            this.Name = "DatabaseEntry";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Write to Database";
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
        private System.Windows.Forms.ComboBox comboTable;
        private System.Windows.Forms.Button btnPlus;
        private System.Windows.Forms.Button btnMinus;
        private System.Windows.Forms.Button btnDoAction;
        private System.Windows.Forms.Label lblSource;
        private System.Windows.Forms.Label lblTarget;
        private System.Windows.Forms.CheckBox cboxHeaderRow;
        private System.Windows.Forms.Label lblReturnRows;
        private System.Windows.Forms.TextBox txtNumRows;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox cBoxTrunc;
        private System.Windows.Forms.Button btnPlusPlus;
    }
}