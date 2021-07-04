namespace MarsAddIn.Forms
{
    partial class DynamicRangeForm
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
            this.rngNameLabel = new System.Windows.Forms.Label();
            this.columnExcludeLabel = new System.Windows.Forms.Label();
            this.rowExcludeLabel = new System.Windows.Forms.Label();
            this.dynamicRangeOKBtn = new System.Windows.Forms.Button();
            this.rngNameTextBox = new System.Windows.Forms.TextBox();
            this.columnExcludeTextBox = new System.Windows.Forms.TextBox();
            this.rowExcludeTextBox = new System.Windows.Forms.TextBox();
            this.columnWidthTextBox = new System.Windows.Forms.TextBox();
            this.rowHeightTextBox = new System.Windows.Forms.TextBox();
            this.columnWidthLabel = new System.Windows.Forms.Label();
            this.rowHeightLabel = new System.Windows.Forms.Label();
            this.globalCheckBox = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // rngNameLabel
            // 
            this.rngNameLabel.AutoSize = true;
            this.rngNameLabel.Location = new System.Drawing.Point(12, 16);
            this.rngNameLabel.Name = "rngNameLabel";
            this.rngNameLabel.Size = new System.Drawing.Size(70, 13);
            this.rngNameLabel.TabIndex = 0;
            this.rngNameLabel.Text = "Range Name";
            // 
            // columnExcludeLabel
            // 
            this.columnExcludeLabel.AutoSize = true;
            this.columnExcludeLabel.Location = new System.Drawing.Point(12, 42);
            this.columnExcludeLabel.Name = "columnExcludeLabel";
            this.columnExcludeLabel.Size = new System.Drawing.Size(101, 13);
            this.columnExcludeLabel.TabIndex = 0;
            this.columnExcludeLabel.Text = "Column (excl/offset)";
            // 
            // rowExcludeLabel
            // 
            this.rowExcludeLabel.AutoSize = true;
            this.rowExcludeLabel.Location = new System.Drawing.Point(12, 68);
            this.rowExcludeLabel.Name = "rowExcludeLabel";
            this.rowExcludeLabel.Size = new System.Drawing.Size(93, 13);
            this.rowExcludeLabel.TabIndex = 0;
            this.rowExcludeLabel.Text = "Rows (excl/offset)";
            // 
            // dynamicRangeOKBtn
            // 
            this.dynamicRangeOKBtn.Location = new System.Drawing.Point(254, 91);
            this.dynamicRangeOKBtn.Name = "dynamicRangeOKBtn";
            this.dynamicRangeOKBtn.Size = new System.Drawing.Size(75, 23);
            this.dynamicRangeOKBtn.TabIndex = 5;
            this.dynamicRangeOKBtn.Text = "OK";
            this.dynamicRangeOKBtn.UseVisualStyleBackColor = true;
            this.dynamicRangeOKBtn.Click += new System.EventHandler(this.dynamicRangeOKBtn_Click);
            // 
            // rngNameTextBox
            // 
            this.rngNameTextBox.Location = new System.Drawing.Point(119, 13);
            this.rngNameTextBox.Name = "rngNameTextBox";
            this.rngNameTextBox.Size = new System.Drawing.Size(348, 20);
            this.rngNameTextBox.TabIndex = 0;
            // 
            // columnExcludeTextBox
            // 
            this.columnExcludeTextBox.Location = new System.Drawing.Point(119, 39);
            this.columnExcludeTextBox.Name = "columnExcludeTextBox";
            this.columnExcludeTextBox.Size = new System.Drawing.Size(134, 20);
            this.columnExcludeTextBox.TabIndex = 1;
            this.columnExcludeTextBox.Text = "0";
            // 
            // rowExcludeTextBox
            // 
            this.rowExcludeTextBox.Location = new System.Drawing.Point(119, 64);
            this.rowExcludeTextBox.Name = "rowExcludeTextBox";
            this.rowExcludeTextBox.Size = new System.Drawing.Size(134, 20);
            this.rowExcludeTextBox.TabIndex = 2;
            this.rowExcludeTextBox.Text = "0";
            // 
            // columnWidthTextBox
            // 
            this.columnWidthTextBox.Location = new System.Drawing.Point(403, 38);
            this.columnWidthTextBox.Name = "columnWidthTextBox";
            this.columnWidthTextBox.Size = new System.Drawing.Size(136, 20);
            this.columnWidthTextBox.TabIndex = 3;
            this.columnWidthTextBox.Text = "-1";
            // 
            // rowHeightTextBox
            // 
            this.rowHeightTextBox.Location = new System.Drawing.Point(403, 64);
            this.rowHeightTextBox.Name = "rowHeightTextBox";
            this.rowHeightTextBox.Size = new System.Drawing.Size(136, 20);
            this.rowHeightTextBox.TabIndex = 4;
            this.rowHeightTextBox.Text = "-1";
            // 
            // columnWidthLabel
            // 
            this.columnWidthLabel.AutoSize = true;
            this.columnWidthLabel.Location = new System.Drawing.Point(324, 42);
            this.columnWidthLabel.Name = "columnWidthLabel";
            this.columnWidthLabel.Size = new System.Drawing.Size(73, 13);
            this.columnWidthLabel.TabIndex = 0;
            this.columnWidthLabel.Text = "Column Width";
            // 
            // rowHeightLabel
            // 
            this.rowHeightLabel.AutoSize = true;
            this.rowHeightLabel.Location = new System.Drawing.Point(324, 68);
            this.rowHeightLabel.Name = "rowHeightLabel";
            this.rowHeightLabel.Size = new System.Drawing.Size(63, 13);
            this.rowHeightLabel.TabIndex = 0;
            this.rowHeightLabel.Text = "Row Height";
            // 
            // globalCheckBox
            // 
            this.globalCheckBox.AutoSize = true;
            this.globalCheckBox.Checked = true;
            this.globalCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.globalCheckBox.Location = new System.Drawing.Point(483, 16);
            this.globalCheckBox.Name = "globalCheckBox";
            this.globalCheckBox.Size = new System.Drawing.Size(56, 17);
            this.globalCheckBox.TabIndex = 6;
            this.globalCheckBox.Text = "Global";
            this.globalCheckBox.UseVisualStyleBackColor = true;
            // 
            // DynamicRangeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(554, 126);
            this.Controls.Add(this.globalCheckBox);
            this.Controls.Add(this.rowHeightTextBox);
            this.Controls.Add(this.columnWidthTextBox);
            this.Controls.Add(this.rowExcludeTextBox);
            this.Controls.Add(this.columnExcludeTextBox);
            this.Controls.Add(this.rngNameTextBox);
            this.Controls.Add(this.dynamicRangeOKBtn);
            this.Controls.Add(this.rowHeightLabel);
            this.Controls.Add(this.columnWidthLabel);
            this.Controls.Add(this.rowExcludeLabel);
            this.Controls.Add(this.columnExcludeLabel);
            this.Controls.Add(this.rngNameLabel);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DynamicRangeForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Dynamic Range";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label rngNameLabel;
        private System.Windows.Forms.Label columnExcludeLabel;
        private System.Windows.Forms.Label rowExcludeLabel;
        private System.Windows.Forms.Button dynamicRangeOKBtn;
        private System.Windows.Forms.TextBox rngNameTextBox;
        private System.Windows.Forms.TextBox columnExcludeTextBox;
        private System.Windows.Forms.TextBox rowExcludeTextBox;
        private System.Windows.Forms.TextBox columnWidthTextBox;
        private System.Windows.Forms.TextBox rowHeightTextBox;
        private System.Windows.Forms.Label columnWidthLabel;
        private System.Windows.Forms.Label rowHeightLabel;
        private System.Windows.Forms.CheckBox globalCheckBox;
    }
}