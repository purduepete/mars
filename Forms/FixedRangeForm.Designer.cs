namespace MarsAddIn.Forms
{
    partial class FixedRangeForm
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
            this.rngNameTextBox = new System.Windows.Forms.TextBox();
            this.rngNameLabel = new System.Windows.Forms.Label();
            this.globalCheckBox = new System.Windows.Forms.CheckBox();
            this.fixedRangeOKBtn = new System.Windows.Forms.Button();
            this.valueLabel = new System.Windows.Forms.Label();
            this.valueTextBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // rngNameTextBox
            // 
            this.rngNameTextBox.Location = new System.Drawing.Point(88, 12);
            this.rngNameTextBox.Name = "rngNameTextBox";
            this.rngNameTextBox.Size = new System.Drawing.Size(240, 20);
            this.rngNameTextBox.TabIndex = 0;
            // 
            // rngNameLabel
            // 
            this.rngNameLabel.AutoSize = true;
            this.rngNameLabel.Location = new System.Drawing.Point(12, 16);
            this.rngNameLabel.Name = "rngNameLabel";
            this.rngNameLabel.Size = new System.Drawing.Size(70, 13);
            this.rngNameLabel.TabIndex = 2;
            this.rngNameLabel.Text = "Range Name";
            // 
            // globalCheckBox
            // 
            this.globalCheckBox.AutoSize = true;
            this.globalCheckBox.Checked = true;
            this.globalCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.globalCheckBox.Location = new System.Drawing.Point(272, 43);
            this.globalCheckBox.Name = "globalCheckBox";
            this.globalCheckBox.Size = new System.Drawing.Size(56, 17);
            this.globalCheckBox.TabIndex = 2;
            this.globalCheckBox.Text = "Global";
            this.globalCheckBox.UseVisualStyleBackColor = true;
            // 
            // fixedRangeOKBtn
            // 
            this.fixedRangeOKBtn.Location = new System.Drawing.Point(135, 64);
            this.fixedRangeOKBtn.Name = "fixedRangeOKBtn";
            this.fixedRangeOKBtn.Size = new System.Drawing.Size(75, 23);
            this.fixedRangeOKBtn.TabIndex = 3;
            this.fixedRangeOKBtn.Text = "OK";
            this.fixedRangeOKBtn.UseVisualStyleBackColor = true;
            this.fixedRangeOKBtn.Click += new System.EventHandler(this.fixedRangeOKBtn_Click);
            // 
            // valueLabel
            // 
            this.valueLabel.AutoSize = true;
            this.valueLabel.Location = new System.Drawing.Point(12, 42);
            this.valueLabel.Name = "valueLabel";
            this.valueLabel.Size = new System.Drawing.Size(34, 13);
            this.valueLabel.TabIndex = 7;
            this.valueLabel.Text = "Value";
            // 
            // valueTextBox
            // 
            this.valueTextBox.Location = new System.Drawing.Point(88, 38);
            this.valueTextBox.Name = "valueTextBox";
            this.valueTextBox.Size = new System.Drawing.Size(168, 20);
            this.valueTextBox.TabIndex = 1;
            // 
            // FixedRangeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(340, 99);
            this.Controls.Add(this.valueTextBox);
            this.Controls.Add(this.valueLabel);
            this.Controls.Add(this.fixedRangeOKBtn);
            this.Controls.Add(this.globalCheckBox);
            this.Controls.Add(this.rngNameTextBox);
            this.Controls.Add(this.rngNameLabel);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FixedRangeForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Fixed Range";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox rngNameTextBox;
        private System.Windows.Forms.Label rngNameLabel;
        private System.Windows.Forms.CheckBox globalCheckBox;
        private System.Windows.Forms.Button fixedRangeOKBtn;
        private System.Windows.Forms.Label valueLabel;
        private System.Windows.Forms.TextBox valueTextBox;
    }
}