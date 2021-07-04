namespace MarsAddIn.Forms
{
    partial class AddToExcel
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
            this.lblAddToExcel = new System.Windows.Forms.Label();
            this.tboxAddToExcel = new System.Windows.Forms.TextBox();
            this.btnAddToExcelCancel = new System.Windows.Forms.Button();
            this.btnAddToExcelOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblAddToExcel
            // 
            this.lblAddToExcel.AutoSize = true;
            this.lblAddToExcel.Location = new System.Drawing.Point(13, 13);
            this.lblAddToExcel.Name = "lblAddToExcel";
            this.lblAddToExcel.Size = new System.Drawing.Size(197, 13);
            this.lblAddToExcel.TabIndex = 0;
            this.lblAddToExcel.Text = "How many columns do you wish to add?";
            // 
            // tboxAddToExcel
            // 
            this.tboxAddToExcel.Location = new System.Drawing.Point(16, 29);
            this.tboxAddToExcel.Name = "tboxAddToExcel";
            this.tboxAddToExcel.Size = new System.Drawing.Size(256, 20);
            this.tboxAddToExcel.TabIndex = 1;
            // 
            // btnAddToExcelCancel
            // 
            this.btnAddToExcelCancel.Location = new System.Drawing.Point(197, 56);
            this.btnAddToExcelCancel.Name = "btnAddToExcelCancel";
            this.btnAddToExcelCancel.Size = new System.Drawing.Size(75, 23);
            this.btnAddToExcelCancel.TabIndex = 2;
            this.btnAddToExcelCancel.Text = "Cancel";
            this.btnAddToExcelCancel.UseVisualStyleBackColor = true;
            this.btnAddToExcelCancel.Click += new System.EventHandler(this.btnAddToExcelCancel_Click);
            // 
            // btnAddToExcelOK
            // 
            this.btnAddToExcelOK.Location = new System.Drawing.Point(116, 56);
            this.btnAddToExcelOK.Name = "btnAddToExcelOK";
            this.btnAddToExcelOK.Size = new System.Drawing.Size(75, 23);
            this.btnAddToExcelOK.TabIndex = 3;
            this.btnAddToExcelOK.Text = "OK";
            this.btnAddToExcelOK.UseVisualStyleBackColor = true;
            this.btnAddToExcelOK.Click += new System.EventHandler(this.btnAddToExcelOK_Click);
            // 
            // AddToExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 91);
            this.Controls.Add(this.btnAddToExcelOK);
            this.Controls.Add(this.btnAddToExcelCancel);
            this.Controls.Add(this.tboxAddToExcel);
            this.Controls.Add(this.lblAddToExcel);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AddToExcel";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Add Columns";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblAddToExcel;
        private System.Windows.Forms.TextBox tboxAddToExcel;
        private System.Windows.Forms.Button btnAddToExcelCancel;
        private System.Windows.Forms.Button btnAddToExcelOK;
    }
}