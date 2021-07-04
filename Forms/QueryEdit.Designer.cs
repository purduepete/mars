namespace MarsAddIn.Forms
{
    partial class QueryEdit
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
            this.txtQuery = new System.Windows.Forms.TextBox();
            this.btnQueryOK = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtQuery
            // 
            this.txtQuery.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtQuery.Location = new System.Drawing.Point(13, 13);
            this.txtQuery.Multiline = true;
            this.txtQuery.Name = "txtQuery";
            this.txtQuery.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtQuery.Size = new System.Drawing.Size(259, 207);
            this.txtQuery.TabIndex = 0;
            // 
            // btnQueryOK
            // 
            this.btnQueryOK.Location = new System.Drawing.Point(73, 226);
            this.btnQueryOK.Name = "btnQueryOK";
            this.btnQueryOK.Size = new System.Drawing.Size(75, 23);
            this.btnQueryOK.TabIndex = 1;
            this.btnQueryOK.Text = "OK";
            this.btnQueryOK.UseVisualStyleBackColor = true;
            this.btnQueryOK.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(154, 226);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 23);
            this.btnClear.TabIndex = 2;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // QueryEdit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnQueryOK);
            this.Controls.Add(this.txtQuery);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(300, 1000);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(300, 39);
            this.Name = "QueryEdit";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Query Editor";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtQuery;
        private System.Windows.Forms.Button btnQueryOK;
        private System.Windows.Forms.Button btnClear;
    }
}