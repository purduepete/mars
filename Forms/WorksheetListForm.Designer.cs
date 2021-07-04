namespace MarsAddIn.Forms
{
    partial class WorksheetListForm
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
            this.worksheetListFormHiddenButton = new System.Windows.Forms.Button();
            this.worksheetListFormVisibleButton = new System.Windows.Forms.Button();
            this.worksheetListFormVHiddenButton = new System.Windows.Forms.Button();
            this.worksheetListFormListView = new System.Windows.Forms.ListView();
            this.Worksheet = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Visibility = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.SuspendLayout();
            // 
            // worksheetListFormHiddenButton
            // 
            this.worksheetListFormHiddenButton.Location = new System.Drawing.Point(28, 330);
            this.worksheetListFormHiddenButton.Name = "worksheetListFormHiddenButton";
            this.worksheetListFormHiddenButton.Size = new System.Drawing.Size(75, 23);
            this.worksheetListFormHiddenButton.TabIndex = 1;
            this.worksheetListFormHiddenButton.Text = "Hidden";
            this.worksheetListFormHiddenButton.UseVisualStyleBackColor = true;
            this.worksheetListFormHiddenButton.Click += new System.EventHandler(this.worksheetListFormHiddenButton_Click);
            // 
            // worksheetListFormVisibleButton
            // 
            this.worksheetListFormVisibleButton.Location = new System.Drawing.Point(282, 330);
            this.worksheetListFormVisibleButton.Name = "worksheetListFormVisibleButton";
            this.worksheetListFormVisibleButton.Size = new System.Drawing.Size(75, 23);
            this.worksheetListFormVisibleButton.TabIndex = 3;
            this.worksheetListFormVisibleButton.Text = "Visible";
            this.worksheetListFormVisibleButton.UseVisualStyleBackColor = true;
            this.worksheetListFormVisibleButton.Click += new System.EventHandler(this.worksheetListFormVisibleButton_Click);
            // 
            // worksheetListFormVHiddenButton
            // 
            this.worksheetListFormVHiddenButton.Location = new System.Drawing.Point(156, 330);
            this.worksheetListFormVHiddenButton.Name = "worksheetListFormVHiddenButton";
            this.worksheetListFormVHiddenButton.Size = new System.Drawing.Size(75, 23);
            this.worksheetListFormVHiddenButton.TabIndex = 4;
            this.worksheetListFormVHiddenButton.Text = "Very Hidden";
            this.worksheetListFormVHiddenButton.UseVisualStyleBackColor = true;
            this.worksheetListFormVHiddenButton.Click += new System.EventHandler(this.worksheetListFormVHiddenButton_Click);
            // 
            // worksheetListFormListView
            // 
            this.worksheetListFormListView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.worksheetListFormListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Worksheet,
            this.Visibility});
            this.worksheetListFormListView.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.worksheetListFormListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.worksheetListFormListView.Location = new System.Drawing.Point(13, 12);
            this.worksheetListFormListView.Name = "worksheetListFormListView";
            this.worksheetListFormListView.ShowGroups = false;
            this.worksheetListFormListView.Size = new System.Drawing.Size(359, 312);
            this.worksheetListFormListView.TabIndex = 6;
            this.worksheetListFormListView.UseCompatibleStateImageBehavior = false;
            this.worksheetListFormListView.SelectedIndexChanged += new System.EventHandler(this.worksheetListFormListView_SelectedIndexChanged);
            this.worksheetListFormListView.DoubleClick += new System.EventHandler(this.worksheetListFormListView_DoubleClick);
            // 
            // WorksheetListForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 361);
            this.Controls.Add(this.worksheetListFormListView);
            this.Controls.Add(this.worksheetListFormVHiddenButton);
            this.Controls.Add(this.worksheetListFormVisibleButton);
            this.Controls.Add(this.worksheetListFormHiddenButton);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "WorksheetListForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Worksheet List";
            this.TopMost = true;
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button worksheetListFormHiddenButton;
        private System.Windows.Forms.Button worksheetListFormVisibleButton;
        private System.Windows.Forms.Button worksheetListFormVHiddenButton;
        private System.Windows.Forms.ListView worksheetListFormListView;
        private System.Windows.Forms.ColumnHeader Worksheet;
        private System.Windows.Forms.ColumnHeader Visibility;
    }
}