namespace DocToSoc
{
    partial class ReportsDialog
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
            this.reportsListBox = new System.Windows.Forms.CheckedListBox();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // reportsListBox
            // 
            this.reportsListBox.FormattingEnabled = true;
            this.reportsListBox.Location = new System.Drawing.Point(21, 12);
            this.reportsListBox.Name = "reportsListBox";
            this.reportsListBox.Size = new System.Drawing.Size(248, 289);
            this.reportsListBox.Sorted = true;
            this.reportsListBox.TabIndex = 1;
            // 
            // okButton
            // 
            this.okButton.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::DocToSoc.Properties.Settings.Default, "okButton", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.okButton.Location = new System.Drawing.Point(49, 310);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 23);
            this.okButton.TabIndex = 0;
            this.okButton.Text = global::DocToSoc.Properties.Settings.Default.okButton;
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(161, 310);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // ReportsDialog
            // 
            this.AcceptButton = this.okButton;
            this.AccessibleRole = System.Windows.Forms.AccessibleRole.Dialog;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(281, 345);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.reportsListBox);
            this.Controls.Add(this.okButton);
            this.Name = "ReportsDialog";
            this.Text = "DocToSoc - Select Reports";
            this.TopMost = true;
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.CheckedListBox reportsListBox;
        private System.Windows.Forms.Button cancelButton;
    }
}