namespace ExcelJsonEditorAddin
{
    partial class StartupControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.JsonLoadDialog = new System.Windows.Forms.OpenFileDialog();
            this.OpenDialogButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // JsonLoadDialog
            // 
            this.JsonLoadDialog.FileName = "jsonLoadDialog";
            // 
            // OpenDialogButton
            // 
            this.OpenDialogButton.Location = new System.Drawing.Point(52, 35);
            this.OpenDialogButton.Name = "OpenDialogButton";
            this.OpenDialogButton.Size = new System.Drawing.Size(136, 39);
            this.OpenDialogButton.TabIndex = 0;
            this.OpenDialogButton.Text = "Open Json";
            this.OpenDialogButton.UseVisualStyleBackColor = true;
            this.OpenDialogButton.Click += new System.EventHandler(this.OpenDialogButton_Click);
            // 
            // StartupControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDark;
            this.Controls.Add(this.OpenDialogButton);
            this.Name = "StartupControl";
            this.Size = new System.Drawing.Size(380, 513);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog JsonLoadDialog;
        private System.Windows.Forms.Button OpenDialogButton;
    }
}
