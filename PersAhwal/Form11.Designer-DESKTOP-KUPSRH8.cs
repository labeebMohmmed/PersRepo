
namespace PersAhwal
{
    partial class Form11
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
            this.userApplicant1 = new PersAhwal.UserApplicant();
            this.SuspendLayout();
            // 
            // userApplicant1
            // 
            this.userApplicant1.Location = new System.Drawing.Point(12, 12);
            this.userApplicant1.Name = "userApplicant1";
            this.userApplicant1.Size = new System.Drawing.Size(1333, 695);
            this.userApplicant1.TabIndex = 0;
            // 
            // Form11
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1373, 725);
            this.Controls.Add(this.userApplicant1);
            this.Name = "Form11";
            this.Text = "Form11";
            this.ResumeLayout(false);

        }

        #endregion

        private UserApplicant userApplicant1;
    }
}