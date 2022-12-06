
namespace PersAhwal
{
    partial class FormDataBase
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
            this.components = new System.ComponentModel.Container();
            this.labTitel = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.Password = new System.Windows.Forms.TextBox();
            this.btnLog = new System.Windows.Forms.Button();
            this.dataGrid = new System.Windows.Forms.DataGridView();
            this.username = new System.Windows.Forms.TextBox();
            this.SeePass = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.appversion = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.timer2 = new System.Windows.Forms.Timer(this.components);
            this.green57 = new System.Windows.Forms.PictureBox();
            this.red57 = new System.Windows.Forms.PictureBox();
            this.labebserver57 = new System.Windows.Forms.Label();
            this.timer3 = new System.Windows.Forms.Timer(this.components);
            this.pass1 = new System.Windows.Forms.TextBox();
            this.pass2 = new System.Windows.Forms.TextBox();
            this.labelpass1 = new System.Windows.Forms.Label();
            this.labelpass2 = new System.Windows.Forms.Label();
            this.timer4 = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.green57)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.red57)).BeginInit();
            this.SuspendLayout();
            // 
            // labTitel
            // 
            this.labTitel.AutoSize = true;
            this.labTitel.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labTitel.Location = new System.Drawing.Point(422, 35);
            this.labTitel.Name = "labTitel";
            this.labTitel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labTitel.Size = new System.Drawing.Size(225, 27);
            this.labTitel.TabIndex = 14;
            this.labTitel.Text = "برنامج الاحوال الشخصية وشؤون الؤعايا";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(331, 14);
            this.label3.Name = "label3";
            this.label3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label3.Size = new System.Drawing.Size(341, 27);
            this.label3.TabIndex = 13;
            this.label3.Text = "القنصلية العامة لجمهورية السودان - جدة المملكة العربية السعودية";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(279, 129);
            this.label2.Name = "label2";
            this.label2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label2.Size = new System.Drawing.Size(72, 27);
            this.label2.TabIndex = 12;
            this.label2.Text = "كلمة المرور:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(305, 88);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(46, 27);
            this.label1.TabIndex = 10;
            this.label1.Text = "الاسم:";
            // 
            // Password
            // 
            this.Password.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Password.Location = new System.Drawing.Point(357, 129);
            this.Password.Name = "Password";
            this.Password.PasswordChar = '*';
            this.Password.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Password.Size = new System.Drawing.Size(252, 35);
            this.Password.TabIndex = 9;
            this.Password.UseSystemPasswordChar = true;
            this.Password.TextChanged += new System.EventHandler(this.Password_TextChanged);
            this.Password.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Password_KeyPress);
            // 
            // btnLog
            // 
            this.btnLog.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnLog.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLog.Location = new System.Drawing.Point(486, 168);
            this.btnLog.Name = "btnLog";
            this.btnLog.Size = new System.Drawing.Size(123, 36);
            this.btnLog.TabIndex = 8;
            this.btnLog.Text = "تسجيل دخول";
            this.btnLog.UseVisualStyleBackColor = false;
            this.btnLog.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // dataGrid
            // 
            this.dataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGrid.Location = new System.Drawing.Point(137, 331);
            this.dataGrid.Name = "dataGrid";
            this.dataGrid.Size = new System.Drawing.Size(553, 239);
            this.dataGrid.TabIndex = 15;
            // 
            // username
            // 
            this.username.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.username.Location = new System.Drawing.Point(357, 88);
            this.username.Name = "username";
            this.username.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.username.Size = new System.Drawing.Size(252, 35);
            this.username.TabIndex = 16;
            this.username.TextChanged += new System.EventHandler(this.employeeName_TextChanged);
            this.username.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.employeeName_KeyPress);
            // 
            // SeePass
            // 
            this.SeePass.AutoSize = true;
            this.SeePass.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SeePass.Location = new System.Drawing.Point(615, 131);
            this.SeePass.Name = "SeePass";
            this.SeePass.Size = new System.Drawing.Size(61, 31);
            this.SeePass.TabIndex = 18;
            this.SeePass.Text = "معاينة";
            this.SeePass.UseVisualStyleBackColor = true;
            this.SeePass.CheckedChanged += new System.EventHandler(this.SeePass_CheckedChanged);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.button1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(357, 168);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(123, 36);
            this.button1.TabIndex = 19;
            this.button1.Text = "تسجيل جديد";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::PersAhwal.Properties.Resources.WhatsApp_Image_2021_02_23_at_3_45_23_AM;
            this.pictureBox1.Location = new System.Drawing.Point(703, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(226, 221);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 17;
            this.pictureBox1.TabStop = false;
            // 
            // appversion
            // 
            this.appversion.AutoSize = true;
            this.appversion.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.appversion.Location = new System.Drawing.Point(859, 258);
            this.appversion.Name = "appversion";
            this.appversion.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.appversion.Size = new System.Drawing.Size(56, 27);
            this.appversion.TabIndex = 23;
            this.appversion.Text = "0.0.0.0";
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // timer2
            // 
            this.timer2.Enabled = true;
            this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
            // 
            // green57
            // 
            this.green57.Image = global::PersAhwal.Properties.Resources.green_circle;
            this.green57.Location = new System.Drawing.Point(12, 258);
            this.green57.Name = "green57";
            this.green57.Size = new System.Drawing.Size(28, 28);
            this.green57.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.green57.TabIndex = 26;
            this.green57.TabStop = false;
            this.green57.Visible = false;
            // 
            // red57
            // 
            this.red57.Image = global::PersAhwal.Properties.Resources.red;
            this.red57.Location = new System.Drawing.Point(12, 258);
            this.red57.Name = "red57";
            this.red57.Size = new System.Drawing.Size(28, 28);
            this.red57.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.red57.TabIndex = 25;
            this.red57.TabStop = false;
            this.red57.Visible = false;
            // 
            // labebserver57
            // 
            this.labebserver57.AutoSize = true;
            this.labebserver57.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labebserver57.Location = new System.Drawing.Point(46, 258);
            this.labebserver57.Name = "labebserver57";
            this.labebserver57.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labebserver57.Size = new System.Drawing.Size(367, 27);
            this.labebserver57.TabIndex = 29;
            this.labebserver57.Text = "مخدم قسمي الاحوال الشخصية وشؤون الرعايا متصلان بشكل صحيح";
            // 
            // timer3
            // 
            this.timer3.Interval = 10000;
            this.timer3.Tick += new System.EventHandler(this.timer3_Tick);
            // 
            // pass1
            // 
            this.pass1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pass1.Location = new System.Drawing.Point(357, 171);
            this.pass1.Name = "pass1";
            this.pass1.PasswordChar = '*';
            this.pass1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.pass1.Size = new System.Drawing.Size(252, 35);
            this.pass1.TabIndex = 31;
            this.pass1.UseSystemPasswordChar = true;
            this.pass1.Visible = false;
            // 
            // pass2
            // 
            this.pass2.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pass2.Location = new System.Drawing.Point(357, 212);
            this.pass2.Name = "pass2";
            this.pass2.PasswordChar = '*';
            this.pass2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.pass2.Size = new System.Drawing.Size(252, 35);
            this.pass2.TabIndex = 32;
            this.pass2.UseSystemPasswordChar = true;
            this.pass2.Visible = false;
            // 
            // labelpass1
            // 
            this.labelpass1.AutoSize = true;
            this.labelpass1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelpass1.Location = new System.Drawing.Point(238, 174);
            this.labelpass1.Name = "labelpass1";
            this.labelpass1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelpass1.Size = new System.Drawing.Size(113, 27);
            this.labelpass1.TabIndex = 33;
            this.labelpass1.Text = "كلمة المرور الجديدة:";
            this.labelpass1.Visible = false;
            // 
            // labelpass2
            // 
            this.labelpass2.AutoSize = true;
            this.labelpass2.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelpass2.Location = new System.Drawing.Point(210, 215);
            this.labelpass2.Name = "labelpass2";
            this.labelpass2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelpass2.Size = new System.Drawing.Size(136, 27);
            this.labelpass2.TabIndex = 34;
            this.labelpass2.Text = "إعادة ادخال كلمة المرور:";
            this.labelpass2.Visible = false;
            // 
            // timer4
            // 
            this.timer4.Enabled = true;
            this.timer4.Tick += new System.EventHandler(this.timer4_Tick);
            // 
            // FormDataBase
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(941, 294);
            this.Controls.Add(this.labelpass2);
            this.Controls.Add(this.labelpass1);
            this.Controls.Add(this.pass2);
            this.Controls.Add(this.labebserver57);
            this.Controls.Add(this.appversion);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.SeePass);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.Password);
            this.Controls.Add(this.username);
            this.Controls.Add(this.dataGrid);
            this.Controls.Add(this.labTitel);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnLog);
            this.Controls.Add(this.green57);
            this.Controls.Add(this.red57);
            this.Controls.Add(this.pass1);
            this.Name = "FormDataBase";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.RightToLeftLayout = true;
            this.Text = "تسجيل دخول برنامج أحوالك";
            this.Load += new System.EventHandler(this.FormDataBase_Load);
            this.MouseEnter += new System.EventHandler(this.FormDataBase_MouseEnter_1);
            this.MouseHover += new System.EventHandler(this.FormDataBase_MouseHover);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.green57)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.red57)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labTitel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox Password;
        private System.Windows.Forms.Button btnLog;
        private System.Windows.Forms.DataGridView dataGrid;
        private System.Windows.Forms.TextBox username;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.CheckBox SeePass;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label appversion;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Timer timer2;
        private System.Windows.Forms.PictureBox green57;
        private System.Windows.Forms.PictureBox red57;
        private System.Windows.Forms.Label labebserver57;
        private System.Windows.Forms.Timer timer3;
        private System.Windows.Forms.TextBox pass1;
        private System.Windows.Forms.TextBox pass2;
        private System.Windows.Forms.Label labelpass1;
        private System.Windows.Forms.Label labelpass2;
        private System.Windows.Forms.Timer timer4;
    }
}