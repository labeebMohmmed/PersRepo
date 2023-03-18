
namespace PersAhwal
{
    partial class Form8
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle21 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle22 = new System.Windows.Forms.DataGridViewCellStyle();
            this.قائمة_النصوص_العامة = new System.Windows.Forms.ComboBox();
            this.label13 = new System.Windows.Forms.Label();
            this.AppType = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.النص = new System.Windows.Forms.TextBox();
            this.panelChar = new System.Windows.Forms.FlowLayoutPanel();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.textBox12 = new System.Windows.Forms.TextBox();
            this.textBox13 = new System.Windows.Forms.TextBox();
            this.textBox14 = new System.Windows.Forms.TextBox();
            this.textBox15 = new System.Windows.Forms.TextBox();
            this.textBox16 = new System.Windows.Forms.TextBox();
            this.عدد_النماذج = new System.Windows.Forms.Button();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.نص_مرجعي = new System.Windows.Forms.Button();
            this.تعيين_كخيار = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.picStar = new System.Windows.Forms.PictureBox();
            this.قائمة_النصوص_الفرعية = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.panelChar.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picStar)).BeginInit();
            this.SuspendLayout();
            // 
            // قائمة_النصوص_العامة
            // 
            this.قائمة_النصوص_العامة.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.قائمة_النصوص_العامة.FormattingEnabled = true;
            this.قائمة_النصوص_العامة.Items.AddRange(new object[] {
            "عدم ممانعة زواج",
            "عدم ممانعة وشهادة كفاءة"});
            this.قائمة_النصوص_العامة.Location = new System.Drawing.Point(922, 41);
            this.قائمة_النصوص_العامة.Name = "قائمة_النصوص_العامة";
            this.قائمة_النصوص_العامة.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.قائمة_النصوص_العامة.Size = new System.Drawing.Size(302, 35);
            this.قائمة_النصوص_العامة.TabIndex = 531;
            this.قائمة_النصوص_العامة.SelectedIndexChanged += new System.EventHandler(this.قائمة_النصوص_SelectedIndexChanged);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(1230, 44);
            this.label13.Name = "label13";
            this.label13.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label13.Size = new System.Drawing.Size(77, 27);
            this.label13.TabIndex = 530;
            this.label13.Text = "القائمة العامة:";
            // 
            // AppType
            // 
            this.AppType.AutoSize = true;
            this.AppType.Checked = true;
            this.AppType.CheckState = System.Windows.Forms.CheckState.Checked;
            this.AppType.Enabled = false;
            this.AppType.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.AppType.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.AppType.Location = new System.Drawing.Point(1113, 1);
            this.AppType.Name = "AppType";
            this.AppType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AppType.Size = new System.Drawing.Size(111, 31);
            this.AppType.TabIndex = 532;
            this.AppType.Text = "مجموع المعاملات";
            this.AppType.UseVisualStyleBackColor = true;
            this.AppType.CheckedChanged += new System.EventHandler(this.AppType_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(1230, 5);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(64, 27);
            this.label1.TabIndex = 533;
            this.label1.Text = "نوع النص:";
            // 
            // النص
            // 
            this.النص.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.النص.Location = new System.Drawing.Point(922, 125);
            this.النص.Multiline = true;
            this.النص.Name = "النص";
            this.النص.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.النص.Size = new System.Drawing.Size(389, 532);
            this.النص.TabIndex = 632;
            this.النص.Visible = false;
            // 
            // panelChar
            // 
            this.panelChar.AutoScroll = true;
            this.panelChar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelChar.Controls.Add(this.textBox9);
            this.panelChar.Controls.Add(this.textBox10);
            this.panelChar.Controls.Add(this.textBox11);
            this.panelChar.Controls.Add(this.textBox12);
            this.panelChar.Controls.Add(this.textBox13);
            this.panelChar.Controls.Add(this.textBox14);
            this.panelChar.Controls.Add(this.textBox15);
            this.panelChar.Controls.Add(this.textBox16);
            this.panelChar.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.panelChar.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panelChar.Location = new System.Drawing.Point(696, 4);
            this.panelChar.Name = "panelChar";
            this.panelChar.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.panelChar.Size = new System.Drawing.Size(219, 653);
            this.panelChar.TabIndex = 833;
            // 
            // textBox9
            // 
            this.textBox9.Enabled = false;
            this.textBox9.Font = new System.Drawing.Font("Arabic Typesetting", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox9.Location = new System.Drawing.Point(3, 3);
            this.textBox9.Multiline = true;
            this.textBox9.Name = "textBox9";
            this.textBox9.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.textBox9.Size = new System.Drawing.Size(211, 59);
            this.textBox9.TabIndex = 354;
            this.textBox9.Text = "1- في حالة النهاية (/ت - كثال أوقع/أوقعت) يستخدم الرمز #*#";
            // 
            // textBox10
            // 
            this.textBox10.Enabled = false;
            this.textBox10.Font = new System.Drawing.Font("Arabic Typesetting", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox10.Location = new System.Drawing.Point(3, 68);
            this.textBox10.Multiline = true;
            this.textBox10.Name = "textBox10";
            this.textBox10.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.textBox10.Size = new System.Drawing.Size(211, 30);
            this.textBox10.TabIndex = 355;
            this.textBox10.Text = "2- الذي/التي #1";
            // 
            // textBox11
            // 
            this.textBox11.Enabled = false;
            this.textBox11.Font = new System.Drawing.Font("Arabic Typesetting", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox11.Location = new System.Drawing.Point(3, 104);
            this.textBox11.Multiline = true;
            this.textBox11.Name = "textBox11";
            this.textBox11.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.textBox11.Size = new System.Drawing.Size(211, 59);
            this.textBox11.TabIndex = 356;
            this.textBox11.Text = "3- في حالة النهاية  (ت/نا-مثال أذنت/أذنا) يستخدم الرمز &&&";
            // 
            // textBox12
            // 
            this.textBox12.Enabled = false;
            this.textBox12.Font = new System.Drawing.Font("Arabic Typesetting", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox12.Location = new System.Drawing.Point(3, 169);
            this.textBox12.Multiline = true;
            this.textBox12.Name = "textBox12";
            this.textBox12.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.textBox12.Size = new System.Drawing.Size(211, 63);
            this.textBox12.TabIndex = 635;
            this.textBox12.Text = "4- في حالة النهاية  (ي/ا- مثال عني/عنا) يستخدم الرمز $$$";
            // 
            // textBox13
            // 
            this.textBox13.Enabled = false;
            this.textBox13.Font = new System.Drawing.Font("Arabic Typesetting", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox13.Location = new System.Drawing.Point(3, 238);
            this.textBox13.Multiline = true;
            this.textBox13.Name = "textBox13";
            this.textBox13.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.textBox13.Size = new System.Drawing.Size(211, 62);
            this.textBox13.TabIndex = 636;
            this.textBox13.Text = "5- في حالة النهاية  (/ت/ا/تا/ن/وا -مثال أوكل/أوكلت/أوكلا) يستخدم الرمز ***";
            // 
            // textBox14
            // 
            this.textBox14.Enabled = false;
            this.textBox14.Font = new System.Drawing.Font("Arabic Typesetting", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox14.Location = new System.Drawing.Point(3, 306);
            this.textBox14.Multiline = true;
            this.textBox14.Name = "textBox14";
            this.textBox14.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.textBox14.Size = new System.Drawing.Size(211, 60);
            this.textBox14.TabIndex = 637;
            this.textBox14.Text = "6- في حالة النهاية  (ه/ها/هما/هما/من/هم- مثال له/لها/لهم) يستخدم الرمز ###";
            // 
            // textBox15
            // 
            this.textBox15.Enabled = false;
            this.textBox15.Font = new System.Drawing.Font("Arabic Typesetting", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox15.Location = new System.Drawing.Point(3, 372);
            this.textBox15.Multiline = true;
            this.textBox15.Name = "textBox15";
            this.textBox15.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.textBox15.Size = new System.Drawing.Size(211, 60);
            this.textBox15.TabIndex = 638;
            this.textBox15.Text = "7- في حالة النهاية  (هو/هي/هما/هن/هم) يستخدم الرمز #2";
            // 
            // textBox16
            // 
            this.textBox16.Enabled = false;
            this.textBox16.Font = new System.Drawing.Font("Arabic Typesetting", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox16.Location = new System.Drawing.Point(3, 438);
            this.textBox16.Multiline = true;
            this.textBox16.Name = "textBox16";
            this.textBox16.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.textBox16.Size = new System.Drawing.Size(211, 63);
            this.textBox16.TabIndex = 639;
            this.textBox16.Text = "8- الاسم tN/رقم الوثيقة tP/الاإصدار tS/الجنس tX/نوع الوثيقة tD/اللقب  tT";
            // 
            // عدد_النماذج
            // 
            this.عدد_النماذج.BackColor = System.Drawing.SystemColors.Control;
            this.عدد_النماذج.FlatAppearance.BorderSize = 0;
            this.عدد_النماذج.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.عدد_النماذج.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.عدد_النماذج.Location = new System.Drawing.Point(15, 663);
            this.عدد_النماذج.Name = "عدد_النماذج";
            this.عدد_النماذج.Size = new System.Drawing.Size(675, 35);
            this.عدد_النماذج.TabIndex = 835;
            this.عدد_النماذج.Text = "عدد المعاملات";
            this.عدد_النماذج.UseVisualStyleBackColor = false;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.AutoScroll = true;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(12, 1);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(678, 656);
            this.flowLayoutPanel1.TabIndex = 836;
            // 
            // نص_مرجعي
            // 
            this.نص_مرجعي.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.نص_مرجعي.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.نص_مرجعي.Location = new System.Drawing.Point(922, 660);
            this.نص_مرجعي.Name = "نص_مرجعي";
            this.نص_مرجعي.Size = new System.Drawing.Size(124, 35);
            this.نص_مرجعي.TabIndex = 837;
            this.نص_مرجعي.Text = "تعيين كنص مرجعي";
            this.نص_مرجعي.UseVisualStyleBackColor = true;
            this.نص_مرجعي.Click += new System.EventHandler(this.نص_مرجعي_Click);
            // 
            // تعيين_كخيار
            // 
            this.تعيين_كخيار.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.تعيين_كخيار.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.تعيين_كخيار.Location = new System.Drawing.Point(1052, 660);
            this.تعيين_كخيار.Name = "تعيين_كخيار";
            this.تعيين_كخيار.Size = new System.Drawing.Size(85, 35);
            this.تعيين_كخيار.TabIndex = 838;
            this.تعيين_كخيار.Text = "تعيين كخيار";
            this.تعيين_كخيار.UseVisualStyleBackColor = true;
            this.تعيين_كخيار.Click += new System.EventHandler(this.تعيين_كخيار_Click);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button1.Location = new System.Drawing.Point(1143, 660);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(55, 35);
            this.button1.TabIndex = 839;
            this.button1.Text = "حذف";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle21.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle21.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle21.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle21.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle21.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle21.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle21.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle21;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle22.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle22.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle22.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle22.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle22.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle22.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle22.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle22;
            this.dataGridView1.Location = new System.Drawing.Point(1006, 377);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.dataGridView1.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.dataGridView1.RowTemplate.DefaultCellStyle.BackColor = System.Drawing.Color.White;
            this.dataGridView1.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Arabic Typesetting", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dataGridView1.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            this.dataGridView1.RowTemplate.DefaultCellStyle.Padding = new System.Windows.Forms.Padding(4);
            this.dataGridView1.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.LightGreen;
            this.dataGridView1.RowTemplate.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.RowTemplate.Height = 30;
            this.dataGridView1.Size = new System.Drawing.Size(192, 103);
            this.dataGridView1.TabIndex = 840;
            this.dataGridView1.Visible = false;
            // 
            // picStar
            // 
            this.picStar.Image = global::PersAhwal.Properties.Resources.star;
            this.picStar.Location = new System.Drawing.Point(922, 1);
            this.picStar.Name = "picStar";
            this.picStar.Size = new System.Drawing.Size(47, 39);
            this.picStar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picStar.TabIndex = 841;
            this.picStar.TabStop = false;
            this.picStar.Visible = false;
            // 
            // قائمة_النصوص_الفرعية
            // 
            this.قائمة_النصوص_الفرعية.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.قائمة_النصوص_الفرعية.FormattingEnabled = true;
            this.قائمة_النصوص_الفرعية.Items.AddRange(new object[] {
            "عدم ممانعة زواج",
            "عدم ممانعة وشهادة كفاءة"});
            this.قائمة_النصوص_الفرعية.Location = new System.Drawing.Point(922, 82);
            this.قائمة_النصوص_الفرعية.Name = "قائمة_النصوص_الفرعية";
            this.قائمة_النصوص_الفرعية.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.قائمة_النصوص_الفرعية.Size = new System.Drawing.Size(302, 35);
            this.قائمة_النصوص_الفرعية.TabIndex = 843;
            this.قائمة_النصوص_الفرعية.SelectedIndexChanged += new System.EventHandler(this.قائمة_النصوص_الفرعية_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(1230, 85);
            this.label2.Name = "label2";
            this.label2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label2.Size = new System.Drawing.Size(83, 27);
            this.label2.TabIndex = 842;
            this.label2.Text = "القائمة الفرعية:";
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button2.Location = new System.Drawing.Point(1201, 660);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(53, 35);
            this.button2.TabIndex = 844;
            this.button2.Text = "تعديل";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button3.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button3.Location = new System.Drawing.Point(1258, 660);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(53, 35);
            this.button3.TabIndex = 845;
            this.button3.Text = "إضافة";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // Form8
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1335, 748);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.قائمة_النصوص_الفرعية);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.picStar);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.تعيين_كخيار);
            this.Controls.Add(this.نص_مرجعي);
            this.Controls.Add(this.flowLayoutPanel1);
            this.Controls.Add(this.عدد_النماذج);
            this.Controls.Add(this.panelChar);
            this.Controls.Add(this.النص);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.AppType);
            this.Controls.Add(this.قائمة_النصوص_العامة);
            this.Controls.Add(this.label13);
            this.Name = "Form8";
            this.Text = "المساعد النصي";
            this.Load += new System.EventHandler(this.Form8_Load);
            this.panelChar.ResumeLayout(false);
            this.panelChar.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picStar)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ComboBox قائمة_النصوص_العامة;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.CheckBox AppType;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox النص;
        private System.Windows.Forms.FlowLayoutPanel panelChar;
        private System.Windows.Forms.TextBox textBox9;
        private System.Windows.Forms.TextBox textBox10;
        private System.Windows.Forms.TextBox textBox11;
        private System.Windows.Forms.TextBox textBox12;
        private System.Windows.Forms.TextBox textBox13;
        private System.Windows.Forms.TextBox textBox14;
        private System.Windows.Forms.TextBox textBox15;
        private System.Windows.Forms.TextBox textBox16;
        private System.Windows.Forms.Button عدد_النماذج;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Button نص_مرجعي;
        private System.Windows.Forms.Button تعيين_كخيار;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.PictureBox picStar;
        private System.Windows.Forms.ComboBox قائمة_النصوص_الفرعية;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
    }
}