
namespace PersAhwal
{
    partial class Form6
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
            this.timer2 = new System.Windows.Forms.Timer(this.components);
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.Ifadaid = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.HijriDate = new System.Windows.Forms.TextBox();
            this.AttendViceConsul = new System.Windows.Forms.ComboBox();
            this.ApplicantSex = new System.Windows.Forms.CheckBox();
            this.label11 = new System.Windows.Forms.Label();
            this.GregorianDate = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.PassIssuedSource = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.ApplicantPassNo = new System.Windows.Forms.TextBox();
            this.labeldoctype = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.ApplicantIdocName = new System.Windows.Forms.TextBox();
            this.labelName = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.ApplicantIqamaNo = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.IqamaIssuedSource = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.checkedViewed = new System.Windows.Forms.CheckBox();
            this.mandoubLabel = new System.Windows.Forms.Label();
            this.mandoubName = new System.Windows.Forms.ComboBox();
            this.AppType = new System.Windows.Forms.CheckBox();
            this.label21 = new System.Windows.Forms.Label();
            this.ConsulateEmployee = new System.Windows.Forms.Label();
            this.btnprintOnly = new System.Windows.Forms.Button();
            this.SaveOnly = new System.Windows.Forms.Button();
            this.btnSavePrint = new System.Windows.Forms.Button();
            this.ResetAll = new System.Windows.Forms.Button();
            this.SearchFile = new System.Windows.Forms.TextBox();
            this.labelArch = new System.Windows.Forms.Label();
            this.ArchivedSt = new System.Windows.Forms.CheckBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.ListSearch = new System.Windows.Forms.TextBox();
            this.SearchDoc = new System.Windows.Forms.Button();
            this.label24 = new System.Windows.Forms.Label();
            this.Comment = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // timer2
            // 
            this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(287, 306);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(97, 27);
            this.label1.TabIndex = 210;
            this.label1.Text = "اسم موقع الإفادة:";
            // 
            // Ifadaid
            // 
            this.Ifadaid.Enabled = false;
            this.Ifadaid.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Ifadaid.Location = new System.Drawing.Point(39, 54);
            this.Ifadaid.Name = "Ifadaid";
            this.Ifadaid.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Ifadaid.Size = new System.Drawing.Size(178, 35);
            this.Ifadaid.TabIndex = 208;
            this.Ifadaid.Text = "ق س ج/160/xyz";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label19.Location = new System.Drawing.Point(226, 57);
            this.label19.Name = "label19";
            this.label19.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label19.Size = new System.Drawing.Size(64, 27);
            this.label19.TabIndex = 207;
            this.label19.Text = "رقم الإفادة:";
            // 
            // HijriDate
            // 
            this.HijriDate.Enabled = false;
            this.HijriDate.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.HijriDate.Location = new System.Drawing.Point(39, 95);
            this.HijriDate.Name = "HijriDate";
            this.HijriDate.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.HijriDate.Size = new System.Drawing.Size(178, 35);
            this.HijriDate.TabIndex = 206;
            // 
            // AttendViceConsul
            // 
            this.AttendViceConsul.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AttendViceConsul.FormattingEnabled = true;
            this.AttendViceConsul.Items.AddRange(new object[] {
            "محمد عثمان عكاشة الحسين",
            "يوسف صديق أبوعاقلة",
            "لبيب محمد أحمد"});
            this.AttendViceConsul.Location = new System.Drawing.Point(28, 303);
            this.AttendViceConsul.Name = "AttendViceConsul";
            this.AttendViceConsul.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AttendViceConsul.Size = new System.Drawing.Size(248, 35);
            this.AttendViceConsul.TabIndex = 204;
            // 
            // ApplicantSex
            // 
            this.ApplicantSex.AutoSize = true;
            this.ApplicantSex.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ApplicantSex.Location = new System.Drawing.Point(1007, 96);
            this.ApplicantSex.Name = "ApplicantSex";
            this.ApplicantSex.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ApplicantSex.Size = new System.Drawing.Size(49, 31);
            this.ApplicantSex.TabIndex = 200;
            this.ApplicantSex.Text = "ذكر";
            this.ApplicantSex.UseVisualStyleBackColor = true;
            this.ApplicantSex.CheckedChanged += new System.EventHandler(this.ApplicantSex_CheckedChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(223, 95);
            this.label11.Name = "label11";
            this.label11.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label11.Size = new System.Drawing.Size(90, 27);
            this.label11.TabIndex = 199;
            this.label11.Text = "التاريخ الهجري:";
            // 
            // GregorianDate
            // 
            this.GregorianDate.Enabled = false;
            this.GregorianDate.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GregorianDate.Location = new System.Drawing.Point(39, 136);
            this.GregorianDate.Name = "GregorianDate";
            this.GregorianDate.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.GregorianDate.Size = new System.Drawing.Size(178, 35);
            this.GregorianDate.TabIndex = 197;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(226, 136);
            this.label12.Name = "label12";
            this.label12.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label12.Size = new System.Drawing.Size(94, 27);
            this.label12.TabIndex = 198;
            this.label12.Text = "التاريخ الميلادي:";
            // 
            // PassIssuedSource
            // 
            this.PassIssuedSource.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PassIssuedSource.Location = new System.Drawing.Point(782, 173);
            this.PassIssuedSource.Name = "PassIssuedSource";
            this.PassIssuedSource.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.PassIssuedSource.Size = new System.Drawing.Size(268, 35);
            this.PassIssuedSource.TabIndex = 195;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(1067, 176);
            this.label7.Name = "label7";
            this.label7.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label7.Size = new System.Drawing.Size(145, 27);
            this.label7.TabIndex = 196;
            this.label7.Text = "مكان إصدار جواز السفر:";
            // 
            // ApplicantPassNo
            // 
            this.ApplicantPassNo.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ApplicantPassNo.Location = new System.Drawing.Point(782, 133);
            this.ApplicantPassNo.Name = "ApplicantPassNo";
            this.ApplicantPassNo.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ApplicantPassNo.Size = new System.Drawing.Size(268, 35);
            this.ApplicantPassNo.TabIndex = 193;
            this.ApplicantPassNo.Text = "P";
            // 
            // labeldoctype
            // 
            this.labeldoctype.AutoSize = true;
            this.labeldoctype.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labeldoctype.Location = new System.Drawing.Point(1061, 133);
            this.labeldoctype.Name = "labeldoctype";
            this.labeldoctype.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labeldoctype.Size = new System.Drawing.Size(100, 27);
            this.labeldoctype.TabIndex = 194;
            this.labeldoctype.Text = "رقم جواز السفر: ";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(1067, 96);
            this.label3.Name = "label3";
            this.label3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label3.Size = new System.Drawing.Size(40, 27);
            this.label3.TabIndex = 191;
            this.label3.Text = "النوع:";
            // 
            // ApplicantIdocName
            // 
            this.ApplicantIdocName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ApplicantIdocName.Location = new System.Drawing.Point(782, 54);
            this.ApplicantIdocName.Name = "ApplicantIdocName";
            this.ApplicantIdocName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ApplicantIdocName.Size = new System.Drawing.Size(268, 35);
            this.ApplicantIdocName.TabIndex = 189;
            // 
            // labelName
            // 
            this.labelName.AutoSize = true;
            this.labelName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelName.Location = new System.Drawing.Point(1067, 60);
            this.labelName.Name = "labelName";
            this.labelName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelName.Size = new System.Drawing.Size(80, 27);
            this.labelName.TabIndex = 190;
            this.labelName.Text = "مقدم الطلب:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(1067, 219);
            this.label2.Name = "label2";
            this.label2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label2.Size = new System.Drawing.Size(70, 27);
            this.label2.TabIndex = 194;
            this.label2.Text = "رقم الاقامة:";
            // 
            // ApplicantIqamaNo
            // 
            this.ApplicantIqamaNo.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ApplicantIqamaNo.Location = new System.Drawing.Point(782, 215);
            this.ApplicantIqamaNo.Name = "ApplicantIqamaNo";
            this.ApplicantIqamaNo.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ApplicantIqamaNo.Size = new System.Drawing.Size(268, 35);
            this.ApplicantIqamaNo.TabIndex = 193;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(1067, 258);
            this.label4.Name = "label4";
            this.label4.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label4.Size = new System.Drawing.Size(116, 27);
            this.label4.TabIndex = 196;
            this.label4.Text = "مكان إصدار الإقامة:";
            // 
            // IqamaIssuedSource
            // 
            this.IqamaIssuedSource.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.IqamaIssuedSource.Location = new System.Drawing.Point(782, 255);
            this.IqamaIssuedSource.Name = "IqamaIssuedSource";
            this.IqamaIssuedSource.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.IqamaIssuedSource.Size = new System.Drawing.Size(268, 35);
            this.IqamaIssuedSource.TabIndex = 195;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(28, 457);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.dataGridView1.Size = new System.Drawing.Size(1204, 192);
            this.dataGridView1.TabIndex = 240;
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            // 
            // checkedViewed
            // 
            this.checkedViewed.AutoSize = true;
            this.checkedViewed.Checked = true;
            this.checkedViewed.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkedViewed.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F);
            this.checkedViewed.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.checkedViewed.Location = new System.Drawing.Point(134, 2);
            this.checkedViewed.Name = "checkedViewed";
            this.checkedViewed.Size = new System.Drawing.Size(151, 33);
            this.checkedViewed.TabIndex = 239;
            this.checkedViewed.Text = "NotViewed";
            this.checkedViewed.UseVisualStyleBackColor = true;
            this.checkedViewed.Visible = false;
            // 
            // mandoubLabel
            // 
            this.mandoubLabel.AutoSize = true;
            this.mandoubLabel.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.mandoubLabel.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.mandoubLabel.Location = new System.Drawing.Point(287, 255);
            this.mandoubLabel.Name = "mandoubLabel";
            this.mandoubLabel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.mandoubLabel.Size = new System.Drawing.Size(79, 27);
            this.mandoubLabel.TabIndex = 238;
            this.mandoubLabel.Text = "اسم المندوب:";
            this.mandoubLabel.Visible = false;
            // 
            // mandoubName
            // 
            this.mandoubName.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.mandoubName.FormattingEnabled = true;
            this.mandoubName.Items.AddRange(new object[] {
            "عثمان محمد أحمد ضي النور",
            "خالد الشيخ دفع الله التجاني"});
            this.mandoubName.Location = new System.Drawing.Point(28, 252);
            this.mandoubName.Name = "mandoubName";
            this.mandoubName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.mandoubName.Size = new System.Drawing.Size(248, 35);
            this.mandoubName.TabIndex = 237;
            this.mandoubName.Visible = false;
            // 
            // AppType
            // 
            this.AppType.AutoSize = true;
            this.AppType.Checked = true;
            this.AppType.CheckState = System.Windows.Forms.CheckState.Checked;
            this.AppType.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.AppType.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.AppType.Location = new System.Drawing.Point(82, 213);
            this.AppType.Name = "AppType";
            this.AppType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AppType.Size = new System.Drawing.Size(170, 31);
            this.AppType.TabIndex = 236;
            this.AppType.Text = "حضور مباشرة إلى القنصلية";
            this.AppType.UseVisualStyleBackColor = true;
            this.AppType.CheckedChanged += new System.EventHandler(this.AppType_CheckedChanged);
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.label21.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label21.Location = new System.Drawing.Point(287, 217);
            this.label21.Name = "label21";
            this.label21.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label21.Size = new System.Drawing.Size(115, 27);
            this.label21.TabIndex = 235;
            this.label21.Text = "طريقة تقديم الطلب:";
            // 
            // ConsulateEmployee
            // 
            this.ConsulateEmployee.AutoSize = true;
            this.ConsulateEmployee.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.ConsulateEmployee.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.ConsulateEmployee.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ConsulateEmployee.Location = new System.Drawing.Point(28, 9);
            this.ConsulateEmployee.Name = "ConsulateEmployee";
            this.ConsulateEmployee.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ConsulateEmployee.Size = new System.Drawing.Size(51, 27);
            this.ConsulateEmployee.TabIndex = 244;
            this.ConsulateEmployee.Text = "الموظف";
            // 
            // btnprintOnly
            // 
            this.btnprintOnly.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnprintOnly.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnprintOnly.Location = new System.Drawing.Point(76, 367);
            this.btnprintOnly.Name = "btnprintOnly";
            this.btnprintOnly.Size = new System.Drawing.Size(43, 71);
            this.btnprintOnly.TabIndex = 404;
            this.btnprintOnly.Text = "طباعة";
            this.btnprintOnly.UseVisualStyleBackColor = false;
            this.btnprintOnly.Visible = false;
            this.btnprintOnly.Click += new System.EventHandler(this.printOnly_Click);
            // 
            // SaveOnly
            // 
            this.SaveOnly.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.SaveOnly.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.SaveOnly.Location = new System.Drawing.Point(26, 367);
            this.SaveOnly.Name = "SaveOnly";
            this.SaveOnly.Size = new System.Drawing.Size(45, 71);
            this.SaveOnly.TabIndex = 403;
            this.SaveOnly.Text = "حفظ";
            this.SaveOnly.UseVisualStyleBackColor = false;
            this.SaveOnly.Visible = false;
            this.SaveOnly.Click += new System.EventHandler(this.SaveOnly_Click);
            // 
            // btnSavePrint
            // 
            this.btnSavePrint.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnSavePrint.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnSavePrint.Location = new System.Drawing.Point(25, 367);
            this.btnSavePrint.Name = "btnSavePrint";
            this.btnSavePrint.Size = new System.Drawing.Size(94, 71);
            this.btnSavePrint.TabIndex = 402;
            this.btnSavePrint.Text = "طباعة وحفظ";
            this.btnSavePrint.UseVisualStyleBackColor = false;
            this.btnSavePrint.Click += new System.EventHandler(this.btnSavePrint_Click);
            // 
            // ResetAll
            // 
            this.ResetAll.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.ResetAll.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ResetAll.Location = new System.Drawing.Point(125, 367);
            this.ResetAll.Name = "ResetAll";
            this.ResetAll.Size = new System.Drawing.Size(92, 71);
            this.ResetAll.TabIndex = 401;
            this.ResetAll.Text = "مسح جميع الحقول";
            this.ResetAll.UseVisualStyleBackColor = false;
            this.ResetAll.Click += new System.EventHandler(this.ResetAll_Click);
            // 
            // SearchFile
            // 
            this.SearchFile.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.SearchFile.Location = new System.Drawing.Point(782, 402);
            this.SearchFile.Name = "SearchFile";
            this.SearchFile.Size = new System.Drawing.Size(319, 35);
            this.SearchFile.TabIndex = 400;
            this.SearchFile.Visible = false;
            // 
            // labelArch
            // 
            this.labelArch.AutoSize = true;
            this.labelArch.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.labelArch.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.labelArch.Location = new System.Drawing.Point(344, 408);
            this.labelArch.Name = "labelArch";
            this.labelArch.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelArch.Size = new System.Drawing.Size(79, 27);
            this.labelArch.TabIndex = 399;
            this.labelArch.Text = "حالة الأرشفة:";
            // 
            // ArchivedSt
            // 
            this.ArchivedSt.AutoSize = true;
            this.ArchivedSt.BackColor = System.Drawing.Color.Red;
            this.ArchivedSt.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.ArchivedSt.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ArchivedSt.Location = new System.Drawing.Point(235, 404);
            this.ArchivedSt.Name = "ArchivedSt";
            this.ArchivedSt.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ArchivedSt.Size = new System.Drawing.Size(98, 31);
            this.ArchivedSt.TabIndex = 398;
            this.ArchivedSt.Text = "غير مؤرشف";
            this.ArchivedSt.UseVisualStyleBackColor = false;
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button4.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button4.Location = new System.Drawing.Point(436, 404);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(114, 33);
            this.button4.TabIndex = 397;
            this.button4.Text = "فتح المستند رقم 2";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button3.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button3.Location = new System.Drawing.Point(1103, 402);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(129, 33);
            this.button3.TabIndex = 396;
            this.button3.Text = "بحث القائمة أدناه";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button2.Location = new System.Drawing.Point(556, 404);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(112, 33);
            this.button2.TabIndex = 395;
            this.button2.Text = "فتح المستند رقم 1";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // ListSearch
            // 
            this.ListSearch.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.ListSearch.Location = new System.Drawing.Point(782, 402);
            this.ListSearch.Name = "ListSearch";
            this.ListSearch.Size = new System.Drawing.Size(319, 35);
            this.ListSearch.TabIndex = 394;
            // 
            // SearchDoc
            // 
            this.SearchDoc.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.SearchDoc.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.SearchDoc.Location = new System.Drawing.Point(674, 404);
            this.SearchDoc.Name = "SearchDoc";
            this.SearchDoc.Size = new System.Drawing.Size(102, 33);
            this.SearchDoc.TabIndex = 393;
            this.SearchDoc.Text = "بحث المستند";
            this.SearchDoc.UseVisualStyleBackColor = true;
            this.SearchDoc.Click += new System.EventHandler(this.SearchDoc_Click);
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.label24.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label24.Location = new System.Drawing.Point(691, 50);
            this.label24.Name = "label24";
            this.label24.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label24.Size = new System.Drawing.Size(44, 27);
            this.label24.TabIndex = 406;
            this.label24.Text = "تعليق:";
            // 
            // Comment
            // 
            this.Comment.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.Comment.Location = new System.Drawing.Point(410, 54);
            this.Comment.Multiline = true;
            this.Comment.Name = "Comment";
            this.Comment.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Comment.Size = new System.Drawing.Size(270, 119);
            this.Comment.TabIndex = 405;
            this.Comment.Text = "لا تعليق";
            // 
            // Form6
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1244, 661);
            this.Controls.Add(this.label24);
            this.Controls.Add(this.Comment);
            this.Controls.Add(this.btnprintOnly);
            this.Controls.Add(this.SaveOnly);
            this.Controls.Add(this.btnSavePrint);
            this.Controls.Add(this.ResetAll);
            this.Controls.Add(this.SearchFile);
            this.Controls.Add(this.labelArch);
            this.Controls.Add(this.ArchivedSt);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.ListSearch);
            this.Controls.Add(this.SearchDoc);
            this.Controls.Add(this.ConsulateEmployee);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.checkedViewed);
            this.Controls.Add(this.mandoubLabel);
            this.Controls.Add(this.mandoubName);
            this.Controls.Add(this.AppType);
            this.Controls.Add(this.label21);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Ifadaid);
            this.Controls.Add(this.label19);
            this.Controls.Add(this.HijriDate);
            this.Controls.Add(this.AttendViceConsul);
            this.Controls.Add(this.ApplicantSex);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.GregorianDate);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.IqamaIssuedSource);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.PassIssuedSource);
            this.Controls.Add(this.ApplicantIqamaNo);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ApplicantPassNo);
            this.Controls.Add(this.labeldoctype);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ApplicantIdocName);
            this.Controls.Add(this.labelName);
            this.Name = "Form6";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Text = "إفادة للإدلة الجنائية";
            this.Load += new System.EventHandler(this.Form6_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Timer timer2;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox Ifadaid;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TextBox HijriDate;
        private System.Windows.Forms.ComboBox AttendViceConsul;
        private System.Windows.Forms.CheckBox ApplicantSex;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox GregorianDate;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox PassIssuedSource;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox ApplicantPassNo;
        private System.Windows.Forms.Label labeldoctype;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox ApplicantIdocName;
        private System.Windows.Forms.Label labelName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox ApplicantIqamaNo;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox IqamaIssuedSource;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.CheckBox checkedViewed;
        private System.Windows.Forms.Label mandoubLabel;
        private System.Windows.Forms.ComboBox mandoubName;
        private System.Windows.Forms.CheckBox AppType;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label ConsulateEmployee;
        private System.Windows.Forms.Button btnprintOnly;
        private System.Windows.Forms.Button SaveOnly;
        private System.Windows.Forms.Button btnSavePrint;
        private System.Windows.Forms.Button ResetAll;
        private System.Windows.Forms.TextBox SearchFile;
        private System.Windows.Forms.Label labelArch;
        private System.Windows.Forms.CheckBox ArchivedSt;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox ListSearch;
        private System.Windows.Forms.Button SearchDoc;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.TextBox Comment;
    }
}