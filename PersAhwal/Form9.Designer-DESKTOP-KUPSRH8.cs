
namespace PersAhwal
{
    partial class Form9
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
            this.DocType = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.AppDocName = new System.Windows.Forms.TextBox();
            this.labelName = new System.Windows.Forms.Label();
            this.AppDocNatio = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.Iqrarid = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.HijriDate = new System.Windows.Forms.TextBox();
            this.AttendViceConsul = new System.Windows.Forms.ComboBox();
            this.ApplicantSex = new System.Windows.Forms.CheckBox();
            this.label11 = new System.Windows.Forms.Label();
            this.GregorianDate = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.IssuedSource = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.AppDocNo = new System.Windows.Forms.TextBox();
            this.labeldoctype = new System.Windows.Forms.Label();
            this.OtherDocName = new System.Windows.Forms.TextBox();
            this.labelOtherName = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.OtherDocType = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.OtherIssuedSource = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.OtherDocNo = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.checkedViewed = new System.Windows.Forms.CheckBox();
            this.mandoubLabel = new System.Windows.Forms.Label();
            this.mandoubName = new System.Windows.Forms.ComboBox();
            this.AppType = new System.Windows.Forms.CheckBox();
            this.label21 = new System.Windows.Forms.Label();
            this.ConsulateEmployee = new System.Windows.Forms.Label();
            this.label24 = new System.Windows.Forms.Label();
            this.Comment = new System.Windows.Forms.TextBox();
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
            this.checkBox1 = new System.Windows.Forms.CheckBox();
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
            // DocType
            // 
            this.DocType.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DocType.FormattingEnabled = true;
            this.DocType.Items.AddRange(new object[] {
            "جواز سفر",
            "رقم وطني",
            "إقامة"});
            this.DocType.Location = new System.Drawing.Point(805, 62);
            this.DocType.Name = "DocType";
            this.DocType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.DocType.Size = new System.Drawing.Size(251, 35);
            this.DocType.TabIndex = 272;
            this.DocType.Text = "جواز سفر";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(1073, 70);
            this.label5.Name = "label5";
            this.label5.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label5.Size = new System.Drawing.Size(118, 27);
            this.label5.TabIndex = 271;
            this.label5.Text = "نوع اثبات الشخصية:";
            // 
            // AppDocName
            // 
            this.AppDocName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AppDocName.Location = new System.Drawing.Point(805, 21);
            this.AppDocName.Name = "AppDocName";
            this.AppDocName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AppDocName.Size = new System.Drawing.Size(251, 35);
            this.AppDocName.TabIndex = 269;
            // 
            // labelName
            // 
            this.labelName.AutoSize = true;
            this.labelName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelName.Location = new System.Drawing.Point(1073, 24);
            this.labelName.Name = "labelName";
            this.labelName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelName.Size = new System.Drawing.Size(103, 27);
            this.labelName.TabIndex = 270;
            this.labelName.Text = "اسم مقدم الطلب:";
            // 
            // AppDocNatio
            // 
            this.AppDocNatio.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AppDocNatio.Location = new System.Drawing.Point(410, 60);
            this.AppDocNatio.Name = "AppDocNatio";
            this.AppDocNatio.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AppDocNatio.Size = new System.Drawing.Size(242, 35);
            this.AppDocNatio.TabIndex = 267;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(669, 60);
            this.label4.Name = "label4";
            this.label4.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label4.Size = new System.Drawing.Size(56, 27);
            this.label4.TabIndex = 268;
            this.label4.Text = "الجنسية:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(663, 228);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(102, 27);
            this.label1.TabIndex = 266;
            this.label1.Text = "اسم موقع الاقرار:";
            // 
            // Iqrarid
            // 
            this.Iqrarid.Enabled = false;
            this.Iqrarid.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Iqrarid.Location = new System.Drawing.Point(43, 24);
            this.Iqrarid.Name = "Iqrarid";
            this.Iqrarid.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Iqrarid.Size = new System.Drawing.Size(178, 35);
            this.Iqrarid.TabIndex = 264;
            this.Iqrarid.Text = "ق س ج/160/xyz";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label19.Location = new System.Drawing.Point(230, 24);
            this.label19.Name = "label19";
            this.label19.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label19.Size = new System.Drawing.Size(70, 27);
            this.label19.TabIndex = 263;
            this.label19.Text = " رقم الإقرار:";
            // 
            // HijriDate
            // 
            this.HijriDate.Enabled = false;
            this.HijriDate.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.HijriDate.Location = new System.Drawing.Point(43, 65);
            this.HijriDate.Name = "HijriDate";
            this.HijriDate.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.HijriDate.Size = new System.Drawing.Size(178, 35);
            this.HijriDate.TabIndex = 262;
            // 
            // AttendViceConsul
            // 
            this.AttendViceConsul.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AttendViceConsul.FormattingEnabled = true;
            this.AttendViceConsul.Items.AddRange(new object[] {
            "محمد عثمان عكاشة الحسين",
            "يوسف صديق أبوعاقلة",
            "لبيب محمد أحمد"});
            this.AttendViceConsul.Location = new System.Drawing.Point(410, 225);
            this.AttendViceConsul.Name = "AttendViceConsul";
            this.AttendViceConsul.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AttendViceConsul.Size = new System.Drawing.Size(242, 35);
            this.AttendViceConsul.TabIndex = 261;
            // 
            // ApplicantSex
            // 
            this.ApplicantSex.AutoSize = true;
            this.ApplicantSex.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ApplicantSex.Location = new System.Drawing.Point(1007, 107);
            this.ApplicantSex.Name = "ApplicantSex";
            this.ApplicantSex.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ApplicantSex.Size = new System.Drawing.Size(49, 31);
            this.ApplicantSex.TabIndex = 260;
            this.ApplicantSex.Text = "ذكر";
            this.ApplicantSex.UseVisualStyleBackColor = true;
            this.ApplicantSex.CheckedChanged += new System.EventHandler(this.ApplicantSex_CheckedChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(230, 65);
            this.label11.Name = "label11";
            this.label11.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label11.Size = new System.Drawing.Size(90, 27);
            this.label11.TabIndex = 259;
            this.label11.Text = "التاريخ الهجري:";
            // 
            // GregorianDate
            // 
            this.GregorianDate.Enabled = false;
            this.GregorianDate.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GregorianDate.Location = new System.Drawing.Point(43, 106);
            this.GregorianDate.Name = "GregorianDate";
            this.GregorianDate.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.GregorianDate.Size = new System.Drawing.Size(178, 35);
            this.GregorianDate.TabIndex = 257;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(230, 106);
            this.label12.Name = "label12";
            this.label12.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label12.Size = new System.Drawing.Size(94, 27);
            this.label12.TabIndex = 258;
            this.label12.Text = "التاريخ الميلادي:";
            // 
            // IssuedSource
            // 
            this.IssuedSource.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.IssuedSource.Location = new System.Drawing.Point(805, 186);
            this.IssuedSource.Name = "IssuedSource";
            this.IssuedSource.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.IssuedSource.Size = new System.Drawing.Size(251, 35);
            this.IssuedSource.TabIndex = 255;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(1073, 186);
            this.label7.Name = "label7";
            this.label7.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label7.Size = new System.Drawing.Size(87, 27);
            this.label7.TabIndex = 256;
            this.label7.Text = "مكان الإصدار:";
            // 
            // AppDocNo
            // 
            this.AppDocNo.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AppDocNo.Location = new System.Drawing.Point(805, 146);
            this.AppDocNo.Name = "AppDocNo";
            this.AppDocNo.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AppDocNo.Size = new System.Drawing.Size(251, 35);
            this.AppDocNo.TabIndex = 253;
            // 
            // labeldoctype
            // 
            this.labeldoctype.AutoSize = true;
            this.labeldoctype.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labeldoctype.Location = new System.Drawing.Point(1073, 146);
            this.labeldoctype.Name = "labeldoctype";
            this.labeldoctype.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labeldoctype.Size = new System.Drawing.Size(105, 27);
            this.labeldoctype.TabIndex = 254;
            this.labeldoctype.Text = "رقم الوثيقة المقدمة:";
            // 
            // OtherDocName
            // 
            this.OtherDocName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OtherDocName.Location = new System.Drawing.Point(410, 19);
            this.OtherDocName.Name = "OtherDocName";
            this.OtherDocName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.OtherDocName.Size = new System.Drawing.Size(242, 35);
            this.OtherDocName.TabIndex = 250;
            // 
            // labelOtherName
            // 
            this.labelOtherName.AutoSize = true;
            this.labelOtherName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelOtherName.Location = new System.Drawing.Point(669, 19);
            this.labelOtherName.Name = "labelOtherName";
            this.labelOtherName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelOtherName.Size = new System.Drawing.Size(123, 27);
            this.labelOtherName.TabIndex = 251;
            this.labelOtherName.Text = "اسم المراد الزواج منها:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(1073, 107);
            this.label3.Name = "label3";
            this.label3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label3.Size = new System.Drawing.Size(40, 27);
            this.label3.TabIndex = 252;
            this.label3.Text = "النوع:";
            // 
            // OtherDocType
            // 
            this.OtherDocType.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OtherDocType.FormattingEnabled = true;
            this.OtherDocType.Items.AddRange(new object[] {
            "جواز سفر",
            "رقم وطني",
            "إقامة"});
            this.OtherDocType.Location = new System.Drawing.Point(410, 102);
            this.OtherDocType.Name = "OtherDocType";
            this.OtherDocType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.OtherDocType.Size = new System.Drawing.Size(242, 35);
            this.OtherDocType.TabIndex = 280;
            this.OtherDocType.Text = "جواز سفر";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(669, 102);
            this.label2.Name = "label2";
            this.label2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label2.Size = new System.Drawing.Size(118, 27);
            this.label2.TabIndex = 279;
            this.label2.Text = "نوع اثبات الشخصية:";
            // 
            // OtherIssuedSource
            // 
            this.OtherIssuedSource.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OtherIssuedSource.Location = new System.Drawing.Point(410, 184);
            this.OtherIssuedSource.Name = "OtherIssuedSource";
            this.OtherIssuedSource.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.OtherIssuedSource.Size = new System.Drawing.Size(242, 35);
            this.OtherIssuedSource.TabIndex = 276;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(669, 184);
            this.label6.Name = "label6";
            this.label6.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label6.Size = new System.Drawing.Size(87, 27);
            this.label6.TabIndex = 277;
            this.label6.Text = "مكان الإصدار:";
            // 
            // OtherDocNo
            // 
            this.OtherDocNo.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OtherDocNo.Location = new System.Drawing.Point(410, 144);
            this.OtherDocNo.Name = "OtherDocNo";
            this.OtherDocNo.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.OtherDocNo.Size = new System.Drawing.Size(242, 35);
            this.OtherDocNo.TabIndex = 274;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(669, 144);
            this.label8.Name = "label8";
            this.label8.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label8.Size = new System.Drawing.Size(105, 27);
            this.label8.TabIndex = 275;
            this.label8.Text = "رقم الوثيقة المقدمة:";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(25, 462);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1315, 201);
            this.dataGridView1.TabIndex = 286;
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            // 
            // checkedViewed
            // 
            this.checkedViewed.AutoSize = true;
            this.checkedViewed.Checked = true;
            this.checkedViewed.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkedViewed.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F);
            this.checkedViewed.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.checkedViewed.Location = new System.Drawing.Point(230, 228);
            this.checkedViewed.Name = "checkedViewed";
            this.checkedViewed.Size = new System.Drawing.Size(151, 33);
            this.checkedViewed.TabIndex = 285;
            this.checkedViewed.Text = "NotViewed";
            this.checkedViewed.UseVisualStyleBackColor = true;
            this.checkedViewed.Visible = false;
            // 
            // mandoubLabel
            // 
            this.mandoubLabel.AutoSize = true;
            this.mandoubLabel.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.mandoubLabel.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.mandoubLabel.Location = new System.Drawing.Point(1073, 258);
            this.mandoubLabel.Name = "mandoubLabel";
            this.mandoubLabel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.mandoubLabel.Size = new System.Drawing.Size(79, 27);
            this.mandoubLabel.TabIndex = 284;
            this.mandoubLabel.Text = "اسم المندوب:";
            this.mandoubLabel.Visible = false;
            // 
            // mandoubName
            // 
            this.mandoubName.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.mandoubName.FormattingEnabled = true;
            this.mandoubName.Items.AddRange(new object[] {
            "محمد عوض قاسم الشيخ",
            "محمود أحمد حامد النور"});
            this.mandoubName.Location = new System.Drawing.Point(805, 258);
            this.mandoubName.Name = "mandoubName";
            this.mandoubName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.mandoubName.Size = new System.Drawing.Size(251, 35);
            this.mandoubName.TabIndex = 283;
            this.mandoubName.Visible = false;
            // 
            // AppType
            // 
            this.AppType.AutoSize = true;
            this.AppType.Checked = true;
            this.AppType.CheckState = System.Windows.Forms.CheckState.Checked;
            this.AppType.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.AppType.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.AppType.Location = new System.Drawing.Point(862, 225);
            this.AppType.Name = "AppType";
            this.AppType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AppType.Size = new System.Drawing.Size(170, 31);
            this.AppType.TabIndex = 282;
            this.AppType.Text = "حضور مباشرة إلى القنصلية";
            this.AppType.UseVisualStyleBackColor = true;
            this.AppType.CheckedChanged += new System.EventHandler(this.AppType_CheckedChanged);
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.label21.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label21.Location = new System.Drawing.Point(1072, 225);
            this.label21.Name = "label21";
            this.label21.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label21.Size = new System.Drawing.Size(115, 27);
            this.label21.TabIndex = 281;
            this.label21.Text = "طريقة تقديم الطلب:";
            // 
            // ConsulateEmployee
            // 
            this.ConsulateEmployee.AutoSize = true;
            this.ConsulateEmployee.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.ConsulateEmployee.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.ConsulateEmployee.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ConsulateEmployee.Location = new System.Drawing.Point(25, 667);
            this.ConsulateEmployee.Name = "ConsulateEmployee";
            this.ConsulateEmployee.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ConsulateEmployee.Size = new System.Drawing.Size(51, 27);
            this.ConsulateEmployee.TabIndex = 290;
            this.ConsulateEmployee.Text = "الموظف";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.label24.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label24.Location = new System.Drawing.Point(228, 150);
            this.label24.Name = "label24";
            this.label24.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label24.Size = new System.Drawing.Size(44, 27);
            this.label24.TabIndex = 428;
            this.label24.Text = "تعليق:";
            // 
            // Comment
            // 
            this.Comment.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.Comment.Location = new System.Drawing.Point(43, 147);
            this.Comment.Multiline = true;
            this.Comment.Name = "Comment";
            this.Comment.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Comment.Size = new System.Drawing.Size(179, 144);
            this.Comment.TabIndex = 427;
            this.Comment.Text = "لا تعليق";
            // 
            // btnprintOnly
            // 
            this.btnprintOnly.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnprintOnly.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnprintOnly.Location = new System.Drawing.Point(80, 385);
            this.btnprintOnly.Name = "btnprintOnly";
            this.btnprintOnly.Size = new System.Drawing.Size(43, 71);
            this.btnprintOnly.TabIndex = 441;
            this.btnprintOnly.Text = "طباعة";
            this.btnprintOnly.UseVisualStyleBackColor = false;
            this.btnprintOnly.Visible = false;
            this.btnprintOnly.Click += new System.EventHandler(this.printOnly_Click);
            // 
            // SaveOnly
            // 
            this.SaveOnly.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.SaveOnly.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.SaveOnly.Location = new System.Drawing.Point(30, 385);
            this.SaveOnly.Name = "SaveOnly";
            this.SaveOnly.Size = new System.Drawing.Size(45, 71);
            this.SaveOnly.TabIndex = 440;
            this.SaveOnly.Text = "حفظ";
            this.SaveOnly.UseVisualStyleBackColor = false;
            this.SaveOnly.Visible = false;
            this.SaveOnly.Click += new System.EventHandler(this.SaveOnly_Click);
            // 
            // btnSavePrint
            // 
            this.btnSavePrint.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnSavePrint.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnSavePrint.Location = new System.Drawing.Point(29, 385);
            this.btnSavePrint.Name = "btnSavePrint";
            this.btnSavePrint.Size = new System.Drawing.Size(94, 71);
            this.btnSavePrint.TabIndex = 439;
            this.btnSavePrint.Text = "طباعة وحفظ";
            this.btnSavePrint.UseVisualStyleBackColor = false;
            this.btnSavePrint.Click += new System.EventHandler(this.btnSavePrint_Click);
            // 
            // ResetAll
            // 
            this.ResetAll.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.ResetAll.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ResetAll.Location = new System.Drawing.Point(129, 385);
            this.ResetAll.Name = "ResetAll";
            this.ResetAll.Size = new System.Drawing.Size(92, 71);
            this.ResetAll.TabIndex = 438;
            this.ResetAll.Text = "مسح جميع الحقول";
            this.ResetAll.UseVisualStyleBackColor = false;
            this.ResetAll.Click += new System.EventHandler(this.ResetAll_Click);
            // 
            // SearchFile
            // 
            this.SearchFile.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.SearchFile.Location = new System.Drawing.Point(890, 421);
            this.SearchFile.Name = "SearchFile";
            this.SearchFile.Size = new System.Drawing.Size(319, 35);
            this.SearchFile.TabIndex = 437;
            this.SearchFile.Visible = false;
            // 
            // labelArch
            // 
            this.labelArch.AutoSize = true;
            this.labelArch.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.labelArch.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.labelArch.Location = new System.Drawing.Point(452, 427);
            this.labelArch.Name = "labelArch";
            this.labelArch.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelArch.Size = new System.Drawing.Size(79, 27);
            this.labelArch.TabIndex = 436;
            this.labelArch.Text = "حالة الأرشفة:";
            // 
            // ArchivedSt
            // 
            this.ArchivedSt.AutoSize = true;
            this.ArchivedSt.BackColor = System.Drawing.Color.Red;
            this.ArchivedSt.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.ArchivedSt.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ArchivedSt.Location = new System.Drawing.Point(343, 423);
            this.ArchivedSt.Name = "ArchivedSt";
            this.ArchivedSt.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ArchivedSt.Size = new System.Drawing.Size(98, 31);
            this.ArchivedSt.TabIndex = 435;
            this.ArchivedSt.Text = "غير مؤرشف";
            this.ArchivedSt.UseVisualStyleBackColor = false;
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button4.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button4.Location = new System.Drawing.Point(544, 423);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(114, 33);
            this.button4.TabIndex = 434;
            this.button4.Text = "فتح المستند رقم 2";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button3.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button3.Location = new System.Drawing.Point(1211, 421);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(129, 33);
            this.button3.TabIndex = 433;
            this.button3.Text = "بحث القائمة أدناه";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button2.Location = new System.Drawing.Point(664, 423);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(112, 33);
            this.button2.TabIndex = 432;
            this.button2.Text = "فتح المستند رقم 1";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // ListSearch
            // 
            this.ListSearch.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.ListSearch.Location = new System.Drawing.Point(890, 421);
            this.ListSearch.Name = "ListSearch";
            this.ListSearch.Size = new System.Drawing.Size(319, 35);
            this.ListSearch.TabIndex = 431;
            // 
            // SearchDoc
            // 
            this.SearchDoc.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.SearchDoc.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.SearchDoc.Location = new System.Drawing.Point(782, 423);
            this.SearchDoc.Name = "SearchDoc";
            this.SearchDoc.Size = new System.Drawing.Size(102, 33);
            this.SearchDoc.TabIndex = 430;
            this.SearchDoc.Text = "بحث المستند";
            this.SearchDoc.UseVisualStyleBackColor = true;
            this.SearchDoc.Click += new System.EventHandler(this.SearchDoc_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F);
            this.checkBox1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.checkBox1.Location = new System.Drawing.Point(593, 370);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(151, 33);
            this.checkBox1.TabIndex = 429;
            this.checkBox1.Text = "NotViewed";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.Visible = false;
            // 
            // Form9
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1368, 742);
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
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.label24);
            this.Controls.Add(this.Comment);
            this.Controls.Add(this.ConsulateEmployee);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.checkedViewed);
            this.Controls.Add(this.mandoubLabel);
            this.Controls.Add(this.mandoubName);
            this.Controls.Add(this.AppType);
            this.Controls.Add(this.label21);
            this.Controls.Add(this.OtherDocType);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.OtherIssuedSource);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.OtherDocNo);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.DocType);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.AppDocName);
            this.Controls.Add(this.labelName);
            this.Controls.Add(this.AppDocNatio);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Iqrarid);
            this.Controls.Add(this.label19);
            this.Controls.Add(this.HijriDate);
            this.Controls.Add(this.AttendViceConsul);
            this.Controls.Add(this.ApplicantSex);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.GregorianDate);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.IssuedSource);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.AppDocNo);
            this.Controls.Add(this.labeldoctype);
            this.Controls.Add(this.OtherDocName);
            this.Controls.Add(this.labelOtherName);
            this.Controls.Add(this.label3);
            this.Name = "Form9";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Text = "شهادة عدم ممانعة زواج";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Timer timer2;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.ComboBox DocType;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox AppDocName;
        private System.Windows.Forms.Label labelName;
        private System.Windows.Forms.TextBox AppDocNatio;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox Iqrarid;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TextBox HijriDate;
        private System.Windows.Forms.ComboBox AttendViceConsul;
        private System.Windows.Forms.CheckBox ApplicantSex;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox GregorianDate;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox IssuedSource;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox AppDocNo;
        private System.Windows.Forms.Label labeldoctype;
        private System.Windows.Forms.TextBox OtherDocName;
        private System.Windows.Forms.Label labelOtherName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox OtherDocType;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox OtherIssuedSource;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox OtherDocNo;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.CheckBox checkedViewed;
        private System.Windows.Forms.Label mandoubLabel;
        private System.Windows.Forms.ComboBox mandoubName;
        private System.Windows.Forms.CheckBox AppType;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label ConsulateEmployee;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.TextBox Comment;
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
        private System.Windows.Forms.CheckBox checkBox1;
    }
}