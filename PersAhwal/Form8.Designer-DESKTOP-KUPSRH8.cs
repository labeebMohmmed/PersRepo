
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
            this.components = new System.ComponentModel.Container();
            this.DocType = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.AppDocName = new System.Windows.Forms.TextBox();
            this.labelName = new System.Windows.Forms.Label();
            this.FacultyName = new System.Windows.Forms.TextBox();
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
            this.UniName = new System.Windows.Forms.TextBox();
            this.labelUNI = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.label3 = new System.Windows.Forms.Label();
            this.timer2 = new System.Windows.Forms.Timer(this.components);
            this.StudyYear = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.StudyLevel = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.MatricNom = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.checkedViewed = new System.Windows.Forms.CheckBox();
            this.mandoubLabel = new System.Windows.Forms.Label();
            this.mandoubName = new System.Windows.Forms.ComboBox();
            this.AppType = new System.Windows.Forms.CheckBox();
            this.label21 = new System.Windows.Forms.Label();
            this.ConsulateEmployee = new System.Windows.Forms.Label();
            this.SearchFile = new System.Windows.Forms.TextBox();
            this.labelArch = new System.Windows.Forms.Label();
            this.ArchivedSt = new System.Windows.Forms.CheckBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.ListSearch = new System.Windows.Forms.TextBox();
            this.SearchDoc = new System.Windows.Forms.Button();
            this.btnprintOnly = new System.Windows.Forms.Button();
            this.SaveOnly = new System.Windows.Forms.Button();
            this.btnSavePrint = new System.Windows.Forms.Button();
            this.ResetAll = new System.Windows.Forms.Button();
            this.label24 = new System.Windows.Forms.Label();
            this.Comment = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // DocType
            // 
            this.DocType.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DocType.FormattingEnabled = true;
            this.DocType.Items.AddRange(new object[] {
            "جواز سفر",
            "رقم وطني",
            "إقامة"});
            this.DocType.Location = new System.Drawing.Point(750, 71);
            this.DocType.Name = "DocType";
            this.DocType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.DocType.Size = new System.Drawing.Size(263, 35);
            this.DocType.TabIndex = 243;
            this.DocType.Text = "جواز سفر";
            this.DocType.SelectedIndexChanged += new System.EventHandler(this.DocType_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(1030, 79);
            this.label5.Name = "label5";
            this.label5.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label5.Size = new System.Drawing.Size(118, 27);
            this.label5.TabIndex = 242;
            this.label5.Text = "نوع اثبات الشخصية:";
            // 
            // AppDocName
            // 
            this.AppDocName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AppDocName.Location = new System.Drawing.Point(750, 30);
            this.AppDocName.Name = "AppDocName";
            this.AppDocName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AppDocName.Size = new System.Drawing.Size(263, 35);
            this.AppDocName.TabIndex = 240;
            this.AppDocName.TextChanged += new System.EventHandler(this.AppDocName_TextChanged);
            // 
            // labelName
            // 
            this.labelName.AutoSize = true;
            this.labelName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelName.Location = new System.Drawing.Point(1030, 33);
            this.labelName.Name = "labelName";
            this.labelName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelName.Size = new System.Drawing.Size(103, 27);
            this.labelName.TabIndex = 241;
            this.labelName.Text = "اسم مقدم الطلب:";
            // 
            // FacultyName
            // 
            this.FacultyName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FacultyName.Location = new System.Drawing.Point(363, 76);
            this.FacultyName.Name = "FacultyName";
            this.FacultyName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.FacultyName.Size = new System.Drawing.Size(256, 35);
            this.FacultyName.TabIndex = 238;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(636, 76);
            this.label4.Name = "label4";
            this.label4.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label4.Size = new System.Drawing.Size(68, 27);
            this.label4.TabIndex = 239;
            this.label4.Text = "اسم الكلية:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(630, 243);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(102, 27);
            this.label1.TabIndex = 237;
            this.label1.Text = "اسم موقع الاقرار:";
            // 
            // Iqrarid
            // 
            this.Iqrarid.Enabled = false;
            this.Iqrarid.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Iqrarid.Location = new System.Drawing.Point(12, 35);
            this.Iqrarid.Name = "Iqrarid";
            this.Iqrarid.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Iqrarid.Size = new System.Drawing.Size(178, 35);
            this.Iqrarid.TabIndex = 235;
            this.Iqrarid.Text = "ق س ج/160/xyz";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label19.Location = new System.Drawing.Point(196, 38);
            this.label19.Name = "label19";
            this.label19.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label19.Size = new System.Drawing.Size(70, 27);
            this.label19.TabIndex = 234;
            this.label19.Text = " رقم الإقرار:";
            // 
            // HijriDate
            // 
            this.HijriDate.Enabled = false;
            this.HijriDate.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.HijriDate.Location = new System.Drawing.Point(12, 76);
            this.HijriDate.Name = "HijriDate";
            this.HijriDate.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.HijriDate.Size = new System.Drawing.Size(178, 35);
            this.HijriDate.TabIndex = 233;
            // 
            // AttendViceConsul
            // 
            this.AttendViceConsul.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AttendViceConsul.FormattingEnabled = true;
            this.AttendViceConsul.Items.AddRange(new object[] {
            "محمد عثمان عكاشة الحسين",
            "يوسف صديق أبوعاقلة",
            "لبيب محمد أحمد"});
            this.AttendViceConsul.Location = new System.Drawing.Point(363, 240);
            this.AttendViceConsul.Name = "AttendViceConsul";
            this.AttendViceConsul.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AttendViceConsul.Size = new System.Drawing.Size(256, 35);
            this.AttendViceConsul.TabIndex = 232;
            // 
            // ApplicantSex
            // 
            this.ApplicantSex.AutoSize = true;
            this.ApplicantSex.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ApplicantSex.Location = new System.Drawing.Point(964, 116);
            this.ApplicantSex.Name = "ApplicantSex";
            this.ApplicantSex.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ApplicantSex.Size = new System.Drawing.Size(49, 31);
            this.ApplicantSex.TabIndex = 231;
            this.ApplicantSex.Text = "ذكر";
            this.ApplicantSex.UseVisualStyleBackColor = true;
            this.ApplicantSex.CheckedChanged += new System.EventHandler(this.ApplicantSex_CheckedChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(196, 76);
            this.label11.Name = "label11";
            this.label11.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label11.Size = new System.Drawing.Size(90, 27);
            this.label11.TabIndex = 230;
            this.label11.Text = "التاريخ الهجري:";
            // 
            // GregorianDate
            // 
            this.GregorianDate.Enabled = false;
            this.GregorianDate.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GregorianDate.Location = new System.Drawing.Point(12, 115);
            this.GregorianDate.Name = "GregorianDate";
            this.GregorianDate.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.GregorianDate.Size = new System.Drawing.Size(178, 35);
            this.GregorianDate.TabIndex = 228;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(196, 115);
            this.label12.Name = "label12";
            this.label12.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label12.Size = new System.Drawing.Size(94, 27);
            this.label12.TabIndex = 229;
            this.label12.Text = "التاريخ الميلادي:";
            // 
            // IssuedSource
            // 
            this.IssuedSource.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.IssuedSource.Location = new System.Drawing.Point(750, 195);
            this.IssuedSource.Name = "IssuedSource";
            this.IssuedSource.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.IssuedSource.Size = new System.Drawing.Size(263, 35);
            this.IssuedSource.TabIndex = 226;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(1030, 195);
            this.label7.Name = "label7";
            this.label7.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label7.Size = new System.Drawing.Size(87, 27);
            this.label7.TabIndex = 227;
            this.label7.Text = "مكان الإصدار:";
            // 
            // AppDocNo
            // 
            this.AppDocNo.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AppDocNo.Location = new System.Drawing.Point(750, 155);
            this.AppDocNo.Name = "AppDocNo";
            this.AppDocNo.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.AppDocNo.Size = new System.Drawing.Size(263, 35);
            this.AppDocNo.TabIndex = 224;
            this.AppDocNo.Text = "P";
            // 
            // labeldoctype
            // 
            this.labeldoctype.AutoSize = true;
            this.labeldoctype.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labeldoctype.Location = new System.Drawing.Point(1030, 155);
            this.labeldoctype.Name = "labeldoctype";
            this.labeldoctype.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labeldoctype.Size = new System.Drawing.Size(105, 27);
            this.labeldoctype.TabIndex = 225;
            this.labeldoctype.Text = "رقم الوثيقة المقدمة:";
            // 
            // UniName
            // 
            this.UniName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UniName.Location = new System.Drawing.Point(363, 35);
            this.UniName.Name = "UniName";
            this.UniName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.UniName.Size = new System.Drawing.Size(256, 35);
            this.UniName.TabIndex = 221;
            // 
            // labelUNI
            // 
            this.labelUNI.AutoSize = true;
            this.labelUNI.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelUNI.Location = new System.Drawing.Point(636, 41);
            this.labelUNI.Name = "labelUNI";
            this.labelUNI.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelUNI.Size = new System.Drawing.Size(74, 27);
            this.labelUNI.TabIndex = 222;
            this.labelUNI.Text = "اسم الجامعة:";
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(1030, 116);
            this.label3.Name = "label3";
            this.label3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label3.Size = new System.Drawing.Size(40, 27);
            this.label3.TabIndex = 223;
            this.label3.Text = "النوع:";
            // 
            // timer2
            // 
            this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
            // 
            // StudyYear
            // 
            this.StudyYear.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.StudyYear.Location = new System.Drawing.Point(363, 158);
            this.StudyYear.Name = "StudyYear";
            this.StudyYear.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.StudyYear.Size = new System.Drawing.Size(256, 35);
            this.StudyYear.TabIndex = 244;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(636, 158);
            this.label6.Name = "label6";
            this.label6.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label6.Size = new System.Drawing.Size(81, 27);
            this.label6.TabIndex = 245;
            this.label6.Text = "العام الدراسي:";
            // 
            // StudyLevel
            // 
            this.StudyLevel.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.StudyLevel.Location = new System.Drawing.Point(363, 117);
            this.StudyLevel.Name = "StudyLevel";
            this.StudyLevel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.StudyLevel.Size = new System.Drawing.Size(256, 35);
            this.StudyLevel.TabIndex = 246;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(636, 117);
            this.label8.Name = "label8";
            this.label8.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label8.Size = new System.Drawing.Size(103, 27);
            this.label8.TabIndex = 247;
            this.label8.Text = "المستوى الدراسي:";
            // 
            // MatricNom
            // 
            this.MatricNom.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MatricNom.Location = new System.Drawing.Point(363, 199);
            this.MatricNom.Name = "MatricNom";
            this.MatricNom.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.MatricNom.Size = new System.Drawing.Size(256, 35);
            this.MatricNom.TabIndex = 248;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(636, 199);
            this.label9.Name = "label9";
            this.label9.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label9.Size = new System.Drawing.Size(78, 27);
            this.label9.TabIndex = 249;
            this.label9.Text = "الرقم الجامعي:";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 437);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1311, 273);
            this.dataGridView1.TabIndex = 259;
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            // 
            // checkedViewed
            // 
            this.checkedViewed.AutoSize = true;
            this.checkedViewed.Checked = true;
            this.checkedViewed.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkedViewed.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F);
            this.checkedViewed.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.checkedViewed.Location = new System.Drawing.Point(575, 345);
            this.checkedViewed.Name = "checkedViewed";
            this.checkedViewed.Size = new System.Drawing.Size(151, 33);
            this.checkedViewed.TabIndex = 258;
            this.checkedViewed.Text = "NotViewed";
            this.checkedViewed.UseVisualStyleBackColor = true;
            this.checkedViewed.Visible = false;
            // 
            // mandoubLabel
            // 
            this.mandoubLabel.AutoSize = true;
            this.mandoubLabel.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.mandoubLabel.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.mandoubLabel.Location = new System.Drawing.Point(1024, 269);
            this.mandoubLabel.Name = "mandoubLabel";
            this.mandoubLabel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.mandoubLabel.Size = new System.Drawing.Size(79, 27);
            this.mandoubLabel.TabIndex = 257;
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
            this.mandoubName.Location = new System.Drawing.Point(750, 267);
            this.mandoubName.Name = "mandoubName";
            this.mandoubName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.mandoubName.Size = new System.Drawing.Size(263, 35);
            this.mandoubName.TabIndex = 256;
            this.mandoubName.Visible = false;
            // 
            // AppType
            // 
            this.AppType.AutoSize = true;
            this.AppType.Checked = true;
            this.AppType.CheckState = System.Windows.Forms.CheckState.Checked;
            this.AppType.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.AppType.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.AppType.Location = new System.Drawing.Point(819, 234);
            this.AppType.Name = "AppType";
            this.AppType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AppType.Size = new System.Drawing.Size(170, 31);
            this.AppType.TabIndex = 255;
            this.AppType.Text = "حضور مباشرة إلى القنصلية";
            this.AppType.UseVisualStyleBackColor = true;
            this.AppType.CheckedChanged += new System.EventHandler(this.AppType_CheckedChanged);
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.label21.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label21.Location = new System.Drawing.Point(1024, 238);
            this.label21.Name = "label21";
            this.label21.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label21.Size = new System.Drawing.Size(115, 27);
            this.label21.TabIndex = 254;
            this.label21.Text = "طريقة تقديم الطلب:";
            // 
            // ConsulateEmployee
            // 
            this.ConsulateEmployee.AutoSize = true;
            this.ConsulateEmployee.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.ConsulateEmployee.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.ConsulateEmployee.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ConsulateEmployee.Location = new System.Drawing.Point(11, 5);
            this.ConsulateEmployee.Name = "ConsulateEmployee";
            this.ConsulateEmployee.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ConsulateEmployee.Size = new System.Drawing.Size(51, 27);
            this.ConsulateEmployee.TabIndex = 263;
            this.ConsulateEmployee.Text = "الموظف";
            // 
            // SearchFile
            // 
            this.SearchFile.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.SearchFile.Location = new System.Drawing.Point(872, 396);
            this.SearchFile.Name = "SearchFile";
            this.SearchFile.Size = new System.Drawing.Size(319, 35);
            this.SearchFile.TabIndex = 420;
            this.SearchFile.Visible = false;
            // 
            // labelArch
            // 
            this.labelArch.AutoSize = true;
            this.labelArch.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.labelArch.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.labelArch.Location = new System.Drawing.Point(434, 402);
            this.labelArch.Name = "labelArch";
            this.labelArch.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelArch.Size = new System.Drawing.Size(79, 27);
            this.labelArch.TabIndex = 419;
            this.labelArch.Text = "حالة الأرشفة:";
            // 
            // ArchivedSt
            // 
            this.ArchivedSt.AutoSize = true;
            this.ArchivedSt.BackColor = System.Drawing.Color.Red;
            this.ArchivedSt.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.ArchivedSt.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ArchivedSt.Location = new System.Drawing.Point(325, 398);
            this.ArchivedSt.Name = "ArchivedSt";
            this.ArchivedSt.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ArchivedSt.Size = new System.Drawing.Size(98, 31);
            this.ArchivedSt.TabIndex = 418;
            this.ArchivedSt.Text = "غير مؤرشف";
            this.ArchivedSt.UseVisualStyleBackColor = false;
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button4.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button4.Location = new System.Drawing.Point(526, 398);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(114, 33);
            this.button4.TabIndex = 417;
            this.button4.Text = "فتح المستند رقم 2";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button3.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button3.Location = new System.Drawing.Point(1193, 396);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(129, 33);
            this.button3.TabIndex = 416;
            this.button3.Text = "بحث القائمة أدناه";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button2.Location = new System.Drawing.Point(646, 398);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(112, 33);
            this.button2.TabIndex = 415;
            this.button2.Text = "فتح المستند رقم 1";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // ListSearch
            // 
            this.ListSearch.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.ListSearch.Location = new System.Drawing.Point(872, 396);
            this.ListSearch.Name = "ListSearch";
            this.ListSearch.Size = new System.Drawing.Size(319, 35);
            this.ListSearch.TabIndex = 414;
            // 
            // SearchDoc
            // 
            this.SearchDoc.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.SearchDoc.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.SearchDoc.Location = new System.Drawing.Point(764, 398);
            this.SearchDoc.Name = "SearchDoc";
            this.SearchDoc.Size = new System.Drawing.Size(102, 33);
            this.SearchDoc.TabIndex = 413;
            this.SearchDoc.Text = "بحث المستند";
            this.SearchDoc.UseVisualStyleBackColor = true;
            this.SearchDoc.Click += new System.EventHandler(this.SearchDoc_Click);
            // 
            // btnprintOnly
            // 
            this.btnprintOnly.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnprintOnly.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnprintOnly.Location = new System.Drawing.Point(62, 360);
            this.btnprintOnly.Name = "btnprintOnly";
            this.btnprintOnly.Size = new System.Drawing.Size(43, 71);
            this.btnprintOnly.TabIndex = 424;
            this.btnprintOnly.Text = "طباعة";
            this.btnprintOnly.UseVisualStyleBackColor = false;
            this.btnprintOnly.Visible = false;
            this.btnprintOnly.Click += new System.EventHandler(this.printOnly_Click);
            // 
            // SaveOnly
            // 
            this.SaveOnly.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.SaveOnly.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.SaveOnly.Location = new System.Drawing.Point(12, 360);
            this.SaveOnly.Name = "SaveOnly";
            this.SaveOnly.Size = new System.Drawing.Size(45, 71);
            this.SaveOnly.TabIndex = 423;
            this.SaveOnly.Text = "حفظ";
            this.SaveOnly.UseVisualStyleBackColor = false;
            this.SaveOnly.Visible = false;
            this.SaveOnly.Click += new System.EventHandler(this.SaveOnly_Click);
            // 
            // btnSavePrint
            // 
            this.btnSavePrint.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnSavePrint.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnSavePrint.Location = new System.Drawing.Point(11, 360);
            this.btnSavePrint.Name = "btnSavePrint";
            this.btnSavePrint.Size = new System.Drawing.Size(94, 71);
            this.btnSavePrint.TabIndex = 422;
            this.btnSavePrint.Text = "طباعة وحفظ";
            this.btnSavePrint.UseVisualStyleBackColor = false;
            this.btnSavePrint.Click += new System.EventHandler(this.btnSavePrint_Click);
            // 
            // ResetAll
            // 
            this.ResetAll.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.ResetAll.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ResetAll.Location = new System.Drawing.Point(111, 360);
            this.ResetAll.Name = "ResetAll";
            this.ResetAll.Size = new System.Drawing.Size(92, 71);
            this.ResetAll.TabIndex = 421;
            this.ResetAll.Text = "مسح جميع الحقول";
            this.ResetAll.UseVisualStyleBackColor = false;
            this.ResetAll.Click += new System.EventHandler(this.ResetAll_Click);
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.label24.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label24.Location = new System.Drawing.Point(196, 161);
            this.label24.Name = "label24";
            this.label24.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label24.Size = new System.Drawing.Size(44, 27);
            this.label24.TabIndex = 426;
            this.label24.Text = "تعليق:";
            // 
            // Comment
            // 
            this.Comment.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.Comment.Location = new System.Drawing.Point(11, 158);
            this.Comment.Multiline = true;
            this.Comment.Name = "Comment";
            this.Comment.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Comment.Size = new System.Drawing.Size(179, 144);
            this.Comment.TabIndex = 425;
            this.Comment.Text = "لا تعليق";
            // 
            // Form8
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1335, 748);
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
            this.Controls.Add(this.MatricNom);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.StudyLevel);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.StudyYear);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.DocType);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.AppDocName);
            this.Controls.Add(this.labelName);
            this.Controls.Add(this.FacultyName);
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
            this.Controls.Add(this.UniName);
            this.Controls.Add(this.labelUNI);
            this.Controls.Add(this.label3);
            this.Name = "Form8";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Text = "الإرتباط ببرنامج دراسي";
            this.Load += new System.EventHandler(this.Form8_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox DocType;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox AppDocName;
        private System.Windows.Forms.Label labelName;
        private System.Windows.Forms.TextBox FacultyName;
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
        private System.Windows.Forms.TextBox UniName;
        private System.Windows.Forms.Label labelUNI;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Timer timer2;
        private System.Windows.Forms.TextBox StudyYear;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox StudyLevel;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox MatricNom;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.CheckBox checkedViewed;
        private System.Windows.Forms.Label mandoubLabel;
        private System.Windows.Forms.ComboBox mandoubName;
        private System.Windows.Forms.CheckBox AppType;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label ConsulateEmployee;
        private System.Windows.Forms.TextBox SearchFile;
        private System.Windows.Forms.Label labelArch;
        private System.Windows.Forms.CheckBox ArchivedSt;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox ListSearch;
        private System.Windows.Forms.Button SearchDoc;
        private System.Windows.Forms.Button btnprintOnly;
        private System.Windows.Forms.Button SaveOnly;
        private System.Windows.Forms.Button btnSavePrint;
        private System.Windows.Forms.Button ResetAll;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.TextBox Comment;
    }
}