
namespace PersAhwal
{
    partial class Form7
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.IqrarType = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnAddDoc = new System.Windows.Forms.Button();
            this.document = new System.Windows.Forms.TextBox();
            this.deleteRow = new System.Windows.Forms.Button();
            this.label24 = new System.Windows.Forms.Label();
            this.Comment = new System.Windows.Forms.TextBox();
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
            this.PersonDesc = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.ConsulateEmployee = new System.Windows.Forms.Label();
            this.timer2 = new System.Windows.Forms.Timer(this.components);
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.checkedViewed = new System.Windows.Forms.CheckBox();
            this.mandoubLabel = new System.Windows.Forms.Label();
            this.mandoubName = new System.Windows.Forms.ComboBox();
            this.AppType = new System.Windows.Forms.CheckBox();
            this.label21 = new System.Windows.Forms.Label();
            this.نوع_الهوية = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.مقدم_الطلب = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.AppWrongName = new System.Windows.Forms.TextBox();
            this.labelWrongName = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.Iqrarid = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.التاريخ_الهجري = new System.Windows.Forms.TextBox();
            this.AttendViceConsul = new System.Windows.Forms.ComboBox();
            this.النوع = new System.Windows.Forms.CheckBox();
            this.label11 = new System.Windows.Forms.Label();
            this.التاريخ_الميلادي = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.مكان_الإصدار = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.رقم_الهوية = new System.Windows.Forms.TextBox();
            this.labeldoctype = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.AppTrueName = new System.Windows.Forms.TextBox();
            this.labelName = new System.Windows.Forms.Label();
            this.PanelFiles = new System.Windows.Forms.Panel();
            this.PanelMain = new System.Windows.Forms.Panel();
            this.التاريخ_الميلادي_off = new System.Windows.Forms.TextBox();
            this.txtEditID2 = new System.Windows.Forms.TextBox();
            this.btnEditID = new System.Windows.Forms.Button();
            this.txtEditID1 = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.PanelFiles.SuspendLayout();
            this.PanelMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // IqrarType
            // 
            this.IqrarType.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.IqrarType.FormattingEnabled = true;
            this.IqrarType.Items.AddRange(new object[] {
            "اثبات اسمان لذات واحدة",
            "اثبات صحة وثائق"});
            this.IqrarType.Location = new System.Drawing.Point(846, 22);
            this.IqrarType.Name = "IqrarType";
            this.IqrarType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.IqrarType.Size = new System.Drawing.Size(266, 35);
            this.IqrarType.TabIndex = 518;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(1127, -28);
            this.label4.Name = "label4";
            this.label4.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label4.Size = new System.Drawing.Size(76, 27);
            this.label4.TabIndex = 517;
            this.label4.Text = "نوع الاجراء:";
            // 
            // btnAddDoc
            // 
            this.btnAddDoc.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.btnAddDoc.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnAddDoc.Location = new System.Drawing.Point(649, 145);
            this.btnAddDoc.Name = "btnAddDoc";
            this.btnAddDoc.Size = new System.Drawing.Size(65, 33);
            this.btnAddDoc.TabIndex = 516;
            this.btnAddDoc.Text = "إضافة";
            this.btnAddDoc.UseVisualStyleBackColor = true;
            this.btnAddDoc.Visible = false;
            this.btnAddDoc.Click += new System.EventHandler(this.btnAddDoc_Click_1);
            // 
            // document
            // 
            this.document.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.document.Location = new System.Drawing.Point(377, 145);
            this.document.Multiline = true;
            this.document.Name = "document";
            this.document.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.document.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.document.Size = new System.Drawing.Size(266, 148);
            this.document.TabIndex = 515;
            this.document.Text = "*";
            this.document.Visible = false;
            // 
            // deleteRow
            // 
            this.deleteRow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.deleteRow.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.deleteRow.Location = new System.Drawing.Point(206, 395);
            this.deleteRow.Name = "deleteRow";
            this.deleteRow.Size = new System.Drawing.Size(39, 71);
            this.deleteRow.TabIndex = 514;
            this.deleteRow.Text = "مسح";
            this.deleteRow.UseVisualStyleBackColor = false;
            this.deleteRow.Visible = false;
            this.deleteRow.Click += new System.EventHandler(this.deleteRow_Click_1);
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.label24.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label24.Location = new System.Drawing.Point(1127, 260);
            this.label24.Name = "label24";
            this.label24.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label24.Size = new System.Drawing.Size(44, 27);
            this.label24.TabIndex = 513;
            this.label24.Text = "تعليق:";
            // 
            // Comment
            // 
            this.Comment.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.Comment.Location = new System.Drawing.Point(846, 264);
            this.Comment.Multiline = true;
            this.Comment.Name = "Comment";
            this.Comment.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Comment.Size = new System.Drawing.Size(266, 119);
            this.Comment.TabIndex = 512;
            this.Comment.Text = "لا تعليق";
            // 
            // btnSavePrint
            // 
            this.btnSavePrint.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnSavePrint.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnSavePrint.Location = new System.Drawing.Point(9, 395);
            this.btnSavePrint.Name = "btnSavePrint";
            this.btnSavePrint.Size = new System.Drawing.Size(94, 71);
            this.btnSavePrint.TabIndex = 509;
            this.btnSavePrint.Text = "طباعة وحفظ";
            this.btnSavePrint.UseVisualStyleBackColor = false;
            this.btnSavePrint.Click += new System.EventHandler(this.btnSavePrint_Click_1);
            // 
            // ResetAll
            // 
            this.ResetAll.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.ResetAll.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ResetAll.Location = new System.Drawing.Point(108, 395);
            this.ResetAll.Name = "ResetAll";
            this.ResetAll.Size = new System.Drawing.Size(92, 71);
            this.ResetAll.TabIndex = 508;
            this.ResetAll.Text = "مسح جميع الحقول";
            this.ResetAll.UseVisualStyleBackColor = false;
            this.ResetAll.Click += new System.EventHandler(this.ResetAll_Click_1);
            // 
            // SearchFile
            // 
            this.SearchFile.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.SearchFile.Location = new System.Drawing.Point(550, 0);
            this.SearchFile.Name = "SearchFile";
            this.SearchFile.Size = new System.Drawing.Size(319, 35);
            this.SearchFile.TabIndex = 507;
            this.SearchFile.Visible = false;
            this.SearchFile.TextChanged += new System.EventHandler(this.SearchFile_TextChanged);
            // 
            // labelArch
            // 
            this.labelArch.AutoSize = true;
            this.labelArch.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.labelArch.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.labelArch.Location = new System.Drawing.Point(112, 6);
            this.labelArch.Name = "labelArch";
            this.labelArch.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelArch.Size = new System.Drawing.Size(79, 27);
            this.labelArch.TabIndex = 506;
            this.labelArch.Text = "حالة الأرشفة:";
            // 
            // ArchivedSt
            // 
            this.ArchivedSt.AutoSize = true;
            this.ArchivedSt.BackColor = System.Drawing.Color.Red;
            this.ArchivedSt.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.ArchivedSt.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ArchivedSt.Location = new System.Drawing.Point(3, 2);
            this.ArchivedSt.Name = "ArchivedSt";
            this.ArchivedSt.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ArchivedSt.Size = new System.Drawing.Size(98, 31);
            this.ArchivedSt.TabIndex = 505;
            this.ArchivedSt.Text = "غير مؤرشف";
            this.ArchivedSt.UseVisualStyleBackColor = false;
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button4.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button4.Location = new System.Drawing.Point(204, 2);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(114, 33);
            this.button4.TabIndex = 504;
            this.button4.Text = "فتح المستند رقم 2";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click_1);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button3.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button3.Location = new System.Drawing.Point(871, 0);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(129, 33);
            this.button3.TabIndex = 503;
            this.button3.Text = "بحث القائمة أدناه";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.button2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.button2.Location = new System.Drawing.Point(324, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(112, 33);
            this.button2.TabIndex = 502;
            this.button2.Text = "فتح المستند رقم 1";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // ListSearch
            // 
            this.ListSearch.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.ListSearch.Location = new System.Drawing.Point(550, 0);
            this.ListSearch.Name = "ListSearch";
            this.ListSearch.Size = new System.Drawing.Size(319, 35);
            this.ListSearch.TabIndex = 501;
            this.ListSearch.TextChanged += new System.EventHandler(this.ListSearch_TextChanged);
            // 
            // SearchDoc
            // 
            this.SearchDoc.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.SearchDoc.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.SearchDoc.Location = new System.Drawing.Point(442, 2);
            this.SearchDoc.Name = "SearchDoc";
            this.SearchDoc.Size = new System.Drawing.Size(102, 33);
            this.SearchDoc.TabIndex = 500;
            this.SearchDoc.Text = "بحث المستند";
            this.SearchDoc.UseVisualStyleBackColor = true;
            this.SearchDoc.Click += new System.EventHandler(this.SearchDoc_Click_1);
            // 
            // PersonDesc
            // 
            this.PersonDesc.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PersonDesc.FormattingEnabled = true;
            this.PersonDesc.Items.AddRange(new object[] {
            "شخصي",
            "وثائق رسمية",
            "ابني",
            "ابنتي",
            "زوجتي",
            "والدي",
            "والدتي",
            "شقيقي",
            "شقيقتي"});
            this.PersonDesc.Location = new System.Drawing.Point(377, 22);
            this.PersonDesc.Name = "PersonDesc";
            this.PersonDesc.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.PersonDesc.Size = new System.Drawing.Size(266, 35);
            this.PersonDesc.TabIndex = 499;
            this.PersonDesc.SelectedIndexChanged += new System.EventHandler(this.PersonDesc_SelectedIndexChanged_1);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(660, 25);
            this.label6.Name = "label6";
            this.label6.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label6.Size = new System.Drawing.Size(138, 27);
            this.label6.TabIndex = 498;
            this.label6.Text = "الشخص موضوع الإقرار:";
            // 
            // ConsulateEmployee
            // 
            this.ConsulateEmployee.AutoSize = true;
            this.ConsulateEmployee.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.ConsulateEmployee.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.ConsulateEmployee.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ConsulateEmployee.Location = new System.Drawing.Point(29, 14);
            this.ConsulateEmployee.Name = "ConsulateEmployee";
            this.ConsulateEmployee.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ConsulateEmployee.Size = new System.Drawing.Size(51, 27);
            this.ConsulateEmployee.TabIndex = 497;
            this.ConsulateEmployee.Text = "الموظف";
            // 
            // timer2
            // 
            this.timer2.Tick += new System.EventHandler(this.timer2_Tick_1);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick_1);
            // 
            // checkedViewed
            // 
            this.checkedViewed.AutoSize = true;
            this.checkedViewed.Checked = true;
            this.checkedViewed.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkedViewed.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F);
            this.checkedViewed.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.checkedViewed.Location = new System.Drawing.Point(73, 305);
            this.checkedViewed.Name = "checkedViewed";
            this.checkedViewed.Size = new System.Drawing.Size(151, 33);
            this.checkedViewed.TabIndex = 495;
            this.checkedViewed.Text = "NotViewed";
            this.checkedViewed.UseVisualStyleBackColor = true;
            this.checkedViewed.Visible = false;
            // 
            // mandoubLabel
            // 
            this.mandoubLabel.AutoSize = true;
            this.mandoubLabel.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.mandoubLabel.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.mandoubLabel.Location = new System.Drawing.Point(263, 225);
            this.mandoubLabel.Name = "mandoubLabel";
            this.mandoubLabel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.mandoubLabel.Size = new System.Drawing.Size(79, 27);
            this.mandoubLabel.TabIndex = 494;
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
            this.mandoubName.Location = new System.Drawing.Point(5, 223);
            this.mandoubName.Name = "mandoubName";
            this.mandoubName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.mandoubName.Size = new System.Drawing.Size(248, 35);
            this.mandoubName.TabIndex = 493;
            this.mandoubName.Visible = false;
            // 
            // AppType
            // 
            this.AppType.AutoSize = true;
            this.AppType.Checked = true;
            this.AppType.CheckState = System.Windows.Forms.CheckState.Checked;
            this.AppType.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.AppType.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.AppType.Location = new System.Drawing.Point(45, 190);
            this.AppType.Name = "AppType";
            this.AppType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AppType.Size = new System.Drawing.Size(170, 31);
            this.AppType.TabIndex = 492;
            this.AppType.Text = "حضور مباشرة إلى القنصلية";
            this.AppType.UseVisualStyleBackColor = true;
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Font = new System.Drawing.Font("Arabic Typesetting", 18F);
            this.label21.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label21.Location = new System.Drawing.Point(263, 194);
            this.label21.Name = "label21";
            this.label21.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label21.Size = new System.Drawing.Size(115, 27);
            this.label21.TabIndex = 491;
            this.label21.Text = "طريقة تقديم الطلب:";
            // 
            // نوع_الهوية
            // 
            this.نوع_الهوية.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.نوع_الهوية.FormattingEnabled = true;
            this.نوع_الهوية.Items.AddRange(new object[] {
            "جواز سفر",
            "رقم وطني"});
            this.نوع_الهوية.Location = new System.Drawing.Point(846, 104);
            this.نوع_الهوية.Name = "نوع_الهوية";
            this.نوع_الهوية.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.نوع_الهوية.Size = new System.Drawing.Size(266, 35);
            this.نوع_الهوية.TabIndex = 490;
            this.نوع_الهوية.Text = "جواز سفر";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(1127, 104);
            this.label5.Name = "label5";
            this.label5.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label5.Size = new System.Drawing.Size(118, 27);
            this.label5.TabIndex = 489;
            this.label5.Text = "نوع اثبات الشخصية:";
            // 
            // مقدم_الطلب
            // 
            this.مقدم_الطلب.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.مقدم_الطلب.Location = new System.Drawing.Point(846, 63);
            this.مقدم_الطلب.Name = "مقدم_الطلب";
            this.مقدم_الطلب.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.مقدم_الطلب.Size = new System.Drawing.Size(266, 35);
            this.مقدم_الطلب.TabIndex = 487;
            this.مقدم_الطلب.TextChanged += new System.EventHandler(this.مقدم_الطلب_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(1127, 65);
            this.label2.Name = "label2";
            this.label2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label2.Size = new System.Drawing.Size(103, 27);
            this.label2.TabIndex = 488;
            this.label2.Text = "اسم مقدم الطلب:";
            // 
            // AppWrongName
            // 
            this.AppWrongName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AppWrongName.Location = new System.Drawing.Point(377, 104);
            this.AppWrongName.Name = "AppWrongName";
            this.AppWrongName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AppWrongName.Size = new System.Drawing.Size(266, 35);
            this.AppWrongName.TabIndex = 485;
            // 
            // labelWrongName
            // 
            this.labelWrongName.AutoSize = true;
            this.labelWrongName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelWrongName.Location = new System.Drawing.Point(660, 104);
            this.labelWrongName.Name = "labelWrongName";
            this.labelWrongName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelWrongName.Size = new System.Drawing.Size(179, 27);
            this.labelWrongName.TabIndex = 486;
            this.labelWrongName.Text = "الاسم الوارد بصورة غير صحيحة:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(263, 267);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(102, 27);
            this.label1.TabIndex = 484;
            this.label1.Text = "اسم موقع الاقرار:";
            // 
            // Iqrarid
            // 
            this.Iqrarid.Enabled = false;
            this.Iqrarid.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Iqrarid.Location = new System.Drawing.Point(8, 59);
            this.Iqrarid.Name = "Iqrarid";
            this.Iqrarid.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Iqrarid.Size = new System.Drawing.Size(249, 35);
            this.Iqrarid.TabIndex = 483;
            this.Iqrarid.Text = "ق س ج/160/xyz";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label19.Location = new System.Drawing.Point(263, 62);
            this.label19.Name = "label19";
            this.label19.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label19.Size = new System.Drawing.Size(70, 27);
            this.label19.TabIndex = 482;
            this.label19.Text = " رقم الإقرار:";
            // 
            // التاريخ_الهجري
            // 
            this.التاريخ_الهجري.Enabled = false;
            this.التاريخ_الهجري.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.التاريخ_الهجري.Location = new System.Drawing.Point(4, 100);
            this.التاريخ_الهجري.Name = "التاريخ_الهجري";
            this.التاريخ_الهجري.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.التاريخ_الهجري.Size = new System.Drawing.Size(258, 35);
            this.التاريخ_الهجري.TabIndex = 481;
            // 
            // AttendViceConsul
            // 
            this.AttendViceConsul.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AttendViceConsul.FormattingEnabled = true;
            this.AttendViceConsul.Items.AddRange(new object[] {
            "محمد عثمان عكاشة الحسين",
            "يوسف صديق أبوعاقلة",
            "لبيب محمد أحمد"});
            this.AttendViceConsul.Location = new System.Drawing.Point(5, 264);
            this.AttendViceConsul.Name = "AttendViceConsul";
            this.AttendViceConsul.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AttendViceConsul.Size = new System.Drawing.Size(248, 35);
            this.AttendViceConsul.TabIndex = 480;
            this.AttendViceConsul.SelectedIndexChanged += new System.EventHandler(this.AttendViceConsul_SelectedIndexChanged);
            // 
            // النوع
            // 
            this.النوع.AutoSize = true;
            this.النوع.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.النوع.Location = new System.Drawing.Point(1063, 149);
            this.النوع.Name = "النوع";
            this.النوع.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.النوع.Size = new System.Drawing.Size(49, 31);
            this.النوع.TabIndex = 479;
            this.النوع.Text = "ذكر";
            this.النوع.UseVisualStyleBackColor = true;
            this.النوع.CheckedChanged += new System.EventHandler(this.ApplicantSex_CheckedChanged_1);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(263, 107);
            this.label11.Name = "label11";
            this.label11.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label11.Size = new System.Drawing.Size(90, 27);
            this.label11.TabIndex = 478;
            this.label11.Text = "التاريخ الهجري:";
            // 
            // التاريخ_الميلادي
            // 
            this.التاريخ_الميلادي.Enabled = false;
            this.التاريخ_الميلادي.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.التاريخ_الميلادي.Location = new System.Drawing.Point(4, 141);
            this.التاريخ_الميلادي.Name = "التاريخ_الميلادي";
            this.التاريخ_الميلادي.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.التاريخ_الميلادي.Size = new System.Drawing.Size(258, 35);
            this.التاريخ_الميلادي.TabIndex = 476;
            this.التاريخ_الميلادي.TextChanged += new System.EventHandler(this.التاريخ_الميلادي_TextChanged);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(263, 144);
            this.label12.Name = "label12";
            this.label12.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label12.Size = new System.Drawing.Size(94, 27);
            this.label12.TabIndex = 477;
            this.label12.Text = "التاريخ الميلادي:";
            // 
            // مكان_الإصدار
            // 
            this.مكان_الإصدار.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.مكان_الإصدار.Location = new System.Drawing.Point(846, 222);
            this.مكان_الإصدار.Name = "مكان_الإصدار";
            this.مكان_الإصدار.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.مكان_الإصدار.Size = new System.Drawing.Size(266, 35);
            this.مكان_الإصدار.TabIndex = 474;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(1129, 222);
            this.label7.Name = "label7";
            this.label7.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label7.Size = new System.Drawing.Size(87, 27);
            this.label7.TabIndex = 475;
            this.label7.Text = "مكان الإصدار:";
            // 
            // رقم_الهوية
            // 
            this.رقم_الهوية.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.رقم_الهوية.Location = new System.Drawing.Point(846, 181);
            this.رقم_الهوية.Name = "رقم_الهوية";
            this.رقم_الهوية.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.رقم_الهوية.Size = new System.Drawing.Size(266, 35);
            this.رقم_الهوية.TabIndex = 472;
            this.رقم_الهوية.Text = "P0";
            // 
            // labeldoctype
            // 
            this.labeldoctype.AutoSize = true;
            this.labeldoctype.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labeldoctype.Location = new System.Drawing.Point(1127, 181);
            this.labeldoctype.Name = "labeldoctype";
            this.labeldoctype.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labeldoctype.Size = new System.Drawing.Size(105, 27);
            this.labeldoctype.TabIndex = 473;
            this.labeldoctype.Text = "رقم الوثيقة المقدمة:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(1127, 149);
            this.label3.Name = "label3";
            this.label3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label3.Size = new System.Drawing.Size(40, 27);
            this.label3.TabIndex = 471;
            this.label3.Text = "النوع:";
            // 
            // AppTrueName
            // 
            this.AppTrueName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AppTrueName.Location = new System.Drawing.Point(377, 63);
            this.AppTrueName.Name = "AppTrueName";
            this.AppTrueName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.AppTrueName.Size = new System.Drawing.Size(266, 35);
            this.AppTrueName.TabIndex = 469;
            // 
            // labelName
            // 
            this.labelName.AutoSize = true;
            this.labelName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelName.Location = new System.Drawing.Point(660, 69);
            this.labelName.Name = "labelName";
            this.labelName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelName.Size = new System.Drawing.Size(96, 27);
            this.labelName.TabIndex = 470;
            this.labelName.Text = "الاسم الصحيح :";
            // 
            // PanelFiles
            // 
            this.PanelFiles.Controls.Add(this.ArchivedSt);
            this.PanelFiles.Controls.Add(this.SearchFile);
            this.PanelFiles.Controls.Add(this.SearchDoc);
            this.PanelFiles.Controls.Add(this.ListSearch);
            this.PanelFiles.Controls.Add(this.button2);
            this.PanelFiles.Controls.Add(this.button3);
            this.PanelFiles.Controls.Add(this.button4);
            this.PanelFiles.Controls.Add(this.labelArch);
            this.PanelFiles.Location = new System.Drawing.Point(282, 11);
            this.PanelFiles.Name = "PanelFiles";
            this.PanelFiles.Size = new System.Drawing.Size(1011, 40);
            this.PanelFiles.TabIndex = 629;
            // 
            // PanelMain
            // 
            this.PanelMain.Controls.Add(this.التاريخ_الميلادي_off);
            this.PanelMain.Controls.Add(this.txtEditID2);
            this.PanelMain.Controls.Add(this.btnEditID);
            this.PanelMain.Controls.Add(this.AttendViceConsul);
            this.PanelMain.Controls.Add(this.deleteRow);
            this.PanelMain.Controls.Add(this.labelName);
            this.PanelMain.Controls.Add(this.IqrarType);
            this.PanelMain.Controls.Add(this.AppTrueName);
            this.PanelMain.Controls.Add(this.ResetAll);
            this.PanelMain.Controls.Add(this.label4);
            this.PanelMain.Controls.Add(this.label3);
            this.PanelMain.Controls.Add(this.btnAddDoc);
            this.PanelMain.Controls.Add(this.labeldoctype);
            this.PanelMain.Controls.Add(this.document);
            this.PanelMain.Controls.Add(this.checkedViewed);
            this.PanelMain.Controls.Add(this.رقم_الهوية);
            this.PanelMain.Controls.Add(this.label7);
            this.PanelMain.Controls.Add(this.label24);
            this.PanelMain.Controls.Add(this.مكان_الإصدار);
            this.PanelMain.Controls.Add(this.Comment);
            this.PanelMain.Controls.Add(this.label12);
            this.PanelMain.Controls.Add(this.التاريخ_الميلادي);
            this.PanelMain.Controls.Add(this.label11);
            this.PanelMain.Controls.Add(this.النوع);
            this.PanelMain.Controls.Add(this.التاريخ_الهجري);
            this.PanelMain.Controls.Add(this.PersonDesc);
            this.PanelMain.Controls.Add(this.label19);
            this.PanelMain.Controls.Add(this.label6);
            this.PanelMain.Controls.Add(this.label1);
            this.PanelMain.Controls.Add(this.labelWrongName);
            this.PanelMain.Controls.Add(this.AppWrongName);
            this.PanelMain.Controls.Add(this.mandoubLabel);
            this.PanelMain.Controls.Add(this.label2);
            this.PanelMain.Controls.Add(this.mandoubName);
            this.PanelMain.Controls.Add(this.مقدم_الطلب);
            this.PanelMain.Controls.Add(this.AppType);
            this.PanelMain.Controls.Add(this.label5);
            this.PanelMain.Controls.Add(this.label21);
            this.PanelMain.Controls.Add(this.نوع_الهوية);
            this.PanelMain.Controls.Add(this.btnSavePrint);
            this.PanelMain.Controls.Add(this.txtEditID1);
            this.PanelMain.Controls.Add(this.Iqrarid);
            this.PanelMain.Location = new System.Drawing.Point(28, 78);
            this.PanelMain.Name = "PanelMain";
            this.PanelMain.Size = new System.Drawing.Size(1250, 481);
            this.PanelMain.TabIndex = 630;
            this.PanelMain.Visible = false;
            // 
            // التاريخ_الميلادي_off
            // 
            this.التاريخ_الميلادي_off.Enabled = false;
            this.التاريخ_الميلادي_off.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.التاريخ_الميلادي_off.Location = new System.Drawing.Point(366, 359);
            this.التاريخ_الميلادي_off.Name = "التاريخ_الميلادي_off";
            this.التاريخ_الميلادي_off.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.التاريخ_الميلادي_off.Size = new System.Drawing.Size(257, 35);
            this.التاريخ_الميلادي_off.TabIndex = 853;
            this.التاريخ_الميلادي_off.Visible = false;
            // 
            // txtEditID2
            // 
            this.txtEditID2.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtEditID2.Location = new System.Drawing.Point(62, 59);
            this.txtEditID2.Name = "txtEditID2";
            this.txtEditID2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.txtEditID2.Size = new System.Drawing.Size(42, 35);
            this.txtEditID2.TabIndex = 854;
            this.txtEditID2.Visible = false;
            // 
            // btnEditID
            // 
            this.btnEditID.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnEditID.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEditID.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEditID.Location = new System.Drawing.Point(5, 60);
            this.btnEditID.Name = "btnEditID";
            this.btnEditID.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnEditID.Size = new System.Drawing.Size(54, 34);
            this.btnEditID.TabIndex = 852;
            this.btnEditID.Text = "تعديل";
            this.btnEditID.UseVisualStyleBackColor = false;
            this.btnEditID.Visible = false;
            this.btnEditID.Click += new System.EventHandler(this.btnEditID_Click);
            // 
            // txtEditID1
            // 
            this.txtEditID1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtEditID1.Location = new System.Drawing.Point(110, 59);
            this.txtEditID1.Name = "txtEditID1";
            this.txtEditID1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.txtEditID1.Size = new System.Drawing.Size(152, 35);
            this.txtEditID1.TabIndex = 853;
            this.txtEditID1.Visible = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.Location = new System.Drawing.Point(27, 49);
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
            this.dataGridView1.Size = new System.Drawing.Size(1255, 632);
            this.dataGridView1.TabIndex = 631;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            // 
            // Form7
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1305, 693);
            this.Controls.Add(this.PanelFiles);
            this.Controls.Add(this.ConsulateEmployee);
            this.Controls.Add(this.PanelMain);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form7";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Text = "إسناد اسمين لذات واحدة";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form7_FormClosed);
            this.Load += new System.EventHandler(this.Form7_Load);
            this.PanelFiles.ResumeLayout(false);
            this.PanelFiles.PerformLayout();
            this.PanelMain.ResumeLayout(false);
            this.PanelMain.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox IqrarType;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnAddDoc;
        private System.Windows.Forms.TextBox document;
        private System.Windows.Forms.Button deleteRow;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.TextBox Comment;
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
        private System.Windows.Forms.ComboBox PersonDesc;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label ConsulateEmployee;
        private System.Windows.Forms.Timer timer2;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.CheckBox checkedViewed;
        private System.Windows.Forms.Label mandoubLabel;
        private System.Windows.Forms.ComboBox mandoubName;
        private System.Windows.Forms.CheckBox AppType;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.ComboBox نوع_الهوية;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox مقدم_الطلب;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox AppWrongName;
        private System.Windows.Forms.Label labelWrongName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox Iqrarid;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TextBox التاريخ_الهجري;
        private System.Windows.Forms.ComboBox AttendViceConsul;
        private System.Windows.Forms.CheckBox النوع;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox التاريخ_الميلادي;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox مكان_الإصدار;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox رقم_الهوية;
        private System.Windows.Forms.Label labeldoctype;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox AppTrueName;
        private System.Windows.Forms.Label labelName;
        private System.Windows.Forms.Panel PanelFiles;
        private System.Windows.Forms.Panel PanelMain;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TextBox txtEditID2;
        private System.Windows.Forms.Button btnEditID;
        private System.Windows.Forms.TextBox txtEditID1;
        private System.Windows.Forms.TextBox التاريخ_الميلادي_off;
    }
}