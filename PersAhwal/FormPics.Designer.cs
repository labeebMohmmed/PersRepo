
namespace PersAhwal
{
    partial class FormPics
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle13 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle14 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormPics));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle15 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle16 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.توضيح_المعاملة = new System.Windows.Forms.Button();
            this.btnAuth = new System.Windows.Forms.Button();
            this.btnSaveEnd = new System.Windows.Forms.Button();
            this.Combo1 = new System.Windows.Forms.ComboBox();
            this.Combo2 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.requiredDocument = new System.Windows.Forms.TextBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.timer2 = new System.Windows.Forms.Timer(this.components);
            this.button2 = new System.Windows.Forms.Button();
            this.docId = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.printPreviewDialog1 = new System.Windows.Forms.PrintPreviewDialog();
            this.button4 = new System.Windows.Forms.Button();
            this.mandoubName = new System.Windows.Forms.ComboBox();
            this.loadPic = new System.Windows.Forms.Button();
            this.reLoadPic = new System.Windows.Forms.Button();
            this.DocType = new System.Windows.Forms.CheckBox();
            this.btnArchived = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnExten = new System.Windows.Forms.Button();
            this.txtIDNo = new System.Windows.Forms.TextBox();
            this.button5 = new System.Windows.Forms.Button();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.بيانات_المندوب = new System.Windows.Forms.TextBox();
            this.التاريخ_الميلادي = new System.Windows.Forms.TextBox();
            this.panelFinalArch = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.nameSave = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.تاريخ_الميلاد = new System.Windows.Forms.TextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.lalProType = new System.Windows.Forms.Label();
            this.panelpicTemp = new System.Windows.Forms.Panel();
            this.btnArchMandoub = new System.Windows.Forms.Button();
            this.commentPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.button1 = new System.Windows.Forms.Button();
            this.تعليق_جديد_Off = new System.Windows.Forms.TextBox();
            this.التعليقات_السابقة_Off = new System.Windows.Forms.TextBox();
            this.تعليق = new System.Windows.Forms.TextBox();
            this.drawPicPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.panelFinalArch.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.commentPanel.SuspendLayout();
            this.drawPicPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.dataGridView1);
            this.panel1.Controls.Add(this.توضيح_المعاملة);
            this.panel1.Location = new System.Drawing.Point(12, 7);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(718, 716);
            this.panel1.TabIndex = 0;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::PersAhwal.Properties.Resources.noImage;
            this.pictureBox1.Location = new System.Drawing.Point(3, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(577, 706);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.LoadCompleted += new System.ComponentModel.AsyncCompletedEventHandler(this.pictureBox1_LoadCompleted);
            this.pictureBox1.BackgroundImageChanged += new System.EventHandler(this.pictureBox1_BackgroundImageChanged);
            this.pictureBox1.LocationChanged += new System.EventHandler(this.pictureBox1_LocationChanged);
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle13.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle13.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle13.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle13.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle13.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle13.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle13;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle14.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle14.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle14.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle14.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle14.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle14.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle14;
            this.dataGridView1.Location = new System.Drawing.Point(357, 253);
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
            this.dataGridView1.Size = new System.Drawing.Size(253, 225);
            this.dataGridView1.TabIndex = 546;
            this.dataGridView1.Visible = false;
            // 
            // توضيح_المعاملة
            // 
            this.توضيح_المعاملة.BackColor = System.Drawing.SystemColors.Control;
            this.توضيح_المعاملة.FlatAppearance.BorderSize = 0;
            this.توضيح_المعاملة.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.توضيح_المعاملة.Font = new System.Drawing.Font("Arabic Typesetting", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.توضيح_المعاملة.Location = new System.Drawing.Point(4, 65);
            this.توضيح_المعاملة.Name = "توضيح_المعاملة";
            this.توضيح_المعاملة.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.توضيح_المعاملة.Size = new System.Drawing.Size(574, 588);
            this.توضيح_المعاملة.TabIndex = 594;
            this.توضيح_المعاملة.Text = "وصف المعاملة";
            this.توضيح_المعاملة.UseVisualStyleBackColor = false;
            // 
            // btnAuth
            // 
            this.btnAuth.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAuth.Location = new System.Drawing.Point(3, 3);
            this.btnAuth.Name = "btnAuth";
            this.btnAuth.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnAuth.Size = new System.Drawing.Size(311, 62);
            this.btnAuth.TabIndex = 509;
            this.btnAuth.Text = "بدء الارشفة";
            this.btnAuth.UseVisualStyleBackColor = true;
            this.btnAuth.Click += new System.EventHandler(this.btnAuth_Click);
            // 
            // btnSaveEnd
            // 
            this.btnSaveEnd.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveEnd.Location = new System.Drawing.Point(736, 668);
            this.btnSaveEnd.Name = "btnSaveEnd";
            this.btnSaveEnd.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnSaveEnd.Size = new System.Drawing.Size(298, 55);
            this.btnSaveEnd.TabIndex = 510;
            this.btnSaveEnd.Text = "حفظ وإنهاء الارشفة";
            this.btnSaveEnd.UseVisualStyleBackColor = true;
            this.btnSaveEnd.Click += new System.EventHandler(this.button1_Click);
            // 
            // Combo1
            // 
            this.Combo1.Font = new System.Drawing.Font("Arabic Typesetting", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Combo1.FormattingEnabled = true;
            this.Combo1.Location = new System.Drawing.Point(1041, 111);
            this.Combo1.Name = "Combo1";
            this.Combo1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Combo1.Size = new System.Drawing.Size(311, 42);
            this.Combo1.TabIndex = 609;
            this.Combo1.Text = "إختر نوع التوكيل";
            this.Combo1.SelectedIndexChanged += new System.EventHandler(this.CombAuthType_SelectedIndexChanged);
            this.Combo1.TextUpdate += new System.EventHandler(this.Combo1_TextUpdate);
            this.Combo1.TextChanged += new System.EventHandler(this.Combo1_TextChanged);
            // 
            // Combo2
            // 
            this.Combo2.Font = new System.Drawing.Font("Arabic Typesetting", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Combo2.FormattingEnabled = true;
            this.Combo2.Location = new System.Drawing.Point(1042, 156);
            this.Combo2.Name = "Combo2";
            this.Combo2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Combo2.Size = new System.Drawing.Size(310, 42);
            this.Combo2.TabIndex = 608;
            this.Combo2.Text = "إختر الإجراء";
            this.Combo2.SelectedIndexChanged += new System.EventHandler(this.Combo2_SelectedIndexChanged);
            this.Combo2.TextChanged += new System.EventHandler(this.Combo2_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arabic Typesetting", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(170, 2);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(138, 31);
            this.label1.TabIndex = 614;
            this.label1.Text = "المواطن مقدم الطلب:";
            // 
            // requiredDocument
            // 
            this.requiredDocument.Enabled = false;
            this.requiredDocument.Font = new System.Drawing.Font("Arabic Typesetting", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.requiredDocument.Location = new System.Drawing.Point(4, 46);
            this.requiredDocument.Multiline = true;
            this.requiredDocument.Name = "requiredDocument";
            this.requiredDocument.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.requiredDocument.Size = new System.Drawing.Size(305, 494);
            this.requiredDocument.TabIndex = 616;
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
            this.timer2.Interval = 1000;
            this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(3, 3);
            this.button2.Name = "button2";
            this.button2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.button2.Size = new System.Drawing.Size(155, 62);
            this.button2.TabIndex = 617;
            this.button2.Text = "اعادة الارشفة";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Visible = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // docId
            // 
            this.docId.Font = new System.Drawing.Font("Arabic Typesetting", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.docId.Location = new System.Drawing.Point(1042, 7);
            this.docId.Multiline = true;
            this.docId.Name = "docId";
            this.docId.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.docId.Size = new System.Drawing.Size(311, 47);
            this.docId.TabIndex = 621;
            this.docId.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.docId.TextChanged += new System.EventHandler(this.docId_TextChanged);
            this.docId.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.docId_KeyPress);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.Location = new System.Drawing.Point(1041, 156);
            this.button3.Name = "button3";
            this.button3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.button3.Size = new System.Drawing.Size(312, 42);
            this.button3.TabIndex = 622;
            this.button3.Text = "أرشفة معاملة نهائية أو إضافة إلى ملف سابق";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Visible = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // printPreviewDialog1
            // 
            this.printPreviewDialog1.AutoScrollMargin = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.AutoScrollMinSize = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.ClientSize = new System.Drawing.Size(400, 300);
            this.printPreviewDialog1.Enabled = true;
            this.printPreviewDialog1.Icon = ((System.Drawing.Icon)(resources.GetObject("printPreviewDialog1.Icon")));
            this.printPreviewDialog1.Name = "printPreviewDialog1";
            this.printPreviewDialog1.Visible = false;
            this.printPreviewDialog1.Load += new System.EventHandler(this.printPreviewDialog1_Load);
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button4.Location = new System.Drawing.Point(1040, 111);
            this.button4.Name = "button4";
            this.button4.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.button4.Size = new System.Drawing.Size(312, 42);
            this.button4.TabIndex = 624;
            this.button4.Text = "إضافة مستندات";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // mandoubName
            // 
            this.mandoubName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mandoubName.FormattingEnabled = true;
            this.mandoubName.Location = new System.Drawing.Point(1104, 247);
            this.mandoubName.Name = "mandoubName";
            this.mandoubName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.mandoubName.Size = new System.Drawing.Size(250, 35);
            this.mandoubName.TabIndex = 625;
            this.mandoubName.SelectedIndexChanged += new System.EventHandler(this.mandoubName_SelectedIndexChanged);
            this.mandoubName.TextChanged += new System.EventHandler(this.mandoubName_TextChanged);
            // 
            // loadPic
            // 
            this.loadPic.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.loadPic.Location = new System.Drawing.Point(3, 69);
            this.loadPic.Name = "loadPic";
            this.loadPic.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.loadPic.Size = new System.Drawing.Size(310, 59);
            this.loadPic.TabIndex = 627;
            this.loadPic.Text = "تحميل من ملف";
            this.loadPic.UseVisualStyleBackColor = true;
            this.loadPic.Click += new System.EventHandler(this.button5_Click);
            // 
            // reLoadPic
            // 
            this.reLoadPic.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.reLoadPic.Location = new System.Drawing.Point(3, 69);
            this.reLoadPic.Name = "reLoadPic";
            this.reLoadPic.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.reLoadPic.Size = new System.Drawing.Size(155, 59);
            this.reLoadPic.TabIndex = 628;
            this.reLoadPic.Text = "إعادة تحميل من ملف";
            this.reLoadPic.UseVisualStyleBackColor = true;
            this.reLoadPic.Visible = false;
            this.reLoadPic.Click += new System.EventHandler(this.reLoadPic_Click);
            // 
            // DocType
            // 
            this.DocType.AutoSize = true;
            this.DocType.Checked = true;
            this.DocType.CheckState = System.Windows.Forms.CheckState.Checked;
            this.DocType.Font = new System.Drawing.Font("Arabic Typesetting", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DocType.Location = new System.Drawing.Point(202, 7);
            this.DocType.Name = "DocType";
            this.DocType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.DocType.Size = new System.Drawing.Size(79, 25);
            this.DocType.TabIndex = 629;
            this.DocType.Text = "أصل المكاتبة";
            this.DocType.UseVisualStyleBackColor = true;
            this.DocType.CheckedChanged += new System.EventHandler(this.DocType_CheckedChanged);
            // 
            // btnArchived
            // 
            this.btnArchived.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnArchived.Location = new System.Drawing.Point(1042, 201);
            this.btnArchived.Name = "btnArchived";
            this.btnArchived.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnArchived.Size = new System.Drawing.Size(136, 42);
            this.btnArchived.TabIndex = 631;
            this.btnArchived.Text = "تعديل كمؤرشف نهائي";
            this.btnArchived.UseVisualStyleBackColor = true;
            this.btnArchived.Visible = false;
            this.btnArchived.Click += new System.EventHandler(this.btnArchived_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDelete.Location = new System.Drawing.Point(1267, 201);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnDelete.Size = new System.Drawing.Size(85, 42);
            this.btnDelete.TabIndex = 632;
            this.btnDelete.Text = "حذف";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Visible = false;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnExten
            // 
            this.btnExten.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExten.Location = new System.Drawing.Point(1184, 201);
            this.btnExten.Name = "btnExten";
            this.btnExten.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnExten.Size = new System.Drawing.Size(75, 42);
            this.btnExten.TabIndex = 633;
            this.btnExten.Text = "تمديد";
            this.btnExten.UseVisualStyleBackColor = true;
            this.btnExten.Visible = false;
            this.btnExten.Click += new System.EventHandler(this.btnExten_Click);
            // 
            // txtIDNo
            // 
            this.txtIDNo.Enabled = false;
            this.txtIDNo.Font = new System.Drawing.Font("Arabic Typesetting", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtIDNo.Location = new System.Drawing.Point(1042, 60);
            this.txtIDNo.Multiline = true;
            this.txtIDNo.Name = "txtIDNo";
            this.txtIDNo.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtIDNo.Size = new System.Drawing.Size(311, 46);
            this.txtIDNo.TabIndex = 634;
            this.txtIDNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtIDNo.TextChanged += new System.EventHandler(this.txtIDNo_TextChanged);
            // 
            // button5
            // 
            this.button5.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button5.Location = new System.Drawing.Point(8, 31);
            this.button5.Name = "button5";
            this.button5.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.button5.Size = new System.Drawing.Size(272, 66);
            this.button5.TabIndex = 635;
            this.button5.Text = "تصدير استمارات للمندوب";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Visible = false;
            this.button5.Click += new System.EventHandler(this.button5_Click_1);
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToDeleteRows = false;
            this.dataGridView2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle15.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle15.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle15.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle15.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle15.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle15.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView2.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle15;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle16.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle16.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle16.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle16.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle16.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle16.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView2.DefaultCellStyle = dataGridViewCellStyle16;
            this.dataGridView2.Location = new System.Drawing.Point(15, 7);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.dataGridView2.RowTemplate.DefaultCellStyle.BackColor = System.Drawing.Color.White;
            this.dataGridView2.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Arabic Typesetting", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dataGridView2.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            this.dataGridView2.RowTemplate.DefaultCellStyle.Padding = new System.Windows.Forms.Padding(4);
            this.dataGridView2.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.LightGreen;
            this.dataGridView2.RowTemplate.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView2.RowTemplate.Height = 30;
            this.dataGridView2.Size = new System.Drawing.Size(715, 722);
            this.dataGridView2.TabIndex = 836;
            this.dataGridView2.Visible = false;
            // 
            // بيانات_المندوب
            // 
            this.بيانات_المندوب.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.بيانات_المندوب.Location = new System.Drawing.Point(78, 85);
            this.بيانات_المندوب.Multiline = true;
            this.بيانات_المندوب.Name = "بيانات_المندوب";
            this.بيانات_المندوب.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.بيانات_المندوب.Size = new System.Drawing.Size(215, 35);
            this.بيانات_المندوب.TabIndex = 845;
            this.بيانات_المندوب.Visible = false;
            // 
            // التاريخ_الميلادي
            // 
            this.التاريخ_الميلادي.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.التاريخ_الميلادي.Location = new System.Drawing.Point(78, 44);
            this.التاريخ_الميلادي.Multiline = true;
            this.التاريخ_الميلادي.Name = "التاريخ_الميلادي";
            this.التاريخ_الميلادي.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.التاريخ_الميلادي.Size = new System.Drawing.Size(215, 35);
            this.التاريخ_الميلادي.TabIndex = 844;
            this.التاريخ_الميلادي.Visible = false;
            // 
            // panelFinalArch
            // 
            this.panelFinalArch.AutoScroll = true;
            this.panelFinalArch.Controls.Add(this.loadPic);
            this.panelFinalArch.Controls.Add(this.btnAuth);
            this.panelFinalArch.Controls.Add(this.reLoadPic);
            this.panelFinalArch.Controls.Add(this.button2);
            this.panelFinalArch.Location = new System.Drawing.Point(1040, 410);
            this.panelFinalArch.Name = "panelFinalArch";
            this.panelFinalArch.Size = new System.Drawing.Size(329, 196);
            this.panelFinalArch.TabIndex = 840;
            this.panelFinalArch.Visible = false;
            // 
            // panel2
            // 
            this.panel2.AutoScroll = true;
            this.panel2.Controls.Add(this.nameSave);
            this.panel2.Controls.Add(this.requiredDocument);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Location = new System.Drawing.Point(1041, 322);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(328, 151);
            this.panel2.TabIndex = 841;
            // 
            // nameSave
            // 
            this.nameSave.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nameSave.Location = new System.Drawing.Point(3, 3);
            this.nameSave.Name = "nameSave";
            this.nameSave.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.nameSave.Size = new System.Drawing.Size(125, 37);
            this.nameSave.TabIndex = 633;
            this.nameSave.Text = "حفظ التعديل";
            this.nameSave.UseVisualStyleBackColor = true;
            this.nameSave.Visible = false;
            this.nameSave.Click += new System.EventHandler(this.nameSave_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arabic Typesetting", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.Black;
            this.label5.Location = new System.Drawing.Point(1263, 65);
            this.label5.Name = "label5";
            this.label5.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label5.Size = new System.Drawing.Size(86, 31);
            this.label5.TabIndex = 648;
            this.label5.Text = "تاريخ الميلاد:";
            this.label5.Visible = false;
            // 
            // تاريخ_الميلاد
            // 
            this.تاريخ_الميلاد.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.تاريخ_الميلاد.Location = new System.Drawing.Point(1044, 65);
            this.تاريخ_الميلاد.Multiline = true;
            this.تاريخ_الميلاد.Name = "تاريخ_الميلاد";
            this.تاريخ_الميلاد.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.تاريخ_الميلاد.Size = new System.Drawing.Size(215, 35);
            this.تاريخ_الميلاد.TabIndex = 843;
            this.تاريخ_الميلاد.Text = "عام-شهر-يوم";
            this.تاريخ_الميلاد.Visible = false;
            this.تاريخ_الميلاد.Click += new System.EventHandler(this.تاريخ_الميلاد_Click);
            this.تاريخ_الميلاد.TextChanged += new System.EventHandler(this.تاريخ_الميلاد_TextChanged);
            // 
            // panel3
            // 
            this.panel3.AutoScroll = true;
            this.panel3.Controls.Add(this.button5);
            this.panel3.Controls.Add(this.DocType);
            this.panel3.Location = new System.Drawing.Point(1056, 618);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(294, 100);
            this.panel3.TabIndex = 844;
            this.panel3.Visible = false;
            // 
            // lalProType
            // 
            this.lalProType.AutoSize = true;
            this.lalProType.Font = new System.Drawing.Font("Arabic Typesetting", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lalProType.ForeColor = System.Drawing.Color.Black;
            this.lalProType.Location = new System.Drawing.Point(1236, 208);
            this.lalProType.Name = "lalProType";
            this.lalProType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.lalProType.Size = new System.Drawing.Size(114, 31);
            this.lalProType.TabIndex = 845;
            this.lalProType.Text = "اختر آلية الأجراء:";
            this.lalProType.Visible = false;
            // 
            // panelpicTemp
            // 
            this.panelpicTemp.AutoScroll = true;
            this.panelpicTemp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelpicTemp.Location = new System.Drawing.Point(597, 10);
            this.panelpicTemp.Name = "panelpicTemp";
            this.panelpicTemp.Size = new System.Drawing.Size(130, 714);
            this.panelpicTemp.TabIndex = 846;
            this.panelpicTemp.Visible = false;
            // 
            // btnArchMandoub
            // 
            this.btnArchMandoub.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnArchMandoub.Location = new System.Drawing.Point(1041, 247);
            this.btnArchMandoub.Name = "btnArchMandoub";
            this.btnArchMandoub.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnArchMandoub.Size = new System.Drawing.Size(57, 35);
            this.btnArchMandoub.TabIndex = 847;
            this.btnArchMandoub.Text = "أرشفة";
            this.btnArchMandoub.UseVisualStyleBackColor = true;
            this.btnArchMandoub.Click += new System.EventHandler(this.btnArchMandoub_Click);
            // 
            // commentPanel
            // 
            this.commentPanel.AutoScroll = true;
            this.commentPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.commentPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.commentPanel.Controls.Add(this.button1);
            this.commentPanel.Controls.Add(this.تعليق_جديد_Off);
            this.commentPanel.Controls.Add(this.التعليقات_السابقة_Off);
            this.commentPanel.Location = new System.Drawing.Point(1040, 478);
            this.commentPanel.Name = "commentPanel";
            this.commentPanel.Size = new System.Drawing.Size(328, 245);
            this.commentPanel.TabIndex = 848;
            this.commentPanel.Visible = false;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.Control;
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(3, 3);
            this.button1.Name = "button1";
            this.button1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.button1.Size = new System.Drawing.Size(308, 35);
            this.button1.TabIndex = 678;
            this.button1.Text = "إضف تعليقا";
            this.button1.UseVisualStyleBackColor = false;
            // 
            // تعليق_جديد_Off
            // 
            this.تعليق_جديد_Off.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.تعليق_جديد_Off.Location = new System.Drawing.Point(3, 44);
            this.تعليق_جديد_Off.Multiline = true;
            this.تعليق_جديد_Off.Name = "تعليق_جديد_Off";
            this.تعليق_جديد_Off.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.تعليق_جديد_Off.Size = new System.Drawing.Size(306, 48);
            this.تعليق_جديد_Off.TabIndex = 671;
            this.تعليق_جديد_Off.Click += new System.EventHandler(this.تعليق_جديد_Off_Click);
            // 
            // التعليقات_السابقة_Off
            // 
            this.التعليقات_السابقة_Off.Enabled = false;
            this.التعليقات_السابقة_Off.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.التعليقات_السابقة_Off.Location = new System.Drawing.Point(3, 98);
            this.التعليقات_السابقة_Off.Multiline = true;
            this.التعليقات_السابقة_Off.Name = "التعليقات_السابقة_Off";
            this.التعليقات_السابقة_Off.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.التعليقات_السابقة_Off.Size = new System.Drawing.Size(306, 528);
            this.التعليقات_السابقة_Off.TabIndex = 679;
            this.التعليقات_السابقة_Off.Click += new System.EventHandler(this.التعليقات_السابقة_Off_Click);
            // 
            // تعليق
            // 
            this.تعليق.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.تعليق.Location = new System.Drawing.Point(78, 3);
            this.تعليق.Multiline = true;
            this.تعليق.Name = "تعليق";
            this.تعليق.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.تعليق.Size = new System.Drawing.Size(215, 35);
            this.تعليق.TabIndex = 846;
            this.تعليق.Visible = false;
            // 
            // drawPicPanel
            // 
            this.drawPicPanel.AutoScroll = true;
            this.drawPicPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.drawPicPanel.Controls.Add(this.تعليق);
            this.drawPicPanel.Controls.Add(this.التاريخ_الميلادي);
            this.drawPicPanel.Controls.Add(this.بيانات_المندوب);
            this.drawPicPanel.Location = new System.Drawing.Point(736, 12);
            this.drawPicPanel.Name = "drawPicPanel";
            this.drawPicPanel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.drawPicPanel.Size = new System.Drawing.Size(298, 650);
            this.drawPicPanel.TabIndex = 849;
            // 
            // FormPics
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1370, 741);
            this.Controls.Add(this.drawPicPanel);
            this.Controls.Add(this.mandoubName);
            this.Controls.Add(this.btnArchMandoub);
            this.Controls.Add(this.panelpicTemp);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.btnSaveEnd);
            this.Controls.Add(this.تاريخ_الميلاد);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panelFinalArch);
            this.Controls.Add(this.btnExten);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnArchived);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.commentPanel);
            this.Controls.Add(this.lalProType);
            this.Controls.Add(this.txtIDNo);
            this.Controls.Add(this.Combo1);
            this.Controls.Add(this.Combo2);
            this.Controls.Add(this.docId);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button4);
            this.Name = "FormPics";
            this.Text = "أرشفة الملفات";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormPics_FormClosed);
            this.Load += new System.EventHandler(this.FormPics_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.panelFinalArch.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.commentPanel.ResumeLayout(false);
            this.commentPanel.PerformLayout();
            this.drawPicPanel.ResumeLayout(false);
            this.drawPicPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnAuth;
        private System.Windows.Forms.Button btnSaveEnd;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ComboBox Combo1;
        private System.Windows.Forms.ComboBox Combo2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox requiredDocument;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Timer timer2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox docId;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.PrintPreviewDialog printPreviewDialog1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.ComboBox mandoubName;
        private System.Windows.Forms.Button loadPic;
        private System.Windows.Forms.Button reLoadPic;
        private System.Windows.Forms.CheckBox DocType;
        private System.Windows.Forms.Button btnArchived;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnExten;
        private System.Windows.Forms.TextBox txtIDNo;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Panel panelFinalArch;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox تاريخ_الميلاد;
        private System.Windows.Forms.Button nameSave;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label lalProType;
        private System.Windows.Forms.Panel panelpicTemp;
        private System.Windows.Forms.TextBox التاريخ_الميلادي;
        private System.Windows.Forms.TextBox بيانات_المندوب;
        private System.Windows.Forms.Button btnArchMandoub;
        private System.Windows.Forms.FlowLayoutPanel commentPanel;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox تعليق_جديد_Off;
        private System.Windows.Forms.TextBox التعليقات_السابقة_Off;
        private System.Windows.Forms.Button توضيح_المعاملة;
        private System.Windows.Forms.TextBox تعليق;
        private System.Windows.Forms.FlowLayoutPanel drawPicPanel;
    }
}