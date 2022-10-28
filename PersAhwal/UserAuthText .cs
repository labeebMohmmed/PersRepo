using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Globalization;
using System.Threading;

namespace PersAhwal
{
    public partial class UserAuthText : UserControl
    {
        public Delegate AppMovePage;
        public Delegate strRightsText;
        public Delegate strRightIndex;
        public Delegate strAuthSubject;
        public Delegate strAuthList2;
        string strRights = "", authList2 = "", NewAuthSubject = "", AuthSubjectValue = "", DPTitle = "المرحوم", LegaceyAttch = "";
        static int[] statistic = new int[100];
        public string specialDataSum = "";
        public string Mentioned = "باسمي";
        public int birthindex = 0;
        public static string[] BirthName = new string[10];
        public static string[] BirthPlace = new string[10];
        public static string[] BirthDate = new string[10];
        public static string[] BirthMother = new string[10];
        public static string[] BirthDecs = new string[10];
        int idShow = 0;
        bool ShowNewApp = false;
        string[] dataGrid = new string[50];
        string[] txtComboOptions = new string[5] { "","","","",""};
        string[] txtCheckOptions = new string[5] { "","","","",""};
        static int[] staticIndex = new int[100];
        static int[] times = new int[100];
        static string[,] preffix = new string[10, 20];
        static string[] Text_statis = new string[5];
        static string[] strListcomb = new string[10] { "", "", "", "", "", "", "", "", "", "" };
        static string[] strListcomb1 = new string[10] { "", "", "", "", "", "", "", "", "", "" };
        string[] dataset = new string[20];
        string strUni = "";
        int x = 0, checkboxNo = 0;
        string LastCol = "", DataSource = "";
        string ListedRightIndex = "", ColName = "", BoxesData = "";
        DataTable checkboxdt;
        string removedDocInfo;
        private int[] checkliststr = new int[100];
        int Nobox = 0, LastID = 0, LastTabIndex = 0;
        Form11 Form11Parameter;
        string spacialCharacter = "";
        int idIndex = -1;
        string[] allList = new string[100];
        public Form11 ParentForm { get; set; }
        string giveUpText = "ول### بموجب هذا التوكيل الحق في مقابلة كافة الجهات المختصة، وإكمال إجراءات التنازل وتحويل السجل في إسم***، وتسديد الرسوم المقررة، والظهور وتمثيلي أمام كافة المحاكم بمختلف أنواعها ودرجاتها، والقيام بكافة الإجراءات التي تتطلب حضور$$$، والتوقيع نيابةً عن$$$ على كافة الأوراق والمستندات اللازمة لذلك ، وأذن&&& لمن يشهد والله خير الشاهدين";
        string generalAuthText = "والوقوف والمقاضاة نيابة عن$$$ أمام كافة المحاكم والنيابات بمختلف أنواعها ودرجاتها، والقيام بكافة الإجراءات التي تتطلب حضور$$$، والتوقيع نيابةً عن$$$ على كافة الأوراق والمستندات اللازمة لذلك، وأذن&&& لمن يشهد والله خير الشاهدين";
        public string RemovedDocInfo
        {
            get { return removedDocInfo; }
            set { removedDocInfo = value; }
        }

        public ComboBox comboBoxAuthValue
        {
            get { return CombAuthType; }
            set { CombAuthType = value; }
        }
        public string AuthList2Value
        {
            get { return authList2; }
            set { authList2 = value; }
        }
        public int  BirthindexValue
        {
            get { return birthindex; }
            set { birthindex = value; }
        }
        public string[] BirthDescValue
        {
            get { return BirthDecs; }
            set { BirthDecs = value; }
        }

        public string[] BirthNameValue
        {
            get { return BirthName; }
            set { BirthName = value; }
        }
        public string[] BirthDateValue
        {
            get { return BirthDate; }
            set { BirthDate = value; }
        }
        public string[] BirthPlaceValue
        {
            get { return BirthPlace; }
            set { BirthPlace = value; }
        }
        public string[] BirthMotherValue
        {
            get { return BirthMother; }
            set { BirthMother = value; }
        }

        public string ColNameValue
        {
            get { return ColName; }
            set { ColName = value; }
        }

        public string NewAuthSubjectValue
        {
            get { return NewAuthSubject; }
            set { NewAuthSubject = value; }
        }
        public string AllBoxesData
        {
            get { return BoxesData; }
            set { BoxesData = value; }
        }

        public ComboBox ComboProcedureValue
        {
            get { return ComboProcedure; }
            set { ComboProcedure = value; }
        }

        public ComboBox comboPropertyTypeValue
        {
            get { return comboPropertyType; }
            set { comboPropertyType = value; }
        }
        

            public Button addNameValue
        {
            get { return LibtnAdd1; }
            set { 
                
                LibtnAdd1 = value; }
        }
        public FlowLayoutPanel PanelItemsboxesValue
        {
            get { return PanelItemsboxes; }
            set { PanelItemsboxes = value; }
        }
        public TextBox txtReviewValue
        {
            get { return txtReview; }
            set { txtReview = value; }
        }

        public TextBox txt1Value
        {
            get { return Vitext1; }
            set { Vitext1 = value; }
        }

        public TextBox txt2Value
        {
            get { return Vitext2; }
            set { Vitext2 = value; }
        }

        public TextBox txt3Value
        {
            get { return Vitext3; }
            set { Vitext3 = value; }
        }

        public TextBox txt4Value
        {
            get { return Vitext4; }
            set { Vitext4 = value; }
        }



        public TextBox txtAddRightValue
        {
            get { return txtAddRight; }
            set { txtAddRight = value; }
        }
        public FlowLayoutPanel panelAuthOptionsValue
        {
            get { return panelAuthOptions; }
            set { panelAuthOptions = value; }
        }
        public FlowLayoutPanel PanelSubItemBoxValue
        {
            get { return PanelSubItemBox; }
            set { PanelSubItemBox = value; }
        }

        public int LastIDValue
        {
            get { return LastID; }
            set { LastID = value; }
        }
        public DataTable checkboxdtValue
        {
            get { return checkboxdt; }
            set { checkboxdt = value; }
        }

        public Label label34Value
        {
            get { return label34; }
            set { label34 = value; }
        }
        public Label label36Value
        {
            get { return label36; }
            set { label36 = value; }
        }


        public UserAuthText()
        {
            InitializeComponent();
            //checkboxdt = new DataTable();
            //for (int x = 0; x < 28; x++) StaredColumns(x);


            comboPropertyType.SelectedIndex = 0;
        }

        private void UserAuthText_Load(object sender, EventArgs e)
        {
            //StaredColumns();

        }
        public void ColumnStatistics(string source, string col, string table)
        {
            int x = 0;
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if(((CheckBox)control).Tag.ToString() == "valid")
                    { 
                    //MessageBox.Show(((CheckBox)control).Text);
                    if (x == 0)
                        times[x]++;
                    if (((CheckBox)control).CheckState == CheckState.Checked) { statistic[x]++; }
                    UpdateColumn(source, col, x + 1, ((CheckBox)control).Text + "_" + statistic[x].ToString() + "_" + times[x].ToString() + "_" + staticIndex[x].ToString() + "_Star", false);
                        x++;
                        if (x == Nobox) return;
                    }
                }
            }
        }
        private void btnSizeSpecial_Click_1(object sender, EventArgs e)
        {

            //CreatestrAuthRight();
            //CreateAuthList2(StrSpecPur);
            //authList2 = LegaceyPreStr + ParentForm.strauthList1 + authList2;
            //strAuthList2.DynamicInvoke(authList2);
            //strAuthSubject.DynamicInvoke(AuthSubjectValue);
            //strRightIndex.DynamicInvoke(ListedRightIndex);
            //strRightsText.DynamicInvoke(strRights);
            AppMovePage.DynamicInvoke(2);
        }

        private void removedDoc(string dataSource, string documentID, string relatedDoc)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET حالة_الارشفة=@حالة_الارشفة, توكيل_مرجعي=@توكيل_مرجعي where رقم_التوكيل = @رقم_التوكيل", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_التوكيل", documentID);
            sqlCmd.Parameters.AddWithValue("@توكيل_مرجعي", relatedDoc);
            sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", "ملغي" + ParentForm.txtGreDateValue.Text);

            sqlCmd.ExecuteNonQuery();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            CreatestrAuthRight();
            collectItemsData();
            removedDocInfo = removedDocSource.Text + "_" + removedDocDate.Text + "_" + removedDocNo.Text;
            if (!ShowNewApp)
            {
                CreateAuthList2(StrSpecPur);
                authList2 = authList2 + LegaceyPreStr;
            }
            else
            {
                
                for (int x = 0; x < 30; x++)
                    StrSpecPur = SuffPrefReplacements(StrSpecPur);
                
                if (CombAuthType.SelectedIndex != 6)
                    authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + StrSpecPur + LegaceyPreStr;
                else authList2 = StrSpecPur + LegaceyPreStr;
            }
               
            
            strAuthList2.DynamicInvoke(authList2);
            strAuthSubject.DynamicInvoke(AuthSubjectValue);
            strRightIndex.DynamicInvoke(ListedRightIndex);
            if (spacialCharacter == "@*@")
            {
                strRights = strRights.Replace("لدى  برقم الايبان ()", "لدى " + Vitext4.Text + " برقم الايبان (" + Vitext5.Text + ")");
            }            
            strRightsText.DynamicInvoke(strRights);
            AppMovePage.DynamicInvoke(4);
        }


        private string SuffPrefReplacements(string text)
        {
            Suffex_preffixList();
            //MessageBox.Show(ParentForm.authList1Value);
            
            if (text.Contains("auth1"))
                return text.Replace("auth1", ParentForm.authList1Value);
            
            if (text.Contains("t1"))
                return text.Replace("t1", Vitext1.Text);
            if (text.Contains("t2"))
                return text.Replace("t2", Vitext2.Text);
            if (text.Contains("t3"))
                return text.Replace("t3", Vitext3.Text);
            if (text.Contains("t4"))
                return text.Replace("t4", Vitext4.Text);
                
            if (text.Contains("t5"))
                return text.Replace("t5", Vitext5.Text);

            if (text.Contains("c1"))
                return text.Replace("c1", Vicheck1.Text);

            if (text.Contains("m1"))
                return text.Replace("m1", Vicombo1.Text);
            if (text.Contains("m2"))
                return text.Replace("m2", Vicombo2.Text);

            if (text.Contains("a1"))
                return text.Replace("a1", LibtnAdd1.Text);

            if (text.Contains("n1"))
                return text.Replace("n1", " " + VitxtDate1VD.Text + "/" + VitxtDate1VM.Text + "/" + VitxtDate1VY.Text + " ");
            if (text.Contains("#*#"))
                return text.Replace("#*#", preffix[ParentForm.intAppcases, 10]);

            if (text.Contains("#1"))
                return text.Replace("#1", preffix[ParentForm.intAppcases, 11]);

            if (text.Contains("#2"))
                return text.Replace("#2", preffix[ParentForm.intAppcases, 12]);    
            if (text.Contains("@*@"))
            {
                spacialCharacter = "@*@";
                return text.Replace("@*@", "لدى  برقم الايبان ()");
            }

            if (text.Contains("#8"))
                return text.Replace("#8", removedDocNo.Text);
            if (text.Contains("#6"))
                return text.Replace("#6", removedDocSource.Text);
            if (text.Contains("#7"))
                return text.Replace("#7", removedDocDate.Text);
            


            if (text.Contains("#3"))
                return text.Replace("#3", preffix[0, 7]);
            if (text.Contains("#4"))
                return text.Replace("#4", preffix[0, 8]);
            if (text.Contains("#5"))
                return text.Replace("#5", preffix[0, 9]);


            if (text.Contains("#$#"))
                return text.Replace("#$#", preffix[ParentForm.intAppcases, 13]);

            if (text.Contains("$$$"))
                return text.Replace("$$$", preffix[ParentForm.intAppcases, 0]);
            if (text.Contains("&&&"))
                return text.Replace("&&&", preffix[ParentForm.intAppcases, 1]);
            if (text.Contains("^^^"))
                return text.Replace("^^^", preffix[ParentForm.intAppcases, 2]);
            if (text.Contains("###"))
                return text.Replace("###", preffix[ParentForm.intAuthcases, 4]);
            if (text.Contains("***"))
                return text.Replace("***", preffix[ParentForm.intAuthcases, 3]);
            else return text;
        }

        private void UpdateColumn(string source, string comlumnName, int id, string data, bool datatype)
        {
            //SqlConnection sqlCon = new SqlConnection(source);
            //string column = "@" + comlumnName;
            //string qurey;
            //if (datatype) qurey = "INSERT INTO TableAuthRights (" + comlumnName + ") values(" + column + ")";
            //else qurey = "UPDATE TableAuthRights SET " + comlumnName + " = " + column + " WHERE ID = @ID";

            //SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            //if (sqlCon.State == ConnectionState.Closed)
            //    sqlCon.Open();
            //sqlCmd.CommandType = CommandType.Text;

            //if (datatype)
            //{
            //    sqlCmd.Parameters.AddWithValue(column, data.Trim());
            //    sqlCmd.ExecuteNonQuery();
            //}
            //else
            //{

            //    sqlCmd.Parameters.AddWithValue("@ID", id);
            //    sqlCmd.Parameters.AddWithValue(column, data.Trim());
            //    sqlCmd.ExecuteNonQuery();
            //}
            //sqlCon.Close();
        }

        private void Suffex_preffixList()
        {
            preffix[0, 1] = "ت";//&&&
            preffix[1, 1] = "ت";
            preffix[2, 1] = "نا";
            preffix[3, 1] = "نا";
            preffix[4, 1] = "نا";
            preffix[5, 1] = "نا";

            preffix[0, 0] = "ي"; //$$$
            preffix[1, 0] = "ي";
            preffix[2, 0] = "نا";
            preffix[3, 0] = "نا";
            preffix[4, 0] = "نا";
            preffix[5, 0] = "نا";

            
            preffix[0, 2] = "ني";//^^^
            preffix[1, 2] = "ني";
            preffix[2, 2] = "نا";
            preffix[3, 2] = "نا";
            preffix[4, 2] = "نا";
            preffix[5, 2] = "نا";

            preffix[0, 3] = "";//***
            preffix[1, 3] = "ت";
            preffix[2, 3] = "ا";
            preffix[3, 3] = "تا";
            preffix[4, 3] = "ن";
            preffix[5, 3] = "وا";

            preffix[0, 4] = "ه";//###
            preffix[1, 4] = "ها";
            preffix[2, 4] = "هما";
            preffix[3, 4] = "هما";
            preffix[4, 4] = "هن";
            preffix[5, 4] = "هم";

            preffix[0, 5] = "";//#6
            preffix[1, 5] = "ة";
            preffix[2, 5] = "ان";
            preffix[3, 5] = "تان";
            preffix[4, 5] = "ات";
            preffix[5, 5] = "ون";

            preffix[0, 6] = "";//#5
            preffix[1, 6] = "ة";
            preffix[2, 6] = "ين";
            preffix[3, 6] = "تين";
            preffix[4, 6] = "ات";
            preffix[5, 6] = "رين";

            preffix[0, 7] = "ينوب";//#3
            preffix[1, 7] = "تنوب";
            preffix[2, 7] = "ينوبا";
            preffix[3, 7] = "تنوبا";
            preffix[4, 7] = "ينبن";
            preffix[5, 7] = "ينوبوا";

            preffix[0, 8] = "يقوم";//#4
            preffix[1, 8] = "تقوم";
            preffix[2, 8] = "يقوما";
            preffix[3, 8] = "تقوما";
            preffix[4, 8] = "يقمن";
            preffix[5, 8] = "يقوموا";

            preffix[0, 9] = "نصيبي";//#5
            preffix[1, 9] = "نصيبي";
            preffix[2, 9] = "نصيبينا";
            preffix[3, 9] = "نصيبينا";
            preffix[4, 9] = "أنصبتنا";
            preffix[5, 9] = "أنصبتنا";

            preffix[0, 10] = "ت";//#*#
            preffix[1, 10] = "";

            preffix[0, 11] = "التي";//#1
            preffix[1, 11] = "الذي";

            preffix[0, 12] = "هو";//#2
            preffix[1, 12] = "هي";
            preffix[2, 12] = "هما";
            preffix[3, 12] = "هما";
            preffix[4, 12] = "هن";
            preffix[5, 12] = "هم";

            preffix[0, 13] = "نت";//#$#
            preffix[1, 13] = "نت";
            preffix[2, 13] = "نا";
            preffix[3, 13] = "نا";
            preffix[4, 13] = "نا";
            preffix[5, 13] = "نا";
        }
        private void StaredColumns(int x)
        {
            checkboxdt = new DataTable();
            int LastID = 0;
            string col = "Col" + x.ToString();
            string col1 = "Col" + (x + 1).ToString(); ;
            string query = "SELECT ID," + col + " FROM TableAuthRights";
            string source = "Data Source=192.168.100.56,49170;Network Library=DBMSSOCN;Initial Catalog=AhwalDataBase;User ID=ConsJeddahAdmin;Password=DataBC0nsJ49170";
            using (SqlConnection con = new SqlConnection(source))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter(query, con))
                {
                    checkboxdt.Clear();
                    sda.Fill(checkboxdt);
                    listchecked = checkboxdt.Rows.Count;
                    LastID = 0;
                    foreach (DataRow row in checkboxdt.Rows)
                    {
                        if (checkboxdt.Rows[LastID][col].ToString().Contains("_") && (!checkboxdt.Rows[LastID][col].ToString().Contains("Star") || checkboxdt.Rows[LastID][col].ToString().Contains("off")))
                        {
                            Text_statis = checkboxdt.Rows[LastID][col].ToString().Split('_');
                            if (Text_statis[0].Contains("توكيل الغير في"))
                                UpdateColumn(source, col, LastID + 1, Text_statis[0] + "_" + Text_statis[1] + "_" + Text_statis[2] + "_" + Text_statis[3] + "_off", false);
                            else
                                UpdateColumn(source, col, LastID + 1, Text_statis[0] + "_" + Text_statis[1] + "_" + Text_statis[2] + "_" + Text_statis[3] + "_Star", false);
                            LastID++;
                        }
                    }
                }
            }

        }

        int listchecked = 0;
        public void PopulateCheckBoxes(string col, string table, string dataSource)
        {
            LastCol = col;
            if (col == "" || table == "" || dataSource == "") return;
            string query = "SELECT ID," + col + " FROM " + table;

            using (SqlConnection con = new SqlConnection(dataSource))
            {

                using (SqlDataAdapter sda = new SqlDataAdapter(query, con))
                {

                    sda.Fill(checkboxdt);
                    listchecked = checkboxdt.Rows.Count;
                    Nobox = 0;
                    foreach (DataRow row in checkboxdt.Rows)
                    {
                        if (checkboxdt.Rows[Nobox][col].ToString() == "" || checkboxdt.Rows[Nobox][col].ToString() == "null") return;
                        //{
                            CheckBox chk = new CheckBox();
                            chk.TabIndex = Nobox;
                            chk.Width = 80;
                            chk.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            if (Nobox == 0) chk.Width = panelAuthOptions.Width - 100;
                            else chk.Width = panelAuthOptions.Width - 130;
                            chk.Height = 33;
                            chk.CheckState = CheckState.Unchecked;
                            chk.Location = new System.Drawing.Point(70, 3 + Nobox * 37);
                            chk.Name = "checkBox" + Nobox.ToString();
                            Text_statis = checkboxdt.Rows[Nobox][col].ToString().Split('_');

                            string text = SuffPrefReplacements(Text_statis[0]);
                            text = SuffPrefReplacements(text);
                        text = SuffPrefReplacements(text);
                        chk.Text = text;
                            chk.Tag = "valid";
                            statistic[Nobox] = Convert.ToInt32(Text_statis[1]);
                            times[Nobox] = Convert.ToInt32(Text_statis[2]);
                            staticIndex[Nobox] = Convert.ToInt32(Text_statis[3]);
                            if (Text_statis[4] == "Star")
                               chk.CheckState = CheckState.Checked;
                            
                            panelAuthOptions.Controls.Add(chk);
                            PictureBox picboxedit = new PictureBox();
                            picboxedit.Image = global::PersAhwal.Properties.Resources.edit;
                            picboxedit.Location = new System.Drawing.Point(55, Nobox * 37);
                            picboxedit.Name = Nobox.ToString();
                            picboxedit.Size = new System.Drawing.Size(24, 26);
                            picboxedit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                            picboxedit.TabIndex = 175 + Nobox;
                            picboxedit.TabStop = false;
                            picboxedit.Click += new System.EventHandler(this.pictureBoxedit_Click);
                            panelAuthOptions.Controls.Add(picboxedit);

                            PictureBox picboxup = new PictureBox();
                            picboxup.Image = global::PersAhwal.Properties.Resources.arrowup;
                            picboxup.Location = new System.Drawing.Point(86, Nobox * 37);
                            picboxup.Name = "Up";
                            picboxup.Size = new System.Drawing.Size(24, 26);
                            picboxup.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                            picboxup.TabIndex = 176 + Nobox;
                            picboxup.TabStop = false;
                            picboxup.Click += new System.EventHandler(this.pictureBoxup_Click);
                            if (Nobox == 0)
                            {
                                picboxup.Visible = false;
                            }
                            if (chk.Text.Contains("لمن يشهد والله خير الشاهدين")) picboxup.Visible = false;
                            panelAuthOptions.Controls.Add(picboxup);

                            PictureBox picboxdown = new PictureBox();
                            picboxdown.Image = global::PersAhwal.Properties.Resources.arrowdown;
                            picboxdown.Location = new System.Drawing.Point(55, Nobox * 37);
                            picboxdown.Name = "Down";
                            picboxdown.Size = new System.Drawing.Size(24, 26);
                            picboxdown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                            picboxdown.TabIndex = 177 + Nobox;
                            picboxdown.TabStop = false;
                            picboxdown.Click += new System.EventHandler(this.pictureBoxdown_Click);
                        if (chk.Text.Contains("ويعتبر التوكيل الصادر") || chk.Text.Contains("لمن يشهد والله خير الشاهدين"))
                            picboxdown.Visible = false;
                            panelAuthOptions.Controls.Add(picboxdown);
                            LastID = Convert.ToInt32(checkboxdt.Rows[Nobox]["ID"].ToString());
                            Nobox++;
                        //}
                    }
                }
            }

        }

        public void pictureBoxup_Click(object sender, EventArgs e)
        {


            PictureBox picbox = (PictureBox)sender;

            string st = "", nd = "";
            bool statest = false, statend = false;
            bool FirstCase = false;

            foreach (Control control in panelAuthOptions.Controls)
            {

                if (control is CheckBox)
                {

                    if (!((CheckBox)control).Text.Contains("والله خير الشاهدين"))
                    {
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 176)
                        {
                            st = ((CheckBox)control).Text;
                            if (((CheckBox)control).CheckState == CheckState.Checked) statest = true;
                            else statest = false;

                        }
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 177)
                        {
                            nd = ((CheckBox)control).Text;
                            if (((CheckBox)control).CheckState == CheckState.Checked) statend = true;
                            else statend = false;

                        }
                        FirstCase = true;
                    }
                    else FirstCase = false;

                }
            }
            int x = 0;

            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (!((CheckBox)control).Text.Contains("والله خير الشاهدين"))
                    {
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 176)
                        {
                            ((CheckBox)control).Text = nd;
                            if (statend) ((CheckBox)control).CheckState = CheckState.Checked;
                            else ((CheckBox)control).CheckState = CheckState.Unchecked;
                            int y = 0;

                            y = statistic[x];
                            statistic[x] = statistic[x - 1];
                            statistic[x - 1] = y;

                            y = staticIndex[x];
                            staticIndex[x] = staticIndex[x - 1];
                            staticIndex[x - 1] = y;
                        }
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 177)
                        {
                            ((CheckBox)control).Text = st;
                            if (statest) ((CheckBox)control).CheckState = CheckState.Checked;
                            else ((CheckBox)control).CheckState = CheckState.Unchecked;
                        }
                        x++;
                    }
                }
            }


        }

        public void pictureBoxdown_Click(object sender, EventArgs e)
        {
            PictureBox picbox = (PictureBox)sender;

            string st = "", nd = "";
            bool statest = false, statend = false; bool FirstCase = false;
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (!((CheckBox)control).Text.Contains("والله خير الشاهدين"))
                    {
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 177)
                        {
                            st = ((CheckBox)control).Text;
                            if (((CheckBox)control).CheckState == CheckState.Checked) statest = true;
                            else statest = false;
                        }
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 176)
                        {
                            nd = ((CheckBox)control).Text;
                            if (((CheckBox)control).CheckState == CheckState.Checked) statend = true;
                            else statend = false;
                        }
                        FirstCase = true;
                    }
                    else FirstCase = false;
                }
            }
            int x = 0, y = 0;

            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (!((CheckBox)control).Text.Contains("والله خير الشاهدين"))
                    {
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 177)
                        {
                            ((CheckBox)control).Text = nd;
                            if (statend) ((CheckBox)control).CheckState = CheckState.Checked;
                            else ((CheckBox)control).CheckState = CheckState.Unchecked;
                        }
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 176)
                        {
                            ((CheckBox)control).Text = st;
                            if (statest) ((CheckBox)control).CheckState = CheckState.Checked;
                            else ((CheckBox)control).CheckState = CheckState.Unchecked;
                            y = statistic[x];
                            statistic[x] = statistic[x + 1];
                            statistic[x + 1] = y;
                            y = staticIndex[x];
                            staticIndex[x] = staticIndex[x + 1];
                            staticIndex[x + 1] = y;
                        }
                        x++;
                    }
                }

            }
        }

        private void btnAddRight_Click_1(object sender, EventArgs e)
        {
            if (txtAddRight.Text != "" && btnAddRight.Text == "إضافة")
            {
                CheckBox chk = new CheckBox();
                chk.TabIndex = Nobox;
                chk.Width = 80;
                chk.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                chk.Width = panelAuthOptions.Width - 130;
                chk.Height = 33;
                chk.Tag = "valid";
                chk.CheckState = CheckState.Checked;
                chk.Location = new System.Drawing.Point(60, 3 + Nobox * 37);
                chk.Name = "checkBox" + Nobox.ToString();
                chk.Text = txtAddRight.Text;
                txtAddRight.Clear();
                statistic[Nobox] = 1;
                times[Nobox] = 1;
                panelAuthOptions.Controls.Add(chk);

                PictureBox picboxedit = new PictureBox();
                picboxedit.Image = global::PersAhwal.Properties.Resources.edit;
                picboxedit.Location = new System.Drawing.Point(55, Nobox * 37);
                picboxedit.Name = "Edit";
                picboxedit.Size = new System.Drawing.Size(24, 26);
                picboxedit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                picboxedit.TabIndex = 175 + Nobox;
                picboxedit.TabStop = false;
                picboxedit.Click += new System.EventHandler(this.pictureBoxedit_Click);
                panelAuthOptions.Controls.Add(picboxedit);

                PictureBox picboxup = new PictureBox();
                picboxup.Image = global::PersAhwal.Properties.Resources.arrowup;
                picboxup.Location = new System.Drawing.Point(76, Nobox * 37);
                picboxup.Name = "Up";
                picboxup.Size = new System.Drawing.Size(24, 26);
                picboxup.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                picboxup.TabIndex = 176 + Nobox;
                picboxup.TabStop = false;
                picboxup.Visible = false;
                picboxup.Click += new System.EventHandler(this.pictureBoxup_Click);
                panelAuthOptions.Controls.Add(picboxup);

                PictureBox picboxdown = new PictureBox();
                picboxdown.Image = global::PersAhwal.Properties.Resources.arrowdown;
                picboxdown.Location = new System.Drawing.Point(45, Nobox * 37);
                picboxdown.Size = new System.Drawing.Size(24, 26);
                picboxdown.Name = "Down";
                picboxdown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                picboxdown.TabIndex = 177 + Nobox; ;
                picboxdown.TabStop = false;
                picboxdown.Visible = false;
                picboxdown.Click += new System.EventHandler(this.pictureBoxdown_Click);

                panelAuthOptions.Controls.Add(picboxdown);

                UpdateColumn(ParentForm.PublicDataSource, LastCol, LastID + 1, chk.Text + "_" + statistic[Nobox].ToString() + "_" + times[Nobox].ToString() + "_" + staticIndex[Nobox].ToString() + "_Off", true);
                Nobox++;
                for (int swap = 0; swap < 2; swap++)

                {
                    SwapText(Nobox - swap);
                    ShowArrows(Nobox, swap);
                }

            }
            else if (txtAddRight.Text != "" && btnAddRight.Text == "تعديل")
            {
                foreach (Control control in panelAuthOptions.Controls)
                {
                    if (control is CheckBox)
                    {
                        if (((CheckBox)control).TabIndex == LastTabIndex)
                        {
                            ((CheckBox)control).Text = txtAddRight.Text;
                            btnAddRight.Text = "إضافة";
                            txtAddRight.Text = "";
                        }
                    }
                }
            }
        }

        private void ShowArrows(int tabindex, int indexMinus)
        {
            foreach (Control control in panelAuthOptions.Controls)
            {

                if (control is PictureBox)
                {

                    if (((PictureBox)control).Name == "Down" && ((PictureBox)control).TabIndex == 177 + tabindex - 3)
                    {
                        ((PictureBox)control).Visible = true;
                    }
                    if (((PictureBox)control).Name == "Up" && ((PictureBox)control).TabIndex == 176 + tabindex - 2 - indexMinus)
                    {
                        ((PictureBox)control).Visible = true;
                    }
                }
            }
        }

        private void SwapText(int tabindex)
        {
            string st = "", nd = "";
            bool statest = false, statend = false;


            foreach (Control control in panelAuthOptions.Controls)
            {

                if (control is CheckBox)
                {

                    if (((CheckBox)control).TabIndex == tabindex - 1)
                    {
                        st = ((CheckBox)control).Text;
                        if (((CheckBox)control).CheckState == CheckState.Checked) statest = true;
                        else statest = false;

                    }
                    if (((CheckBox)control).TabIndex == tabindex - 2)
                    {
                        nd = ((CheckBox)control).Text;
                        if (((CheckBox)control).CheckState == CheckState.Checked) statend = true;
                        else statend = false;

                    }
                }
            }
            int x = 0;

            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (((CheckBox)control).TabIndex == tabindex - 1)
                    {
                        ((CheckBox)control).Text = nd;
                        if (statend) ((CheckBox)control).CheckState = CheckState.Checked;
                        else ((CheckBox)control).CheckState = CheckState.Unchecked;
                        int y = 0;

                        y = statistic[x];
                        statistic[x] = statistic[x - 1];
                        statistic[x - 1] = y;

                        y = staticIndex[x];
                        staticIndex[x] = staticIndex[x - 1];
                        staticIndex[x - 1] = y;
                    }
                    if (((CheckBox)control).TabIndex == tabindex - 2)
                    {
                        ((CheckBox)control).Text = st;
                        if (statest) ((CheckBox)control).CheckState = CheckState.Checked;
                        else ((CheckBox)control).CheckState = CheckState.Unchecked;
                    }
                    x++;
                }

            }
        }

        public void deleteItemsAO()
        {
            checkboxdt = new DataTable();
            checkboxdt.Clear();
            Nobox = 0;
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    ((CheckBox)control).Visible = false;
                    ((CheckBox)control).CheckState = CheckState.Unchecked;
                    ((CheckBox)control).Tag = "dispoase";
                }

                if (control is PictureBox)
                {
                    ((PictureBox)control).Visible = false;
                }
            }
        }

        private void StrDivorcePurpose(string text)
        {
            StrSpecPur = text;
        }

        string StrSpecPur = "";

        private string textValue2()
        {
            string text2;
            switch (CombAuthType.Text)
            {
                case "قطعة أرض سكنية":

                    text2 = " القطعة:";
                    break;
                case "ساقية":

                    text2 = " الساقية الواقع بالقطعة:";
                    break;
                case "قطعة أرض حيازة":

                    text2 = " القطعة:";
                    break;
                case "عقار":

                    text2 = " العقار:";
                    break;
                default:
                    text2 = "";
                    break;
            }
            return text2;
        }

        private string textValue()
        {
            string text = "";
            switch (CombAuthType.Text)
            {
                case "قطعة أرض سكنية":
                    text = "قطعة الأرض السكنية";

                    break;
                case "ساقية":
                    text = "ساقية";

                    break;
                case "قطعة أرض حيازة":
                    text = "حيازة";

                    break;
                case "عقار":
                    text = "عقار";

                    break;
                default:
                    text = "";
                    break;
            }
            return text;
        }

        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName)
        {
            combbox.Visible = true;
            label33.Visible = true;
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;


                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                foreach (DataRow dataRow in table.Rows)
                {                    
                    if(dataRow[comlumnName].ToString() != "" && dataRow[comlumnName].ToString() != "null") combbox.Items.Add(dataRow[comlumnName].ToString().Trim());
                }
                saConn.Close();
            }
        }

        private void fillCheckBox(ComboBox box, string source, string comlumnName, string tableName)
        {
            box.Visible = true;
            box.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();
                string query = "select " + comlumnName + " from " + tableName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                int x = 0;
                for (x = 0; x < 10; x++) dataset[x] = "";
                x = 0;
                foreach (DataRow dataRow in table.Rows)
                {
                    if (String.IsNullOrEmpty(dataRow[comlumnName].ToString())) return; 
                    dataset[x] = dataRow[comlumnName].ToString();                    
                    box.Items.Add(dataset[x]);
                    x++;
                }
                saConn.Close();
            }
        }

        private bool checkColumnName(string colNo)
        {
            SqlConnection sqlCon = new SqlConnection(ParentForm.PublicDataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableListCombo", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(dataRow["COLUMN_NAME"].ToString()))
                {
                    if (dataRow["COLUMN_NAME"].ToString() == colNo)
                        return true;
                }
            }
            return false;
        }

        private void CombAuthType_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (checkColumnName(CombAuthType.Text.Replace(" ", "_")))
            {
                ComboProcedure.Items.Clear();
                newFillComboBox1(ComboProcedure, ParentForm.PublicDataSource, CombAuthType.SelectedIndex.ToString(), "العربية");
                if (CombAuthType.SelectedIndex == 6)
                {
                    label7.Visible = comboPropertyType.Visible = true;
                    LegaceyPreStr = " في التركة المذكورة أعلاه";
                }
                else
                {
                    PanelSubItemBox.Visible = label7.Visible = comboPropertyType.Visible = false;
                    LegaceyPreStr = "";
                }
                 
                if (CombAuthType.SelectedIndex != 16) 
                    fileComboBox(ParentForm.txtAttendVCValue, ParentForm.PublicDataSource, "ArabicAttendVC", "TableListCombo");            
                else fileComboBox(ParentForm.txtAttendVCValue, ParentForm.PublicDataSource, "EnglishAttendVC", "TableListCombo");
                return;
            }
            if (ComboProcedure.Items.Count > 0) ComboProcedure.Items.Clear();
            restShowingItems();
            SepratBoxes();
            LegaceyPreStr = "";
            if (CombAuthType.SelectedIndex == 0)
            {
                txtReviewValue.Location = new System.Drawing.Point(336, 76);
                txtReviewValue.Size = new System.Drawing.Size(828, 171);
                txtAddRightValue.Location = new System.Drawing.Point(336, 253);
                txtAddRightValue.Size = new System.Drawing.Size(828, 409);
                for (int j = 0; j < 5; j++) generalAuthText = SuffPrefReplacements(generalAuthText);
                txtAddRight.Text = generalAuthText;
                label34Value.Visible = false;
                label36Value.Text = "إضافة نصوص التوكيل:";
                ParentForm.checkBox1Value.CheckState = CheckState.Unchecked;
            }else ParentForm.checkBox1Value.CheckState = CheckState.Checked;

            if (CombAuthType.SelectedIndex >= 1 && CombAuthType.SelectedIndex <= 4)
            {
                fileComboBox(ComboProcedure, ParentForm.PublicDataSource, "Row1Attach", "TableListCombo");
            }
            if (CombAuthType.SelectedIndex == 5)
            {
                fileComboBox(ComboProcedure, ParentForm.PublicDataSource, "Row1Attach", "TableListCombo");
            }
            if (CombAuthType.Text.Contains("زواج"))
            {
                DPTitle = "آنسة";
                fileComboBox(ComboProcedure, ParentForm.PublicDataSource, "RowMerrageAttach", "TableListCombo");
            }
            if (CombAuthType.Text.Contains("ورثة"))
            {
                DPTitle = "المرحوم";
                fileComboBox(ComboProcedure, ParentForm.PublicDataSource, "RowLegacyAttach", "TableListCombo");
            }
            if (CombAuthType.Text.Contains("سيارة"))
            {
                SepratBoxes();
                fileComboBox(ComboProcedure, ParentForm.PublicDataSource, "RowCarAttach", "TableListCombo");
            }
            if (CombAuthType.Text.Contains("طلاق"))
            {
                SepratBoxes();
                fileComboBox(ComboProcedure, ParentForm.PublicDataSource, "RowDeforceAttach", "TableListCombo");
            }
            if (CombAuthType.Text.Contains("جامعية"))
            {
               
                SepratBoxes();
                fillCheckBox(ComboProcedure, ParentForm.PublicDataSource, "RowUniversityAttach", "TableListCombo");
            }
            if (CombAuthType.Text.Contains("ميلاد"))
            {
                SepratBoxes();
                ComboProcedure.Items.Add("استخراج وتوثيق");
                ComboProcedure.Items.Add(" استخراج وتوثيق بدل فاقد");
            }
            if (CombAuthType.Text.Contains("بالتنازل"))
            {
                SepratBoxes();
                fillCheckBox(ComboProcedure, ParentForm.PublicDataSource, "GiveAway", "TableListCombo");
            }
            if (CombAuthType.Text.Contains("حساب بنكي"))
            {
                SepratBoxes();
                fillCheckBox(ComboProcedure, ParentForm.PublicDataSource, "BankAccount", "TableListCombo");
            }

            if (CombAuthType.Text.Contains("تأمين"))
            {
                SepratBoxes();
                ComboProcedure.Items.Add("استلام تأمين");
            }

        }

        public void LegaceyBox(string v1, string v2, string v3, string v4, string v5, string v6, string v7, string v8, string v9)
        {
            
            if (v1 != "")
            {
                Leglab1.Text = v1;
                Leglab1.Visible = true;
                LegtextBox1.Visible = true;
            }
            if (v2 != "")
            {
                Leglab2.Text = v2;
                Leglab2.Visible = true;
                LegtextBox2.Visible = true;
            }
            if (v3 != "")
            {
                Leglab3.Text = v3;
                Leglab3.Visible = true;
                LegtextBox3.Visible = true;
            }
            if (v4 != "")
            {
                Leglab4.Text = v4;
                Leglab4.Visible = true;
                LegtextBox4.Visible = true;
            }
            if (v5 != "")
            {
                Leglab5.Text = v5;
                Leglab5.Visible = true;
                LegtextBox5.Visible = true;
            }
            if (v6 != "")
            {
                Leglab6.Text = v6;
                Leglab6.Visible = true;
                LegtextBox6.Visible = true;
            }
            if (v7 != "")
            {
                Leglab7.Text = v7;
                Leglab7.Visible = true;
                LegtextBox7.Visible = true;
            }
            if (v8 != "")
            {
                Leglab8.Text = v8;
                Leglab8.Visible = true;
                LegtextBox8.Visible = true;
            }
            if (v9 != "")
            {
                Leglabel9.Text = v9;
                Leglabel9.Visible = true;
                LegtxtBoxGeneral.Visible = true;
            }


        }
        //private void DetermineCheckBox(string v1, int s1, string v2, int s2, string v3, int s3, string v4, int s4, string v5, int s5, string v6, string v7, string v81, string[] v8, string button, string v91,string[] v9)
        //{
        //    restShowingItems();
        //    if (v1 != "")
        //    {
        //        labelVitext1.Text = v1;
        //        labeltxt1.Visible = true;
        //        txt1.Width = s1;
        //        txt1.Visible = true;
        //    }
        //    if (v2 != "")
        //    {
        //        labelVitext2.Text = v2;
        //        labeltxt2.Visible = true;
        //        txt2.Width = s2;
        //        txt2.Visible = true;
        //    }
        //    if (v3 != "")
        //    {
        //        labelVitext3.Text = v3;
        //        labeltxt3.Visible = true;
        //        txt3.Width = s3;
        //        txt3.Visible = true;
        //    }
        //    if (v4 != "")
        //    {
        //        labelVitext4.Text = v4;
        //        labeltxt4.Visible = true;
        //        txt4.Width = s4;
        //        txt4.Visible = true;
        //    }
        //    if (v5 != "")
        //    {
        //        labelVitext5.Text = v5;
        //        labeltxt5.Visible = true;
        //        txt5.Width = s5;
        //        txt5.Visible = true;
        //    }


        //    if (v6 != "")
        //    {
        //        DPTitle = dataGrid[12];
        //        labelVicheck1.Text = v6;
        //        labelVicheck1.Visible = true;
        //        if(DPTitle.Contains("_")) Vicheck1.Text = DPTitle.Split('_')[0];
        //        else Vicheck1.Text = DPTitle;
        //        Vicheck1.Visible = true;
        //    }

        //    if (v7 != "")
        //    {
        //        labeldate1.Visible = true;
        //        lblD1.Visible = true; 
        //        VitxtDate1VD.Visible = true; 
        //        lalM1.Visible = true; 
        //        VitxtDate1VM.Visible = true; 
        //        lalY1.Visible = true; 
        //        VitxtDate1VY.Visible = true;
        //        labeldate1.Text = v7;
        //    }


        //    if (v81 != "")
        //    {
        //        labelcomb1.Visible = true;
        //        Vicombo1.Visible = true;
        //        labelcomb1.Text = v81;

        //        Vicombo1.Items.Clear();
        //        for (int x = 0; x < v8.Length; x++)
        //            Vicombo1.Items.Add(v8[x]);
        //        Vicombo1.SelectedIndex = 0;
        //    }
        //    if (button != "")
        //    {
        //        VibtnAdd1.Text = button;
        //        VibtnAdd1.Visible = true; 
        //    }
        //    if (v91 != "")
        //    {
        //        labelcomb12.Visible = true;
        //        Vicombo2.Visible = true;
        //        labelcomb12.Text = v91;

        //        Vicombo2.Items.Clear();
        //        for (int x = 0; x < v9.Length; x++)
        //            Vicombo2.Items.Add(v9[x]);
        //        Vicombo2.SelectedIndex = 0;
        //    }
        //}

        private void restShowingItems()
        {
            foreach (Control control in PanelItemsboxes.Controls)
            {
                control.Visible = false;
            }            
        }

        private void SepratBoxes()
        {
            txtReview.Location = new System.Drawing.Point(338, 162);
            txtReview.Height = 80;
            PanelItemsboxes.Visible = true;
            label8.Location = new System.Drawing.Point(1172, 76);
        }

        private void radioSelectAll_CheckedChanged_1(object sender, EventArgs e)
        {
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox && ((CheckBox)control).Tag.ToString() == "valid")
                {
                    ((CheckBox)control).CheckState = CheckState.Checked;
                }
            }
        }

        private void radiounSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    ((CheckBox)control).CheckState = CheckState.Unchecked;
                }
            }
        }

        private void CombAuthType_MouseClick(object sender, MouseEventArgs e)
        {



        }

        private void comboPropertyType_SelectedIndexChanged_1(object sender, EventArgs e)
        {

            if (comboPropertyType.Text.Contains("عقار"))
            {
                PanelSubItemBox.Visible = true;
                //MessageBox.Show(comboPropertyType.Text);
                LegaceyBox("رقم العقار", "رقم المربع:", "المساحة:", "الحي:", "المدينة:", "", "", "", "");
            }

            else if (comboPropertyType.Text.Contains("مركبة"))
            {
                PanelSubItemBox.Visible = true;
                //MessageBox.Show(comboPropertyType.Text);
                LegaceyBox("نوع السيارة", "للون:", "رقم اللوحة:", "رقم الشاسية:", "سنة الموديل", "", "", "", "");
            }
            else if (comboPropertyType.Text.Contains("أخرى"))
            {
                PanelSubItemBox.Visible = true;
                //MessageBox.Show(comboPropertyType.Text);
                LegaceyBox("", "", "", "", "", "", "", "", "وصف الورثة");
            }
            else
            {
                PanelSubItemBox.Visible = false;
                if (CombAuthType.SelectedIndex == 6)
                    LegaceyPreStr = " في التركة المذكورة أعلاه";
                else LegaceyPreStr = "";
            }
            }
        int LegaceyIndex = 0;
        string LegaceyItem = "";
        string LegaceyPreStr = "";
        

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ComboProcedure_TextChanged_1(object sender, EventArgs e)
        {
            //ComboProcedure_Text();
            
        }

        public void ComboProcedure_Text()
        {
            deleteItemsAO();
            ShowNewApp = false;
            label7.Visible = false;
            if(comboPropertyType.Items.Count > 0) comboPropertyType.SelectedIndex = 0;
            txtReview.Text = "";
            foreach (Control control in PanelItemsboxes.Controls)
            {
                control.Visible = false;
                control.Text = "";
                if (control is ComboBox)
                {
                    ((ComboBox)control).Items.Clear();
                }
                else if (control is CheckBox) ((CheckBox)control).CheckState = CheckState.Unchecked;
            }
            
            strRights = "";
            ColName = "Col0";
            CreateBoxesWithData(ComboProcedure.Text.Trim(), "", false);            
            PopulateCheckBoxes(ColName, "TableAuthRight", ParentForm.PublicDataSource);
        }

        private void addName_Click(object sender, EventArgs e)
        {            
            BirthName[birthindex] = Vitext1.Text;
            BirthPlace[birthindex] = Vitext2.Text;
            BirthDate[birthindex] = Vitext3.Text;
            BirthMother[birthindex] = Vitext4.Text;
            
            Vitext1.Text = Vitext2.Text = Vitext3.Text = Vitext4.Text = "";
            if (birthindex == 0) specialDataSum = BirthName[birthindex] + "_" + BirthPlace[birthindex] + "_" + BirthDate[birthindex] + "_" + BirthMother[birthindex] + "_" + BirthDecs[birthindex];
            else specialDataSum = specialDataSum  + "*" +BirthName[birthindex] + "_" + BirthPlace[birthindex] + "_" + BirthDate[birthindex] + "_" + BirthMother[birthindex] + "_" + BirthDecs[birthindex];
            
            if (birthindex == 0 && Vicombo2.SelectedIndex > 0)
            {
                
                if (Vicombo2.SelectedIndex == 1)
                {

                    Mentioned = "لابني";
                    
                }
                else if (Vicombo2.SelectedIndex == 2)
                {

                    Mentioned = "لابنتي";
                    
                }
            }
            else if (birthindex == 1 && Vicombo2.SelectedIndex > 0)
            {
                if (Vicombo2.SelectedIndex == 1 && Mentioned == "لابني")
                {
                    Mentioned = "لابنيَّ";
                    
                }
                else if (Vicombo2.SelectedIndex == 2 && Mentioned == "لابنتي")
                {
                    Mentioned = "لابنتيَّ";
                    
                }
                else
                {
                    Mentioned = "لابنائي";
                    
                }
                
            }
            else if (birthindex >= 2 && Vicombo2.SelectedIndex > 0)
            {
                if (Vicombo2.SelectedIndex == 2 && Mentioned == "لابنتيَّ")
                {
                    Mentioned = "لبناتي";
                }
                else
                {
                    Mentioned = "لأبنائي";
                    
                }

            }

            BirthDecs[birthindex] = Mentioned;
            
            birthindex++;
            idShow = birthindex;
            LibtnAdd1.Text = "اضافة (" + idShow.ToString() + "/" + birthindex.ToString() + ")" + "   ";
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            
            if (idShow < birthindex )
            {
              

                if (Vitext1.Text != "")
                {                    
                    string[] strLines = specialDataSum.Split('*');
                    strLines[idShow] = Vitext1.Text + "_" + Vitext2.Text + "_" + Vitext3.Text + "_" + Vitext4.Text;
                    specialDataSum = strLines[0];
                    for (int x = 1; x < birthindex; x++) specialDataSum = specialDataSum + "*" + strLines[x];
                }
                idShow++;
                Vitext1.Text = BirthName[idShow];
                Vitext2.Text = BirthPlace[idShow];
                Vitext3.Text = BirthDate[idShow];
                Vitext4.Text = BirthMother[idShow];
                LibtnAdd1.Text = "اضافة (" + idShow.ToString() + "/" + birthindex.ToString() + ")";
            }
            else
            {
                Vitext1.Text = Vitext2.Text =Vitext3.Text =Vitext4.Text = "";
            }

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (idShow > 0)
            {

                if (Vitext1.Text != "")
                {
                    string[] strLines = specialDataSum.Split('*');
                    strLines[idShow] = Vitext1.Text + "_" + Vitext2.Text + "_" + Vitext3.Text + "_" + Vitext4.Text;
                    specialDataSum = strLines[0];
                    for (int x = 1; x < birthindex; x++)
                    {
                        specialDataSum = specialDataSum + "*" + strLines[x];
                        
                    }
                    
                }
                idShow--;
                Vitext1.Text = BirthName[idShow];
                Vitext2.Text = BirthPlace[idShow];
                Vitext3.Text = BirthDate[idShow];
                Vitext4.Text = BirthMother[idShow];

                LibtnAdd1.Text = "اضافة (" + idShow.ToString() + "/" + birthindex.ToString() + ")";
            }
        }

        private void ComboProcedure_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboProcedure_Text();
        }

        private void datasumFamily(string FamilyMemberList, int idlist, int totalNo)
        {
            string[] memberdata = FamilyMemberList.Split('/');

            BirthNameValue[idlist] = memberdata[0];
            BirthPlaceValue[idlist] = memberdata[1];
            BirthDateValue[idlist] = memberdata[2];
            BirthMotherValue[idlist] = memberdata[3];
            

        }

        

        private void txtReview_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtReview_MouseHover(object sender, EventArgs e)
        {
            if (ParentForm.checkBox1Value.CheckState == CheckState.Checked && !ShowNewApp)
            {
                string other = "الآنسة";
                if (ComboProcedure.Text.Trim() == "عقد قران شخصي")
                {
                    StrSpecPur = " في عقد قراني على " + Vicheck1.Text + " /" + Vitext1.Text;

                }
                else if (ComboProcedure.Text.Trim() == "عقد قران غير شخصي")
                {
                    StrSpecPur = " في عقد قران  " + Vicombo2.Text + " " + Vitext2.Text + " " + Vicheck1.Text + " /" + Vitext1.Text + " على السيد " + Vicombo1.Text + "/ " + Vitext3.Text;

                }
                else if (ComboProcedure.Text.Trim().Contains("وثيقة تصادق"))
                {
                    other = "زوجتي/ ";
                    if (ParentForm.strAppMaleFemaleList[0] == "أنثى") other = "زوجي/ ";
                    StrSpecPur = " في إستخراج وثيقة تصادق على زواجي من " + other + Vitext1.Text + " بتاريخ: " + Vitext2.Text + "/" + Vitext3.Text + "/" + Vitext4.Text;

                }
                else if (ComboProcedure.Text.Trim().Contains("قسيمة زواج"))
                {
                    other = "زوجتي/ ";
                    if (ParentForm.strAppMaleFemaleList[0] == "أنثى") other = "زوجي/ ";
                    StrSpecPur = " في إستخراج قسيمة زواجي من " + other + Vitext1.Text + " بتاريخ: " + Vitext2.Text + "/" + Vitext3.Text + "/" + Vitext4.Text;
                }
                string DivNo = "زوجتي";
                //

                if (CombAuthType.Text == "طلاق")
                {

                    if (ParentForm.strAppMaleFemaleList[0] != "ذكر")
                    {
                        if (Vicombo1.Text == "طلقة ثالثة") DivNo = "مطلقتي "; else DivNo = "زوجتي ";
                    }
                    else
                    {
                        if (Vicombo1.Text == "طلقة ثالثة") DivNo = "مطلقي "; else DivNo = "زوجي ";
                    }
                    if (ComboProcedure.SelectedIndex == 0)
                    {
                        if (DivNo == "زوجتي " || DivNo == "مطلقتي ")
                            StrSpecPur = " إيقاع " + Vicombo1.Text + " على " + DivNo + " السيدة/ " + Vitext1.Text;
                    }
                    else if (ComboProcedure.SelectedIndex == 1)
                        StrSpecPur = " إستخراج قسيمة طلاقي من " + DivNo + Vitext1.Text + " التي أوقعت عليها " + Vicombo1.Text + " بتاريخ: " + Vitext2.Text + "/" + Vitext3.Text + "/" + Vitext4.Text;
                }

                rightModify("@*@");

                if (!ShowNewApp)
                {
                    
                    CreateAuthList2(StrSpecPur);
                    authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + LegaceyPreStr + authList2;
                }
            }
            else {
                txtReviewBody();                
        }
        }

        private void txtReviewBody()
        {
            txtReview.Text = StrSpecPur + LegaceyPreStr;
            for (int x = 0; x < 40; x++)
                txtReview.Text = SuffPrefReplacements(txtReview.Text);
            authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + txtReview.Text;
        }

        private void btnAddLegacey_Click_1(object sender, EventArgs e)
        {
            string str = "";
            if (comboPropertyType.Text.Contains("مركبة"))
            {

                if (LegtextBox1.Text != "") str = "في السيارة من نوع " + LegtextBox1.Text;
                if (LegtextBox5.Text != "") str = str + " موديل العام (" + LegtextBox5.Text + ")";
                if (LegtextBox2.Text != "") str = str + "باللون " + LegtextBox2.Text;
                if (LegtextBox3.Text != "") str = str + " ورقم لوحة (" + LegtextBox3.Text + " )";
                if (LegtextBox4.Text != "") str = str + "وشاسيه بالرقم (" + LegtextBox4.Text + ") ";
                if (LegaceyIndex == 0 || LegaceyPreStr == " في التركة المذكورة أعلاه")
                {
                    LegaceyIndex = 0;
                    LegaceyPreStr = str;
                }
                else LegaceyPreStr = LegaceyPreStr + " و " + str;
            }
            else if (comboPropertyType.Text.Contains("عقار"))
            {

                if (LegtextBox1.Text != "") str = "في العقار بالرقم (" + LegtextBox1.Text;
                if (LegtextBox2.Text != "") str = str + ") بمربع رقم (" + LegtextBox2.Text + ")";
                if (LegtextBox3.Text != "") str = str + ") البالغ مساحتها(" + LegtextBox3.Text + "م.م)";
                if (LegtextBox4.Text != "") str = str + " ب" + LegtextBox4.Text + "-" + LegtextBox5.Text + " )";
                if (LegaceyIndex == 0 || LegaceyPreStr == " في التركة المذكورة أعلاه")
                {
                    LegaceyIndex = 0;
                    LegaceyPreStr = str;
                }
                else LegaceyPreStr = LegaceyPreStr + " و " + str;
            }
            else if (comboPropertyType.Text.Contains("أخرى"))
            {
                if (LegaceyIndex == 0 || LegaceyPreStr == " في التركة المذكورة أعلاه")
                {
                    LegaceyIndex = 0;
                    LegaceyPreStr = " في " + LegtxtBoxGeneral.Text;
                }                
                else LegaceyPreStr = LegaceyPreStr + " وفي " + LegtxtBoxGeneral.Text;
            }
            else {
                LegaceyPreStr = " في التركة المذكورة أعلاه";
            }
            LegtextBox1.Text = LegtextBox2.Text = LegtextBox3.Text = LegtextBox4.Text = LegtextBox5.Text = LegtxtBoxGeneral.Text = "";
            LegaceyIndex++;
            //txtReviewBody();
        }

        private void rightModify(string text)
        {
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (((CheckBox)control).Text.Contains(text))
                    {
                        ((CheckBox)control).Text = SuffPrefReplacements(((CheckBox)control).Text);
                    }
                }
            }
        }

        private void removedDocNo_TextChanged(object sender, EventArgs e)
        {            
            loadInfo(removedDocNo.Text);
            removeText();
        }

        private void removeText() {
            string orgText = SuffPrefReplacements("ويعتبر التوكيل الصادر من #6 بتاريخ #7 بالرقم #8 لاغٍ،");
            orgText = SuffPrefReplacements(orgText);
            orgText = SuffPrefReplacements(orgText);
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (((CheckBox)control).Text.Contains("ويعتبر التوكيل الصادر"))
                    {
                        ((CheckBox)control).Text = orgText;
                        ((CheckBox)control).CheckState = CheckState.Checked;
                    }
                }
            }
        }
        private string loadInfo(string documenNo)
        {
            SqlConnection sqlCon = new SqlConnection(ParentForm.PublicDataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT التاريخ_الميلادي from TableAuth where رقم_التوكيل=@رقم_التوكيل", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_التوكيل", documenNo);

            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string rowCnt = "";

            foreach (DataRow row in dtbl.Rows)
            {
                if (row["التاريخ_الميلادي"].ToString() != "")
                {
                    removedDocDate.Text = row["التاريخ_الميلادي"].ToString().Replace("-","/");
                    removedDocSource.Text = "القنصلية العامة لجمهورية السودان بجدة ";
                }
            }
            return rowCnt;

        }

        private void removedDocSource_TextChanged(object sender, EventArgs e)
        {
            removeText();
        }

        private void removedDocDate_TextChanged(object sender, EventArgs e)
        {
            removeText();
        }

        private void checkSexType_CheckedChanged_1(object sender, EventArgs e)
        {
            if (!ShowNewApp) {
                if (CombAuthType.Text.Contains("ورثة"))
                {
                    if (Vicheck1.CheckState == CheckState.Unchecked) { DPTitle = "المرحوم"; Vicheck1.Text = "ذكر"; }
                    else { DPTitle = "المرحومة"; Vicheck1.Text = "أنثى"; }
                }
                else if (CombAuthType.Text.Contains("زواج"))
                {
                    if (Vicheck1.CheckState == CheckState.Unchecked) { Vicheck1.Text = "الآنسة"; }
                    else
                    {
                        DPTitle = "السيدة";
                        Vicheck1.Text = "السيدة";
                    }
                }
            }
            else {
                if (!dataGrid[12].Contains("_")) return;
                if (Vicheck1.CheckState == CheckState.Unchecked) { Vicheck1.Text = dataGrid[12].Split('_')[0]; }
                else
                {
                    Vicheck1.Text = dataGrid[12].Split('_')[1];
                }
            
            }
        }

        private void Vicheck_CheckedChanged(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 1)
                for (int index = 0; index < dataGridView1.Rows.Count - 1; index++)
                    if (idIndex == Convert.ToInt32(dataGridView1.Rows[index].Cells[0].Value.ToString()))
                    {
                        CheckBox checkBox = (CheckBox)sender;
                        string optionscheck = dataGridView1.Rows[index].Cells[checkBox.Name.Replace("Vi", "options")].Value.ToString();
                        if (optionscheck.Contains("_"))
                        {
                            if (checkBox.Checked)
                                checkBox.Text = optionscheck.Split('_')[1];
                            else checkBox.Text = optionscheck.Split('_')[0];
                        }
                    }
        }

        private void LibtnAdd1_Click(object sender, EventArgs e)
        {

        }

        private void txtReviewBody_TextChanged(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;
            NewAuthSubject = txt.Text;
        }

        private bool ShowRowNo()
        {
            //label1,lenght1,label2,lenght2,label3,lenght3,label4,lenght4,label5,lenght5,
            //labelcheck,optionscheck,12
            //labelcomb1,optionscombo1,lenghtscombo1,labelcomb2,optionsVicombo2,lenghtsVicombo2,13
            //labelbtn,lenghtsbtn,19
            //dateYN,dateType,TextModel,ColRight,ColName 21

            SqlConnection sqlCon = new SqlConnection(ParentForm.PublicDataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("ContextViewSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ColName", "");
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            //dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            sqlCon.Close();
            //MessageBox.Show(ComboProcedure.Text + "-" + CombAuthType.SelectedIndex.ToString());
            //MessageBox.Show(ComboProcedure.Text + "-" + CombAuthType.SelectedIndex.ToString());
            for (int id = 0; id < dataGridView1.Rows.Count-1; id++)
            {
                
                if (dataGridView1.Rows[id].Cells[25].Value.ToString() == ComboProcedure.Text + "-" + CombAuthType.SelectedIndex.ToString())
                {
                    for (int col = 0; col < 26; col++)
                    {
                        dataGrid[col] = dataGridView1.Rows[id].Cells[col].Value.ToString();
                    }
                    
                    if (dataGrid[21] == "لا") dataGrid[21] = "";
                    ColName = dataGridView1.Rows[id].Cells[24].Value.ToString();
                    StrSpecPur = dataGridView1.Rows[id].Cells[23].Value.ToString();
                    
                    int s1 = 50;
                    if (dataGrid[2] != "") s1 = Convert.ToInt32(dataGrid[2]);
                    int s2 = 50;
                    if (dataGrid[4] != "") s2 = Convert.ToInt32(dataGrid[4]);
                    int s3 = 50;
                    if (dataGrid[6] != "") s3 = Convert.ToInt32(dataGrid[6]);
                    int s4 = 50;
                    if (dataGrid[8] != "") s4 = Convert.ToInt32(dataGrid[8]);
                    int s5 = 50;
                    if (dataGrid[10] != "") s5 = Convert.ToInt32(dataGrid[10]);
                    idIndex = Convert.ToInt32(dataGridView1.Rows[id].Cells[0].Value.ToString());
                    flllPanelItemsboxes(idIndex);
                    return true;
                }
            }
            return false;
        }


        private void flllPanelItemsboxes(int idIndex)
        {
            if (dataGridView1.Rows.Count > 1)
            {
                for (int index = 0; index < dataGridView1.Rows.Count - 1; index++)
                    if (idIndex == Convert.ToInt32(dataGridView1.Rows[index].Cells[0].Value.ToString()))
                    {
                        foreach (Control Lcontrol in PanelItemsboxes.Controls)
                            try
                            {
                                if (Lcontrol.Name.StartsWith("L"))
                                {
                                    Lcontrol.Text = dataGridView1.Rows[index].Cells[Lcontrol.Name.Replace("L", "")].Value.ToString();
                                    if (Lcontrol.Text != "")
                                    {
                                        Lcontrol.Visible = true;
                                        foreach (Control Vcontrol in PanelItemsboxes.Controls)
                                        {
                                            //MessageBox.Show("View = " + Vcontrol.Name + " - Label=" + Lcontrol.Name.Replace("L", "V"));
                                            if (Vcontrol.Name.Trim() == Lcontrol.Name.Replace("L", "V").Trim())
                                            {
                                                //MessageBox.Show(Lcontrol.Name + "Length");
                                                Vcontrol.Visible = true;
                                                string size = dataGridView1.Rows[index].Cells[Lcontrol.Name.Replace("L", "") + "Length"].Value.ToString();
                                                Vcontrol.Width = Convert.ToInt32(size);
                                                if (Convert.ToInt32(size) >= 700)
                                                {
                                                    if (Vcontrol is TextBox) ((TextBox)Vcontrol).Multiline = true;
                                                    Vcontrol.Height = 150;
                                                }
                                                //MessageBox.Show(Lcontrol.Name + "Length");
                                            }
                                            if (Vcontrol.Name.Contains(Lcontrol.Name.Replace("L", "V") + "V") || Vcontrol.Name.Contains(Lcontrol.Name.Replace("L", "V") + "L"))
                                            {
                                                Vcontrol.Visible = true;

                                                //MessageBox.Show(Vcontrol.Name + " - " + Lcontrol.Name.Replace("L", "V"));
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(Lcontrol.Name.Replace("L", ""));
                            }

                        for (int x = 1; x < 6; x++)
                        {
                            try
                            {
                                txtCheckOptions[x - 1] = dataGridView1.Rows[index].Cells["optionscheck" + x.ToString()].Value.ToString();
                            }
                            catch (Exception X)
                            {
                            }
                        }
                        if (txtComboOptions[0] != "") Vicheck1.Text = txtComboOptions[0].Split()[0];                     
                        if (txtComboOptions[1] != "") Vicheck2.Text = txtComboOptions[1].Split()[0];                     
                        if (txtComboOptions[2] != "") Vicheck3.Text = txtComboOptions[2].Split()[0];                     
                        if (txtComboOptions[3] != "") Vicheck4.Text = txtComboOptions[3].Split()[0];                     
                        if (txtComboOptions[4] != "") Vicheck5.Text = txtComboOptions[4].Split()[0];                     

                            
                        for (int x = 1; x < 6; x++)
                        {
                            try
                            {
                                txtComboOptions[x - 1] = dataGridView1.Rows[index].Cells["icomboOption" + x.ToString()].Value.ToString();
                            }
                            catch (Exception X)
                            {
                            }
                        }
                        if (txtComboOptions[0] != "")
                        {
                            Vicombo1.Items.Clear();
                            for (int x = 0; x < txtComboOptions[0].Split('_').Length; x++)
                                if(txtComboOptions[0].Split('_')[x] !="") Vicombo1.Items.Add(txtComboOptions[0].Split('_')[x]);
                            Vicombo1.SelectedIndex = 0;
                        }
                        if (txtComboOptions[1] != "")
                        {
                            Vicombo2.Items.Clear();
                            for (int x = 0; x < txtComboOptions[1].Split('_').Length; x++)
                                if (txtComboOptions[1].Split('_')[x] != "") Vicombo2.Items.Add(txtComboOptions[1].Split('_')[x]);
                            Vicombo2.SelectedIndex = 0;
                        }
                        if (txtComboOptions[2] != "")
                        {
                            Vicombo3.Items.Clear();
                            for (int x = 0; x < txtComboOptions[2].Split('_').Length; x++)
                                if (txtComboOptions[2].Split('_')[x] != "") Vicombo3.Items.Add(txtComboOptions[2].Split('_')[x]);
                            Vicombo3.SelectedIndex = 0;
                        }
                        if (txtComboOptions[3] != "")
                        {
                            Vicombo4.Items.Clear();
                            for (int x = 0; x < txtComboOptions[3].Split('_').Length; x++)
                                if (txtComboOptions[3].Split('_')[x] != "") Vicombo4.Items.Add(txtComboOptions[3].Split('_')[x]);
                            Vicombo4.SelectedIndex = 0;
                        }
                        if (txtComboOptions[4] != "")
                        {
                            Vicombo5.Items.Clear();
                            for (int x = 0; x < txtComboOptions[4].Split('_').Length; x++)
                                if (txtComboOptions[4].Split('_')[x] != "") Vicombo5.Items.Add(txtComboOptions[4].Split('_')[x]);
                            Vicombo5.SelectedIndex = 0;
                        }


                    }
            }


        }
        public void CreateBoxesWithData(string textdata, string textItems, bool database)

        {
            int x = 0;
            string[] str6Listcomb = new string[4] { "صفة الزوج", "السيد", "موكلي السيد","" };
            string[] SI = new string[9];
            string text = "";
            pictureBox1.Visible = pictureBox2.Visible = false;
            x = 0;
            
            if (textItems.Contains("_")) SI = textItems.Split('_');
            
            strListcomb1 = new string[10] { "", "", "", "", "", "", "", "", "", "" };

            if (ShowRowNo()) 
            {
                ShowNewApp = true;
                if (database && !ParentForm.ArchData)
                {
                    if (!String.IsNullOrEmpty(textItems))
                    {
                        if (SI.Length == 10)
                            DetermineData(SI[0], SI[1], SI[2], SI[3], SI[4], SI[5], SI[6], SI[7], SI[8], SI[9]);
                        else
                            DetermineData(SI[0], SI[1], SI[2], SI[3], SI[4], SI[5], SI[6]);
                    }
                }
                return;
            }
            
            if (!String.IsNullOrEmpty(textItems))
            {
                if (SI.Length == 10)
                    DetermineData(SI[0], SI[1], SI[2], SI[3], SI[4], SI[5], SI[6], SI[7], SI[8], SI[9]);
                else
                    DetermineData(SI[0], SI[1], SI[2], SI[3], SI[4], SI[5], SI[6]);
            }
            //if (ParentForm.ArchData) { ComboProcedure_Text(); }

        }

        private void newFillComboBox1(ComboBox combbox, string source, string id, string Language)
        {
            combbox.Visible = true;
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();
                string query = "select ColName,ColRight,Lang from TableAddContext";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                foreach (DataRow dataRow in table.Rows)
                {
                    if (dataRow["Lang"].ToString() == Language && dataRow["ColRight"].ToString() != "" && !String.IsNullOrEmpty(dataRow["ColName"].ToString()) && dataRow["ColName"].ToString().Contains("-"))
                    {
                        if (dataRow["ColName"].ToString().Split('-')[1].All(char.IsDigit))
                        {
                            try
                            {
                                if (id == dataRow["ColName"].ToString().Split('-')[1])
                                {
                                    combbox.Items.Add(dataRow["ColName"].ToString().Split('-')[0]);
                                }
                            }
                            catch (Exception exp)
                            {

                            }
                        }
                    }
                }
                saConn.Close();
            }
            if (combbox.Items.Count > 0) combbox.SelectedIndex = 0;
        }

        public void DetermineData(string v1, string v2, string v3, string v4, string v5, string v6, string v7)
        {
            Vitext1.Text = v1;
            Vitext2.Text = v2;
            Vitext3.Text = v3;
            Vitext4.Text = v4;
            Vitext5.Text = v5;
            Vicheck1.Text = v6;
            if (v7.Contains("-"))
            {
                VitxtDate1VD.Text = v7.Split('-')[0];
                VitxtDate1VM.Text = v7.Split('-')[1];
                VitxtDate1VY.Text = v7.Split('-')[2];
            }
            Vicombo1.Text = v7;
            
        }

        public void DetermineData(string v1, string v2, string v3, string v4, string v5, string v6, string v7, string v8, string v9, string v10)
        {
            Vitext1.Text = v1;
            Vitext2.Text = v2;
            Vitext3.Text = v3;
            Vitext4.Text = v4;
            Vitext5.Text = v5;
            Vicheck1.Text = v6;
            if (v7.Contains("-"))
            {
                VitxtDate1VD.Text = v7.Split('-')[0];
                VitxtDate1VM.Text = v7.Split('-')[1];
                VitxtDate1VY.Text = v7.Split('-')[2];
            }
            Vicombo1.Text = v8;
            Vicombo1.Text = v10;
            LibtnAdd1.Text = v9;
            //MessageBox.Show(v1);
        }

        public void pictureBoxedit_Click(object sender, EventArgs e)
        {
            PictureBox picbox = (PictureBox)sender;
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (((CheckBox)control).TabIndex == Convert.ToInt32(picbox.Name))
                    {
                        txtAddRight.Text = ((CheckBox)control).Text;
                        btnAddRight.Text = "تعديل";
                        LastTabIndex = ((CheckBox)control).TabIndex;

                    }
                }
            }
        }


        private void CreatestrAuthRight()
        {

            int xindex = 0;
            strRights = "";
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (((CheckBox)control).CheckState == CheckState.Checked)
                    {
                        if (xindex == 0)
                        {
                            ListedRightIndex = "1";
                            strRights = ((CheckBox)control).Text;
                        }
                        else
                        {
                            ListedRightIndex = ListedRightIndex + "_1";
                            strRights = strRights + " " + ((CheckBox)control).Text;
                        }

                    }
                    xindex++;
                }
            }
        }

        private void CreateAuthList2(string specificStr)
        {

            switch (CombAuthType.SelectedIndex)
            {
                case 0:
                    if(ParentForm.strauthList1.Length != 0)
                        authList2 = NewAuthSubject.Replace(ParentForm.strauthList1, "");
                    break;
                case 1:


                    authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] +  " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr + " بالرقم (" + Vitext1.Text + ") بمربع رقم (" + Vitext2.Text + ") البالغ مساحتها(" + Vitext3.Text + "م.م) ب" + Vitext4.Text + " - " + Vitext5.Text + " ";
                    break;
                case 2:

                    authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr + " بالرقم (" + Vitext1.Text + ") بمربع رقم (" + Vitext2.Text + ") البالغ مساحتها(" + Vitext3.Text + "م.م) ب" + Vitext4.Text + " - " + Vitext5.Text + " ";
                    break;
                case 3:

                    authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr + " بالرقم (" + Vitext1.Text + ") بمربع رقم (" + Vitext2.Text + ") البالغ مساحتها(" + Vitext3.Text + "م.م) ب" + Vitext4.Text + " - " + Vitext5.Text + " ";
                    break;
                case 5:
                    //سيارة

                    authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr + " من نوع " + Vitext1.Text + " موديل العام (" + Vitext5.Text + ") باللون " + Vitext2.Text + " ورقم لوحة (" + Vitext3.Text + " )وشاسيه بالرقم (" + Vitext4.Text + ") ";
                    break;
                case 6:
                    //


                    LegaceyPreStr = " وبصفت" + preffix[ParentForm.intAppcases, 0] + " ضمن ورثة " + DPTitle + " " + Vitext1.Text + "، بموجب الإعلام الشرعي رقم (" + Vitext2.Text + ") الصادر من محكمة " + Vitext3.Text + " والتركة بالرقم (" + Vitext4.Text + ")";
                    authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + LegaceyAttch + specificStr;
                    break;
                case 7:
                    //طلاق                    

                    authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr;
                    break;
                case 8:
                    //زواج                    

                    authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr;
                    break;
                case 9:
                    //جامعة                 
                    //foreach (Control control in PanelSubItemBox.Controls)
                    //{
                    //    if (control is CheckBox && ((CheckBox)control).CheckState == CheckState.Checked)
                    //    {
                    //        if (x == 0) strUni = ((CheckBox)control).Text;
                    //        else strUni = strUni + "و" + ((CheckBox)control).Text;
                    //        x++;
                    //    }
                    //}
                    if(ComboProcedure.Text.Contains("دراسة"))
                    authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + " في " + strUni + " بجامعة " + Vitext1.Text + " كلية " + Vitext2.Text + " ب" + Vitext3.Text;
                    else authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + " في " + Vitext1.Text;
                    break;
                case 10:
                    //شهادة ميلاد                    

                    authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + " في استخراج شهادة ميلاد " + Mentioned + " بالتفاصيل الاتي";                    
                    break;
                case 13:
                    //شهادة ميلاد                    

                    authList2 = " ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + " في استلام مبلغ مالي يخصني من شركة " + Vitext2.Text + " بشأن " + Vitext3.Text; 
                    break;

                default:
                    break;
            }
            collectItemsData();
        }

        private void collectItemsData()
        {
            BoxesData = Vitext1.Text + "_" + Vitext2.Text + "_" + Vitext3.Text + "_" + Vitext4.Text + "_" + Vitext5.Text + "_" + Vicheck1.Text + "_" + VitxtDate1VD.Text + "/" + VitxtDate1VM.Text + "/" + VitxtDate1VY.Text + "_" + Vicombo1.Text + "_" + LibtnAdd1.Text + "_" + Vicombo2.Text;            
        }
    }
}

