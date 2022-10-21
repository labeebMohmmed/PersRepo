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
        string strRights = "", authList2 = "", NewAuthSubject = "", AuthSubjectValue = "", DPTitle = "المرحوم";
        static int[] statistic = new int[100];
        static int[] staticIndex = new int[100];
        static int[] times = new int[100];
        static string[,] preffix = new string[10, 20];
        static string[] Text_statis = new string[5];
        static string[] strListcomb = new string[10] { "","","","","","","","","",""};
        
        string LastCol = "", DataSource = "";
        string ListedRightIndex = "", ColName = "";
        DataTable checkboxdt;
        private int[] checkliststr = new int[100];
        int Nobox = 0, LastID = 0, LastTabIndex = 0;
        Form11 Form11Parameter;
        public Form11 ParentForm { get; set; }

        public ComboBox comboBoxAuthValue
        {
            get { return CombAuthType; }
            set { CombAuthType = value; }
        }

        public UserAuthText()
        {
            InitializeComponent();            
            checkboxdt = new DataTable();
            //for(int y= 1;y<15;y++)
            //StaredColumns(y);
        }

        private void UserAuthText_Load(object sender, EventArgs e)
        {
            //StaredColumns();
            
        }

        private void btnSizeSpecial_Click(object sender, EventArgs e)
        {

            CreatestrAuthRight();
            CreateAuthList2(StrSpecPur);
            strAuthList2.DynamicInvoke(authList2);
            strAuthSubject.DynamicInvoke(AuthSubjectValue);
            strRightIndex.DynamicInvoke(ListedRightIndex);
            strRightsText.DynamicInvoke(strRights);
            AppMovePage.DynamicInvoke(2);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CreatestrAuthRight();
            CreateAuthList2(StrSpecPur);
            strAuthList2.DynamicInvoke(authList2);
            strAuthSubject.DynamicInvoke(AuthSubjectValue);
            strRightIndex.DynamicInvoke(ListedRightIndex);
            strRightsText.DynamicInvoke(strRights);
            AppMovePage.DynamicInvoke(4);
        }

        private string SuffPrefReplacements(string text)
        {
            Suffex_preffixList();
            if (text.Contains("@@@"))
                return text.Replace("@@@", preffix[ParentForm.intAppcases, 1]);
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
            SqlConnection sqlCon = new SqlConnection(source);
            string column = "@" + comlumnName;
            string qurey;
            if (datatype) qurey = "INSERT INTO TableAuthRights (" + comlumnName + ") values(" + column + ")";
            else qurey = "UPDATE TableAuthRights SET " + comlumnName + " = " + column + " WHERE ID = @ID";

            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;

            if (datatype)
            {
                sqlCmd.Parameters.AddWithValue(column, data.Trim());
                sqlCmd.ExecuteNonQuery();
            }
            else
            {

                sqlCmd.Parameters.AddWithValue("@ID", id);
                sqlCmd.Parameters.AddWithValue(column, data.Trim());
                sqlCmd.ExecuteNonQuery();
            }
            sqlCon.Close();
        }

        private void Suffex_preffixList()
        {

            preffix[0, 0] = "ي"; //$$$
            preffix[1, 0] = "ي";
            preffix[2, 0] = "نا";
            preffix[3, 0] = "نا";
            preffix[4, 0] = "نا";
            preffix[5, 0] = "نا";

            preffix[0, 1] = "ت";//&&&
            preffix[1, 1] = "ت";
            preffix[2, 1] = "نا";
            preffix[3, 1] = "نا";
            preffix[4, 1] = "نا";
            preffix[5, 1] = "نا";

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

            preffix[0, 5] = "";
            preffix[1, 5] = "ة";
            preffix[2, 5] = "ان";
            preffix[3, 5] = "تان";
            preffix[4, 5] = "ات";
            preffix[5, 5] = "ون";

            preffix[0, 6] = "";
            preffix[1, 6] = "ة";
            preffix[2, 6] = "ين";
            preffix[3, 6] = "تين";
            preffix[4, 6] = "ات";
            preffix[5, 6] = "رين";

            preffix[0, 7] = "ينوب";
            preffix[1, 7] = "تنوب";
            preffix[2, 7] = "ينوبا";
            preffix[3, 7] = "تنوبا";
            preffix[4, 7] = "ينبن";
            preffix[5, 7] = "ينوبوا";

            preffix[0, 8] = "يقوم";
            preffix[1, 8] = "تقوم";
            preffix[2, 8] = "يقوما";
            preffix[3, 8] = "تقوما";
            preffix[4, 8] = "يقمن";
            preffix[5, 8] = "يقوموا";

        }
        private void StaredColumns(int x)
        {
            checkboxdt = new DataTable();
                int LastID = 0;
                string col = "Col" + x;
                string query = "SELECT ID," + col + " FROM TableAuthRight";
                string source = "Data Source=192.168.100.123,49170;Network Library=DBMSSOCN;Initial Catalog=AhwalDataBase;User ID=ConsJeddah;Password=DBaseC@nsJ0d103";
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
                            if (checkboxdt.Rows[LastID][col].ToString().Contains("_"))
                            {
                                Text_statis = checkboxdt.Rows[LastID][col].ToString().Split('_');  
                            if(Text_statis[0].Contains("توكيل الغير في")) 
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
        private void PopulateCheckBoxes(string col, string table)
        {
            LastCol = col;

            string query = "SELECT ID," + col + " FROM " + table;

            using (SqlConnection con = new SqlConnection(ParentForm.PublicDataSource))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter(query, con))
                {
                    sda.Fill(checkboxdt);
                    listchecked = checkboxdt.Rows.Count;

                    foreach (DataRow row in checkboxdt.Rows)
                    {
                        if (checkboxdt.Rows[Nobox][col].ToString().Contains("_"))
                        {
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
                            chk.Text = text;
                            chk.Tag = "valid";
                            statistic[Nobox] = Convert.ToInt32(Text_statis[1]);
                            times[Nobox] = Convert.ToInt32(Text_statis[2]);
                            staticIndex[Nobox] = Convert.ToInt32(Text_statis[3]);

                            for (int i = 0; i < listchecked; i++)
                            {//
                                if (staticIndex[Nobox] == checkliststr[i] || Text_statis[4] == "Star")
                                    chk.CheckState = CheckState.Checked;
                            }
                            chk.CheckedChanged += new EventHandler(CheckBox_Checked);
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
                            picboxup.Name = Nobox.ToString();
                            picboxup.Size = new System.Drawing.Size(24, 26);
                            picboxup.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                            picboxup.TabIndex = 176 + Nobox;
                            picboxup.TabStop = false;
                            picboxup.Click += new System.EventHandler(this.pictureBoxup_Click);
                            if (Nobox == 0)
                            {
                                picboxup.Visible = false;
                            }
                            panelAuthOptions.Controls.Add(picboxup);

                            PictureBox picboxdown = new PictureBox();
                            picboxdown.Image = global::PersAhwal.Properties.Resources.arrowdown;
                            picboxdown.Location = new System.Drawing.Point(55, Nobox * 37);
                            picboxdown.Name = "pictureBoxdown" + Nobox.ToString();
                            picboxdown.Size = new System.Drawing.Size(24, 26);
                            picboxdown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                            picboxdown.TabIndex = 177 + Nobox;
                            picboxdown.TabStop = false;
                            picboxdown.Click += new System.EventHandler(this.pictureBoxdown_Click);
                            if (Nobox == listchecked - 1) picboxdown.Visible = false;
                            panelAuthOptions.Controls.Add(picboxdown);
                            LastID = Convert.ToInt32(checkboxdt.Rows[Nobox]["ID"].ToString());
                            Nobox++;
                        }
                    }
                }
            }

        }

        private void pictureBoxup_Click(object sender, EventArgs e)
        {

            PictureBox picbox = (PictureBox)sender;
            string st = "", nd = "";
            bool statest = false, statend = false;
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (((CheckBox)control).TabIndex == picbox.TabIndex - 176 && !((CheckBox)control).Text.Contains("والله خير الشاهدين"))
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

                }
            }
            int x = 0, y = 0;
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (((CheckBox)control).TabIndex == picbox.TabIndex - 176 && !((CheckBox)control).Text.Contains("والله خير الشاهدين"))
                    {
                        ((CheckBox)control).Text = nd;
                        if (statend) ((CheckBox)control).CheckState = CheckState.Checked;
                        else ((CheckBox)control).CheckState = CheckState.Unchecked;
                        y = statistic[x];
                        statistic[x] = statistic[x - 1];
                        statistic[x - 1] = y;

                        y = staticIndex[x];
                        staticIndex[x] = staticIndex[x - 1];
                        staticIndex[x - 1] = y;
                    }
                    if (((CheckBox)control).TabIndex == picbox.TabIndex - 177 && !((CheckBox)control).Text.Contains("والله خير الشاهدين"))
                    {
                        ((CheckBox)control).Text = st;
                        if (statest) ((CheckBox)control).CheckState = CheckState.Checked;
                        else ((CheckBox)control).CheckState = CheckState.Unchecked;

                    }
                    x++;
                }

            }


        }

        private void pictureBoxdown_Click(object sender, EventArgs e)
        {
            PictureBox picbox = (PictureBox)sender;
            string st = "", nd = "";
            bool statest = false, statend = false;
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
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
                }
            }
            int x = 0, y = 0;
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
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

        private void btnAddRight_Click(object sender, EventArgs e)
        {
            MessageBox.Show("UserAuthText" + ParentForm.PublicDataSource);
        }
        private void deleteItemsAO()
        {
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

        string StrSpecPur = "";
        private string StrSpecificPurpose()
        {

            string text = "", text2="";

            

            switch (ComboProcedure.SelectedIndex)
            {
                case 0:

                    StrSpecPur = " في بيع " + textValue();
                    break;
                case 2:
                    StrSpecPur = " في شراء " + textValue();
                    break;

                case 1:
                    StrSpecPur = " في إجراءات سحب القرعة والإستلام لقطعة الأرض السكنية بالخطة الإسكانية لأراضي مدينة " + textValue();
                    break;

                case 3:
                    StrSpecPur = " في فك الحجز وبيع " + textValue();
                    break;
                case 4:
                    StrSpecPur = " في الإشراف على " + textValue();
                    break;
                case 5:
                    StrSpecPur = " في البناء وإدخال خدمات المياه والكهرباء والصرف الصحي  ل" + textValue();
                    break;
                case 6:
                    StrSpecPur = " في رفع دعاوي أمام كافة المحاكم والنيابات بمختلف أنواعها ودرجاتها وتمثيلي في الدعاوى المرفوعة مني أو ضدي والوقوف والمقاضاة بشأن كل ما يتعلق ب" + textValue();
                    break;
                case 7:
                    StrSpecPur = " في حجز " + textValue();
                    break;
                case 9:
                    StrSpecPur = " في هبة " + textValue();
                    break;
                case 8:
                    StrSpecPur = " في رهن " + textValue();
                    break;
                case 11:
                    StrSpecPur = " في إستخراج شهادة بحث بغرض التأكد ل" + textValue();
                    break;
                case 12:
                    StrSpecPur = " في إستخراج شهادة بحث بغرض الرهن ل" + textValue();
                    break;
                case 13:
                    StrSpecPur = " في إستخراج شهادة بحث بغرض الحجز ل" + textValue();
                    break;
                case 14:
                    StrSpecPur = " في إستخراج شهادة بحث بغرض البيع ل" + textValue();
                    break;

                default:
                    StrSpecPur = "";
                    break;
            }
            return textValue2();
        }

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
            string text="";
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
                    combbox.Items.Add(dataRow[comlumnName].ToString());
                }
                saConn.Close();
            }
        }

        private void CombAuthType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CombAuthType.SelectedIndex == 0)
            {
                txtReview.Location = new System.Drawing.Point(273, 76);
                txtReview.Height = 160;
                PanelItemsboxes.Visible = false;
                label8.Location = new System.Drawing.Point(1114, 76);
            }

            if (CombAuthType.SelectedIndex >= 1 && CombAuthType.SelectedIndex <= 4)
            {
                SepratBoxes();
                fileComboBox(ComboProcedure, ParentForm.PublicDataSource, "Row1Attach", "TableListCombo");                
            }
            if (CombAuthType.SelectedIndex == 5)
            {
                SepratBoxes();
                DetermineCheckBox("اسم المتوفى", 230, "رقم الاعلام الشرعي:", 200, "اسم المحكمة:", 230, "رقم التركة:", 200, "", 80, DPTitle, "", strListcomb);
                fileComboBox(ComboProcedure, ParentForm.PublicDataSource, "Row1Attach", "TableListCombo");
            }
            if (CombAuthType.SelectedIndex == 6)
            {
                SepratBoxes();
                string[] str6Listcomb = new string [3] { "طلقة أولى","طلقة ثانية","طلقة ثالثة"};
                
                DetermineCheckBox("اسم المطلقة", 230, "", 200, "", 230, "", 200, "", 80, "", "تاريخ الطلاق", str6Listcomb);
                fileComboBox(ComboProcedure, ParentForm.PublicDataSource, "Row1Attach", "TableListCombo");
            }
        }
        private void DetermineCheckBox(string v1, int s1, string v2, int s2, string v3, int s3, string v4, int s4, string v5, int s5, string v6, string v7, string[] v8)
        {
            foreach (Control control in PanelItemsboxes.Controls)
            {
                if (control is TextBox)
                {
                    ((TextBox)control).Visible = false;
                }
                if (control is Label)
                {
                    ((Label)control).Visible = false;
                }
            }
            if (v1 != "")
            {
                label1.Text = v1;
                label1.Visible = true;
                txt1.Width = s1;
                txt1.Visible = true;
            }
            if (v2 != "")
            {
                label2.Text = v2;
                label2.Visible = true;
                txt2.Width = s2;
                txt2.Visible = true;
            }
            if (v3 != "")
            {
                label3.Text = v3;
                label3.Visible = true;
                txt3.Width = s3;
                txt3.Visible = true;
            }
            if (v4 != "")
            {
                label4.Text = v4;
                label4.Visible = true;
                txt4.Width = s4;
                txt4.Visible = true;
            }
            if (v5 != "")
            {
                label5.Text = v5;
                label5.Visible = true;
                txt5.Width = s5;
                txt5.Visible = true;
            }

            checkSexType.Visible = false;
            DT.Visible = false;
            comboBox1.Visible = false;
            if (v6 != "")
            {
                label9.Visible = true;
                checkSexType.Visible = true;
            }

            if (v7 != "")
            {
                label6.Visible = true;
                DT.Visible = true;
                label6.Text = v6;
            }


            if (v8[0] != "")
            {
                label10.Visible = true;
                comboBox1.Visible = true;


                for (int x = 0; x < v8.Length; x++)
                    comboBox1.Items.Add(v8[x]);
                        
            }

        }
        private void SepratBoxes()
        {
            txtReview.Location = new System.Drawing.Point(273, 167);
            txtReview.Height = 80;
            PanelItemsboxes.Visible = true;
            label8.Location = new System.Drawing.Point(1114, 167);
        }

        private void radioSelectAll_CheckedChanged(object sender, EventArgs e)
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

        private void CombAuthType_MouseEnter(object sender, EventArgs e)
        {
            //CombAuthType.Items.Clear();
            //foreach (DataRow dataRow in ParentForm.UserComboTexttable.Rows)
            //{
            //    CombAuthType.Items.Add(dataRow["AuthTypes"].ToString());
            //}
        }

        private void checkSexType_CheckedChanged(object sender, EventArgs e)
        {
            if (checkSexType.CheckState == CheckState.Unchecked) { DPTitle = "المرحوم"; checkSexType.Text = "ذكر"; }
            else { DPTitle = "المرحومة"; checkSexType.Text = "أنثى"; }
        }

        
        private void txtReviewBody_TextChanged(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;
            NewAuthSubject = txt.Text;
        }

        private void ComboProcedure_SelectedIndexChanged(object sender, EventArgs e)
        {
            deleteItemsAO();
            strRights = "";

            string text = StrSpecificPurpose();
            if (ComboProcedure.SelectedIndex == 0)
            { ColName = "Col14";
                DetermineCheckBox("رقم" + text, 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "","", strListcomb);
            }
            else
            {
                ColName = "Col" + (ComboProcedure.SelectedIndex - 1).ToString();
                DetermineCheckBox("رقم" + text, 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
            }
            PopulateCheckBoxes(ColName, "TableAuthRights");
        }

        private void pictureBoxedit_Click(object sender, EventArgs e)
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

        private void CheckBox_Checked(object sender, EventArgs e)
        {

            CheckBox chk = (sender as CheckBox);
            if (chk.Checked)
            {

                //strRights = strRights + chk.Text + "،";
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
                        if (xindex == 0) ListedRightIndex = staticIndex[xindex].ToString();
                        else ListedRightIndex = ListedRightIndex + "_" + staticIndex[xindex].ToString();
                        strRights = strRights + ((CheckBox)control).Text;

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
                    authList2 = NewAuthSubject.Replace(ParentForm.strauthList1, "");
                    break;
                case 1:
                    AuthSubjectValue = txt1.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt2.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt3.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt4.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt5.Text;
                    authList2 = "ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr + "  بالرقم (" + txt1.Text + ") بمربع رقم (" + txt2.Text + ") البالغ مساحتها(" + txt3.Text + "م.م) ب" + txt4.Text + " - " + txt5.Text + " ";
                    break;
                case 2:
                    AuthSubjectValue = txt1.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt2.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt3.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt4.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt5.Text;
                    authList2 = "ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr + "  بالرقم (" + txt1.Text + ") بمربع رقم (" + txt2.Text + ") البالغ مساحتها(" + txt3.Text + "م.م) ب" + txt4.Text + " - " + txt5.Text + " ";
                    break;
                case 3:
                    AuthSubjectValue = txt1.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt2.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt3.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt4.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt5.Text;
                    authList2 = "ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr + "  بالرقم (" + txt1.Text + ") بمربع رقم (" + txt2.Text + ") البالغ مساحتها(" + txt3.Text + "م.م) ب" + txt4.Text + " - " + txt5.Text + " ";
                    break;
                case 5:
                    AuthSubjectValue = txt1.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt2.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt3.Text;
                    AuthSubjectValue = AuthSubjectValue + "_" + txt4.Text;
                    authList2 = " وبصفت" + preffix[ParentForm.intAppcases, 0] + " ضمن ورثة " + DPTitle + " " + txt1.Text + "، بموجب الإعلام الشرعي رقم (" + txt2.Text + ") الصادر من محكمة " + txt3.Text + " والتركة بالرقم (" + txt4.Text + ")";
                    break;
                default:
                    break;
            }
        }

    }
}
