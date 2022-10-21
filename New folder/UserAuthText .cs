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
        static int[] staticIndex = new int[100];
        static int[] times = new int[100];
        static string[,] preffix = new string[10, 20];
        static string[] Text_statis = new string[5];
        static string[] strListcomb = new string[10] { "","","","","","","","","",""};
        string[] dataset = new string[20];
        string strUni = "";
        int x = 0;
        string LastCol = "", DataSource = "";
        string ListedRightIndex = "", ColName = "", BoxesData = "";
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
        
        
        public UserAuthText()
        {
            InitializeComponent();            
            checkboxdt = new DataTable();  
            for(int x = 1;x<26;x++) StaredColumns(x);
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
                    if (x == 0)
                        times[x]++;
                    if (((CheckBox)control).CheckState == CheckState.Checked) { statistic[x]++; }
                    UpdateColumn(source, col, x + 1, ((CheckBox)control).Text + "_" + statistic[x].ToString() + "_" + times[x].ToString() + "_" + staticIndex[x].ToString(), false);
                    x++;
                }
            }
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

            preffix[0, 9] = "نصيبي";
            preffix[1, 9] = "نصيبي";
            preffix[2, 9] = "نصيبينا";
            preffix[3, 9] = "نصيبينا";
            preffix[4, 9] = "أنصبتنا";
            preffix[5, 9] = "أنصبتنا";

        }
        private void StaredColumns(int x)
        {
            checkboxdt = new DataTable();
                int LastID = 0;
                string col = "Col" + x.ToString();
                string col1 = "Col" + (x+1).ToString(); ;
            string query = "SELECT ID," + col + " FROM TableAuthRights";
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
                            if (checkboxdt.Rows[LastID][col].ToString().Contains("_") && (!checkboxdt.Rows[LastID][col].ToString().Contains("Star") || checkboxdt.Rows[LastID][col].ToString().Contains("off")))
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
        public void PopulateCheckBoxes(string col, string table, string dataSource)
        {
            LastCol = col;
           
            string query = "SELECT ID," + col + " FROM " + table;

            using (SqlConnection con = new SqlConnection(dataSource))
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
                            {
                                if ( Text_statis[4] == "Star")
                                    chk.CheckState = CheckState.Checked;
                            }
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
                            if (chk.Text.Contains("الحق في توكيل الغير")|| chk.Text.Contains("لمن يشهد والله خير الشاهدين")) picboxdown.Visible = false;
                            
                            panelAuthOptions.Controls.Add(picboxdown);
                            LastID = Convert.ToInt32(checkboxdt.Rows[Nobox]["ID"].ToString());
                            Nobox++;
                        }
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
                
                if (control is CheckBox )
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
                    } else FirstCase = false;

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
                if (control is CheckBox )
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

        private void btnAddRight_Click(object sender, EventArgs e)
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

                //UpdateColumn(DataSource, LastCol, LastID + 1, chk.Text + "_" + statistic[Nobox].ToString() + "_" + times[Nobox].ToString() + "_" + staticIndex[Nobox].ToString() + "_Star", true);
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

        private void ShowArrows(int tabindex,int indexMinus)
        {
            foreach (Control control in panelAuthOptions.Controls)
            {

                if (control is PictureBox)
                {

                    if (((PictureBox)control).Name == "Down" && ((PictureBox)control).TabIndex == 177 + tabindex -3)
                    {
                        ((PictureBox)control).Visible = true;
                    }
                    if (((PictureBox)control).Name == "Up" && ((PictureBox)control).TabIndex == 176 + tabindex - 2- indexMinus)
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
                    
                    if (((CheckBox)control).TabIndex == tabindex-1)
                    {
                        st = ((CheckBox)control).Text;
                        if (((CheckBox)control).CheckState == CheckState.Checked) statest = true;
                        else statest = false;
                        
                    }
                    if (((CheckBox)control).TabIndex == tabindex-2)
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
                        if (((CheckBox)control).TabIndex == tabindex-1)
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
                        if (((CheckBox)control).TabIndex == tabindex-2)
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
                    combbox.Items.Add(dataRow[comlumnName].ToString().Trim());
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
                    dataset[x] = dataRow[comlumnName].ToString();
                    box.Items.Add(dataset[x]);
                    x++;
                }
                saConn.Close();
            }
        }

        private void CombAuthType_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboProcedure.Text = "إختر الإجراء";
            restShowingItems();
            SepratBoxes();
            LegaceyPreStr = "";
            if (CombAuthType.SelectedIndex == 0)
            {
                txtReview.Location = new System.Drawing.Point(273, 76);
                txtReview.Height = 160;
                PanelItemsboxes.Visible = false;
                label8.Location = new System.Drawing.Point(1114, 76);
            }

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
        }

        public void LegaceyBox(string v1, string v2, string v3,  string v4,  string v5,  string v6, string v7, string v8, string v9)
        {            
            if (v1 != "")
            {
                lab1.Text = v1;
                lab1.Visible = true;                
                textBox1.Visible = true;
            }
            if (v2 != "")
            {
                lab2.Text = v2;
                lab2.Visible = true;                
                textBox2.Visible = true;
            }
            if (v3 != "")
            {
                lab3.Text = v3;
                lab3.Visible = true;
                textBox3.Visible = true;
            }
            if (v4 != "")
            {
                lab4.Text = v4;
                lab4.Visible = true;
                textBox4.Visible = true;
            }
            if (v5 != "")
            {
                lab5.Text = v5;
                lab5.Visible = true;
                textBox5.Visible = true;
            }
            if (v6 != "")
            {
                lab6.Text = v3;
                lab6.Visible = true;
                textBox6.Visible = true;
            }
            if (v7 != "")
            {
                lab7.Text = v4;
                lab7.Visible = true;
                textBox7.Visible = true;
            }
            if (v8 != "")
            {
                lab8.Text = v5;
                lab8.Visible = true;
                textBox8.Visible = true;
            }
            if (v9 != "")
            {
                label11.Text = v5;
                label11.Visible = true;
                txtBoxGeneral.Visible = true;
            }


        }
        private void DetermineCheckBox(string v1, int s1, string v2, int s2, string v3, int s3, string v4, int s4, string v5, int s5, string v6, string v7, string[] v8)
        {
            restShowingItems();
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

            
            if (v6 != "")
            {
                label9.Text = v6;
                label9.Visible = true;
                checkSexType.Text = DPTitle;
                checkSexType.Visible = true;
            }

            if (v7 != "")
            {
                label6.Visible = true;
                DT.Visible = true;
                label7.Text = v7;
            }


            if (v8[0] != "")
            {
                label10.Visible = true;
                comboBox1.Visible = true;
                label10.Text = v8[0];
                
                comboBox1.Items.Clear();
                for (int x = 1; x < v8.Length; x++)
                    comboBox1.Items.Add(v8[x]);
                comboBox1.SelectedIndex = 0;
            }

        }

        private void restShowingItems()
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
            checkSexType.Visible = false;
            DT.Visible = false;
            comboBox1.Visible = false;
            PanelSubItemBox.Visible = false;
        }

        private void SepratBoxes()
        {
            txtReview.Location = new System.Drawing.Point(338, 162);
            txtReview.Height = 80;
            PanelItemsboxes.Visible = true;
            label8.Location = new System.Drawing.Point(1172, 76);
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

        private void comboPropertyType_SelectedIndexChanged(object sender, EventArgs e)
        {
            PanelSubItemBox.Visible = true;
            if (comboPropertyType.Text.Contains("عقار"))
            {
                //MessageBox.Show(comboPropertyType.Text);
                LegaceyBox("رقم العقار" , "رقم المربع:",  "المساحة:", "الحي:",  "المدينة:",  "", "","","");
            }
            
            else if (comboPropertyType.Text.Contains("سيارة"))
            {
                //MessageBox.Show(comboPropertyType.Text);
                LegaceyBox("نوع السيارة", "للون:", "رقم اللوحة:",  "رقم الشاسية:",  "سنة الموديل", "", "", "","");
            }
            else 
            {
                //MessageBox.Show(comboPropertyType.Text);
                LegaceyBox("", "", "", "", "", "", "", "", "وصف الورثة");
            }
        }
        int LegaceyIndex = 0;
        string LegaceyItem = "";
        string LegaceyPreStr = "";
        private void btnAddLegacey_Click(object sender, EventArgs e)
        {
            if (comboPropertyType.Text.Contains("سيارة"))
            {
                if (LegaceyIndex == 0)
                    StrSpecPur = "سيارة من نوع " + textBox1.Text + " موديل العام (" + textBox5.Text + ") باللون " + textBox2.Text + " ورقم لوحة (" + textBox3.Text + " )وشاسيه بالرقم (" + textBox4.Text + ") ";
                else StrSpecPur = StrSpecPur + "وسيارة من نوع " + textBox1.Text + " موديل العام (" + textBox5.Text + ") باللون " + textBox2.Text + " ورقم لوحة (" + textBox3.Text + " )وشاسيه بالرقم (" + textBox4.Text + ") ";
            }
            else if (comboPropertyType.Text.Contains("عقار"))
            {
                if (LegaceyIndex == 0)
                    StrSpecPur = "عقار بالرقم (" + textBox1.Text + ") بمربع رقم (" + textBox2.Text + ") البالغ مساحتها(" + textBox3.Text + "م.م) ب" + textBox4.Text + " - " + textBox5.Text + " ";
                else
                    StrSpecPur = StrSpecPur + "وعقار بالرقم (" + textBox1.Text + ") بمربع رقم (" + textBox2.Text + ") البالغ مساحتها(" + textBox3.Text + "م.م) ب" + textBox4.Text + " - " + textBox5.Text + " ";
            }
            else {
                if (LegaceyIndex == 0) StrSpecPur = txtBoxGeneral.Text; 
                else StrSpecPur = StrSpecPur + " و "+txtBoxGeneral.Text;

            }
            textBox1.Text = textBox2.Text = textBox3.Text = textBox4.Text = textBox5.Text = txtBoxGeneral.Text="";
            LegaceyIndex++;

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void ComboProcedure_TextChanged(object sender, EventArgs e)
        {
            deleteItemsAO();
            label7.Visible = false;
            comboPropertyType.Visible = false;
            txtReview.Text = "";
            foreach (Control control in PanelItemsboxes.Controls)
            {
                control.Visible = false;
                if (control is TextBox) ((TextBox)control).Text = "";
                if (control is CheckBox) ((CheckBox)control).CheckState = CheckState.Unchecked;
            }
            strRights = "";
            ColName = "Col0";
            CreateBoxesWithData(ComboProcedure.Text, "", false);
            
            PopulateCheckBoxes(ColName, "TableAuthRights", ParentForm.PublicDataSource);
        }

        private void ComboProcedure_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtReview_MouseHover(object sender, EventArgs e)
        {
            
            string other = "الآنسة";
            if (ComboProcedure.Text == "عقد قران شخصي")
            {
                StrSpecPur = " في عقد قراني على " + checkSexType.Text + " /" + txt1.Text;                

            }
            else if (ComboProcedure.Text == "عقد قران غير شخصي")
            {
                StrSpecPur = " في عقد قران  " + txt2.Text + " "+checkSexType.Text + " /" + txt1.Text + " على " + comboBox1.Text + "/ " + txt3.Text;
                
            }
            else if (ComboProcedure.Text.Contains("وثيقة تصادق"))
            {
                 other = "زوجتي/ ";
                if (ParentForm.strAppMaleFemaleList[0] == "أنثى") other = "زوجي/ ";
                StrSpecPur = " في إستخراج وثيقة تصادق على زواجي من " + other +  txt1.Text +" بتاريخ: " + DT.Text;
                
            }
            else if (ComboProcedure.Text.Contains("وثيقة زواج"))
            {
                other = "زوجتي/ ";
                if (ParentForm.strAppMaleFemaleList[0] == "أنثى") other = "زوجي/ ";
                StrSpecPur = " في إستخراج قسيمة زواجي من " + other + txt1.Text+ " بتاريخ: " + DT.Text;
            }
            string DivNo = "زوجتي";
            //

            if (CombAuthType.Text == "طلاق")
            {
                
                if (ParentForm.strAppMaleFemaleList[0] != "")
                {
                    if (comboBox1.Text == "طلقة ثالثة") DivNo = "مطلقتي "; else DivNo = "زوجتي ";
                }
                else
                {
                    if (comboBox1.Text == "طلقة ثالثة") DivNo = "مطلقي "; else DivNo = "زوجي ";
                }                
                if (ComboProcedure.SelectedIndex == 0)
                {                    
                    if (DivNo == "زوجتي " || DivNo == "مطلقتي ")
                        StrSpecPur = " إيقاع " + comboBox1.Text + " على " + DivNo +" السيدة/ " +txt1.Text;
                }
                else if (ComboProcedure.SelectedIndex == 1)
                    StrSpecPur = " إستخراج قسيمة طلاقي من " + DivNo + txt1.Text + " التي أوقعت عليها " + comboBox1.Text + " بتاريخ: " + DT.Text;                
            }
            CreateAuthList2(StrSpecPur);
            
            txtReview.Text = LegaceyPreStr + ParentForm.strauthList1 + authList2;
        }

        
        private void checkSexType_CheckedChanged(object sender, EventArgs e)
        {
            if (CombAuthType.Text.Contains("ورثة"))
            {
                if (checkSexType.CheckState == CheckState.Unchecked) { DPTitle = "المرحوم"; checkSexType.Text = "ذكر"; }
                else { DPTitle = "المرحومة"; checkSexType.Text = "أنثى"; }
            }
            else if (CombAuthType.Text.Contains("زواج")) 
            {
                if (checkSexType.CheckState == CheckState.Unchecked) {checkSexType.Text = "الآنسة"; }
                else {
                    DPTitle = "السيدة"; 
                    checkSexType.Text = "السيدة"; }
            }
        }

        
        private void txtReviewBody_TextChanged(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;
            NewAuthSubject = txt.Text;
        }


        
        public void CreateBoxesWithData(string textdata,string textItems, bool database)

        {
            
            int x = 0;
            string[] str6Listcomb = new string[3] { "صفة الزوج", "السيد", "موكلي السيد" };
            string[] SI = new string[9];
            string text = "";
            for (x = 0; x < 8; x++) SI[x] = "";

            x = 0;
            if (textItems.Contains("_")) SI = textItems.Split('_');
            
            switch (textdata)
            {
                case "عقد قران شخصي":
                    ColName = "Col23";
                    StrSpecPur = " في عقد قراني على ";
                    DPTitle = "آنسة";
                    DetermineCheckBox("اسم الطرف الآخر", 210, "", 200, "", 230, "", 200, "", 80, "الحالة الإجتماعية للمراد الزواج منها", "", strListcomb);                    
                    break;
                case "عقد قران غير شخصي":
                    StrSpecPur = " في عقد قران  ";
                    ColName = "Col23";
                    str6Listcomb = new string[3] { "صفة المراد الزواج منها", "الآنسة", "السيدة"};
                    DetermineCheckBox("اسم الزوجة", 200, "", 200, "إسم الزوج", 100, "", 200, "", 80, "الحالة الإجتماعية للمراد الزواج منها", "", str6Listcomb);
                    break;
                case "وثيقة تصادق على زواج":
                    StrSpecPur = " في إستخراج وثيقة تصادق على زواجي من  ";
                    ColName = "Col23";
                    DetermineCheckBox("اسم الطرف الآخر", 210, "", 200, "", 230, "", 200, "", 80, "", "تاريخ الزواج", strListcomb);
                    break;
                case "استخراج قسيمة طلاق":
                    ColName = "Col24";
                    StrSpecPur = "استخراج قسيمة طلاق";
                    str6Listcomb[0] = "عدد الطلقات";
                    str6Listcomb[1] = "طلقة ثانية";
                    str6Listcomb[2] = "طلقة ثالثة";
                    DetermineCheckBox("اسم المطلق" + ParentForm.strAppMaleFemaleList[0], 210, "", 200, "", 230, "", 200, "", 80, "", "تاريخ وقوع الطلاق", str6Listcomb);
                    break;
                case "قسيمة زواج":
                    StrSpecPur = " في إستخراج وثيقة زواجي من  ";
                    ColName = "Col23";
                    DetermineCheckBox("اسم الطرف الآخر", 210, "", 200, "", 230, "", 200, "", 80, "", "تاريخ الزواج:", strListcomb);
                    break;
                case "ايقاع طلاق":
                    ColName = "Col24";
                    str6Listcomb = new string[4] { "عدد الطلقات", "طلقة أولى", "طلقة ثانية", "طلقة ثالثة" };
                    DetermineCheckBox("اسم المطلق" + ParentForm.strAppMaleFemaleList[0], 210, "", 200, "", 230, "", 200, "", 80, "", "", str6Listcomb);
                    break;
                case "ورثة - استلام":
                    LegaceyAttch = "في استلام " + preffix[ParentForm.intAppcases, 9] + " في ";
                    label7.Visible = true;
                    LegaceyIndex = 0;
                    comboPropertyType.Visible = true;
                    ColName = "Col25";
                    DetermineCheckBox("اسم المتوفى", 230, "رقم الاعلام الشرعي:", 200, "اسم المحكمة:", 230, "رقم التركة:", 200, "", 80, DPTitle, "", strListcomb);
                    break;
                case "ورثة - الوقوف والمقاضاة":
                    LegaceyAttch = "في الوقوف والمقاضاة بشأن " + preffix[ParentForm.intAppcases, 9] + " في ";
                    label7.Visible = true;
                    LegaceyIndex = 0;
                    comboPropertyType.Visible = true;
                    ColName = "Col25";
                    DetermineCheckBox("اسم المتوفى", 230, "رقم الاعلام الشرعي:", 200, "اسم المحكمة:", 230, "رقم التركة:", 200, "", 80, DPTitle, "", strListcomb);
                    break;
                case "ورثة - تنازل":
                    LegaceyAttch = "في التنازل عن " + preffix[ParentForm.intAppcases, 9] + " في ";
                    label7.Visible = true;
                    LegaceyIndex = 0;
                    comboPropertyType.Visible = true;
                    ColName = "Col25";
                    DetermineCheckBox("اسم المتوفى", 230, "رقم الاعلام الشرعي:", 200, "اسم المحكمة:", 230, "رقم التركة:", 200, "", 80, DPTitle, "", strListcomb);
                    break;
                case "ورثة - تصرف ناقل للملكية":
                    LegaceyAttch = "التصرف بجميع التصرفات الناقلة للملكية " + preffix[ParentForm.intAppcases, 9] + " في ";
                    label7.Visible = true;
                    LegaceyIndex = 0;
                    comboPropertyType.Visible = true;
                    ColName = "Col25";
                    DetermineCheckBox("اسم المتوفى", 230, "رقم الاعلام الشرعي:", 200, "اسم المحكمة:", 230, "رقم التركة:", 200, "", 80, DPTitle, "", strListcomb);
                    break;
                case "ورثة - الإشراف":
                    LegaceyAttch = "في الإشراف على " + preffix[ParentForm.intAppcases, 9] + " في ";
                    label7.Visible = true;
                    LegaceyIndex = 0;
                    comboPropertyType.Visible = true;
                    ColName = "Col25";
                    DetermineCheckBox("اسم المتوفى", 230, "رقم الاعلام الشرعي:", 200, "اسم المحكمة:", 230, "رقم التركة:", 200, "", 80, DPTitle, "", strListcomb);
                    break;
                case "بيع ارض":
                    ColName = "Col0";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في بيع " + textValue();
                    break;
                case "شراء ارض":                    
                    ColName = "Col2";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في شراء " + textValue();
                    break;
                case "خطة اسكانية":
                    ColName = "Col1";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في إجراءات سحب القرعة والإستلام لقطعة الأرض السكنية بالخطة الإسكانية لأراضي مدينة " + textValue();
                    break;
                case "فك حجز وبيع":
                    ColName = "Col5";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في فك الحجز وبيع " + textValue();
                    break;
                case "إشراف":
                    ColName = "Col4";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في الإشراف على " + textValue();
                    break;
                case "إدخال خدمات":
                    ColName = "Col18";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في البناء وإدخال خدمات المياه والكهرباء والصرف الصحي  ل" + textValue();
                    break;
                case "تقاضي":
                    ColName = "Col6";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في الوقوف والمقاضاة أمام كافة المحاكم والنيابات بمختلف أنواعها ودرجاتها وتمثيلي في الدعاوى المرفوعة مني أو ضدي والوقوف والمقاضاة بشأن كل ما يتعلق ب" + textValue();
                    break;
                case "حجز":
                    ColName = "Col19";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في حجز " + textValue();
                    break;
                case "هبة":
                    ColName = "Col9";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في هبة " + textValue();
                    break;
                case "رهن":
                    ColName = "Col8";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في رهن " + textValue();
                    break;
                case "شهادة بحث بغرض التأكد":
                    ColName = "Col10";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في إستخراج شهادة بحث بغرض التأكد ل " + textValue();
                    break;
                case "شهادة بحث بغرض الرهن":
                    ColName = "Col8";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في إستخراج شهادة بحث بغرض الرهن ل " + textValue();
                    break;
                case "شهادة بحث بغرض الهبة":
                    ColName = "Col19";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في إستخراج شهادة بحث بغرض الحجز ل " + textValue();
                    break;
                case "شهادة بحث بغرض البيع":
                    ColName = "Col13";
                    DetermineCheckBox("رقم" + textValue2(), 80, "رقم المربع:", 80, "المساحة:", 80, "الحي:", 80, "المدينة:", 80, "", "", strListcomb);
                    StrSpecPur = " في إستخراج شهادة بحث بغرض البيع ل " + textValue();
                    break;
                case "سيارة - التخارج":
                    StrSpecPur = " في التخارج عن سيارة " + text;
                    ColName = "Col17";
                    DetermineCheckBox("نوع السيارة", 120, "للون:", 100, "رقم اللوحة:", 120, "رقم الشاسية:", 270, "سنة الموديل:", 80, "", "", strListcomb);
                    break;
                case "سيارة - استلام":
                    StrSpecPur = " في استلام سيارة " + text;
                    ColName = "Col18";
                    DetermineCheckBox("نوع السيارة", 120, "للون:", 100, "رقم اللوحة:", 120, "رقم الشاسية:", 270, "سنة الموديل:", 80, "", "", strListcomb);
                    break;
                case "سيارة - الاشراف":
                    StrSpecPur = " في الإشراف على سيارة " + text;
                    ColName = "Col19";
                    DetermineCheckBox("نوع السيارة", 120, "للون:", 100, "رقم اللوحة:", 120, "رقم الشاسية:", 270, "سنة الموديل:", 80, "", "", strListcomb);
                    break;
                case "سيارة - تقاضي":
                    ColName = "Col19";
                    StrSpecPur = " في الوقوف والمقاضاة بشأن سيارة " + text;
                    DetermineCheckBox("نوع السيارة", 120, "للون:", 100, "رقم اللوحة:", 120, "رقم الشاسية:", 270, "سنة الموديل:", 80, "", "", strListcomb);
                    break;
                case "سيارة - تخليص جمركي":
                    StrSpecPur = " في إجراءات التخليص الجمركي لسيارة " + text;
                    ColName = "Col21";
                    DetermineCheckBox("نوع السيارة", 120, "للون:", 100, "رقم اللوحة:", 120, "رقم الشاسية:", 270, "سنة الموديل:", 80, "", "", strListcomb);
                    break;
                case "سيارة - بيع":
                    StrSpecPur = " في بيع سيارة " + text;
                    ColName = "Col21";
                    DetermineCheckBox("نوع السيارة", 120, "للون:", 100, "رقم اللوحة:", 120, "رقم الشاسية:", 270, "سنة الموديل:", 80, "", "", strListcomb);
                    break;
                case "دراسة جامعية":
                    ColName = "Col16";
                    DetermineCheckBox("إسم الجامعة:", 220, "اسم الكلية:", 200, "الدولة:", 120, "", 270, "", 80, "", "", strListcomb);
                    if (ComboProcedure.Text.Contains("خيارات متعددة"))
                        foreach (Control control in PanelSubItemBox.Controls)
                        {
                            control.Visible = false;
                        }
                    PanelSubItemBox.Visible = true;
                    foreach (Control control in PanelSubItemBox.Controls)
                    {

                        if (control is CheckBox && ((CheckBox)control).Name.Contains("checkBox"))
                        {
                            if (dataset[x] != "" && !dataset[x].Contains("خيارات متعددة"))
                            {
                                ((CheckBox)control).Text = dataset[x];
                                ((CheckBox)control).Visible = true;
                            }
                            x++;
                        }
                    }
                    break;
            }
            
            //if(database) 
            DetermineData(SI[0], SI[1], SI[2], SI[3], SI[4], SI[5], SI[6]);
            //else MessageBox.Show(textdata + "  " + SI[0] + "   " + SI[1] + "   " + SI[2] + "   " + SI[3] + "   " + SI[4] + "   " + SI[5]);
        }

        public void DetermineData(string v1, string v2, string v3, string v4, string v5, string v6, string v7)
        {
            txt1.Text = "30";
            txt2.Text = v2;
            txt3.Text = v3;
            txt4.Text = v4;
            txt5.Text = v5;
            checkSexType.Text = v6; 
            if (v6 != "")
            {                
                string[] YearMonthDay = v6.Split('-');
                int year, month, date;
                year = Convert.ToInt16(YearMonthDay[2]);
                month = Convert.ToInt16(YearMonthDay[0]);
                date = Convert.ToInt16(YearMonthDay[1]);
                DT.Value = new DateTime(year, month, date);
            }
            comboBox1.Text= v7;
            MessageBox.Show(v1);
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
                        if (xindex == 0) ListedRightIndex = "1";
                        else ListedRightIndex = ListedRightIndex + "_1";
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
                    
                    
                    authList2 = "ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr + " بالرقم (" + txt1.Text + ") بمربع رقم (" + txt2.Text + ") البالغ مساحتها(" + txt3.Text + "م.م) ب" + txt4.Text + " - " + txt5.Text + " ";
                    break;
                case 2:
                    
                    authList2 = "ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr + " بالرقم (" + txt1.Text + ") بمربع رقم (" + txt2.Text + ") البالغ مساحتها(" + txt3.Text + "م.م) ب" + txt4.Text + " - " + txt5.Text + " ";
                    break;
                case 3:
                    
                    authList2 = "ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr + " بالرقم (" + txt1.Text + ") بمربع رقم (" + txt2.Text + ") البالغ مساحتها(" + txt3.Text + "م.م) ب" + txt4.Text + " - " + txt5.Text + " ";
                    break;
                case 5:
                    //سيارة
                    
                    authList2 = "ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr + " من نوع " + txt1.Text + " موديل العام (" + txt5.Text + ") باللون " + txt2.Text + " ورقم لوحة (" + txt3.Text + " )وشاسيه بالرقم (" + txt4.Text + ") ";
                    break;
                case 6:
                    //

                    
                    LegaceyPreStr = " وبصفت" + preffix[ParentForm.intAppcases, 0] + " ضمن ورثة " + DPTitle + " " + txt1.Text + "، بموجب الإعلام الشرعي رقم (" + txt2.Text + ") الصادر من محكمة " + txt3.Text + " والتركة بالرقم (" + txt4.Text + ")";
                    authList2 = "ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + LegaceyAttch+ specificStr;
                    break;
                case 7:
                    //طلاق                    
                    
                    authList2 = "ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr ;
                    break;
                case 8:
                    //زواج                    
                    
                    authList2 = "ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + specificStr;
                    break;
                case 9:
                    //جامعة                 
                    foreach (Control control in PanelSubItemBox.Controls)
                    {
                        if (control is CheckBox && ((CheckBox)control).CheckState == CheckState.Checked)
                        {
                            if (x == 0) strUni = ((CheckBox)control).Text;
                            else strUni = strUni + "و" + ((CheckBox)control).Text;
                            x++;
                        }
                    }
                    authList2 = "ل" + preffix[ParentForm.intAuthcases, 7] + " ع" + preffix[ParentForm.intAppcases, 2] + " و" + preffix[ParentForm.intAuthcases, 8] + " مقام" + preffix[ParentForm.intAppcases, 0] + " في " + strUni + " بجامعة " + txt1.Text + " كلية " + txt2.Text + " ب" + txt3.Text;
                    break;
                default:
                    break;
            }
            collectItemsData();
        }

        private void collectItemsData()
        {
            BoxesData = txt1.Text;
            BoxesData = BoxesData + "_" + txt2.Text;
            BoxesData = BoxesData + "_" + txt3.Text;
            BoxesData = BoxesData + "_" + txt4.Text;
            BoxesData = BoxesData + "_" + txt5.Text;
            BoxesData = BoxesData + "_" + checkSexType.Text;
            BoxesData = BoxesData + "_" + DT.Text;
            BoxesData = BoxesData + "_" + comboBox1.Text;
        }
    }
}
