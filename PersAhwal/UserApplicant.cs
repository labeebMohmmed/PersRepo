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
using System.IO;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace PersAhwal
{
    public partial class UserApplicant : UserControl
    {
        string[] MaleFemale = new string[6];
        string[] AuthMaleFemale = new string[6];
        
        string[] EngMaleFemale = new string[6];
        string[] EngAuthMaleFemale = new string[6];
        
        string[] WitNessList = new string[4];
        string[] AuthNameList = new string[6];
        string[] AppNameList = new string[6];
        string[] AppDocNoList = new string[6];
        string[] AppIssueList = new string[6];
        string[] AppDocTypeList = new string[6];
        string[] AuthNationalityList = new string[6];
        string[] AuthNationalityIDList = new string[6];

        int appPersonlCount = 1, AuthPersonlCount = 1;
        DataTable Textboxtable;
        AutoCompleteStringCollection autoComplete;

        public Delegate AppMovePage;


        public Delegate strValueWit;

        public int strUAAuthPersonlCount
        {
            get { return AuthPersonlCount; }
            set { AuthPersonlCount = value; }
        }
        
        public Label strlabelrelateddoc
        {
            get { return labelrelateddoc; }
            set { labelrelateddoc = value; }
        }
        public TextBox strtxtPreRelated
        {
            get { return txtPreRelated; }
            set { txtPreRelated = value; }
        }
        
        public Label strlabebremovedDoc
        {
            get { return labebremovedDoc; }
            set { labebremovedDoc = value; }
        }
        public Label strlabelDocIsRemoved
        {
            get { return labelDocIsRemoved; }
            set { labelDocIsRemoved = value; }
        }
        public TextBox strtxtPreRemoved
        {
            get { return txtPreRemoved; }
            set { txtPreRemoved = value; }
        }
        
        public TextBox strالمهنة
        {
            get { return المهنة; }
            set { المهنة = value; }
        }
        public TextBox strتاريخ_الميلاد
        {
            get { return تاريخ_الميلاد; }
            set { تاريخ_الميلاد = value; }
        }
        public FlowLayoutPanel strFlowLayoutPanel1
        {
            get { return flowLayoutPanel1; }
            set { flowLayoutPanel1 = value; }
        }
        public int strUAappPersonlCount
        {
            get { return appPersonlCount; }
            set { appPersonlCount = value; }
        }

        public string[] strAppnameList
        {
            get { return AppNameList; }
            set { AppNameList = value; }
        }

        public string[] strAuthNationalityList
        {
            get { return AuthNationalityList; }
            set { AuthNationalityList = value; }
        }


        public string[] strAuthNationalityIDList
        {
            get { return AuthNationalityIDList; }
            set { AuthNationalityIDList = value; }
        }
        public string[] strUAAppDocNoList
        {
            get { return AppDocNoList; }
            set { AppDocNoList = value; }
        }
        public string[] strUAAppIssueList
        {
            get { return AppIssueList; }
            set { AppIssueList = value; }
        }
        public string[] strUAMaleFemale
        {
            get { return MaleFemale; }
            set { MaleFemale = value; }
        }
        public string[] strUAAppDocTypeList
        {
            get { return AppDocTypeList; }
            set { AppDocTypeList = value; }
        }
        public string[] strUAAuthMaleFemale
        {
            get { return AuthMaleFemale; }
            set { AuthMaleFemale = value; }
        }
        public string[] strEngAuthMaleFemale
        {
            get { return EngAuthMaleFemale; }
            set { EngAuthMaleFemale = value; }
        }
        public string[] strUAWitNessList
        {
            get { return WitNessList; }
            set { WitNessList = value; }
        }
        public string[] strUAAuthNameList
        {
            get { return AuthNameList; }
            set { AuthNameList = value; }
        }
        public string[] strUAAppnameList
        {
            get { return AppNameList; }
            set { AppNameList = value; }
        }

        public FlowLayoutPanel PanelAppValue
        {
            get { return Panelapp; }
            set { Panelapp = value; }
        }

        public FlowLayoutPanel PanelAuthValue
        {
            get { return PanelAuthPers; }
            set { PanelAuthPers = value; }
        }

        public FlowLayoutPanel PanelWitValue
        {
            get { return PanelWit; }
            set { PanelWit = value; }
        }

        public Form11 ParentFormApp { get; set; }
        public UserApplicant()
        {
            InitializeComponent();
            Panelapp.Height = 82;
            for (int x = 0; x < 6; x++)
            {
                AuthMaleFemale[x] = MaleFemale[x] = "ذكر";
                EngAuthMaleFemale[x] = EngMaleFemale[x] = "Mr.";
            }

            
        }



        private void AppInfoList()
        {
            int y1 = 0, y2 = 0, y3 = 0, y4 = 0;
            foreach (Control control in Panelapp.Controls)
            {
                if (control is TextBox)
                {

                    if (((TextBox)control).Name.Contains("AppName")) { AppNameList[y1] = ((TextBox)control).Text; y1++; }
                    if (((TextBox)control).Name.Contains("DocNo")) { AppDocNoList[y2] = ((TextBox)control).Text; y2++; }
                    if (((TextBox)control).Name.Contains("DocIssue")) { AppIssueList[y3] = ((TextBox)control).Text; y3++; }
                }
                if (control is ComboBox)
                {
                    if (((ComboBox)control).Name.Contains("DocType")) { AppDocTypeList[y4] = ((ComboBox)control).Text; y4++; }
                }
            }

        }

        private void AuthInfoList()
        {
            int y = 0, yy = 0, yyy = 0;
            foreach (Control control in PanelAuthPers.Controls)
            {
                if (control is TextBox)
                {
                    if (((TextBox)control).Name.Contains("txtAuthPerson")) { AuthNameList[y] = ((TextBox)control).Text; y++; }
                }
                if (control is ComboBox)
                {
                    if (((ComboBox)control).Name.Contains("nantionality")) { AuthNationalityList[yy] = ((ComboBox)control).Text; yy++; }
                }
                if (control is TextBox)
                {
                    if (((TextBox)control).Name.Contains("nantionalityID")) { AuthNationalityIDList[yyy] = ((TextBox)control).Text; yyy++; }
                }
            }
        }


        private void AppName1_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkSexType1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkSexType1.CheckState == CheckState.Unchecked) 
                checkSexType1.Text = "ذكر";
            else if (checkSexType1.CheckState == CheckState.Checked)
            {
                checkSexType1.Text = "أنثى";
                MaleFemale[0] = "أنثى";
            }
        }

        private void checkSexType2_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkSexType2.CheckState == CheckState.Unchecked) checkSexType2.Text = "ذكر";
            else
            {
                checkSexType2.Text = "أنثى";
                MaleFemale[1] = "أنثى";
            }
        }

        private void checkSexType3_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkSexType3.CheckState == CheckState.Unchecked) checkSexType3.Text = "ذكر";
            else
            {
                checkSexType3.Text = "أنثى";
                MaleFemale[2] = "أنثى";
            }
        }

        private void checkSexType4_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkSexType4.CheckState == CheckState.Unchecked) checkSexType4.Text = "ذكر";
            else
            {
                checkSexType4.Text = "أنثى";
                MaleFemale[3] = "أنثى";
            }
        }

        private void checkSexType5_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkSexType5.CheckState == CheckState.Unchecked) checkSexType5.Text = "ذكر";
            else
            {
                checkSexType5.Text = "أنثى";
                MaleFemale[4] = "أنثى";
            }
        }

        private void checkSexType6_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkSexType6.CheckState == CheckState.Unchecked) checkSexType6.Text = "ذكر";
            else
            {
                checkSexType6.Text = "أنثى";
                MaleFemale[5] = "أنثى";
            }
        }
        private void pictureBox1_Click_1(object sender, EventArgs e)
        {
            if (appPersonlCount == 1)
            {
                //PanelAuthPers.Location = new System.Drawing.Point(479, 150 + Panelapp.Height);
                //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount) * 82);
                Panelapp.Height += 82;
                appPersonlCount += 1;

            }
        }

        private void pictureBox2_Click_1(object sender, EventArgs e)
        {
            if (appPersonlCount > 1)
            {
                AppName1.Text = AppName2.Text;
                AppName2.Text = AppName3.Text;
                AppName3.Text = AppName4.Text;
                AppName4.Text = AppName5.Text;
                AppName5.Text = AppName6.Text;
                AppName6.Text = "";

                checkSexType1.CheckState = checkSexType2.CheckState;
                checkSexType2.CheckState = checkSexType3.CheckState;
                checkSexType3.CheckState = checkSexType4.CheckState;
                checkSexType4.CheckState = checkSexType5.CheckState;
                checkSexType5.CheckState = checkSexType6.CheckState;
                checkSexType6.CheckState = CheckState.Unchecked;

                DocType1.SelectedIndex = DocType2.SelectedIndex;
                DocType2.SelectedIndex = DocType3.SelectedIndex;
                DocType3.SelectedIndex = DocType4.SelectedIndex;
                DocType4.SelectedIndex = DocType5.SelectedIndex;
                DocType5.SelectedIndex = DocType6.SelectedIndex;
                DocType6.SelectedIndex = 0;

                DocNo1.Text = DocNo2.Text;
                DocNo2.Text = DocNo3.Text;
                DocNo3.Text = DocNo4.Text;
                DocNo4.Text = DocNo5.Text;
                DocNo5.Text = DocNo6.Text;
                DocNo6.Text = "";

                DocIssue1.Text = DocIssue1.Text;
                DocIssue2.Text = DocIssue3.Text;
                DocIssue3.Text = DocIssue4.Text;
                DocIssue4.Text = DocIssue5.Text;
                DocIssue5.Text = DocIssue6.Text;



                DocIssue6.Text = "";
                Panelapp.Height -= 82;
                appPersonlCount -= 1;


            }
            else
            {
                AppName1.Text = "";

                DocType1.SelectedIndex = 0;

                DocNo1.Text = "";

                DocIssue1.Text = "";
                appPersonlCount = 1;
            }
            //PanelAuthPers.Location = new System.Drawing.Point(479, 150 + Panelapp.Height - 82);
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);

        }

        private void pictureBox3_Click_1(object sender, EventArgs e)
        {
            if (appPersonlCount == 2)
            {

                //PanelAuthPers.Location = new System.Drawing.Point(479, 150 + Panelapp.Height);
                //PanelWit.Location = new System.Drawing.Point(443, 207 + AuthPersonlCount * 41 + (appPersonlCount - 1) * 82);
                //appPanel.Height = 3 * 82;
                appPersonlCount += 1;
                Panelapp.Height += 82;
            }
        }

        private void pictureBox4_Click_1(object sender, EventArgs e)
        {
            AppName2.Text = AppName3.Text;
            AppName3.Text = AppName4.Text;
            AppName4.Text = AppName5.Text;
            AppName5.Text = AppName6.Text;
            AppName6.Text = "";

            DocType2.SelectedIndex = DocType3.SelectedIndex;
            DocType3.SelectedIndex = DocType4.SelectedIndex;
            DocType4.SelectedIndex = DocType5.SelectedIndex;
            DocType5.SelectedIndex = DocType6.SelectedIndex;
            DocType6.SelectedIndex = 0;

            DocNo2.Text = DocNo3.Text;
            DocNo3.Text = DocNo4.Text;
            DocNo4.Text = DocNo5.Text;
            DocNo5.Text = DocNo6.Text;
            DocNo6.Text = "";

            DocIssue2.Text = DocIssue3.Text;
            DocIssue3.Text = DocIssue4.Text;
            DocIssue4.Text = DocIssue5.Text;
            DocIssue5.Text = DocIssue6.Text;
            DocIssue6.Text = "";
            Panelapp.Height -= 82;
            appPersonlCount -= 1;
        //    PanelAuthPers.Location = new System.Drawing.Point(479, 150 + Panelapp.Height - 82);
        //    PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
        }

        private void pictureBox5_Click_1(object sender, EventArgs e)
        {
            if (appPersonlCount == 3)
            {
                //PanelAuthPers.Location = new System.Drawing.Point(479, 150 + Panelapp.Height);
                //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);

                //appPanel.Height = 3 * 82;
                appPersonlCount += 1;
                Panelapp.Height += 82;
            }
        }

        private void pictureBox6_Click_1(object sender, EventArgs e)
        {
            AppName3.Text = AppName4.Text;
            AppName4.Text = AppName5.Text;
            AppName5.Text = AppName6.Text;
            AppName6.Text = "";


            DocType3.SelectedIndex = DocType4.SelectedIndex;
            DocType4.SelectedIndex = DocType5.SelectedIndex;
            DocType5.SelectedIndex = DocType6.SelectedIndex;
            DocType6.SelectedIndex = 0;


            DocNo3.Text = DocNo4.Text;
            DocNo4.Text = DocNo5.Text;
            DocNo5.Text = DocNo6.Text;
            DocNo6.Text = "";


            DocIssue3.Text = DocIssue4.Text;
            DocIssue4.Text = DocIssue5.Text;
            DocIssue5.Text = DocIssue6.Text;
            DocIssue6.Text = "";
            Panelapp.Height -= 82;
            appPersonlCount -= 1;
            //PanelAuthPers.Location = new System.Drawing.Point(479, 150 + Panelapp.Height - 82);
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
        }

        private void pictureBox7_Click_1(object sender, EventArgs e)
        {
            if (appPersonlCount == 4)
            {
                //PanelAuthPers.Location = new System.Drawing.Point(479, 150 + Panelapp.Height);
                //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
                appPersonlCount += 1;
                Panelapp.Height += 82;
            }
        }

        private void pictureBox8_Click_1(object sender, EventArgs e)
        {
            AppName4.Text = AppName5.Text;
            AppName5.Text = AppName6.Text;
            AppName6.Text = "";



            DocType4.SelectedIndex = DocType5.SelectedIndex;
            DocType5.SelectedIndex = DocType6.SelectedIndex;
            DocType6.SelectedIndex = 0;



            DocNo4.Text = DocNo5.Text;
            DocNo5.Text = DocNo6.Text;
            DocNo6.Text = "";



            DocIssue4.Text = DocIssue5.Text;
            DocIssue5.Text = DocIssue6.Text;
            DocIssue6.Text = "";
            Panelapp.Height -= 82;
            appPersonlCount -= 1;
            //PanelAuthPers.Location = new System.Drawing.Point(479, 150 + Panelapp.Height - 82);
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
        }

        private void pictureBox9_Click_1(object sender, EventArgs e)
        {
            if (appPersonlCount == 5)
            {
                //PanelAuthPers.Location = new System.Drawing.Point(479, 150 + Panelapp.Height);
                //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
                //appPanel.Height = 3 * 82;
                appPersonlCount += 1;
                Panelapp.Height += 82;
            }
        }

        private void pictureBox10_Click_1(object sender, EventArgs e)
        {
            AppName5.Text = AppName6.Text;
            AppName6.Text = "";
            DocType5.SelectedIndex = DocType6.SelectedIndex;
            DocType6.SelectedIndex = 0;
            DocNo5.Text = DocNo6.Text;
            DocNo6.Text = "";
            DocIssue5.Text = DocIssue6.Text;
            DocIssue6.Text = "";
            Panelapp.Height -= 82;
            appPersonlCount -= 1;
            //PanelAuthPers.Location = new System.Drawing.Point(479, 150 + Panelapp.Height - 82);
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
        }

        private void pictureBox12_Click_1(object sender, EventArgs e)
        {
            AppName6.Text = "";
            DocType6.SelectedIndex = 0;
            DocNo6.Text = "";
            DocIssue6.Text = "";
            Panelapp.Height -= 82;
            //PanelAuthPers.Location = new System.Drawing.Point(479, 150 + Panelapp.Height - 82);
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            AppInfoList();
            AuthInfoList();
            witInfoList();
            ParentFormApp.AppDataFilled = true;
            AppMovePage.DynamicInvoke(3);
        }
        private void updateGenName(string idDoc, string birth, string job, string source)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string query = "update TableAuth set تاريخ_الميلاد=N'" + birth + "',المهنة=N'"+job+"' where ID = '" + idDoc + "'";
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void witInfoList()
        {
            WitNessList[0] = txtWitName1.Text;
            WitNessList[1] = txtWitName2.Text;
            WitNessList[2] = txtWitPass1.Text;
            WitNessList[3] = txtWitPass2.Text;
        }

        private void btnSizeSpecial_Click_1(object sender, EventArgs e)
        {
            AppInfoList();
            AuthInfoList();
            witInfoList();
            AppMovePage.DynamicInvoke(1);
        }

        private void pictureBox11_Click_1(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 1)
            {
                //PanelWit.Location = new System.Drawing.Point(443, 207 + AuthPersonlCount * 41 + (appPersonlCount - 1) * 82);
                PanelAuthPers.Height += 82;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox14_Click_1(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 2)
            {
                //PanelWit.Location = new System.Drawing.Point(443, 207 + AuthPersonlCount * 41 + (appPersonlCount - 1) * 82);
                PanelAuthPers.Height += 82;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox16_Click_1(object sender, EventArgs e)
        {
            //PanelWit.Location = new System.Drawing.Point(443, 207 + AuthPersonlCount * 41 + (appPersonlCount - 1) * 82);
            if (AuthPersonlCount == 3)
            {
                PanelAuthPers.Height += 82;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox18_Click_1(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 4)
            {
                //PanelWit.Location = new System.Drawing.Point(443, 207 + AuthPersonlCount * 41 + (appPersonlCount - 1) * 82);
                PanelAuthPers.Height += 82;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox20_Click_1(object sender, EventArgs e)
        {
            if (AuthPersonlCount == 5)
            {
                //PanelWit.Location = new System.Drawing.Point(443, 207 + AuthPersonlCount * 41 + (appPersonlCount - 1) * 82);
                PanelAuthPers.Height += 82;
                AuthPersonlCount += 1;
            }
        }

        private void pictureBox23_Click(object sender, EventArgs e)
        {
            //PanelWit.Location = new System.Drawing.Point(443, 207 + AuthPersonlCount * 41 + (appPersonlCount - 1) * 82);
            if (AuthPersonlCount == 6)
            {
                PanelAuthPers.Height += 82;
                AuthPersonlCount += 1;
            }
        }



        private void pictureBox13_Click_1(object sender, EventArgs e)
        {
            if (AuthPersonlCount > 1)
            {
                nantionalityID1.Text = nantionalityID2.Text;
                nantionalityID2.Text = nantionalityID3.Text;
                nantionalityID3.Text = nantionalityID4.Text;
                nantionalityID4.Text = nantionalityID5.Text;
                nantionalityID5.Text = "";

                nantionality1.Text = nantionality2.Text;
                nantionality2.Text = nantionality3.Text;
                nantionality3.Text = nantionality4.Text;
                nantionality4.Text = nantionality5.Text;
                nantionality5.Text = "";


                txtAuthPerson1.Text = txtAuthPerson2.Text;
                txtAuthPerson2.Text = txtAuthPerson3.Text;
                txtAuthPerson3.Text = txtAuthPerson4.Text;
                txtAuthPerson4.Text = txtAuthPerson5.Text;
                txtAuthPerson5.Text = "";

                txtAuthPersonsex1.CheckState = txtAuthPersonsex2.CheckState;
                txtAuthPersonsex2.CheckState = txtAuthPersonsex3.CheckState;
                txtAuthPersonsex3.CheckState = txtAuthPersonsex4.CheckState;
                txtAuthPersonsex4.CheckState = txtAuthPersonsex5.CheckState;
                txtAuthPersonsex5.CheckState = CheckState.Unchecked;

                PanelAuthPers.Height -= 82;
                AuthPersonlCount -= 1;
                //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
            }
            else
            {
                nantionalityID5.Text = "";
                nantionality5.Text = "";
                txtAuthPerson1.Text = "";
                PanelAuthPers.Height = 82;
                txtAuthPersonsex1.CheckState = CheckState.Unchecked;
                AuthPersonlCount = 1;
                //PanelWit.Location = new System.Drawing.Point(443, 207 + (appPersonlCount - 1) * 82);
            }

        }

        private void pictureBox15_Click_1(object sender, EventArgs e)
        {
            nantionalityID2.Text = nantionalityID3.Text;
            nantionalityID3.Text = nantionalityID4.Text;
            nantionalityID4.Text = nantionalityID5.Text;
            nantionalityID5.Text = "";

            
            nantionality2.Text = nantionality3.Text;
            nantionality3.Text = nantionality4.Text;
            nantionality4.Text = nantionality5.Text;
            nantionality5.Text = "";

            txtAuthPerson2.Text = txtAuthPerson3.Text;
            txtAuthPerson3.Text = txtAuthPerson4.Text;
            txtAuthPerson4.Text = txtAuthPerson5.Text;
            txtAuthPerson5.Text = "";


            txtAuthPersonsex2.CheckState = txtAuthPersonsex3.CheckState;
            txtAuthPersonsex3.CheckState = txtAuthPersonsex4.CheckState;
            txtAuthPersonsex4.CheckState = txtAuthPersonsex5.CheckState;
            txtAuthPersonsex5.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 82;
            AuthPersonlCount -= 1;
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
        }

        private void pictureBox17_Click_1(object sender, EventArgs e)
        {
            nantionalityID3.Text = nantionalityID4.Text;
            nantionalityID4.Text = nantionalityID5.Text;
            nantionalityID5.Text = "";

            nantionality3.Text = nantionality4.Text;
            nantionality4.Text = nantionality5.Text;
            nantionality5.Text = "";

            txtAuthPerson3.Text = txtAuthPerson4.Text;
            txtAuthPerson4.Text = txtAuthPerson5.Text;
            txtAuthPerson5.Text = "";

            txtAuthPersonsex3.CheckState = txtAuthPersonsex4.CheckState;
            txtAuthPersonsex4.CheckState = txtAuthPersonsex5.CheckState;
            txtAuthPersonsex5.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 82;
            AuthPersonlCount -= 1;
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
        }

        private void pictureBox19_Click_1(object sender, EventArgs e)
        {
            nantionalityID4.Text = nantionalityID5.Text;
            nantionalityID5.Text = "";

            nantionality4.Text = nantionality5.Text;
            nantionality5.Text = "";

            txtAuthPerson4.Text = txtAuthPerson5.Text;
            txtAuthPerson5.Text = "";


            txtAuthPersonsex4.CheckState = txtAuthPersonsex5.CheckState;
            txtAuthPersonsex5.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 82;
            AuthPersonlCount -= 1;
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
        }

        private void pictureBox21_Click_1(object sender, EventArgs e)
        {
            nantionalityID5.Text = "";

            nantionality5.Text = "";

            txtAuthPerson5.Text = "";


            txtAuthPersonsex5.CheckState = CheckState.Unchecked;

            PanelAuthPers.Height -= 82;
            AuthPersonlCount -= 1;
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
        }

        private void txtAuthPersonsex1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (txtAuthPersonsex1.CheckState == CheckState.Unchecked) txtAuthPersonsex1.Text = "ذكر";
            else
            {
                txtAuthPersonsex1.Text = "أنثى";
                AuthMaleFemale[0] = "أنثى";
            }
        }

        private void txtAuthPersonsex2_CheckedChanged_1(object sender, EventArgs e)
        {
            if (txtAuthPersonsex2.CheckState == CheckState.Unchecked) txtAuthPersonsex2.Text = "ذكر";
            else
            {
                txtAuthPersonsex2.Text = "أنثى";
                AuthMaleFemale[1] = "أنثى";
            }
        }

        private void txtAuthPersonsex3_CheckedChanged_1(object sender, EventArgs e)
        {
            if (txtAuthPersonsex3.CheckState == CheckState.Unchecked) txtAuthPersonsex3.Text = "ذكر";
            else
            {
                txtAuthPersonsex3.Text = "أنثى";
                AuthMaleFemale[2] = "أنثى";
            }
        }

        private void txtAuthPersonsex4_CheckedChanged_1(object sender, EventArgs e)
        {
            if (txtAuthPersonsex4.CheckState == CheckState.Unchecked) txtAuthPersonsex4.Text = "ذكر";
            else
            {
                txtAuthPersonsex4.Text = "أنثى";
                AuthMaleFemale[3] = "أنثى";
            }
        }

        private void txtAuthPersonsex5_CheckedChanged_1(object sender, EventArgs e)
        {
            if (txtAuthPersonsex5.CheckState == CheckState.Unchecked) txtAuthPersonsex5.Text = "ذكر";
            else
            {
                txtAuthPersonsex5.Text = "أنثى";
                AuthMaleFemale[4] = "أنثى";
            }
        }
        public int Authcases()
        {
            if (AuthPersonlCount == 1 && txtAuthPersonsex1.CheckState == CheckState.Unchecked) return 0;
            else if (AuthPersonlCount == 1 && txtAuthPersonsex1.CheckState == CheckState.Checked) return 1;
            else if (AuthPersonlCount == 2 && txtAuthPersonsex1.CheckState == CheckState.Unchecked && txtAuthPersonsex2.CheckState == CheckState.Unchecked) return 2;

            else if (AuthPersonlCount == 2 && txtAuthPersonsex1.CheckState == CheckState.Checked && txtAuthPersonsex2.CheckState == CheckState.Unchecked) return 2;
            else if (AuthPersonlCount == 2 && txtAuthPersonsex1.CheckState == CheckState.Unchecked && txtAuthPersonsex2.CheckState == CheckState.Checked) return 2;

            else if (AuthPersonlCount == 2 && txtAuthPersonsex1.CheckState == CheckState.Checked && txtAuthPersonsex2.CheckState == CheckState.Checked) return 3;
            else if (AuthPersonlCount == 3 && txtAuthPersonsex1.CheckState == CheckState.Checked && txtAuthPersonsex2.CheckState == CheckState.Checked && txtAuthPersonsex3.CheckState == CheckState.Checked) return 4;
            else return 5;
        }

        
        private void UserApplicant_Load(object sender, EventArgs e)
        {
            
        }

        private void autoCompleteTextBox(TextBox combbox, string source, string comlumnName, string tableName)
        {
            using (SqlConnection saConn = new SqlConnection(source))
            {
                try
                {
                    saConn.Open();

                    string query = "select " + comlumnName + " from " + tableName;
                    SqlCommand cmd = new SqlCommand(query, saConn);
                    cmd.ExecuteNonQuery();

                    Textboxtable = new DataTable();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                    dataAdapter.Fill(Textboxtable);

                    foreach (DataRow dataRow in Textboxtable.Rows)
                    {
                        autoComplete.Add(dataRow[comlumnName].ToString());
                    }
                    combbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                    combbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                    combbox.AutoCompleteCustomSource = autoComplete;
                    saConn.Close();
                }
                catch (Exception ex) { }
            }
        }
        private void UserApplicant_Click(object sender, EventArgs e)
        {
            Panelapp.Height = 82;
            PanelAuthPers.Height = 82;
            //PanelAuthPers.Location = new System.Drawing.Point(479, 150);
            //PanelWit.Location = new System.Drawing.Point(443, 207);
        }

        private void Panelapp_MouseEnter_1(object sender, EventArgs e)
        {
            Panelapp.Height = appPersonlCount * 82;
            //PanelAuthPers.Location = new System.Drawing.Point(479, 150 + (appPersonlCount - 1) * 82);
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
        }

        private void PanelAuthPers_MouseEnter_1(object sender, EventArgs e)
        {
            PanelAuthPers.Height = AuthPersonlCount * 82;
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (appPersonlCount - 1) * 82);
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
        }

        private void DocType1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            SetDocNoText(DocType1.SelectedIndex, DocNo1, labeldoctype1);
        }

        private void SetDocNoText(int selectedIndex, TextBox text, Label label)
        {
            switch (DocType1.SelectedIndex)
            {
                case 0:
                    text.Text = "P0";
                    label.Text = "جواز سفر";
                    break;
                case 1:
                    text.Text = "";
                    label.Text = "إقامة";
                    break;
                case 2:
                    text.Text = "";
                    label.Text = "رقم وطني";
                    break;
            }
        }

        private void DocType2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            SetDocNoText(DocType1.SelectedIndex, DocNo2, labeldoctype2);
        }

        private void DocType3_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            SetDocNoText(DocType1.SelectedIndex, DocNo3, labeldoctype3);
        }

        private void DocType4_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            SetDocNoText(DocType1.SelectedIndex, DocNo4, labeldoctype4);
        }

        private void DocType5_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            SetDocNoText(DocType1.SelectedIndex, DocNo5, labeldoctype5);
        }

        private void DocType6_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            SetDocNoText(DocType1.SelectedIndex, DocNo6, labeldoctype6);
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        public int Appcases()
        {
            if (appPersonlCount == 1 && checkSexType1.CheckState == CheckState.Unchecked) return 0;
            else if (appPersonlCount == 1 && checkSexType1.CheckState == CheckState.Checked) return 1;
            else if (appPersonlCount == 2 && checkSexType1.CheckState == CheckState.Unchecked && checkSexType2.CheckState == CheckState.Unchecked) return 2;

            else if (appPersonlCount == 2 && checkSexType1.CheckState == CheckState.Checked && checkSexType2.CheckState == CheckState.Unchecked) return 2;
            else if (appPersonlCount == 2 && checkSexType1.CheckState == CheckState.Unchecked && checkSexType2.CheckState == CheckState.Checked) return 2;

            else if (appPersonlCount == 2 && checkSexType1.CheckState == CheckState.Checked && checkSexType2.CheckState == CheckState.Checked) return 3;
            else if (appPersonlCount == 3 && checkSexType1.CheckState == CheckState.Checked && checkSexType2.CheckState == CheckState.Checked && checkSexType2.CheckState == CheckState.Checked) return 4;
            else return 5;
        }

        private void OnceSettings()
        {



        }

        private void GroupFile(Control.ControlCollection controls, string text1, string text2, string text3, string text4, string text5)
        {
            foreach (Control control in controls)
            {
                autoComplete = new AutoCompleteStringCollection();

                if (control is TextBox && control.Name.Contains(text1))
                {
                    if (text2 != "") autoCompleteTextBox(((TextBox)control), ParentFormApp.PublicDataSource, text2, "TableAuth");
                    if (text3 != "") autoCompleteTextBox(((TextBox)control), ParentFormApp.PublicDataSource, text3, "TableAuth");
                    if (text4 != "") autoCompleteTextBox(((TextBox)control), ParentFormApp.PublicDataSource, text4, "TableAuth");
                    if (text5 != "") autoCompleteTextBox(((TextBox)control), ParentFormApp.PublicDataSource, text5, "TableAuth");
                }
            }
        }
        
        private void AppName1_MouseEnter(object sender, EventArgs e)
        {
            //GroupFile(Panelapp.Controls, "AppName", "الموكَّل", "الشاهد_الأول", "الشاهد_الثاني", "");
        }

        private void UserApplicant_MouseEnter(object sender, EventArgs e)
        {

            //PanelAuthPers.Location = new System.Drawing.Point(479, 150 + (appPersonlCount - 1) * 82);
            //PanelWit.Location = new System.Drawing.Point(443, 207 + (AuthPersonlCount - 1) * 41 + (appPersonlCount - 1) * 82);
        }

        private void UserApplicant_MouseHover(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            button7.Enabled = false;
            //MessageBox.Show(ParentFormApp.idAuthTable);
            FillDatafromGenArch("data1", ParentFormApp.idAuthTable, "TableAuth");
            button7.Enabled = true;
        }
        private void OpenFile(int id, int fileNo)
        {
            string query;

            SqlConnection Con = new SqlConnection(ParentFormApp.DataSource);
            if (fileNo == 1)
            {
                query = "select Data1, Extension1,ارشفة_المستندات from TableAuth where ID=@id";
            }
            else if (fileNo == 2)
            {
                query = "select Data2, Extension2,المكاتبة_النهائية from TableAuth where ID=@id";
            }
            else query = "select Data3, Extension3,DocxData from TableAuth where ID=@id";
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@Id", SqlDbType.Int).Value = id;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                if (fileNo == 1)
                {
                    var name = reader["ارشفة_المستندات"].ToString();
                    if (string.IsNullOrEmpty(name)) return;
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ss")) + ext;
                    while (File.Exists(NewFileName)) NewFileName = name.Replace(ext, DateTime.Now.ToString("ss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }
                else if (fileNo == 2)
                {

                    var name = reader["المكاتبة_النهائية"].ToString();
                    if (string.IsNullOrEmpty(name)) return;
                    var Data = (byte[])reader["Data2"];
                    var ext = reader["Extension2"].ToString();
                    var NewFileName = name.Split('.')[0] + DateTime.Now.ToString("ss") + "." + name.Split('.')[1];
                    while (File.Exists(NewFileName)) NewFileName = name.Replace(ext, DateTime.Now.ToString("ss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }
                else
                {
                    var name = reader["DocxData"].ToString();
                    if (string.IsNullOrEmpty(name)) return;
                    var Data = (byte[])reader["Data3"];
                    var ext = name.Split('.')[1];
                    var NewFileName = name.Split('.')[0] + DateTime.Now.ToString("ss") +"."+ name.Split('.')[1];
                    while(File.Exists(NewFileName)) NewFileName = name.Replace(ext, DateTime.Now.ToString("ss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    openFile3(NewFileName, id);
                    System.Diagnostics.Process.Start(NewFileName);
                    
                }

            }
            Con.Close();


        }
        private void openFile3(string filename3, int id)
        {
            SqlConnection sqlCon = new SqlConnection(ParentFormApp.PublicDataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET DocxData=@DocxData,fileUpload=@fileUpload WHERE ID = @ID", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.Parameters.AddWithValue("@DocxData", filename3);
            sqlCmd.Parameters.AddWithValue("@fileUpload", "No");
            sqlCmd.ExecuteNonQuery();



        }
        private void button3_Click(object sender, EventArgs e)
        {

            //OpenFile(Convert.ToInt32(ParentFormApp.iddocAuthTable ), 2);
            button3.Enabled = false;
            //MessageBox.Show(ParentFormApp.idAuthTable);
            FillDatafromGenArch("data2",ParentFormApp.idAuthTable, "TableAuth");
            button3.Enabled = true;
        }

        void FillDatafromGenArch(string doc, string id, string table)
        {
            SqlConnection sqlCon = new SqlConnection(ParentFormApp.DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableGeneralArch where  رقم_المرجع='" + id + "' and نوع_المستند='" + doc + "' and docTable='" + table + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow reader in dtbl.Rows)
            {
                var name = reader["المستند"].ToString();
                var Data = (byte[])reader["Data1"];
                var ext = reader["Extension1"].ToString();
                var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                File.WriteAllBytes(NewFileName, Data);
                System.Diagnostics.Process.Start(NewFileName);
            }


            sqlCon.Close();
        }
        private void txtAuthPerson1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFile(Convert.ToInt32(ParentFormApp.idAuthTable), 3);
            
        }

        private void autoCompleteComboBox(ComboBox combbox, string source, string comlumnName, string tableName)
        {
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();
                foreach (DataRow dataRow in table.Rows)
                {
                    autoComplete.Add(dataRow[comlumnName].ToString());
                }
                combbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                combbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                combbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }

        private void lang_CheckedChanged(object sender, EventArgs e)
        {
            if (lang.CheckState == CheckState.Unchecked)
            {
                lang.Text = "العربية";
                labeltitle12.Visible = labeltitle13.Visible = false;
                PanelWit.Width = 640;
                fileComboBox(ParentFormApp.txtAttendVCValue, ParentFormApp.PublicDataSource, "ArabicAttendVC", "TableListCombo");
            }
            else if (lang.CheckState == CheckState.Checked)
            {
                lang.Text = "الانجليزية";
                labeltitle12.Visible = labeltitle13.Visible = true;
                PanelWit.Width = 759;
                fileComboBox(ParentFormApp.txtAttendVCValue, ParentFormApp.PublicDataSource, "EnglishAttendVC", "TableListCombo");
                
            }
            EnglishTitle(Convert.ToBoolean(lang.CheckState), Panelapp);
            EnglishTitle(Convert.ToBoolean(lang.CheckState), PanelAuthPers);
            EnglishTitle(Convert.ToBoolean(lang.CheckState), PanelWit);
            
        }

        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName)
        {
            combbox.Visible = true;
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
                    if (!String.IsNullOrEmpty(dataRow[comlumnName].ToString())) combbox.Items.Add(dataRow[comlumnName].ToString());
                }
                saConn.Close();
            }
        }

        private void txtWitName1_TextChanged(object sender, EventArgs e)
        {

        }

        private void التاريخ_ValueChanged(object sender, EventArgs e)
        {
            
        }
        string lastInput2 = "";
        private void تاريخ_الميلاد_TextChanged(object sender, EventArgs e)
        {
            if (تاريخ_الميلاد.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(تاريخ_الميلاد.Text, 1, 2));
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //تاريخ_الميلاد.Text = "";
                    تاريخ_الميلاد.Text = SpecificDigit(تاريخ_الميلاد.Text, 3, 10);
                    return;
                }
            }
            
            if (تاريخ_الميلاد.Text.Length == 11)
            {
                تاريخ_الميلاد.Text = lastInput2; return;
            }
            if (تاريخ_الميلاد.Text.Length == 10) return;
            if (تاريخ_الميلاد.Text.Length == 4) تاريخ_الميلاد.Text = "-" + تاريخ_الميلاد.Text;
            else if (تاريخ_الميلاد.Text.Length == 7) تاريخ_الميلاد.Text = "-" + تاريخ_الميلاد.Text;
            lastInput2 = تاريخ_الميلاد.Text;
        }
        private string SpecificDigit(string text, int Firstdigits, int Lastdigits)
        {
            char[] characters = text.ToCharArray();
            string firstNchar = "";
            int z = 0;
            for (int x = Firstdigits - 1; x < Lastdigits && x < text.Length; x++)
            {
                firstNchar = firstNchar + characters[x];

            }
            return firstNchar;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (تاريخ_الميلاد.Text.Length != 10)
            {
                MessageBox.Show("يرجى إدخال تاريخ ميلاد مقدم الطلب أولا");
                return;
            }

            if (المهنة.Text == "")
            {
                MessageBox.Show("يرجى إختيار مهنة مقدم الطلب"); return;
            }
            
            updateGenName(ParentFormApp.idAuthTable, تاريخ_الميلاد.Text, المهنة.Text, ParentFormApp.DataSource);
            تاريخ_الميلاد.Text = المهنة.Text = "";
            btnSizeSpecial.PerformClick();
        }

        private void تاريخ_الميلاد_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                button4.PerformClick();
            }
        }

        private void EnglishTitle(bool visibility, Control controls) {
            
            foreach (Control control in controls.Controls)
            {
                if (control is ComboBox)
                {
                    if (((ComboBox)control).Name.Contains("combTitle"))
                    {
                        ((ComboBox)control).Width = 55;
                        ((ComboBox)control).Visible = visibility;
                    }
                }
                if (control is CheckBox)
                {
                    if (((CheckBox)control).Name.Contains("checkSexType"))
                    {
                        ((CheckBox)control).Visible = !visibility;
                    }
                }
            }
        }
        

    }
}
