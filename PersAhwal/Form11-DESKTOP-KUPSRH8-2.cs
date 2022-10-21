using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Data.SqlClient;
using System.Configuration;
using System.Globalization;
using System.Threading;
using System.IO;
namespace PersAhwal
{
    public partial class Form11 : Form
    {
        public delegate void AppNameFun(string[] values);
        private event AppNameFun AppNamePointer;
        public delegate void AppDocNoFun(string[] values);
        private event AppDocNoFun AppDocNoPointer;
        public delegate void AppDocIssueFun(string[] values);
        private event AppDocIssueFun AppDocIssuePointer;
        public delegate void AppDocTypeFun(string[] values);
        private event AppDocTypeFun AppDocTypePointer;
        public delegate void AppSexFun(string[] values);
        private event AppSexFun AppSexPointer;
        public delegate void AppMaleFemalFun(string[] values);
        private event AppMaleFemalFun AppMaleFemalPointer;
        public delegate void AppMovePageFun(int values);
        private event AppMovePageFun AppMovePagePointer;
        public delegate void DataMovePageFun(int values);
        private event DataMovePageFun DataMovePagePointer;
        public delegate void AuthNameFun(string[] values);
        private event AuthNameFun AuthNamePointer;
        public delegate void AuthSexFun(string[] values);
        private event AuthSexFun AuthSexPointer;
        public delegate void AuthMaleFemalFun(string[] values);
        private event AuthMaleFemalFun AuthMaleFemalPointer;
        public delegate void AuthMovePageFun(int values);
        private event AuthMovePageFun AuthMovePagePointer;
        public delegate void WitValuesFun(string[] values);
        private event WitValuesFun WitPointer;
        public delegate void WitMovePageFun(int values);
        private event WitMovePageFun WitMovePagePointer;

        public delegate void AuthcasesFun(int values);
        private event AuthcasesFun AuthcasesPointer;

        public delegate void AppcasesFun(int values);
        private event AppcasesFun AppcasesPointer;

        public delegate void AuthCountFun(int values);
        private event AuthCountFun AuthCountPointer;

        public delegate void AppCountsFun(int values);
        private event AppCountsFun AppCountsPointer;

        public delegate void strRightsFun(string values);
        private event strRightsFun strRightsPointer;

        public delegate void strRightsIndexFun(string values);
        private event strRightsIndexFun stRightsIndexPointer;

        public delegate void strAuthList2Fun(string values);
        private event strAuthList2Fun strAuthList2Pointer;

        public delegate void strAuthSubjectFun(string values);
        private event strAuthSubjectFun strAuthSubjectPointer;
        
        
        private string DataSource = "Data Source = (LocalDB)\\MSSQLLocalDB;Initial Catalog = myDataBase; Integrated Security = True";

        private static string[] AppnameList = new string[6];
        private static string[] AppMaleFemaleList = new string[6];
        private static string[] AppSexList = new string[6];
        private static string[] AppissueList = new string[6] { "", "", "", "", "", "" };
        private static string[] AppDocTypeList = new string[6];
        private static string[] AppDocNoList = new string[6];
        private static string[] AuthMaleFemale = new string[6];
        private static string[] AuthSex = new string[6];
        private static string[] AuthNames = new string[6];
        private static string[] WitValuesList = new string[4];
        string DBAuthPersNames = "", DBAuthPersSexs = "", DBAppPersNames = "", DBAppDoctype = "", DBAppDocNo = "", DBAppDocIssue = "", DBAppSexs = "";
        static string[,] preffix = new string[10, 20];
        string FilespathIn, FilespathOut;
        private int AuthCount, AppCounts;
        private int Authcases, Appcases;
        DataTable UserTexttable;
        DataTable Textboxtable;
        AutoCompleteStringCollection autoComplete;

        string strRights = "", authList1 = "", authList2 = "", AuthSubject="", ListedRightIndex = "";
        public DataTable UserComboTexttable
        {
            get { return UserTexttable; }
            set { UserTexttable = value; }
        }

        public int intAuthCount
        {
            get { return AuthCount; }
            set { AuthCount = value; }
        }

        public int intAppcases
        {
            get { return Appcases; }
            set { Appcases = value; }
        }

        public int intAuthcases
        {
            get { return Authcases; }
            set { Authcases = value; }
        }

        public int intAppCounts
        {
            get { return AppCounts; }
            set { AppCounts = value; }
        }

        public string PublicDataSource
        {
            get { return DataSource; }
            set { DataSource = value; }
        }
        public string[] strAppnameList
        {
            get { return AppnameList; }
            set { AppnameList = value; }
        }

        public string[] strAppMaleFemaleList
        {
            get { return AppMaleFemaleList; }
            set { AppMaleFemaleList = value; }
        }

        public string[] strAppSexList
        {
            get { return AppSexList; }
            set { AppSexList = value; }
        }

        public string[] strAuthNames
        {
            get { return AuthNames; }
            set { AuthNames = value; }
        }

        public string[] strAuthMaleFemale
        {
            get { return AuthMaleFemale; }
            set { AuthMaleFemale = value; }
        }

        public string[] strAuthSex
        {
            get { return AuthSex; }
            set { AuthSex = value; }
        }

        public string strauthList1
        {
            get { return authList1; }
            set { authList1 = value; }
        }
        
        public Form11 ParentForm { get; set; }
        public Form11(string source, string filespathIn, string filespathOut)
        {
            DataSource = source;
            InitializeComponent();
            FilespathIn = filespathIn;
            FilespathOut = filespathOut;
            this.userAuthText1.ParentForm = this;
            this.userApplicant1.ParentFormApp = this;
            Suffex_preffixList();
            foreach (Control control in this.Controls)
            {
                if (control is UserControl)
                {
                    ((UserControl)control).Hide();
                }
            }
            userAuthText1.Show();
            txtAttendVC.SelectedIndex = 2;

            AppNamePointer += new AppNameFun(ValueAppName);
            userApplicant1.strValueName = AppNamePointer;

            AppDocNoPointer += new AppDocNoFun(AppDocNo);
            userApplicant1.strValueDocNo = AppDocNoPointer;

            AppDocTypePointer += new AppDocTypeFun(AppDocType);
            userApplicant1.strValueDocType = AppDocTypePointer;

            AppDocIssuePointer += new AppDocIssueFun(AppDocIusse);
            userApplicant1.strValueIssue = AppDocIssuePointer;

            AppDocTypePointer += new AppDocTypeFun(AppDocType);
            userApplicant1.strValueDocType = AppDocTypePointer;

            AppMaleFemalPointer += new AppMaleFemalFun(AppMaleFemale);
            userApplicant1.strValueMaleFemal = AppMaleFemalPointer;

            AppSexPointer += new AppSexFun(AppSex);
            userApplicant1.strValueSex = AppSexPointer;

            AppMovePagePointer += new AppMovePageFun(AppMovePage);
            userApplicant1.AppMovePage = AppMovePagePointer;

            DataMovePagePointer += new DataMovePageFun(AppMovePage);
            userDataView1.DataMovePage = DataMovePagePointer;

            AuthMovePagePointer += new AuthMovePageFun(AppMovePage);
            userAuthText1.AppMovePage = AuthMovePagePointer;

            AuthNamePointer += new AuthNameFun(AuthName);
            userApplicant1.strAuthName = AuthNamePointer;

            AuthSexPointer += new AuthSexFun(AuthSexvalue);
            userApplicant1.strValueSex = AuthSexPointer;

            AuthMaleFemalPointer += new AuthMaleFemalFun(AuthMaleFemalevalue);
            userApplicant1.strValueAuthMaleFemal = AuthMaleFemalPointer;

            WitPointer += new WitValuesFun(Witvalue);
            userApplicant1.strValueWit = WitPointer;

            //WitMovePagePointer += new WitMovePageFun(WitMovePage);
            //userWitNess1.witMovePage = WitMovePagePointer;

            AuthcasesPointer += new AuthcasesFun(AuthcasesData);
            userApplicant1.delgeteAuthcases = AuthcasesPointer;

            AppcasesPointer += new AppcasesFun(AppcasesData);
            userApplicant1.delgeteAppcases = AppcasesPointer;

            AuthCountPointer += new AuthCountFun(AuthCountData);
            userApplicant1.AuthCount = AuthCountPointer;

            AppCountsPointer += new AppCountsFun(AppCountsData);
            userApplicant1.ApplCount = AppCountsPointer;

            strRightsPointer += new strRightsFun(strRightsData);
            userAuthText1.strRightsText = strRightsPointer;

            stRightsIndexPointer += new strRightsIndexFun(strRightsIndex);
            userAuthText1.strRightIndex = stRightsIndexPointer;

            strAuthList2Pointer += new strAuthList2Fun(strAuthListValue);
            userAuthText1.strAuthList2 = strAuthList2Pointer;

            strAuthSubjectPointer += new strAuthSubjectFun(strAuthSubjectValue);
            userAuthText1.strAuthSubject = strAuthSubjectPointer;
            
            
        }
        private void strAuthSubjectValue(string values)
        {
            AuthSubject = values;
        }

        private void strAuthListValue(string values)
        {
            authList2 = values;
        }

        private void strRightsIndex(string values)
        {
            ListedRightIndex = values;
        }
        private void strRightsData(string values)
        {
            strRights = values;
        }
        
        private void btnAddRight_Click(object sender, EventArgs e)
        {
            
        }
               
        private void AuthCountData(int values)
        {
            AuthCount = values;
        }

        private void AppcasesData(int values)
        {
            Appcases = values;
        }
        private void AuthcasesData(int values)
        {
            Authcases = values;
        }

        private void AppCountsData(int values)
        {
            AppCounts = values;
        }

        private void ValueAppName(string[] values)
        {            
            for (int x = 0; x < 6; x++) AppnameList[x] = values[x];
        }
        private void AppDocNo(string[] values)
        {            
            for (int x = 0; x < 6; x++) AppDocNoList[x] = values[x];
        }

        private void AppDocType(string[] values)
        {
            for (int x = 0; x < 6; x++) AppDocTypeList[x] = values[x];
        }

        private void AppDocIusse(string[] values)
        {
            for (int x = 0; x < 6; x++) AppissueList[x] = values[x];
        }

        private void AppMaleFemale(string[] values)
        {
            for (int x = 0; x < 6; x++) AppMaleFemaleList[x] = values[x];
        }

        private void AppSex(string[] values)
        {
            for (int x = 0; x < 6; x++) AppSexList[x] = values[x];
        }

        private void AppMovePage(int value)
        {
            userShow(value);
        }

        private void AuthName(string[] values)
        {
            for (int x = 0; x < 6; x++) AuthNames[x] = values[x];
        }

        private void AuthMaleFemalevalue(string[] values)
        {
            for (int x = 0; x < 6; x++) AuthMaleFemale[x] = values[x];
        }

        private void AuthSexvalue(string[] values)
        {
            for (int x = 0; x < 6; x++) AuthSex[x] = values[x];
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            userDataView1.Show();
            flowLayoutPanel1.Visible = false;
            flowLayoutPanel2.Visible = false;
            flowLayoutPanel3.Visible = false;
            btnPrevious.Visible = false;
            btnNext.Visible = false;
        }

        private void btnSaveOnly_Click(object sender, EventArgs e)
        {
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new HijriCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            int Ddiffer = HijriDateDifferment(DataSource, true);
            int Mdiffer = HijriDateDifferment(DataSource, false);
            string Stringdate, Stringmonth, StrHijriDate;
            StrHijriDate = DateTime.Now.ToString("dd-MM-yyyy");
            string[] YearMonthDay = StrHijriDate.Split('-');
            int year, month, date;
            year = Convert.ToInt16(YearMonthDay[2]);
            month = Convert.ToInt16(YearMonthDay[1]) + Mdiffer;
            date = Convert.ToInt16(YearMonthDay[0]) + Ddiffer;
            if (month < 10) Stringmonth = "0" + month.ToString();
            else Stringmonth = month.ToString();
            if (date < 10) Stringdate = "0" + date.ToString();
            else Stringdate = date.ToString();
            txtHijDate.Text = Stringdate + "-" + Stringmonth + "-" + year.ToString();

        }
        private int HijriDateDifferment(string source, bool daymonth)
        {
            int differment = 0;
            string query;
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                if (daymonth) query = "select hijriday from TableSettings";
                else query = "select hijrimonth from TableSettings";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    if (daymonth) differment = Convert.ToInt32(reader["hijriday"].ToString());
                    else differment = Convert.ToInt32(reader["hijrimonth"].ToString());
                }

                saConn.Close();
            }
            return differment;
        }

        private void Form11_Load(object sender, EventArgs e)
        {
            fileComboBox(userAuthText1.comboBoxAuthValue,DataSource, "AuthTypes", "TableListCombo");
            GroupFile(userApplicant1.PanelAppValue.Controls, "AppName", "الموكَّل", "الشاهد_الأول", "الشاهد_الثاني", "مقدم_الطلب");            
            GroupFile(userApplicant1.PanelAuthValue.Controls, "txtAuthPerson", "الموكَّل", "الشاهد_الأول", "الشاهد_الثاني", "مقدم_الطلب");
            GroupFile(userApplicant1.PanelAppValue.Controls, "DocNo", "رقم_الهوية", "هوية_الأول", "هوية_الثاني", "");
            GroupFile(userApplicant1.PanelWitValue.Controls, "txtWitName", "الموكَّل", "الشاهد_الأول", "الشاهد_الثاني", "مقدم_الطلب");
            GroupFile(userApplicant1.PanelWitValue.Controls, "txtWitPass", "رقم_الهوية", "هوية_الأول", "هوية_الثاني", "");
            GroupFile(userApplicant1.PanelAppValue.Controls, "DocIssue", "مكان_الإصدار", "","","");

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
                    combbox.Items.Add(dataRow[comlumnName].ToString());
                }
                saConn.Close();
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            txtGreDate.Text = DateTime.Now.ToString("dd-MM-yyyy");
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            foreach (Control control in this.Controls)
            {
                if (control is UserControl)
                {
                    ((UserControl)control).Hide();
                }
            }
            flowLayoutPanel1.Visible = true;
            flowLayoutPanel2.Visible = true;
            flowLayoutPanel3.Visible = true;
            btnPrevious.Visible = true;
            btnNext.Visible = true;
        }

        private void AuthMovePage(int value)
        {
            userShow(value);
        }

        private void Witvalue(string[] values)
        {

            for (int x = 0; x < 4; x++) WitValuesList[x] = values[x];
            
        }

        private void WitMovePage(int value)
        {
            userShow(value);
        }

        private void AppType_CheckedChanged(object sender, EventArgs e)
        {
            mandoubVisibilty();
        }
        private void mandoubVisibilty()
        {
            if (AppType.CheckState == CheckState.Checked)
            {
                AppType.Text = "حضور مباشرة إلى القنصلية";
                mandoubName.Visible = false;
                mandoubLabel.Visible = false;
                label1.Visible = true;
            }
            else
            {
                AppType.Text = "عن طريق أحد مندوبي القنصلية";
                mandoubName.Visible = true;
                mandoubLabel.Visible = true;
                label1.Visible = false;
            }
        }

        private void btnSavePrint_Click(object sender, EventArgs e)
        {
            btnSavePrint.Enabled = false;
            CreateAuth(true);
            btnSavePrint.Enabled = true;
            //ColumnStatistics(DataSource, LastCol, "TableAuthRights");
            //StartClearReset();


        }
        private void txtAppAuthPerson()
        {            
            DBAuthPersNames = AuthNames[0];
            DBAuthPersSexs = AuthSex[0];
            DBAppPersNames = AppnameList[0];
            DBAppDoctype = AppDocTypeList[0];
            DBAppDocNo = AppDocNoList[0];
            DBAppDocIssue = AppissueList[0];
            DBAppSexs = AppSexList[0];

            for (int x = 1; x < AuthCount; x++)
            {
                DBAuthPersNames = DBAuthPersNames + "_" + AuthNames[x];
                DBAuthPersSexs = DBAuthPersSexs + "_" + AuthMaleFemale[x];
            }
            for (int x = 1; x < AppCounts; x++)
            {
                DBAppPersNames = DBAppPersNames + "_" + AppnameList[x];
                DBAppDoctype = DBAppDoctype + "_" + AppDocTypeList[x];
                DBAppDocNo = DBAppDocNo + "_" + AppDocNoList[x];
                DBAppDocIssue = DBAppDocIssue + "_" + AppissueList[x];
                DBAppSexs = DBAppSexs + "_" + AppSexList[x];
            }
        }
        private void CreateAuth(bool save)
        {
            txtAppAuthPerson();
            string Docxroutefile;
            if (AppType.CheckState == CheckState.Checked) Docxroutefile = FilespathIn + "AuthMulti.docx";
            else Docxroutefile = FilespathIn + "MandoubAuthMulti.docx";
            if (AppCounts == 1)
            {
                if (AppType.CheckState == CheckState.Checked)
                    Docxroutefile = FilespathIn + "AuthSingle.docx";
                else Docxroutefile = FilespathIn + "MandoubAuthSingle.docx";
            }

            string docxouput = FilespathOut + AppnameList[0] + DateTime.Now.ToString("ssmmhh") + ".docx";
            string pdfouput = FilespathOut + AppnameList[0] + DateTime.Now.ToString("ssmmhh") + ".pdf";

            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();

            object objCurrentCopy = Docxroutefile;

            Word.Document oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);
            oBMicroWord.Selection.Find.ClearFormatting();
            oBMicroWord.Selection.Find.Replacement.ClearFormatting();

            int x = 0;

            authList1 = " السيد" + AuthMaleFemale[0] + "/ " + AuthNames[0] + " ";
            if (AuthCount > 1)
            {
                authList1 = " كل من السيد" + AuthMaleFemale[0] + "/ " + AuthNames[0];
                for (x = 1; x < AuthCount; x++)
                {
                    authList1 = authList1 + " والسيد" + AuthMaleFemale[0] + "/ " + AuthNames[x];
                }
            }


            if (AppCounts == 1)
            {
                object ParaMaleFemale1 = "MarkMaleFemale1";
                object ParaAppName1 = "MarkAppName1";
                object ParaDocType = "MarkDocType";
                object ParaDocNo = "MarkDocNo";
                object ParaDocIssue = "MarkDocIssue";
                object ParaAppName2 = "MarkAppName2";

                Word.Range BookMaleFemale1 = oBDoc.Bookmarks.get_Item(ref ParaMaleFemale1).Range;
                Word.Range BookAppName1 = oBDoc.Bookmarks.get_Item(ref ParaAppName1).Range;
                Word.Range BookDocType = oBDoc.Bookmarks.get_Item(ref ParaDocType).Range;
                Word.Range BookDocNo = oBDoc.Bookmarks.get_Item(ref ParaDocNo).Range;
                Word.Range BookDocIssue = oBDoc.Bookmarks.get_Item(ref ParaDocIssue).Range;
                Word.Range BookAppName2 = oBDoc.Bookmarks.get_Item(ref ParaAppName2).Range;

                BookMaleFemale1.Text = AppMaleFemaleList[0];
                BookAppName2.Text = BookAppName1.Text = AppnameList[0];
                BookDocType.Text = AppDocTypeList[0];
                BookDocNo.Text = AppDocNoList[0];
                BookDocIssue.Text = AppissueList[0];

                object rangeMaleFemale1 = BookMaleFemale1;
                object rangeAppName1 = BookAppName1;
                object rangeDocNo = BookDocNo;
                object rangeDocType = BookDocType;
                object rangeDocIssue = BookDocIssue;
                object rangeAppName2 = BookAppName2;


                oBDoc.Bookmarks.Add("MarkMaleFemale1", ref rangeMaleFemale1);
                oBDoc.Bookmarks.Add("MarkAppName1", ref rangeAppName1);
                oBDoc.Bookmarks.Add("MarkDocNo", ref rangeDocNo);
                oBDoc.Bookmarks.Add("MarkDocType", ref rangeDocType);
                oBDoc.Bookmarks.Add("MarkDocIssue", ref rangeDocIssue);
                oBDoc.Bookmarks.Add("MarkAppName2", ref rangeAppName2);

            }
            else
            {
                Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
                for (x = 0; x < AppCounts; x++)
                {
                    table.Rows.Add();
                    table.Rows[x + 2].Cells[1].Range.Text = x.ToString();
                    table.Rows[x + 2].Cells[2].Range.Text = AppnameList[x];
                    table.Rows[x + 2].Cells[3].Range.Text = AppDocNoList[x];
                    table.Rows[x + 2].Cells[4].Range.Text = AppissueList[x];
                }
                object ParaIntroPart1 = "MarkIntroPart1";
                object ParaIntroPart2 = "MarkIntroPart2";

                Word.Range BookIntroPart1 = oBDoc.Bookmarks.get_Item(ref ParaIntroPart1).Range;
                Word.Range BookIntroPart2 = oBDoc.Bookmarks.get_Item(ref ParaIntroPart2).Range;

                if (AppCounts > 2)
                    BookIntroPart1.Text = "نحن المواطنون الموقعون";
                else BookIntroPart1.Text = "نحن المواطن" + preffix[Appcases, 5] + " الموقع" + preffix[Appcases, 5] + " ";
                BookIntroPart2.Text = "المقيم" + preffix[Appcases, 5] + " بالمملكة العربية السعودية، وبكامل قوانا العقلية، وبطوعنا واختيارنا وحالتنا المعتبرة شرعاً وقانوناً ";

                object rangeIntroPart1 = BookIntroPart1;
                object rangeIntroPart2 = BookIntroPart2;

                oBDoc.Bookmarks.Add("MarkIntroPart1", ref rangeIntroPart1);
                oBDoc.Bookmarks.Add("MarkIntroPart2", ref rangeIntroPart2);
            }

            if (save)
            {

                //Save2DataBase(DBAppSexs, DBAppPersNames, DBAppDocNo, DBAppDoctype, DBAppDocIssue, DBAuthPersNames, DBAuthPersSexs, ColName);
            }

            object ParaAuthNo = "MarkAuthNo";
            object ParaHijriData = "MarkHijriData";
            object ParaGreData = "MarkGreData";
            object ParaAuthBody1part1 = "MarkAuthBody1part1";
            object ParaAuthBody1part2 = "MarkAuthBody1part2";
            object ParaAuthBody1part3 = "MarkAuthBody1part3";
            object ParaAuthBody2 = "MarkAuthBody2";
            object ParaAttendVC1 = "MarkAttendVC1";
            object ParaAttendVC2 = "MarkAttendVC2";
            object ParaAuthorization = "MarkAuthorization";
            object ParaWitName1 = "MarkWitName1";
            object ParaWitName2 = "MarkWitName2";
            object ParaWitPass1 = "MarkWitPass1";
            object ParaWitPass2 = "MarkWitPass2";

            Word.Range BookAuthNo = oBDoc.Bookmarks.get_Item(ref ParaAuthNo).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookAuthBody1part1 = oBDoc.Bookmarks.get_Item(ref ParaAuthBody1part1).Range;
            Word.Range BookAuthBody1part2 = oBDoc.Bookmarks.get_Item(ref ParaAuthBody1part2).Range;
            Word.Range BookAuthBody1part3 = oBDoc.Bookmarks.get_Item(ref ParaAuthBody1part3).Range;
            Word.Range BookAuthBody2 = oBDoc.Bookmarks.get_Item(ref ParaAuthBody2).Range;
            Word.Range BookAttendVC1 = oBDoc.Bookmarks.get_Item(ref ParaAttendVC1).Range;
            Word.Range BookAttendVC2 = oBDoc.Bookmarks.get_Item(ref ParaAttendVC2).Range;
            Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;
            Word.Range BookWitName1 = oBDoc.Bookmarks.get_Item(ref ParaWitName1).Range;
            Word.Range BookWitName2 = oBDoc.Bookmarks.get_Item(ref ParaWitName2).Range;
            Word.Range BookWitPass1 = oBDoc.Bookmarks.get_Item(ref ParaWitPass1).Range;
            Word.Range BookWitPass2 = oBDoc.Bookmarks.get_Item(ref ParaWitPass2).Range;


            BookAuthNo.Text = txtAuthNo.Text;
            BookHijriData.Text = txtHijDate.Text;
            BookGreData.Text = txtGreDate.Text;

            BookAuthBody1part1.Text = preffix[Appcases, 1];
            BookAuthBody1part2.Text = authList1;
            BookAuthBody1part3.Text = authList2;

            if (AppType.CheckState == CheckState.Checked)
            {
                BookAuthorization.Text = " المواطن" + preffix[Appcases, 6] + " المذكور" + preffix[Appcases, 6] + " أعلاه قد حضر" + preffix[Appcases, 3] + " ووقع" + preffix[Appcases, 3] + " بتوقيع" + preffix[Appcases, 4] + " على هذا التوكيل وذلك بعد تلاوته علي" + preffix[Appcases, 4] + " وبعد أن فهم" + preffix[Appcases, 3] + " مضمونه ومحتواه";
            }
            else
            {
                BookAuthorization.Text = " المواطن" + preffix[Appcases, 6] + " المذكور" + preffix[Appcases, 6] + " أعلاه قد حضر" + preffix[Appcases, 3] + "  ووقع" + preffix[Appcases, 3] + " بتوقيع" + preffix[Appcases, 4] + " على هذا التوكيل أمام مندوب الجالية لدى القنصلية السيد/ " + mandoubName.Text + " وذلك بعد تلاوته علي" + preffix[Appcases, 4] + " وبعد أن فهم" + preffix[Appcases, 3] + " مضمونه ومحتواه";

            }



            BookAuthBody2.Text = strRights;
            BookAttendVC1.Text = BookAttendVC2.Text = txtAttendVC.Text;
            
            BookWitName1.Text = WitValuesList[0];
            BookWitName2.Text = WitValuesList[1];
            BookWitPass1.Text = WitValuesList[2];
            BookWitPass2.Text = WitValuesList[3];

            object rangeAuthNo = BookAuthNo;
            object rangeHijriData = BookHijriData;
            object rangeGreData = BookGreData;
            object rangeAuthBody1part1 = BookAuthBody1part1;
            object rangeAuthBody1part2 = BookAuthBody1part2;
            object rangeAuthBody1part3 = BookAuthBody1part3;
            object rangeAuthBody2 = BookAuthBody2;
            object rangeAttendVC1 = BookAttendVC1;
            object rangeAttendVC2 = BookAttendVC2;
            object rangeAuthorization = BookAuthorization;
            object rangeWitName1 = BookWitName1;
            object rangeWitName2 = BookWitName2;
            object rangeWitPass1 = BookWitPass1;
            object rangeWitPass2 = BookWitPass2;

            oBDoc.Bookmarks.Add("MarkAuthNo", ref rangeAuthNo);
            oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkAuthBody1part1", ref rangeAuthBody1part1);
            oBDoc.Bookmarks.Add("MarkAuthBody1part2", ref rangeAuthBody1part2);
            oBDoc.Bookmarks.Add("MarkAuthBody1part3", ref rangeAuthBody1part3);
            oBDoc.Bookmarks.Add("MarkAuthBody2", ref rangeAuthBody2);
            oBDoc.Bookmarks.Add("MarkAttendVC1", ref rangeAttendVC1);
            oBDoc.Bookmarks.Add("MarkAttendVC2", ref rangeAttendVC2);
            oBDoc.Bookmarks.Add("Markuthorization", ref rangeAuthorization);
            oBDoc.Bookmarks.Add("MarkWitName1", ref rangeWitName1);
            oBDoc.Bookmarks.Add("MarkWitName2", ref rangeWitName2);
            oBDoc.Bookmarks.Add("MarkWitPass1", ref rangeWitPass1);
            oBDoc.Bookmarks.Add("MarkWitPass2", ref rangeWitPass2);
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(pdfouput);
        }


        private void userAuthText1_Load(object sender, EventArgs e)
        {

        }
        private void GroupFile(Control.ControlCollection controls, string text1, string text2, string text3, string text4, string text5)
        {
            foreach (Control control in controls)
            {
                autoComplete = new AutoCompleteStringCollection();

                if (control is TextBox && control.Name.Contains(text1))
                {
                    if (text2 != "") autoCompleteTextBox(((TextBox)control), DataSource, text2, "TableAuth");
                    if (text3 != "") autoCompleteTextBox(((TextBox)control), DataSource, text3, "TableAuth");
                    if (text4 != "") autoCompleteTextBox(((TextBox)control), DataSource, text4, "TableAuth");
                    if (text5 != "") autoCompleteTextBox(((TextBox)control), DataSource, text5, "TableAuth");
                }
            }
        }

        private void autoCompleteTextBox(TextBox combbox, string source, string comlumnName, string tableName)
        {
            using (SqlConnection saConn = new SqlConnection(source))
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
        }

        private void userShow(int value)
        {
            foreach (Control control in this.Controls)
            {
                if (control is UserControl)
                {
                    ((UserControl)control).Hide();
                }
            }
            switch (value)
            {
                case 1:
                    userDataView1.Show();
                    flowLayoutPanel1.Visible = false;
                    flowLayoutPanel2.Visible = false;
                    flowLayoutPanel3.Visible = false;
                    btnPrevious.Visible = false;
                    btnNext.Visible = false;
                    break;
                case 2:                   
                                       
                    userApplicant1.Show();
                    flowLayoutPanel1.Visible = false;
                    flowLayoutPanel2.Visible = false;
                    flowLayoutPanel3.Visible = false;
                    btnPrevious.Visible = false;
                    btnNext.Visible = false;
                    break;
                case 3:
                    
                    userAuthText1.Show();
                    flowLayoutPanel1.Visible = false;
                    flowLayoutPanel2.Visible = false;
                    flowLayoutPanel3.Visible = false;
                    btnPrevious.Visible = false;
                    btnNext.Visible = false;
                    break;
                case 4:
                    flowLayoutPanel1.Visible = true;
                    flowLayoutPanel2.Visible = true;
                    flowLayoutPanel3.Visible = true;
                    btnPrevious.Visible = true;
                    btnNext.Visible = true;
                    break;
            }
        }
        private string SuffPrefReplacements(string text)
        {
            Suffex_preffixList();
            if (text.Contains("@@@"))
                return text.Replace("@@@", preffix[Appcases, 1]);
            if (text.Contains("$$$"))
                return text.Replace("$$$", preffix[Appcases, 0]);
            if (text.Contains("&&&"))
                return text.Replace("&&&", preffix[Appcases, 1]);
            if (text.Contains("^^^"))
                return text.Replace("^^^", preffix[Appcases, 2]);
            if (text.Contains("###"))
                return text.Replace("###", preffix[Authcases, 4]);
            if (text.Contains("***"))
                return text.Replace("***", preffix[Authcases, 3]);
            else return text;
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
    }
}
