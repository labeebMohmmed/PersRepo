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
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Office2010.Excel;
using Color = System.Drawing.Color;

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


        public string DataSource = "Data Source = (LocalDB)\\MSSQLLocalDB;Initial Catalog = myDataBase; Integrated Security = True";
        public string DataSource56 = "Data Source = (LocalDB)\\MSSQLLocalDB;Initial Catalog = myDataBase; Integrated Security = True";

        private static string[] AppnameList = new string[6];
        private static string[] AppMaleFemaleList = new string[6];
        private static string[] AppissueList = new string[6] { "", "", "", "", "", "" };
        private static string[] AppDocTypeList = new string[6];
        private static string[] AppDocNoList = new string[6];
        private static string[] AuthMaleFemale = new string[6];
        private static string[] AuthNames = new string[6];
        private static string[] WitValuesList = new string[4];
        string relatedDoc;
        static string[] Rights = new string[100];
        static string[] colIDs= new string[100];
        string[] ListedRight = new string[100];
        string DBAuthPersNames = "", DBAuthMaleFemale = "", DBAppPersNames = "", DBAppDoctype = "", DBAppDocNo = "", DBAppDocIssue = "", ListedRightIndex = "", DBAppMaleFemale = "";
        string ConsulateEmpName = "";
        string AuthNoPart2 = "";
        public bool ArchData = false;
        static string[,] preffix = new string[10, 20];
        string FilespathIn, FilespathOut;
        //private int AuthCount, AppCounts;
        public bool AppDataFilled = false;
            bool NewData = false;
        private int Authcases, Appcases;
        int MessageDocNo = 0;
        string MessageNo;
        
        DataTable UserTexttable;
        DataTable Textboxtable;
        AutoCompleteStringCollection autoComplete;

        string strRights = "", authList1 = "", authList2 = "", AuthSubject = "";
        public string idAuthTable = "", iddocAuthTable = "", DocxdataArch = "";

        public TextBox txtGreDateValue
        {
            get { return GregorianDate; }
            set { GregorianDate = value; }
            
        }
        public int ATVCValue
        {
            get { return ATVC; }
            set { ATVC = value; }
        }

        
        public string authList1Value
        {
            get { return authList1; }
            set { authList1 = value; }
        }

        public ComboBox txtAttendVCValue
        {
            get { return txtAttendVC; }
            set { txtAttendVC = value; }
        }
        
        public bool NewDataValue
        {
            get { return NewData; }
            set { NewData = value; }
        }

        public TextBox txtAuthNoValue
        {
            get { return txtAuthNo; }
            set { txtAuthNo = value; }
        }
        public DataTable UserComboTexttable
        {
            get { return UserTexttable; }
            set { UserTexttable = value; }
        }

        public int intAuthCount
        {
            get { return userApplicant1.strUAAuthPersonlCount; }
            set { userApplicant1.strUAAuthPersonlCount = value; }
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
            get { return userApplicant1.strUAappPersonlCount; }
            set { userApplicant1.strUAappPersonlCount = value; }
        }

        public string strrelatedDoc
        {
            get { return relatedDoc; }
            set { relatedDoc = value; }
        }

        public string[] strAppDocNolist
        {
            get { return AppDocNoList; }
            set { AppDocNoList = value; }
        }
        public string[] strAppDocTypelist
        {
            get { return AppDocTypeList; }
            set { AppDocTypeList = value; }
        }
        public string[] strAppissueList
        {
            get { return AppissueList; }
            set { AppissueList = value; }
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


        public string strauthList1
        {
            get { return authList1; }
            set { authList1 = value; }
        }

        public FlowLayoutPanel MainPanel
        {
            get { return flowLayoutPanel1; }
            set { flowLayoutPanel1 = value; }
        }


        public FlowLayoutPanel subMainPanel
        {
            get { return flowLayoutPanel2; }
            set { flowLayoutPanel2 = value; }
        }

        public ComboBox MandoubNameValue
        {
            get { return mandoubName; }
            set { mandoubName = value; }
        }

        public string JobpositionValue
        {
            get { return Jobposition; }
            set { Jobposition = value; }
        }

        public Label MandoubLabelValue
        {
            get { return mandoubLabel; }
            set { mandoubLabel = value; }
        }

        public CheckBox checkBox1Value
        {
            get { return checkBox1; }
            set { checkBox1 = value; }
        }

        public string ServerTag
        {
            get { return Server; }
            set { Server = value; }
        }

        public Form11 ParentForm { get; set; }
        public Form11 ParentData { get; set; }
        string Jobposition;
        int Rowid = 0;
        public int rowIDValue;
        public string AuthNoValue = "";
        int ATVC = 0;
        string Server = "M";
        public Form11(int Atvc,int rowid,string AuthNo,string source57,string source56, string filespathIn, string filespathOut, string EmpName, string jobposition, string greDate, string hijriDate)
        {
            DataSource = source57;
            DataSource56 = source56;
            if (DataSource.Contains("56")) Server = "U";
            Console.WriteLine("Form11  "+DataSource);
            rowIDValue = rowid;
            AuthNoValue = AuthNo;
            ATVC = Atvc;
            InitializeComponent();
            FilespathIn = filespathIn;
            FilespathOut = filespathOut;
            ConsulateEmpName = EmpName;
            Jobposition = jobposition;
            GregorianDate.Text = greDate;
            txtHijDate.Text = hijriDate;
           // MessageBox.Show(hijriDate);
            this.userAuthText1.ParentForm = this;
            this.userApplicant1.ParentFormApp = this;
            this.userDataView1.ParentData = this;
            if (Jobposition.Contains("قنصل"))
                userDataView1.deleteRowValue.Visible = true; 
            else userDataView1.deleteRowValue.Visible = false;
            Suffex_preffixList();
            foreach (Control control in this.Controls)
            {
                if (control is UserControl)
                {
                    ((UserControl)control).Hide();
                }
            }
            ComboAuthDestin.SelectedIndex = 0;
            userDataView1.Show();
            txtAttendVC.SelectedIndex = 2;

            AppMovePagePointer += new AppMovePageFun(AppMovePage);
            userApplicant1.AppMovePage = AppMovePagePointer;

            DataMovePagePointer += new DataMovePageFun(AppMovePage);
            userDataView1.DataMovePage = DataMovePagePointer;

            AuthMovePagePointer += new AuthMovePageFun(AppMovePage);
            userAuthText1.AppMovePage = AuthMovePagePointer;

            //WitMovePagePointer += new WitMovePageFun(WitMovePage);
            //userWitNess1.witMovePage = WitMovePagePointer;

            strRightsPointer += new strRightsFun(strRightsData);
            userAuthText1.strRightsText = strRightsPointer;

            stRightsIndexPointer += new strRightsIndexFun(strRightsIndex);
            userAuthText1.strRightIndex = stRightsIndexPointer;

            strAuthList2Pointer += new strAuthList2Fun(strAuthListValue);
            userAuthText1.strAuthList2 = strAuthList2Pointer;

            strAuthSubjectPointer += new strAuthSubjectFun(strAuthSubjectValue);
            userAuthText1.strAuthSubject = strAuthSubjectPointer;
            //userDataView1.FillDataGridView(DataSource);
            //txtAuthNo.Text = AuthNoPart1 + (userDataView1.rowCount + 2).ToString() ;
            if (AuthNo != "" && rowid != -1)
            {
                //userDataView1.FillDataGridView(DataSource);
                userDataView1.ListSearchValue.Text = AuthNo;
            }
           
        }


        //private void OpenFile(int id, int fileNo)
        //{
        //    string query;

        //    SqlConnection Con = new SqlConnection(DataSource);
        //    if (fileNo == 1)
        //    {
        //        query = "select Data1, Extension1,ارشفة_المستندات from TableAuth where ID=@id";
        //    }
        //    else if (fileNo == 2)
        //    {
        //        query = "select Data2, Extension2,المكاتبة_النهائية from TableAuth where ID=@id";
        //    }
        //    else query = "select Data3, Extension3,DocxData from TableAuth where ID=@id";
        //    SqlCommand sqlCmd1 = new SqlCommand(query, Con);
        //    sqlCmd1.Parameters.Add("@Id", SqlDbType.Int).Value = id;
        //    if (Con.State == ConnectionState.Closed)
        //        Con.Open();

        //    var reader = sqlCmd1.ExecuteReader();
        //    if (reader.Read())
        //    {
        //        if (fileNo == 1)
        //        {
        //            var name = reader["ارشفة_المستندات"].ToString();
        //            var Data = (byte[])reader["Data1"];
        //            var ext = reader["Extension1"].ToString();
        //            var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
        //            File.WriteAllBytes(NewFileName, Data);
        //            System.Diagnostics.Process.Start(NewFileName);
        //        }
        //        else if (fileNo == 2)
        //        {
        //            var name = reader["المكاتبة_النهائية"].ToString();
        //            var Data = (byte[])reader["Data2"];
        //            var ext = reader["Extension2"].ToString();
        //            var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
        //            File.WriteAllBytes(NewFileName, Data);
        //            System.Diagnostics.Process.Start(NewFileName);
        //        }
        //        else
        //        {
        //            var name = reader["DocxData"].ToString();
        //            var Data = (byte[])reader["Data3"];
        //            var ext = reader["Extension3"].ToString();
        //            var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
        //            File.WriteAllBytes(NewFileName, Data);
        //            openFile3(NewFileName, id);
        //            System.Diagnostics.Process.Start(NewFileName);
        //        }

        //    }
        //    Con.Close();


        //}

        private void openFile3(string filename3, int id)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET DocxData=@DocxData,fileUpload=@fileUpload WHERE ID = @ID", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.Parameters.AddWithValue("@DocxData", filename3);
            sqlCmd.Parameters.AddWithValue("@fileUpload", "No");
            sqlCmd.ExecuteNonQuery();



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

        private void AppMovePage(int value)
        {
            userShow(value);
        }


        private void btnPrevious_Click(object sender, EventArgs e)
        {

            userAuthText1.Show();
            flowLayoutPanel1.Visible = false;
            flowLayoutPanel2.Visible = false;
            flowLayoutPanel3.Visible = false;
            btnPrevious.Visible = false;
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
            
            //timer1.Enabled = false;
        }


        private string[] getColList(string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('"+ table+"')", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string[] allList = new string[dtbl.Rows.Count];
            int i = 0;
            foreach (DataRow row in dtbl.Rows)
            {
                allList[i] = row["name"].ToString();
                i++;
            }
            return allList;

        }

        private string loadRerNo(int id)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT رقم_التوكيل from TableAuth where ID=@ID", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ID", id);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string rowCnt = "";
            foreach (DataRow row in dtbl.Rows)
            {
                rowCnt = (Convert.ToInt32(row["رقم_التوكيل"].ToString().Split('/')[3]) + 1).ToString();
            }
            return rowCnt;

        }

        private int loadIDNo()
        {


            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from TableAuth order by ID desc", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string rowCnt = "0";
            foreach (DataRow row in dtbl.Rows)
            {
                rowCnt = row["ID"].ToString();
            }
            return Convert.ToInt32(rowCnt);

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
            fileComboBox(userAuthText1.comboBoxAuthValue, DataSource, "AuthTypes", "TableListCombo");
            fileComboBox(ComboAuthDestin, DataSource, "ArabCountries", "TableListCombo");

            GroupFile(userApplicant1.PanelAppValue.Controls, "AppName", "الموكَّل", "الشاهد_الأول", "الشاهد_الثاني", "مقدم_الطلب","المهنة");
            autoCompleteTextBox(userApplicant1.strالمهنة, DataSource, "jobs", "TableListCombo");
            //GroupFile(userApplicant1.PanelAuthValue.Controls, "txtAuthPerson", "الموكَّل", "الشاهد_الأول", "الشاهد_الثاني", "مقدم_الطلب");
            //GroupFile(userApplicant1.PanelAppValue.Controls, "DocNo", "رقم_الهوية", "هوية_الأول", "هوية_الثاني", "");
            //GroupFile(userApplicant1.PanelWitValue.Controls, "txtWitName", "الموكَّل", "الشاهد_الأول", "الشاهد_الثاني", "مقدم_الطلب");
            //GroupFile(userApplicant1.PanelWitValue.Controls, "txtWitPass", "رقم_الهوية", "هوية_الأول", "هوية_الثاني", "");
            //GroupFile(userApplicant1.PanelAppValue.Controls, "DocIssue", "مكان_الإصدار", "", "", "");
            fileComboBoxMan(mandoubName, DataSource, "MandoubNames", "TableListCombo");
            fileComboBox(txtAttendVC, DataSource, "ArabicAttendVC", "TableListCombo");            
            txtAttendVC.SelectedIndex = ATVC;


        }

        private void fileComboBoxMan(ComboBox combbox, string source, string comlumnName, string tableName)
        {
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
                    if (!string.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                    {
                        combbox.Items.Add(dataRow[comlumnName].ToString());
                    }
                }
                saConn.Close();
            }
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

        private void timer2_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            GregorianDate.Text = DateTime.Now.ToString("MM-dd-yyyy");
            timer2.Enabled = false;
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

        }

        private void AuthMovePage(int value)
        {
            userShow(value);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (userApplicant1.strUAWitNessList[2].Contains("p") || userApplicant1.strUAWitNessList[2].Contains("P"))
            {
                userApplicant1.strUAWitNessList[2].Replace("p", "P");
                if (userApplicant1.strUAWitNessList[2].Length != 9) MessageBox.Show("خطاء في رقم وثيق الشاهد الأول");
            }
            if (userApplicant1.strUAWitNessList[3].Contains("p") || userApplicant1.strUAWitNessList[2].Contains("P"))
            {
                userApplicant1.strUAWitNessList[3].Replace("p", "P");
                if (userApplicant1.strUAWitNessList[3].Length != 9) MessageBox.Show("خطاء في رقم وثيق الشاهد الثاني");
            }
            btnSavePrint.Enabled = false;
            CreateAuth(true, false);
            this.Close();
            btnSavePrint.Enabled = true;
            foreach (Control control in this.Controls)
            {
                if (control is UserControl)
                {
                    ((UserControl)control).Hide();
                }
            }
            flowLayoutPanel1.Visible = false;
            flowLayoutPanel2.Visible = false;
            flowLayoutPanel3.Visible = false;
            btnPrevious.Visible = false;
            userDataView1.Show();
            userAuthText1.ColumnStatistics(DataSource, userAuthText1.ColNameValue, "TableAuthRights");
            StartClearReset();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (userApplicant1.strUAWitNessList[2].Contains("p") || userApplicant1.strUAWitNessList[2].Contains("P"))
            {
                userApplicant1.strUAWitNessList[2].Replace("p", "P");
                if (userApplicant1.strUAWitNessList[2].Length != 9) MessageBox.Show("خطاء في رقم وثيق الشاهد الأول");
            }
            if (userApplicant1.strUAWitNessList[3].Contains("p") || userApplicant1.strUAWitNessList[2].Contains("P"))
            {
                userApplicant1.strUAWitNessList[3].Replace("p", "P");
                if (userApplicant1.strUAWitNessList[3].Length != 9) MessageBox.Show("خطاء في رقم وثيق الشاهد الثاني");
            }
            btnSavePrint.Enabled = false;
            CreateAuth(false, true);
            btnSavePrint.Enabled = true;
            foreach (Control control in this.Controls)
            {
                if (control is UserControl)
                {
                    ((UserControl)control).Hide();
                }
            }
            flowLayoutPanel1.Visible = false;
            flowLayoutPanel2.Visible = false;
            flowLayoutPanel3.Visible = false;
            btnPrevious.Visible = false;
            userDataView1.Show();
            userAuthText1.ColumnStatistics(DataSource, userAuthText1.ColNameValue, "TableAuthRights");
            StartClearReset();
        }
        private void btnSavePrint_Click(object sender, EventArgs e)
        {
            if (userApplicant1.strUAWitNessList[2].Contains("p") || userApplicant1.strUAWitNessList[2].Contains("P"))
            {
                userApplicant1.strUAWitNessList[2].Replace("p", "P");
                if (userApplicant1.strUAWitNessList[2].Length != 9) MessageBox.Show("خطاء في رقم وثيق الشاهد الأول");
            }
            if (userApplicant1.strUAWitNessList[3].Contains("p") || userApplicant1.strUAWitNessList[2].Contains("P"))
            {
                userApplicant1.strUAWitNessList[3].Replace("p", "P");
                if (userApplicant1.strUAWitNessList[3].Length != 9) MessageBox.Show("خطاء في رقم وثيق الشاهد الثاني");
            }
            //MessageBox.Show(ComboAuthDestin.Text);
            if (txtAuthNo.Text == "")
                txtAuthNo.Text = "ق س ج/80/" + GregorianDate.Text.Split('-')[2].Replace("20", "") + "/12/" + loadRerNo(loadIDNo());
            btnSavePrint.Enabled = false;
            CreateAuth(true, true);
            this.Close();
            btnSavePrint.Enabled = true;
            foreach (Control control in this.Controls)
            {
                if (control is UserControl)
                {
                    ((UserControl)control).Hide();
                }
            }
            flowLayoutPanel1.Visible = false;
            flowLayoutPanel2.Visible = false;
            flowLayoutPanel3.Visible = false;
            btnPrevious.Visible = false;
            userDataView1.Show();
            //userAuthText1.ColumnStatistics(DataSource, userAuthText1.ColNameValue, "TableAuthRights");
            NewData = false;
            StartClearReset();
        }

        private void StartClearReset()
        {
            AppDataFilled = false;
            userDataView1.FillDataGridView(DataSource);
            txtAuthNo.Text = "";
            for (int x = 0; x < 6; x++) userApplicant1.strUAAuthMaleFemale[x] = userApplicant1.strUAMaleFemale[x] = "ذكر";
            userApplicant1.PanelAppValue.Height = 82;
            foreach (Control control in userApplicant1.PanelAppValue.Controls)
            {
                if (control is TextBox)
                {

                    if (((TextBox)control).Name.Contains("AppName")) { ((TextBox)control).Text = ""; }
                    if (((TextBox)control).Name.Contains("DocNo")) { ((TextBox)control).Text = "P0"; }
                    if (((TextBox)control).Name.Contains("DocIssue")) { ((TextBox)control).Text = ""; }
                }
                if (control is ComboBox)
                {
                    if (((ComboBox)control).Name.Contains("DocType")) { ((ComboBox)control).Text = "";
                        ((ComboBox)control).SelectedIndex = 0;
                    }
                }
                if (control is CheckBox)
                {
                    if (((CheckBox)control).Name.Contains("checkSexType")) { ((CheckBox)control).Text = "ذكر"; ((CheckBox)control).CheckState = CheckState.Unchecked; }
                }
            }

            foreach (Control control in userApplicant1.PanelAuthValue.Controls)
            {
                if (control is TextBox)
                {
                    if (((TextBox)control).Name.Contains("txtAuthPerson")) { ((TextBox)control).Text = ""; }
                }
                if (control is CheckBox)
                {
                    if (((CheckBox)control).Name.Contains("txtAuthPersonsex")) { ((CheckBox)control).Text = "ذكر"; ((CheckBox)control).CheckState = CheckState.Unchecked; }
                }
            }
            foreach (Control control in userApplicant1.PanelWitValue.Controls)
            {
                if (control is TextBox)
                {
                    if (((TextBox)control).Name.Contains("txtWitName")) { ((TextBox)control).Text = ""; }
                    if (((TextBox)control).Name.Contains("txtWitPass")) { ((TextBox)control).Text = "P0"; }
                }
            }
            userApplicant1.PanelAuthValue.Height = 41;
            userApplicant1.strUAAuthPersonlCount = 1;
            userApplicant1.strUAappPersonlCount = 1;
            userAuthText1.deleteItemsAO();
            //userAuthText1.comboBoxAuthValue.SelectedIndex = 0;
            userAuthText1.comboBoxAuthValue.Text = "";
            //userAuthText1.ComboProcedureValue.Text = "";
            userAuthText1.ComboProcedureValue.Items.Clear();
            userAuthText1.PanelSubItemBoxValue.Visible = false;
            userAuthText1.comboPropertyTypeValue.Items.Clear();
            userAuthText1.comboPropertyTypeValue.Text = "";
            userAuthText1.txtReviewValue.Text = "";
            txtAttendVC.SelectedIndex = 2;
            txtComment.Text = "لا تعليق";
            strRights = "";
            AppType.CheckState = CheckState.Checked;
            userDataView1.ArchivedStValue.Visible = false;
            userDataView1.labelArchValue.Visible = false;
            btnSavePrint.Visible = true;
            btnSavePrint.Text = "طباعة وحفظ";
            userDataView1.fileloadedValue = false;

            userDataView1.FillDataGridView(DataSource);

        }


        private void StartNewAuth()
        {
            for (int x = 0; x < 6; x++) userApplicant1.strUAAuthMaleFemale[x] = userApplicant1.strUAMaleFemale[x] = "ذكر";
            userApplicant1.PanelAppValue.Height = 82;
            foreach (Control control in userApplicant1.PanelAppValue.Controls)
            {
                if (control is TextBox)
                {
                    if (((TextBox)control).Name.Contains("AppName")) { ((TextBox)control).Text = ""; }
                    if (((TextBox)control).Name.Contains("DocNo")) { ((TextBox)control).Text = "P0"; }
                    if (((TextBox)control).Name.Contains("DocIssue")) { ((TextBox)control).Text = ""; }
                }
                if (control is ComboBox)
                {
                    if (((ComboBox)control).Name.Contains("DocType")) { ((ComboBox)control).Text = ""; }
                    ((ComboBox)control).SelectedIndex = 0;
                }
                if (control is CheckBox)
                {
                    if (((CheckBox)control).Name.Contains("checkSexType")) { ((CheckBox)control).Text = "ذكر"; ((CheckBox)control).CheckState = CheckState.Unchecked; }
                }
            }

            foreach (Control control in userApplicant1.PanelAuthValue.Controls)
            {
                if (control is TextBox)
                {
                    if (((TextBox)control).Name.Contains("txtAuthPerson")) { ((TextBox)control).Text = ""; }
                }
                if (control is CheckBox)
                {
                    if (((CheckBox)control).Name.Contains("txtAuthPersonsex")) { ((CheckBox)control).Text = "ذكر"; ((CheckBox)control).CheckState = CheckState.Unchecked; }
                }
            }
            foreach (Control control in userApplicant1.PanelWitValue.Controls)
            {
                if (control is TextBox)
                {
                    if (((TextBox)control).Name.Contains("txtWitName")) { ((TextBox)control).Text = ""; }
                    if (((TextBox)control).Name.Contains("txtWitPass")) { ((TextBox)control).Text = "P0"; }
                }
            }
            checkBox1.CheckState = CheckState.Checked;
            userApplicant1.PanelAuthValue.Height = 41;
            userApplicant1.strUAAuthPersonlCount = 1;
            userApplicant1.strUAappPersonlCount = 1;
            userAuthText1.deleteItemsAO();
            userAuthText1.comboBoxAuthValue.SelectedIndex = 0;
            userAuthText1.comboBoxAuthValue.Text = "";
            userAuthText1.ComboProcedureValue.Text = "";
            userAuthText1.ComboProcedureValue.Items.Clear();
            userAuthText1.PanelSubItemBoxValue.Visible = false;
            userAuthText1.comboPropertyTypeValue.Items.Clear();
            userAuthText1.comboPropertyTypeValue.Text = "";
            userAuthText1.txtReviewValue.Text = "";
            txtAttendVC.SelectedIndex = 2;
            txtComment.Text = "لا تعليق";
            strRights = "";
            AppType.CheckState = CheckState.Checked;
            userDataView1.ArchivedStValue.Visible = false;
            userDataView1.labelArchValue.Visible = false;
            btnSavePrint.Visible = true;
            btnSavePrint.Text = "طباعة وحفظ";
            userDataView1.fileloadedValue = false;

            userDataView1.FillDataGridView(DataSource);

        }

        private void txtAuthNo_TextChanged(object sender, EventArgs e)
        {

        }

        private void ComboAuthDestin_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"D:\rights.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;


            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    //MessageBox.Show(str);
                }
            }

            //xlWorkBook.Close(true, null, null);
            //xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private void userDataView1_Load(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.CheckState == CheckState.Checked)
            {
                userAuthText1.txtReviewValue.Location = new System.Drawing.Point(336, 162);
                userAuthText1.txtReviewValue.Size = new System.Drawing.Size(828, 85);
                userAuthText1.txtAddRightValue.Location = new System.Drawing.Point(387, 253);
                userAuthText1.txtAddRightValue.Size = new System.Drawing.Size(777, 84);
                userAuthText1.label34Value.Visible = true;
                userAuthText1.label36Value.Text = "إضافة نص جديد:";
                checkBox1.Text = "صيغة معتمدة";
            }
                    else
            {
                userAuthText1.txtReviewValue.Location = new System.Drawing.Point(336, 76);
                userAuthText1.txtReviewValue.Size = new System.Drawing.Size(828, 171);
                userAuthText1.txtAddRightValue.Location = new System.Drawing.Point(336, 253);
                userAuthText1.txtAddRightValue.Size = new System.Drawing.Size(828, 409);
                userAuthText1.label34Value.Visible = false;
                userAuthText1.label36Value.Text = "إضافة نصوص التوكيل:";
                checkBox1.Text = "صيغة عامة";
            }

            
                
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void userAuthText1_Load(object sender, EventArgs e)
        {
           
            ;
        }

        private void WitMovePage(int value)
        {
            userShow(value);
        }

        private void ResetAll_Click_1(object sender, EventArgs e)
        {
            StartClearReset();
        }

        private void AppType_CheckedChanged(object sender, EventArgs e)
        {
            mandoubVisibilty();
        }

        private void ComboAuthDestin_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (ComboAuthDestin.SelectedIndex != 0) {
                string name;
                if (DBAppPersNames.Contains("_")) name = DBAppPersNames.Replace("_", " و");
                else name = DBAppPersNames;
                CreateMessageWord(name, ComboAuthDestin.Text, txtAuthNo.Text,"توكيلا", preffix[userApplicant1.Appcases(), 9], GregorianDate.Text, txtHijDate.Text, txtAttendVC.Text);
            }
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


        private void txtAppAuthPerson()
        {
            DBAuthPersNames = userApplicant1.strUAAuthNameList[0];
            DBAuthMaleFemale = userApplicant1.strUAAuthMaleFemale[0];
            DBAppPersNames = userApplicant1.strUAAppnameList[0];
            DBAppDoctype = userApplicant1.strUAAppDocTypeList[0];
            DBAppDocNo = userApplicant1.strUAAppDocNoList[0];
            DBAppDocIssue = userApplicant1.strUAAppIssueList[0];
            DBAppMaleFemale = userApplicant1.strUAMaleFemale[0];

            for (int x = 1; x < userApplicant1.strUAAuthPersonlCount; x++)
            {
                DBAuthPersNames = DBAuthPersNames + "_" + userApplicant1.strUAAuthNameList[x];
                DBAuthMaleFemale = DBAuthMaleFemale + "_" + userApplicant1.strUAAuthMaleFemale[x];
            }
            for (int x = 1; x < userApplicant1.strUAappPersonlCount; x++)
            {
                DBAppPersNames = DBAppPersNames + "_" + userApplicant1.strUAAppnameList[x];
                DBAppDoctype = DBAppDoctype + "_" + userApplicant1.strUAAppDocTypeList[x];
                DBAppDocNo = DBAppDocNo + "_" + userApplicant1.strUAAppDocNoList[x];
                DBAppDocIssue = DBAppDocIssue + "_" + userApplicant1.strUAAppIssueList[x];
                DBAppMaleFemale = DBAppMaleFemale + "_" + userApplicant1.strUAAppIssueList[x];
            }
            
        }
        private void CreateAuth(bool save, bool print)
        {
            txtAppAuthPerson();
            if (userAuthText1.comboBoxAuthValue.SelectedIndex != 16)
                createAuth1();
            else EngcreateAuth1();
            txtAttendVC.SelectedIndex = ATVCValue;
            string docxouput = FilespathOut + userApplicant1.strUAAppnameList[0] + DateTime.Now.ToString("ssmmhh") + ".docx";
            string pdfouput = FilespathOut + userApplicant1.strUAAppnameList[0] + DateTime.Now.ToString("ssmmhh") + ".pdf";
            if (save)
            {
                Save2DataBase(docxouput, DBAppMaleFemale, DBAppPersNames, DBAppDocNo, DBAppDoctype, DBAppDocIssue, DBAuthPersNames, DBAuthMaleFemale, userAuthText1.ColNameValue);
            }
            string Docxroutefile, RouteFile;

            if (AppType.CheckState == CheckState.Checked)
                RouteFile = FilespathIn + "AuthMulti.docx";
            else
                RouteFile = FilespathIn + "MandoubAuthMulti.docx";
            
            if (userApplicant1.strUAappPersonlCount == 1)
            {
                if (AppType.CheckState == CheckState.Checked)
                {
                    RouteFile = FilespathIn + "AuthSingle.docx";
                    if(userAuthText1.comboBoxAuthValue.Text == "شهادة ميلاد")
                        RouteFile = FilespathIn + "newAuthbirth.docx";
                    if (userAuthText1.comboBoxAuthValue.SelectedIndex == 6)
                        RouteFile = FilespathIn + "AuthSingleCopy.docx";
                }
                else
                {
                    RouteFile = FilespathIn + "MandoubAuthSingle.docx";
                }
            }
            if (userAuthText1.comboBoxAuthValue.Text.Contains("أجنبية")) 
            {
                if (userApplicant1.strUAappPersonlCount == 1)
                {
                    if(userApplicant1.strUAWitNessList[0] != "")
                    RouteFile = FilespathIn + "EngAuthSingle1.docx";
                    else 
                        RouteFile = FilespathIn + "EngAuthSingle2.docx";

                }
                else
                    RouteFile = FilespathIn + "EngAuthMulti.docx";
            }
            //MessageBox.Show(RouteFile);
            FileInfo fileInfo = new FileInfo(RouteFile);
            if (fileInfo.IsReadOnly) fileInfo.IsReadOnly = false;
            Docxroutefile = FilespathOut + userApplicant1.strUAAppnameList[0]+ DateTime.Now.ToString("mmss") + ".docx";
            System.IO.File.Copy(RouteFile, Docxroutefile);

            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();

            object objCurrentCopy = Docxroutefile;

            Word.Document oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);
            oBMicroWord.Selection.Find.ClearFormatting();
            oBMicroWord.Selection.Find.Replacement.ClearFormatting();



            colIDs[3] = userApplicant1.strUAAppnameList[0];
            colIDs[4] = ConsulateEmpName.Trim();
            colIDs[5] = AppType.Text;
            colIDs[6] = mandoubName.Text;
            if (userApplicant1.strUAappPersonlCount == 1)
            {
                object ParaMaleFemale1 = "MarkMaleFemale1";
                object ParaMaleFemale2 = "MarkMaleFemale2";
                object ParaMaleFemale3 = "MarkMaleFemale3";
                object ParaAppName1 = "MarkAppName1";
                object ParaDocType = "MarkDocType";
                object ParaDocNo = "MarkDocNo";
                object ParaDocIssue = "MarkDocIssue";
                object ParaAppName2 = "MarkAppName2";

                Word.Range BookMaleFemale1 = oBDoc.Bookmarks.get_Item(ref ParaMaleFemale1).Range;
                Word.Range BookMaleFemale2 = oBDoc.Bookmarks.get_Item(ref ParaMaleFemale2).Range;
                
                Word.Range BookAppName1 = oBDoc.Bookmarks.get_Item(ref ParaAppName1).Range;
                Word.Range BookDocType = oBDoc.Bookmarks.get_Item(ref ParaDocType).Range;
                Word.Range BookDocNo = oBDoc.Bookmarks.get_Item(ref ParaDocNo).Range;
                Word.Range BookDocIssue = oBDoc.Bookmarks.get_Item(ref ParaDocIssue).Range;
                Word.Range BookAppName2 = oBDoc.Bookmarks.get_Item(ref ParaAppName2).Range;

                
                if(userApplicant1.strUAMaleFemale[0] == "ذكر")
                    BookMaleFemale1.Text = BookMaleFemale2.Text = "";
                else BookMaleFemale1.Text = BookMaleFemale2.Text = "ة";
                BookAppName2.Text = BookAppName1.Text = userApplicant1.strUAAppnameList[0];
                BookDocType.Text = userApplicant1.strUAAppDocTypeList[0];
                BookDocNo.Text = userApplicant1.strUAAppDocNoList[0];
                BookDocIssue.Text = userApplicant1.strUAAppIssueList[0];

                object rangeMaleFemale1 = BookMaleFemale1;
                object rangeMaleFemale2 = BookMaleFemale2;
                
                object rangeAppName1 = BookAppName1;
                object rangeDocNo = BookDocNo;
                object rangeDocType = BookDocType;
                object rangeDocIssue = BookDocIssue;
                object rangeAppName2 = BookAppName2;


                oBDoc.Bookmarks.Add("MarkMaleFemale1", ref rangeMaleFemale1);
                oBDoc.Bookmarks.Add("MarkMaleFemale2", ref rangeMaleFemale2);
                
                oBDoc.Bookmarks.Add("MarkAppName1", ref rangeAppName1);
                oBDoc.Bookmarks.Add("MarkDocNo", ref rangeDocNo);
                oBDoc.Bookmarks.Add("MarkDocType", ref rangeDocType);
                oBDoc.Bookmarks.Add("MarkDocIssue", ref rangeDocIssue);
                oBDoc.Bookmarks.Add("MarkAppName2", ref rangeAppName2);

                
                if (RouteFile == FilespathIn + "newAuthbirth.docx")
                {
                    Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
                    for (int x = 0; x < userAuthText1.birthindex; x++)
                    {
                        table.Rows.Add();
                        table.Rows[x + 2].Cells[1].Range.Text = (x+1).ToString();
                        table.Rows[x + 2].Cells[2].Range.Text = userAuthText1.BirthNameValue[x];
                        table.Rows[x + 2].Cells[3].Range.Text = userAuthText1.BirthPlaceValue[x];
                        table.Rows[x + 2].Cells[4].Range.Text = userAuthText1.BirthDateValue[x];
                        table.Rows[x + 2].Cells[5].Range.Text = userAuthText1.BirthMotherValue[x];
                    }

                }
                if (!userAuthText1.comboBoxAuthValue.Text.Contains("أجنبية"))
                {
                    Word.Range BookMaleFemale3 = oBDoc.Bookmarks.get_Item(ref ParaMaleFemale3).Range;
                    if (userApplicant1.strUAMaleFemale[0] == "ذكر")
                        BookMaleFemale3.Text = "";
                    else BookMaleFemale3.Text = "ة";
                    object rangeMaleFemale3 = BookMaleFemale3;
                    oBDoc.Bookmarks.Add("MarkMaleFemale3", ref rangeMaleFemale3);
                    
                }
            }
            else
            {
                Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
                for (int x = 0; x < userApplicant1.strUAappPersonlCount; x++)
                {
                    table.Rows.Add();
                    table.Rows[x + 2].Cells[1].Range.Text = (x+1).ToString();
                    table.Rows[x + 2].Cells[2].Range.Text = userApplicant1.strUAAppnameList[x];
                    table.Rows[x + 2].Cells[3].Range.Text = userApplicant1.strUAAppDocNoList[x];
                    table.Rows[x + 2].Cells[4].Range.Text = userApplicant1.strUAAppIssueList[x];
                }
                object ParaIntroPart1 = "MarkIntroPart1";
                object ParaIntroPart2 = "MarkIntroPart2";

                Word.Range BookIntroPart1 = oBDoc.Bookmarks.get_Item(ref ParaIntroPart1).Range;
                Word.Range BookIntroPart2 = oBDoc.Bookmarks.get_Item(ref ParaIntroPart2).Range;

                if (userApplicant1.strUAappPersonlCount > 2)
                    BookIntroPart1.Text = "نحن المواطنون الموقعون";
                else BookIntroPart1.Text = "نحن المواطن" + preffix[Appcases, 5] + " الموقع" + preffix[Appcases, 5] + " ";
                BookIntroPart2.Text = "المقيم" + preffix[Appcases, 5] + " بالمملكة العربية السعودية، وبكامل قوانا العقلية، وبطوعنا واختيارنا وحالتنا المعتبرة شرعاً وقانوناً ";

                object rangeIntroPart1 = BookIntroPart1;
                object rangeIntroPart2 = BookIntroPart2;

                oBDoc.Bookmarks.Add("MarkIntroPart1", ref rangeIntroPart1);
                oBDoc.Bookmarks.Add("MarkIntroPart2", ref rangeIntroPart2);
            }

            

            object ParaAuthNo = "MarkAuthNo";
            object ParaHijriData = "MarkHijriData";
            object ParaGreData = "MarkGreData";
            object ParaAuthBody1part1 = "MarkAuthBody1part1";
            object ParaAuthBody1part2 = "MarkAuthBody1part2";
            object ParaAuthBody1part3 = "MarkAuthBody1part3";
            object ParaAuthBody2 = "MarkAuthBody2";

            
            
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

            Word.Range BookWitName1 = oBDoc.Bookmarks.get_Item(ref ParaWitName1).Range;
            Word.Range BookWitName2 = oBDoc.Bookmarks.get_Item(ref ParaWitName2).Range;
            Word.Range BookWitPass1 = oBDoc.Bookmarks.get_Item(ref ParaWitPass1).Range;
            Word.Range BookWitPass2 = oBDoc.Bookmarks.get_Item(ref ParaWitPass2).Range;

             
            BookAuthNo.Text = colIDs[0] =txtAuthNo.Text;
            BookHijriData.Text = txtHijDate.Text;
            BookGreData.Text = colIDs[2] = GregorianDate.Text;

            BookAuthBody1part1.Text = "";
            BookAuthBody1part2.Text = authList1;
            

            BookAuthBody1part3.Text = userAuthText1.AuthList2Value.Replace("auth1", authList1);
            
            string[] strmandoub = new string[2];
            strmandoub = mandoubName.Text.Split('-');

            

            if (checkBox1.CheckState == CheckState.Unchecked)
            {
                BookAuthBody2.Text = userAuthText1.txtAddRightValue.Text;
                BookAuthBody1part3.Text = userAuthText1.txtReviewValue.Text.Replace(authList1,"");
            }
            else
            {
                BookAuthBody2.Text = strRights;
                //BookAuthBody1part3.Text = userAuthText1.txtReviewValue.Text;
                BookAuthBody1part3.Text = userAuthText1.AuthList2Value.Replace("aith1", authList1);
            }

            

            BookWitName1.Text = userApplicant1.strUAWitNessList[0];
            BookWitName2.Text = userApplicant1.strUAWitNessList[1];
            BookWitPass1.Text = userApplicant1.strUAWitNessList[2];
            BookWitPass2.Text = userApplicant1.strUAWitNessList[3];

            object rangeAuthNo = BookAuthNo;
            object rangeHijriData = BookHijriData;
            object rangeGreData = BookGreData;
            object rangeAuthBody1part1 = BookAuthBody1part1;
            object rangeAuthBody1part2 = BookAuthBody1part2;
            object rangeAuthBody1part3 = BookAuthBody1part3;
            object rangeAuthBody2 = BookAuthBody2;
            
            
            
            object rangeWitName1 = BookWitName1;
            object rangeWitName2 = BookWitName2;
            object rangeWitPass1 = BookWitPass1;
            object rangeWitPass2 = BookWitPass2;

            

            if (ComboAuthDestin.Text == "جمهورية السودان" && !userAuthText1.comboBoxAuthValue.Text.Contains("أجنبية")) {
                object ParaAuthKhartoum = "AuthKhartoum";
                
                Word.Range BookAuthKhartoum = oBDoc.Bookmarks.get_Item(ref ParaAuthKhartoum).Range;
                if (userAuthText1.comboBoxAuthValue.Text == "إقرار بالتنازل")
                    BookAuthKhartoum.Text = "لا يعتمد هذا الاقرار ما لم يتم توثيقه خلال عام من تاريخ إصدارة من خارجية جمهورية السودان";
                else BookAuthKhartoum.Text = "لا يعتمد هذا التوكيل ما لم يتم توثيقه خلال عام من تاريخ إصدارة من خارجية جمهورية السودان";

                object rangeAuthKhartoum = BookAuthKhartoum;
                
                oBDoc.Bookmarks.Add("AuthKhartoum", ref rangeAuthKhartoum);
               
            }
            
            if (print)
            {
                
                if (AppType.CheckState == CheckState.Checked)
                {
                    object ParaAttendVC1 = "AuthAttendVC1";

                    Word.Range BookAttendVC1 = oBDoc.Bookmarks.get_Item(ref ParaAttendVC1).Range;

                    BookAttendVC1.Text = txtAttendVC.Text;

                    object rangeAttendVC1 = BookAttendVC1;
                    try
                    {
                        oBDoc.Bookmarks.Add("AuthAttendVC1", ref rangeAttendVC1);
                    }
                    catch (Exception e) { }

                }

                if (!userAuthText1.comboBoxAuthValue.Text.Contains("أجنبية"))
                {
                    object ParaAttendVC2 = "MarkAttendVC2";
                    Word.Range BookAttendVC2 = oBDoc.Bookmarks.get_Item(ref ParaAttendVC2).Range;
                    BookAttendVC2.Text = txtAttendVC.Text;
                    object rangeAttendVC2 = BookAttendVC2;
                    oBDoc.Bookmarks.Add("MarkAttendVC2", ref rangeAttendVC2);

                    object ParaAuthorization = "MarkAuthorization";
                    Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;

                    if (userAuthText1.comboBoxAuthValue.Text == "إقرار بالتنازل")
                    {
                        BookAuthBody1part2.Text.Replace("السيد", "للسيد");
                        if (userApplicant1.strUAAuthPersonlCount > 1)
                            BookAuthBody1part2.Text.Replace("كل", "لكل");
                        BookAuthBody1part1.Text = preffix[Appcases, 11];
                        if (AppType.CheckState == CheckState.Checked)
                        {
                            BookAuthorization.Text = " المواطن" + preffix[Appcases, 6] + " المذكور" + preffix[Appcases, 6] + " أعلاه قد حضر" + preffix[Appcases, 3] + " أمامي ووقع" + preffix[Appcases, 3] + " بتوقيع" + preffix[Appcases, 4] + " على هذا الاقرار في حضور الشهود المذكورين أعلاه وذلك بعد تلاوته علي" + preffix[Appcases, 4] + " وبعد أن فهم" + preffix[Appcases, 3] + " مضمونه ومحتواه";
                        }
                        else
                        {
                            BookAuthorization.Text = " المواطن" + preffix[Appcases, 6] + " المذكور" + preffix[Appcases, 6] + " أعلاه قد حضر" + preffix[Appcases, 3] + " أمامي ووقع" + preffix[Appcases, 3] + " بتوقيع" + preffix[Appcases, 4] + " على هذا الاقرار في حضور الشهود المذكورين أعلاه " + " بعد تلاوته علي" + preffix[Appcases, 4] + " وبعد أن فهم" + preffix[Appcases, 3] + " مضمونه ومحتواه" + " وذلك أمام مندوب جالية منطقة " + strmandoub[1] + " السيد/ " + strmandoub[0] + " بموجب التفويض الممنوح له من القنصلية العامة ";
                        }
                    }
                    else
                    {
                        if (userAuthText1.comboBoxAuthValue.SelectedIndex != 6)
                            BookAuthBody1part1.Text = preffix[Appcases, 10];
                        else
                        {
                            BookAuthBody1part2.Text = BookAuthBody1part1.Text = "";
                                                        
                        }
                        if (AppType.CheckState == CheckState.Checked)
                        {
                            BookAuthorization.Text = " المواطن" + preffix[Appcases, 6] + " المذكور" + preffix[Appcases, 6] + " أعلاه قد حضر" + preffix[Appcases, 3] + " ووقع" + preffix[Appcases, 3] + " بتوقيع" + preffix[Appcases, 4] + " على هذا التوكيل في حضور الشهود المذكورين أعلاه وذلك بعد تلاوته علي" + preffix[Appcases, 4] + " وبعد أن فهم" + preffix[Appcases, 3] + " مضمونه ومحتواه";
                        }
                        else
                        {
                            try
                            {
                                BookAuthorization.Text = " المواطن" + preffix[Appcases, 6] + " المذكور" + preffix[Appcases, 6] + " أعلاه قد حضر" + preffix[Appcases, 3] + " ووقع" + preffix[Appcases, 3] + " بتوقيع" + preffix[Appcases, 4] + " على هذا التوكيل في حضور الشهود المذكورين أعلاه " + " بعد تلاوته علي" + preffix[Appcases, 4] + " وبعد أن فهم" + preffix[Appcases, 3] + " مضمونه ومحتواه" + " وذلك أمام مندوب جالية منطقة " + strmandoub[1] + " السيد/ " + strmandoub[0] + " بموجب التفويض الممنوح له من القنصلية العامة ";
                            }
                            catch (Exception ex) { }
                        }
                    }
                    object rangeAuthorization = BookAuthorization;
                    oBDoc.Bookmarks.Add("Markuthorization", ref rangeAuthorization);
                }

                oBDoc.Bookmarks.Add("MarkAuthNo", ref rangeAuthNo);
                oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);
                oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
                oBDoc.Bookmarks.Add("MarkAuthBody1part1", ref rangeAuthBody1part1);
                oBDoc.Bookmarks.Add("MarkAuthBody1part2", ref rangeAuthBody1part2);
                oBDoc.Bookmarks.Add("MarkAuthBody1part3", ref rangeAuthBody1part3);
                oBDoc.Bookmarks.Add("MarkAuthBody2", ref rangeAuthBody2);
                
                
                oBDoc.Bookmarks.Add("MarkWitName1", ref rangeWitName1);
                oBDoc.Bookmarks.Add("MarkWitName2", ref rangeWitName2);
                oBDoc.Bookmarks.Add("MarkWitPass1", ref rangeWitPass1);
                oBDoc.Bookmarks.Add("MarkWitPass2", ref rangeWitPass2);


                object ParaAuthIqrar = "AuthIqrar";

                Word.Range BookAuthIqrar = oBDoc.Bookmarks.get_Item(ref ParaAuthIqrar).Range;

                if (userAuthText1.comboBoxAuthValue.Text.Contains("أجنبية"))
                    BookAuthIqrar.Text = "AUTHORIZATION";
                else if (userAuthText1.comboBoxAuthValue.Text == "إقرار بالتنازل")
                    BookAuthIqrar.Text = "إقرار بالتنازل";
                else BookAuthIqrar.Text = "توكيل";

                object rangeAuthIqrar = BookAuthIqrar;
                oBDoc.Bookmarks.Add("AuthIqrar", ref rangeAuthIqrar);
                oBDoc.SaveAs2(docxouput);
                oBDoc.Close(false, oBMiss);
                oBMicroWord.Quit(false, false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
                System.Diagnostics.Process.Start(docxouput);
                object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
                
                addarchives(colIDs);
            }
        }
        private void addarchives(string[] text)
        {
            string[] allList = getColList("archives");
            string strList = "";
            for (int i = 1; i < allList.Length; i++)
            {
                if (i == 1) strList = "@" + allList[1];
                else strList = strList + ",@" + allList[i];
            }
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            Console.WriteLine(strList);
            SqlCommand sqlCommand = new SqlCommand("insert into archives values ("+ strList+")", sqlConnection);
            //SqlCommand sqlCommand = new SqlCommand("insert into archives (docID, employName,archiveStat,databaseID,appType,appOldNew) " +
            //    "values (@docID, @employName,@archiveStat,@databaseID,@appType,@appOldNew)", sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            for (int i = 1; i < allList.Length; i++)
                sqlCommand.Parameters.AddWithValue("@"+ allList[i], text[i-1]);
            sqlCommand.ExecuteNonQuery();
        }

        private void AddMandoub_Click(object sender, EventArgs e)
        {

        }

        

        private void CreateMessageWord(string ApplicantName, string EmbassySource, string IqrarNo, string MessageType, string ApplicantSex, string GregorianDate, string HijriDate, string ViseConsul)
        {
            string ActiveCopy;
            string ReportName = DateTime.Now.ToString("mmss");
            string routeDoc = FilespathIn + "MessageCap.docx";
            loadMessageNo();
            ActiveCopy = FilespathOut + "Message" + ApplicantName + ReportName + ".docx";
            if (!File.Exists(ActiveCopy))
            {
                System.IO.File.Copy(routeDoc, ActiveCopy);
                object oBMiss2 = System.Reflection.Missing.Value;
                Word.Application oBMicroWord2 = new Word.Application();
                Word.Document oBDoc2 = oBMicroWord2.Documents.Open(ActiveCopy, oBMiss2);
                Object ParacapitalMessage = "MarkcapitalMessage";
                Object ParaMApplicantName = "MarkApplicantName";
                Object ParaMassageIqrarNo = "MarkMassageIqrarNo";
                Object ParaMassageTitle = "MarkMassageTitle";
                Object ParaMassageNo = "MarkMassageNo";
                Object ParaApliSex = "MarkApliSex";
                Object ParaHijriDate = "MarkHijriDate";
                Object ParaDateGre = "MarkDateGre";
                Object ParaGregorDate2 = "MarkGregorDate2";
                Object ParaViseConsul1 = "MarkViseConsul1";


                Word.Range BookMApplicantName = oBDoc2.Bookmarks.get_Item(ref ParaMApplicantName).Range;
                Word.Range BookcapitalMessage = oBDoc2.Bookmarks.get_Item(ref ParacapitalMessage).Range;
                Word.Range BookMassageIqrarNo = oBDoc2.Bookmarks.get_Item(ref ParaMassageIqrarNo).Range;
                Word.Range BookMassageNo = oBDoc2.Bookmarks.get_Item(ref ParaMassageNo).Range;
                Word.Range BookApliSex = oBDoc2.Bookmarks.get_Item(ref ParaApliSex).Range;
                Word.Range BookDateGre = oBDoc2.Bookmarks.get_Item(ref ParaDateGre).Range;
                Word.Range BookHijriDate = oBDoc2.Bookmarks.get_Item(ref ParaHijriDate).Range;
                Word.Range BookGregorDate2 = oBDoc2.Bookmarks.get_Item(ref ParaGregorDate2).Range;
                Word.Range BookMassageTitle = oBDoc2.Bookmarks.get_Item(ref ParaMassageTitle).Range;
                Word.Range BookViseConsul1 = oBDoc2.Bookmarks.get_Item(ref ParaViseConsul1).Range;

                BookMApplicantName.Text = ApplicantName;
                BookcapitalMessage.Text = EmbassySource;
                BookMassageNo.Text = MessageNo + (MessageDocNo + 1).ToString();
                BookMassageIqrarNo.Text = IqrarNo;
                BookApliSex.Text = ApplicantSex;
                BookGregorDate2.Text = BookDateGre.Text = GregorianDate;
                BookHijriDate.Text = HijriDate;
                BookViseConsul1.Text = ViseConsul;
                BookMassageTitle.Text = MessageType;

                object rangeViseConsul1 = BookViseConsul1;
                object rangeMApplicantName = BookMApplicantName;
                object rangecapitalMessage = BookcapitalMessage;
                object rangeMassageIqrarNo = BookMassageIqrarNo;
                object rangeMassageNo = BookMassageNo;
                object rangeApliSex = BookApliSex;
                object rangeDateGre = BookDateGre;
                object rangeHijriDate = BookHijriDate;
                object rangeGregorDate2 = BookGregorDate2;
                object rangeMassageTitle = BookMassageTitle;


                oBDoc2.Bookmarks.Add("MarkViseConsul1", ref rangeViseConsul1);
                oBDoc2.Bookmarks.Add("MarkApplicantName", ref rangeMApplicantName);
                oBDoc2.Bookmarks.Add("MarkcapitalMessage", ref rangecapitalMessage);
                oBDoc2.Bookmarks.Add("MarkMassageIqrarNo", ref rangeMassageIqrarNo);
                oBDoc2.Bookmarks.Add("MarkMassageNo", ref rangeMassageNo);
                oBDoc2.Bookmarks.Add("MarkApliSex", ref rangeApliSex);
                oBDoc2.Bookmarks.Add("MarkDateGre", ref rangeDateGre);
                oBDoc2.Bookmarks.Add("MarkGregorDate2", ref rangeGregorDate2);
                oBDoc2.Bookmarks.Add("MarkHijiData", ref rangeHijriDate);
                oBDoc2.Bookmarks.Add("MarkMassageTitle", ref rangeMassageTitle);

                oBDoc2.Activate();
                oBDoc2.Save();
                oBMicroWord2.Visible = true;
                NewMessageNo();
            }

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET حالة_الارشفة=@حالة_الارشفة  WHERE ID = @ID", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", userDataView1.DatasumValue[9]);            
            sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", "غير مؤرشف");
            sqlCmd.ExecuteNonQuery();
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET نوع_التوكيل=@نوع_التوكيل,إجراء_التوكيل=@إجراء_التوكيل  WHERE ID = @ID", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", userDataView1.DatasumValue[9]);
            sqlCmd.Parameters.AddWithValue("@نوع_التوكيل", userAuthText1.comboBoxAuthValue.Text.Trim());
            sqlCmd.Parameters.AddWithValue("@إجراء_التوكيل", userAuthText1.ComboProcedureValue.Text);
            
            sqlCmd.ExecuteNonQuery();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET نوع_التوكيل=@نوع_التوكيل,حالة_الارشفة=@حالة_الارشفة  WHERE ID = @ID", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", userDataView1.DatasumValue[9]);
            sqlCmd.Parameters.AddWithValue("@نوع_التوكيل", "توكيل بصيغة عامة");
            sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", "غير مؤرشف");
            sqlCmd.ExecuteNonQuery();
            this.Close();
        }

        private void Form11_FormClosed(object sender, FormClosedEventArgs e)
        {
            string primeryLink = @"D:\PrimariFiles\";
            if (!Directory.Exists(@"D:\"))
                {
                string appFileName = Environment.GetCommandLineArgs()[0];
                string directory = Path.GetDirectoryName(appFileName);
                directory = directory + @"\";
                primeryLink = directory + @"PrimariFiles\";                
            }
            dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Allowed");
        }
        private void dataSourceWrite(string dataSourcepath, string text)
        {
            using (FileStream fs = File.Create(dataSourcepath))
            {
                string dataasstring = text;
                byte[] info = new UTF8Encoding(true).GetBytes(dataasstring);
                fs.Write(info, 0, info.Length);
                fs.Close();
            }
        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private void mandoubLabel_Click(object sender, EventArgs e)
        {

        }

        private void mandoubName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void loadMessageNo()
        {
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select MessageNo  from TableSettings where ID=@id", Con);
            sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();

            if (reader.Read())
            {
                MessageDocNo = Convert.ToInt32(reader["MessageNo"].ToString());
            }

            Con.Close();


        }
        private void NewMessageNo()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    SqlCommand sqlCmd = new SqlCommand("UPDATE TableSettings SET MessageNo=@MessageNo WHERE ID=@ID", sqlCon);
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@ID", 1);
                    sqlCmd.Parameters.AddWithValue("@MessageNo", MessageDocNo + 1);
                    sqlCmd.ExecuteNonQuery();
                    sqlCon.Close();

                }

                catch (Exception ex)
                {
                    MessageBox.Show("الوصول لقاعدة البيانات غير متاح");
                }
                finally
                {
                }
        }

        private void createAuth1()
        {
            int x = 0;
            string authtitle = "";
                if (userApplicant1.strUAAuthMaleFemale[0] == "أنثى") authtitle = "ة"; else authtitle = "";
            authList1 = " السيد" + authtitle + "/ " + userApplicant1.strUAAuthNameList[0] + " ";
            string authtitle1 = "إقامة رقم ";
            if (userApplicant1.strAuthNationalityList[0].Contains("سعود")) authtitle1 = "هوية وطينة رقم ";
            if (userApplicant1.strAuthNationalityList[0] != "")
                authList1 = " السيد" + authtitle + "/ " + userApplicant1.strUAAuthNameList[0] + " (" + userApplicant1.strAuthNationalityList[0] + ") حامل" + authtitle + " "+authtitle1 + " " + userApplicant1.strAuthNationalityIDList[0] + " ";
            if (userApplicant1.strUAAuthPersonlCount > 1)
            {
                authList1 = " كل من السيد" + authtitle + "/ " + userApplicant1.strUAAuthNameList[0];
                authtitle1 = "إقامة رقم ";
                if (userApplicant1.strAuthNationalityList[0].Contains("سعود")) authtitle1 = "هوية وطينة رقم ";
                if (userApplicant1.strAuthNationalityList[0] != "")
                    authList1 = " السيد" + authtitle + "/ " + userApplicant1.strUAAuthNameList[0] + " (" + userApplicant1.strAuthNationalityList[0] + ") حامل" + authtitle + " "+ authtitle1 + " " + userApplicant1.strAuthNationalityIDList[0] + " ";
                for ( x = 1; x < userApplicant1.strUAAuthPersonlCount; x++)
                {
                    authtitle1 = "إقامة رقم ";
                    if (userApplicant1.strAuthNationalityList[0].Contains("سعود")) authtitle1 = "هوية وطينة رقم";
                    if (userApplicant1.strUAAuthMaleFemale[x] == "أنثى") authtitle = "ة"; else authtitle = "";
                    authList1 = authList1 + " والسيد" + authtitle + "/ " + userApplicant1.strUAAuthNameList[x] +"(" + userApplicant1.strAuthNationalityList[x] + ") حامل" + authtitle + " " + authtitle1 + " " + userApplicant1.strAuthNationalityIDList[x] + " "; ;
                }
            }

        }

        private void EngcreateAuth1()
        {
            authList1 = userApplicant1.strEngAuthMaleFemale[0] + " " + userApplicant1.strUAAuthNameList[0] + " ";
            string authtitle1 = "Resident ID No. ";            
            if (userApplicant1.strEngAuthMaleFemale[0].Contains("Sauid")) 
                authtitle1 = "National ID No. ";
            if (userApplicant1.strAuthNationalityIDList[0] != "")
                authList1 = userApplicant1.strEngAuthMaleFemale[0] + " " + userApplicant1.strUAAuthNameList[0] + " (" + userApplicant1.strAuthNationalityList[0] + ") holder " + authtitle1 + " " + userApplicant1.strAuthNationalityIDList[0] + " ";
            if (userApplicant1.strUAAuthPersonlCount > 1)
            {
                for (int x = 1; x < userApplicant1.strUAAuthPersonlCount-1; x++)
                {
                    authtitle1 = "Resident ID No.";
                    if (userApplicant1.strAuthNationalityIDList[x].Contains("Sauid"))
                        authtitle1 = "National ID No. ";
                    authList1 = userApplicant1.strEngAuthMaleFemale[x] + " " + userApplicant1.strUAAuthNameList[x] + " ";
                    if (userApplicant1.strAuthNationalityIDList[0] != "")
                        authList1 = authList1 +", "+ userApplicant1.strEngAuthMaleFemale[x] + " " + userApplicant1.strUAAuthNameList[x] + "(" + userApplicant1.strAuthNationalityList[x] + ") holder " + authtitle1 + " " + userApplicant1.strAuthNationalityIDList[x] + " "; ;
                }
                authList1 = userApplicant1.strEngAuthMaleFemale[userApplicant1.strUAAuthPersonlCount - 1] + " " + userApplicant1.strUAAuthNameList[userApplicant1.strUAAuthPersonlCount - 1] + " ";
                authList1 = authList1 + " and " + userApplicant1.strEngAuthMaleFemale[userApplicant1.strUAAuthPersonlCount - 1] + " " + userApplicant1.strUAAuthNameList[userApplicant1.strUAAuthPersonlCount - 1] + "(" + userApplicant1.strAuthNationalityList[userApplicant1.strUAAuthPersonlCount - 1] + ") holder " + authtitle1 + " " + userApplicant1.strAuthNationalityIDList[userApplicant1.strUAAuthPersonlCount - 1] + " "; ;
            }

        }
        private void Save2DataBase(string filename3,string AppGenders, string Appname, string AppdocNo, string AppdocType, string AppdocIssue, string Authname, string AuthGender, string columnName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            if (!NewData && DocxdataArch == "")
            {
                SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableAuth (مقدم_الطلب,رقم_التوكيل, النوع, نوع_الهوية, رقم_الهوية, مكان_الإصدار, الموكَّل, جنس_الموكَّل, نوع_التوكيل, موضوع_التوكيل, اضافة_الموضوع, حقوق_التوكيل, التاريخ_الميلادي, التاريخ_الهجري, موقع_التوكيل, المعالجة, طريقة_الطلب, اسم_الموظف, اسم_المندوب, وجهة_التوكيل, توكيل_مرجعي, الشاهد_الأول, هوية_الأول, الشاهد_الثاني, هوية_الثاني, تعليق, حالة_الارشفة, specialData, إجراء_التوكيل) values(@مقدم_الطلب,@رقم_التوكيل, @النوع, @نوع_الهوية, @رقم_الهوية, @مكان_الإصدار, @الموكَّل, @جنس_الموكَّل, @نوع_التوكيل, @موضوع_التوكيل, @اضافة_الموضوع, @حقوق_التوكيل, @التاريخ_الميلادي, @التاريخ_الهجري, @موقع_التوكيل, @المعالجة, @طريقة_الطلب, @اسم_الموظف, @اسم_المندوب, @وجهة_التوكيل, @توكيل_مرجعي, @الشاهد_الأول, @هوية_الأول, @الشاهد_الثاني, @هوية_الثاني, @تعليق, @حالة_الارشفة, @specialData, @إجراء_التوكيل) ", sqlCon);
                //SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableAuth (مقدم_الطلب, النوع, نوع_الهوية, رقم_الهوية, مكان_الإصدار, الموكَّل, جنس_الموكَّل, نوع_التوكيل, موضوع_التوكيل, اضافة_الموضوع, حقوق_التوكيل, التاريخ_الميلادي, التاريخ_الهجري, موقع_التوكيل, المعالجة, طريقة_الطلب, اسم_الموظف, اسم_المندوب, وجهة_التوكيل, توكيل_مرجعي, الشاهد_الأول, هوية_الأول, الشاهد_الثاني, هوية_الثاني, Data1, Extension1, ارشفة_المستندات, Data2, Extension2, المكاتبة_النهائية, تعليق, حالة_الارشفة, specialData, Data3, Extension3, DocxData, إجراء_التوكيل) values(@مقدم_الطلب, @النوع, @نوع_الهوية, @رقم_الهوية, @مكان_الإصدار, @الموكَّل, @جنس_الموكَّل, @نوع_التوكيل, @موضوع_التوكيل, @اضافة_الموضوع, @حقوق_التوكيل, @التاريخ_الميلادي, @التاريخ_الهجري, @موقع_التوكيل, @المعالجة, @طريقة_الطلب, @اسم_الموظف, @اسم_المندوب, @وجهة_التوكيل, @توكيل_مرجعي, @الشاهد_الأول, @هوية_الأول, @الشاهد_الثاني, @هوية_الثاني, @Data1, @Extension1, @ارشفة_المستندات, @Data2, @Extension2, @المكاتبة_النهائية, @تعليق, @حالة_الارشفة, @specialData, @Data3, @Extension3, @DocxData, @إجراء_التوكيل) ", sqlCon);
                //SqlCommand sqlCmd = new SqlCommand("TableAuthAddorEdit", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@ID", 0);
                sqlCmd.Parameters.AddWithValue("@mode", "Add");
                //MessageBox.Show(txtAuthNo.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@رقم_التوكيل", txtAuthNo.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@مقدم_الطلب", Appname.Trim());
                sqlCmd.Parameters.AddWithValue("@النوع", AppGenders.Trim());
                sqlCmd.Parameters.AddWithValue("@نوع_الهوية", AppdocType.Trim());
                sqlCmd.Parameters.AddWithValue("@رقم_الهوية", AppdocNo.Trim());
                sqlCmd.Parameters.AddWithValue("@مكان_الإصدار", AppdocIssue.Trim());
                sqlCmd.Parameters.AddWithValue("@الموكَّل", Authname.Trim());
                sqlCmd.Parameters.AddWithValue("@جنس_الموكَّل", AuthGender.Trim());
                sqlCmd.Parameters.AddWithValue("@نوع_التوكيل", userAuthText1.comboBoxAuthValue.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@رقم_العمود", columnName.Trim());
                sqlCmd.Parameters.AddWithValue("@موضوع_التوكيل", userAuthText1.AllBoxesData.Trim());
                sqlCmd.Parameters.AddWithValue("@اضافة_الموضوع", userAuthText1.NewAuthSubjectValue.Trim());
                sqlCmd.Parameters.AddWithValue("@حقوق_التوكيل", ListedRightIndex.Trim());
                sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", GregorianDate.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@التاريخ_الهجري", txtHijDate.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@موقع_التوكيل", txtAttendVC.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@المعالجة", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                sqlCmd.Parameters.AddWithValue("@طريقة_الطلب", AppType.Text);
                sqlCmd.Parameters.AddWithValue("@اسم_الموظف", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                sqlCmd.Parameters.AddWithValue("@اسم_المندوب", mandoubName.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@وجهة_التوكيل", ComboAuthDestin.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@توكيل_مرجعي", "");
                sqlCmd.Parameters.AddWithValue("@الشاهد_الأول", userApplicant1.strUAWitNessList[0]);
                sqlCmd.Parameters.AddWithValue("@هوية_الأول", userApplicant1.strUAWitNessList[2]);
                sqlCmd.Parameters.AddWithValue("@الشاهد_الثاني", userApplicant1.strUAWitNessList[1]);
                sqlCmd.Parameters.AddWithValue("@هوية_الثاني", userApplicant1.strUAWitNessList[3]);
                sqlCmd.Parameters.AddWithValue("@Extension3", ".docx");
                sqlCmd.Parameters.AddWithValue("@fileUpload", "No");
                sqlCmd.Parameters.AddWithValue("@تعليق", txtComment.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", "غير مؤرشف");
                sqlCmd.Parameters.AddWithValue("@إجراء_التوكيل", userAuthText1.ComboProcedureValue.Text);
                if (checkBox1.CheckState == CheckState.Unchecked)
                    sqlCmd.Parameters.AddWithValue("@specialData", "صيغة غير معتمدة" + "_" + userAuthText1.txtReviewValue.Text + "_" + userAuthText1.txtAddRightValue.Text);
                else
                    sqlCmd.Parameters.AddWithValue("@specialData", userAuthText1.specialDataSum);
                sqlCmd.ExecuteNonQuery();
            }
            else
            {
                SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET رقم_التوكيل = @رقم_التوكيل, مقدم_الطلب = @مقدم_الطلب, النوع = @النوع, نوع_الهوية = @نوع_الهوية, رقم_الهوية = @رقم_الهوية, مكان_الإصدار = @مكان_الإصدار, الموكَّل = @الموكَّل, جنس_الموكَّل = @جنس_الموكَّل, نوع_التوكيل = @نوع_التوكيل, موضوع_التوكيل = @موضوع_التوكيل, اضافة_الموضوع = @اضافة_الموضوع, حقوق_التوكيل = @حقوق_التوكيل, التاريخ_الميلادي = @التاريخ_الميلادي, التاريخ_الهجري = @التاريخ_الهجري, موقع_التوكيل = @موقع_التوكيل, المعالجة = @المعالجة, طريقة_الطلب = @طريقة_الطلب, اسم_الموظف = @اسم_الموظف, اسم_المندوب = @اسم_المندوب, وجهة_التوكيل = @وجهة_التوكيل,  الشاهد_الأول = @الشاهد_الأول, هوية_الأول = @هوية_الأول, الشاهد_الثاني = @الشاهد_الثاني, هوية_الثاني = @هوية_الثاني, تعليق = @تعليق, حالة_الارشفة = @حالة_الارشفة, specialData = @specialData, إجراء_التوكيل = @إجراء_التوكيل,DocxData=@DocxData,Extension3=@Extension3,fileUpload=@fileUpload,itext1=@itext1,itext2=@itext2,itext3=@itext3,itext4=@itext4,itext5=@itext5,icheck1=@icheck1,itxtDate1=@itxtDate1,icombo1=@icombo1,icombo2=@icombo2,ibtnAdd1=@ibtnAdd1  WHERE ID = @ID", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@ID", userDataView1.DatasumValue[9]);
                sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                sqlCmd.Parameters.AddWithValue("@رقم_التوكيل", txtAuthNo.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@مقدم_الطلب", Appname.Trim());
                sqlCmd.Parameters.AddWithValue("@النوع", AppGenders.Trim());
                sqlCmd.Parameters.AddWithValue("@نوع_الهوية", AppdocType.Trim());
                sqlCmd.Parameters.AddWithValue("@رقم_الهوية", AppdocNo.Trim());
                sqlCmd.Parameters.AddWithValue("@مكان_الإصدار", AppdocIssue.Trim());
                sqlCmd.Parameters.AddWithValue("@الموكَّل", Authname.Trim());
                sqlCmd.Parameters.AddWithValue("@جنس_الموكَّل", AuthGender.Trim());
                sqlCmd.Parameters.AddWithValue("@رقم_العمود", columnName.Trim());
                sqlCmd.Parameters.AddWithValue("@نوع_التوكيل", userAuthText1.comboBoxAuthValue.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@موضوع_التوكيل", userAuthText1.AllBoxesData.Trim());
                sqlCmd.Parameters.AddWithValue("@اضافة_الموضوع", userAuthText1.NewAuthSubjectValue.Trim());
                sqlCmd.Parameters.AddWithValue("@حقوق_التوكيل", ListedRightIndex.Trim());
                sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", GregorianDate.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@التاريخ_الهجري", txtHijDate.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@موقع_التوكيل", txtAttendVC.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@المعالجة", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                sqlCmd.Parameters.AddWithValue("@طريقة_الطلب", AppType.Text);
                sqlCmd.Parameters.AddWithValue("@اسم_الموظف", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                sqlCmd.Parameters.AddWithValue("@اسم_المندوب", mandoubName.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@وجهة_التوكيل", ComboAuthDestin.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@DocxData", filename3);
                sqlCmd.Parameters.AddWithValue("@Extension3", ".docx");
                sqlCmd.Parameters.AddWithValue("@fileUpload", "No");
                
                sqlCmd.Parameters.AddWithValue("@الشاهد_الأول", userApplicant1.strUAWitNessList[0]);
                sqlCmd.Parameters.AddWithValue("@هوية_الأول", userApplicant1.strUAWitNessList[2]);
                sqlCmd.Parameters.AddWithValue("@الشاهد_الثاني", userApplicant1.strUAWitNessList[1]);
                sqlCmd.Parameters.AddWithValue("@هوية_الثاني", userApplicant1.strUAWitNessList[3]);
                string filePath1 = FilespathIn + "text1.txt";
                string filePath2 = FilespathIn + "text2.txt";
                string filePath3 = FilespathIn + "text3.txt";
                
                sqlCmd.Parameters.AddWithValue("@تعليق", txtComment.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", "غير مؤرشف");
                sqlCmd.Parameters.AddWithValue("@إجراء_التوكيل", userAuthText1.ComboProcedureValue.Text);
                if (checkBox1.CheckState == CheckState.Unchecked)
                    sqlCmd.Parameters.AddWithValue("@specialData", "صيغة غير معتمدة" + "_" + userAuthText1.txtReviewValue.Text + "_" + userAuthText1.txtAddRightValue.Text);
                else
                    sqlCmd.Parameters.AddWithValue("@specialData", userAuthText1.specialDataSum);
                string[] str = userAuthText1.AllBoxesData.Split('_');
                if (str.Length == 10)
                {
                    sqlCmd.Parameters.AddWithValue("@itext1", str[0]);
                    sqlCmd.Parameters.AddWithValue("@itext2", str[1]);
                    sqlCmd.Parameters.AddWithValue("@itext3", str[2]);
                    sqlCmd.Parameters.AddWithValue("@itext4", str[3]);
                    sqlCmd.Parameters.AddWithValue("@itext5", str[4]);
                    sqlCmd.Parameters.AddWithValue("@icheck1", str[5]);
                    sqlCmd.Parameters.AddWithValue("@itxtDate1", str[6]);
                    sqlCmd.Parameters.AddWithValue("@icombo1", str[7]);
                    sqlCmd.Parameters.AddWithValue("@ibtnAdd1", str[8]);
                    sqlCmd.Parameters.AddWithValue("@icombo2", str[9]);
                }
                else {
                    sqlCmd.Parameters.AddWithValue("@itext1", "");
                    sqlCmd.Parameters.AddWithValue("@itext2", "");
                    sqlCmd.Parameters.AddWithValue("@itext3", "");
                    sqlCmd.Parameters.AddWithValue("@itext4", "");
                    sqlCmd.Parameters.AddWithValue("@itext5", "");
                    sqlCmd.Parameters.AddWithValue("@icheck1", "");
                    sqlCmd.Parameters.AddWithValue("@itxtDate1", "");
                    sqlCmd.Parameters.AddWithValue("@icombo1", "");
                    sqlCmd.Parameters.AddWithValue("@ibtnAdd1", "");
                    sqlCmd.Parameters.AddWithValue("@icombo2", "");
                }
            
                sqlCmd.ExecuteNonQuery();
            }
            
            if (userAuthText1.RemovedDocInfo.Split('_')[0] != "" && userAuthText1.RemovedDocInfo.Split('_')[1] != "")
            {                
                removedDoc1(DataSource, userAuthText1.RemovedDocInfo.Split('_')[2], txtAuthNo.Text.Trim());
                removedDoc2(DataSource, userAuthText1.RemovedDocInfo, txtAuthNo.Text.Trim());
            }
            userDataView1.FillDataGridView(DataSource);
        }


        private void splitCol(int id, string str)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET text1=@text1,text2=@text2,text3=@text3,text4=@text4,text5=@text5,check1=@check1,txtD1=@txtD1,txtD2=@txtD2,combo1=@combo1,combo2=@combo2,addName1=@addName1 WHERE ID = @id", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@text1", str.Split('_')[0]);
            sqlCmd.Parameters.AddWithValue("@text2", str.Split('_')[1]);
            sqlCmd.Parameters.AddWithValue("@text3", str.Split('_')[2]);
            sqlCmd.Parameters.AddWithValue("@text4", str.Split('_')[3]);
            sqlCmd.Parameters.AddWithValue("@text5", str.Split('_')[4]);
            sqlCmd.Parameters.AddWithValue("@check1", str.Split('_')[5]);
            sqlCmd.Parameters.AddWithValue("@txtD1", str.Split('_')[6]);
            sqlCmd.Parameters.AddWithValue("@txtD2", str.Split('_')[7]);
            sqlCmd.Parameters.AddWithValue("@combo1", str.Split('_')[8]);
            sqlCmd.Parameters.AddWithValue("@addName1", str.Split('_')[9]);
            sqlCmd.Parameters.AddWithValue("@combo2", str.Split('_')[10]);
            sqlCmd.ExecuteNonQuery();
        }

        private void removedDoc1(string dataSource, string removedDoc, string removingDoc)
        {            
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET حالة_الارشفة=@حالة_الارشفة, توكيل_مرجعي=@توكيل_مرجعي where رقم_التوكيل = @رقم_التوكيل", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_التوكيل", removedDoc);
            sqlCmd.Parameters.AddWithValue("@توكيل_مرجعي", removingDoc);
            sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", "ملغي" + GregorianDate.Text.Replace("-","/"));
            
            sqlCmd.ExecuteNonQuery();
        }

        private void removedDoc2(string dataSource, string removedDoc, string removingDoc)
        {            
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET المكاتبات_الملغية=@المكاتبات_الملغية where رقم_التوكيل = @رقم_التوكيل", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_التوكيل", removingDoc);
            sqlCmd.Parameters.AddWithValue("@المكاتبات_الملغية", removedDoc);
            sqlCmd.ExecuteNonQuery();
        }

        private void GroupFile(Control.ControlCollection controls, string text1, string text2, string text3, string text4, string text5, string text6)
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
            idAuthTable = userDataView1.DatasumValue[9];
            iddocAuthTable = txtAuthNo.Text = userDataView1.DatasumValue[19];
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
                    
                    break;
                case 2:
                    
                    int x = 0, y = 0, z = 0, a = 0, b = 0;
                    if (NewData && !AppDataFilled)
                    {
                        if (userDataView1.DatasumValue[15].Contains("ملغي"))
                        {
                            userApplicant1.strFlowLayoutPanel1.BackColor = Color.LightCoral;
                            userApplicant1.strlabelDocIsRemoved.Visible = true;
                        } else 
                        {
                            userApplicant1.strFlowLayoutPanel1.BackColor = Color.White;
                            userApplicant1.strlabelDocIsRemoved.Visible = false;
                        }
                        
                        if (userDataView1.DatasumValue[20] != "")
                        {
                            userApplicant1.strtxtPreRelated.Visible = true;
                            userApplicant1.strlabelrelateddoc.Visible = true;
                            userApplicant1.strtxtPreRelated.Text = userDataView1.DatasumValue[20];
                        }
                        else {
                            userApplicant1.strtxtPreRelated.Visible = false;
                            userApplicant1.strlabelrelateddoc.Visible = false;
                            userApplicant1.strtxtPreRelated.Text = "";
                        }

                        if (userDataView1.DatasumValue[21] != "")
                        {
                            userApplicant1.strtxtPreRemoved.Visible = true;
                            userApplicant1.strlabebremovedDoc.Visible = true;
                            userApplicant1.strtxtPreRemoved.Text = userDataView1.DatasumValue[21];
                        }
                        else {
                            userApplicant1.strtxtPreRemoved.Visible = false;
                            userApplicant1.strlabebremovedDoc.Visible = false;
                            userApplicant1.strtxtPreRemoved.Text = "";
                        }

                        foreach (Control chk in userApplicant1.PanelAppValue.Controls)
                        {
                            if (chk is TextBox && x < intAppCounts && ((TextBox)chk).Name.Contains("AppName"))
                            {
                                ((TextBox)chk).Text = AppnameList[x];
                                userApplicant1.PanelAppValue.Height = userApplicant1.strUAappPersonlCount * 82;
                                x++;
                            }
                            if (chk is CheckBox && y < intAppCounts)
                            {
                                if (AppMaleFemaleList[y] == "ذكر" || AppMaleFemaleList[y] == "") ((CheckBox)chk).CheckState = CheckState.Unchecked;
                                else ((CheckBox)chk).CheckState = CheckState.Checked;
                                y++;
                            }
                            if (chk is ComboBox && z < intAppCounts && ((ComboBox)chk).Name.Contains("DocType"))
                            {
                                ((ComboBox)chk).Text = AppDocTypeList[z];
                                z++;
                            }
                            if (chk is TextBox && a < intAppCounts && ((TextBox)chk).Name.Contains("DocNo"))
                            {
                                ((TextBox)chk).Text = AppDocNoList[a];
                                a++;
                            }
                            if (chk is TextBox && b < intAppCounts && ((TextBox)chk).Name.Contains("DocIssue"))
                            {
                                ((TextBox)chk).Text = AppissueList[b];
                                b++;
                            }

                        }
                        
                        x = 0; y = 0;
                        foreach (Control chk in userApplicant1.PanelAuthValue.Controls)
                        {
                            if (chk is TextBox && x < userApplicant1.strUAAuthPersonlCount)
                            {
                                ((TextBox)chk).Text = AuthNames[x];
                                x++;
                            }
                            if (chk is CheckBox && x < userApplicant1.strUAAuthPersonlCount)
                            {
                                if (AuthMaleFemale[y] == "ذكر") ((CheckBox)chk).CheckState = CheckState.Unchecked;
                                else ((CheckBox)chk).CheckState = CheckState.Checked;
                                y++;
                            }
                        }
                        
                        foreach (Control conrol in userAuthText1.PanelItemsboxesValue.Controls)
                        {
                            if (conrol is TextBox && x < userApplicant1.strUAAuthPersonlCount)
                            {
                                ((TextBox)conrol).Text = AuthNames[x];
                                x++;
                            }

                        }
                        
                        foreach (Control conrol in userApplicant1.PanelWitValue.Controls)
                        {
                            if (conrol is TextBox && ((TextBox)conrol).Name == "txtWitName1")
                            {
                                userApplicant1.strUAWitNessList[0] = ((TextBox)conrol).Text = userDataView1.DatasumValue[10];
                            }
                            if (conrol is TextBox && ((TextBox)conrol).Name == "txtWitPass1")
                            {
                                userApplicant1.strUAWitNessList[1] = ((TextBox)conrol).Text = userDataView1.DatasumValue[11];
                            }
                            if (conrol is TextBox && ((TextBox)conrol).Name == "txtWitName2")
                            {
                                userApplicant1.strUAWitNessList[2] = ((TextBox)conrol).Text = userDataView1.DatasumValue[12];
                            }
                            if (conrol is TextBox && ((TextBox)conrol).Name == "txtWitPass2")
                            {
                                userApplicant1.strUAWitNessList[3] = ((TextBox)conrol).Text = userDataView1.DatasumValue[13];
                            }
                        }
                        DocxdataArch = userDataView1.DatasumValue[9];
                        
                    } else DocxdataArch = "";
                    if (userDataView1.NewAuth)
                    {
                        
                        StartNewAuth();
                        
                    }

                    colIDs[1] = userDataView1.DatasumValue[9];
                    //MessageBox.Show(userDataView1.DatasumValue[18]);
                    if (AppnameList[0] == "" )
                    {
                        FillDatafromGenArch("data1", userDataView1.DatasumValue[9], "TableAuth");
                        btnSaveOnly.Visible = btnprintOnly.Visible = false;
                        
                    }
                    
                    if (AppnameList[0] == "" && userDataView1.DatasumValue[18] != "")
                    {
                        ArchData = true;
                        colIDs[7] = "new";
                        //MessageBox.Show(userDataView1.DatasumValue[9]);
                        //OpenFile(Convert.ToInt32(userDataView1.DatasumValue[9]), 1);
                        
                        
                        btnSaveOnly.Visible = btnprintOnly.Visible = false;
                        
                    }
                    else {
                        colIDs[7] = "old";
                        
                        btnSaveOnly.Visible = btnprintOnly.Visible = true;                        
                    }
                    userDataView1.NewAuth = false;
                    userApplicant1.Show();
                    flowLayoutPanel1.Visible = false;
                    flowLayoutPanel2.Visible = false;
                    flowLayoutPanel3.Visible = false;
                    btnPrevious.Visible = false;
                    AppDataFilled = true;
                    break;
                case 3:
                    createAuth1();
                    intAuthcases = userApplicant1.Authcases();
                    intAppcases = userApplicant1.Appcases();
                    if (userAuthText1.comboBoxAuthValue.SelectedIndex != 16)
                        createAuth1();
                    else EngcreateAuth1();
                    if (NewData && !ArchData)
                    {
                        userAuthText1.checkboxdtValue = new DataTable();
                        userAuthText1.comboBoxAuthValue.Text = userDataView1.DatasumValue[0];
                        

                        if (userAuthText1.comboBoxAuthValue.Text == "توكيل بصيغة غير مدرجة" || userAuthText1.comboBoxAuthValue.Text == "توكيل بصيغة عامة")
                            flowLayoutPanel3.Size = new System.Drawing.Size(303, 112);
                        else flowLayoutPanel3.Size = new System.Drawing.Size(303, 56);


                        userAuthText1.PopulateCheckBoxes(userDataView1.DatasumValue[1], "TableAuthRights", DataSource);
                        userAuthText1.ComboProcedureValue.Text = userDataView1.DatasumValue[16];
                        userAuthText1.CreateBoxesWithData(userDataView1.DatasumValue[16], userDataView1.DatasumValue[2], true);
                        //MessageBox.Show(userDataView1.DatasumValue[17]);
                        if (userDataView1.DatasumValue[17] != "" && !userDataView1.DatasumValue[17].Contains("صيغة غير معتمدة"))
                        {
                            string[] memberdata = userDataView1.DatasumValue[0].Split('*');
                            
                            for (int X = 0; X < userDataView1.DatasumValue[0].Split('*').Length; X++)
                            {
                                
                                if (userDataView1.DatasumValue[0].Split('*')[X].Contains("_"))
                                {
                                    string[] data = memberdata[X].Split('_');
                                    userAuthText1.txt1Value.Text = userAuthText1.BirthNameValue[X] = data[0];
                                    userAuthText1.txt2Value.Text = userAuthText1.BirthPlaceValue[X] = data[1];
                                    userAuthText1.txt3Value.Text = userAuthText1.BirthDateValue[X] = data[2];
                                    userAuthText1.txt4Value.Text = userAuthText1.BirthMotherValue[X] = data[3];
                                    userAuthText1.BirthindexValue = X;
                                    if (userAuthText1.BirthindexValue == 0) userAuthText1.specialDataSum = userAuthText1.BirthNameValue[userAuthText1.BirthindexValue] + "_" + userAuthText1.BirthPlaceValue[userAuthText1.BirthindexValue] + "_" + userAuthText1.BirthDateValue[userAuthText1.BirthindexValue] + "_" + userAuthText1.BirthMotherValue[userAuthText1.BirthindexValue];
                                    else userAuthText1.specialDataSum = userAuthText1.specialDataSum + "*" + userAuthText1.BirthNameValue[userAuthText1.BirthindexValue] + "_" + userAuthText1.BirthPlaceValue[userAuthText1.BirthindexValue] + "_" + userAuthText1.BirthDateValue[userAuthText1.BirthindexValue] + "_" + userAuthText1.BirthMotherValue[userAuthText1.BirthindexValue];
                                    userAuthText1.Mentioned = userAuthText1.BirthDescValue[X] = data[4];
                                    userAuthText1.addNameValue.Text = "اضافة (" + memberdata.Length.ToString() + "/" + memberdata.Length.ToString() + ")";
                                }

                            }
                        }
                        else if (userDataView1.DatasumValue[17].Contains("صيغة غير معتمدة"))
                        {
                            checkBox1.CheckState = CheckState.Unchecked;
                            userAuthText1.txtReviewValue.Text = userDataView1.DatasumValue[17].Split('_')[1];
                            userAuthText1.txtAddRightValue.Text = userDataView1.DatasumValue[17].Split('_')[2];
                        }
                    }
                    else {
                        
                        Suffex_preffixList();
                        userAuthText1.txtReviewValue.Text = authList1 + " ل" + preffix[userApplicant1.Authcases(), 7] + " ع" + preffix[userApplicant1.Appcases(), 2] + " و" + preffix[userApplicant1.Authcases(), 8] + " مقام" + preffix[userApplicant1.Appcases(), 0] + " في ";
                    }
                    if (checkBox1.CheckState == CheckState.Checked)
                    {
                        userAuthText1.txtReviewValue.Location = new System.Drawing.Point(336, 162);
                        userAuthText1.txtReviewValue.Size = new System.Drawing.Size(828, 85);
                        userAuthText1.txtAddRightValue.Location = new System.Drawing.Point(387, 253);
                        userAuthText1.txtAddRightValue.Size = new System.Drawing.Size(777, 84);
                        userAuthText1.label34Value.Visible = true;
                        userAuthText1.label36Value.Text = "إضافة نص جديد:";
                    }
                    else
                    {
                        userAuthText1.txtReviewValue.Location = new System.Drawing.Point(336, 76);
                        userAuthText1.txtReviewValue.Size = new System.Drawing.Size(828, 171);
                        userAuthText1.txtAddRightValue.Location = new System.Drawing.Point(336, 253);
                        userAuthText1.txtAddRightValue.Size = new System.Drawing.Size(828, 409);
                        userAuthText1.label34Value.Visible = false;
                        userAuthText1.label36Value.Text = "إضافة نصوص التوكيل:";
                    }
                    userAuthText1.Show();
                    flowLayoutPanel1.Visible = false;
                    flowLayoutPanel2.Visible = false;
                    flowLayoutPanel3.Visible = false;
                    btnPrevious.Visible = false;

                    break;
                case 4:
                    flowLayoutPanel1.Visible = true;
                    flowLayoutPanel2.Visible = true;
                    flowLayoutPanel3.Visible = true;
                    btnPrevious.Visible = true;

                    break;
            }
        }

        void FillDatafromGenArch(string doc, string id, string table)
        {            
            SqlConnection sqlCon = new SqlConnection(DataSource);
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
                if (name == "") return;
                var Data = (byte[])reader["Data1"];
                var ext = reader["Extension1"].ToString();
                var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                File.WriteAllBytes(NewFileName, Data);
                System.Diagnostics.Process.Start(NewFileName);
            }


            sqlCon.Close();
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
            preffix[2, 0] = "ا";
            preffix[3, 0] = "ا";
            preffix[4, 0] = "ا";
            preffix[5, 0] = "ا";

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
            preffix[5, 6] = "ين";

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

            preffix[0, 9] = "";
            preffix[1, 9] = "ة";
            preffix[2, 9] = "ين";
            preffix[3, 9] = "ان";
            preffix[4, 9] = "ات";
            preffix[5, 9] = "ين";

            preffix[0, 10] = "أوكلت";//&&&
            preffix[1, 10] = "أوكلت";
            preffix[2, 10] = "أوكلنا";
            preffix[3, 10] = "أوكلنا";
            preffix[4, 10] = "أوكلنا";
            preffix[5, 10] = "أوكلنا";

            preffix[0, 11] = "تنازلت تنازلاً نهائياً";//&&&
            preffix[1, 11] = "تنازلت تنازلاً نهائياً";
            preffix[2, 11] = "تنازلنا تنازلاً نهائياً";
            preffix[3, 11] = "تنازلنا تنازلاً نهائياً";
            preffix[4, 11] = "تنازلنا تنازلاً نهائياً";
            preffix[5, 11] = "تنازلنا تنازلاً نهائياً";

        }
    }
}
