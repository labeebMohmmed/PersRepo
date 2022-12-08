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
using System.Threading;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Xceed.Words.NET;
using System.Diagnostics;
using Xceed.Document.NET;
using System.Globalization;
using System.Threading;
using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using WIA;
using Image = System.Drawing.Image;
using DocumentFormat.OpenXml.Office2010.Excel;
using Color = System.Drawing.Color;
using System.Runtime.InteropServices.ComTypes;

namespace PersAhwal
{
    //https://www.youtube.com/watch?v=sHJVusS5Qz0
    //https://doc.xceed.com/xceed-document-libraries-for-net/Code_Snippets.html



    public partial class MainForm : Form
    {
        DataTable dataRowTable;
        static string[] queryNewYear = new string[15];
        string DataSource, DataSource56, DataSource57, GregorianDate, HijriDate, GregorianDateReport;
        string[] quorterS = new string[12];
        string[] quorterE = new string[12];
        string[] ProType = new string[100];
        int userId = 1;
        int imagecount = 0;
        string CurrentVersion = "0.0.0.0.O";
        bool rowFound = false;
        bool onUpdate = false;
        string[] PathImages = new string[100];
        private SqlConnection sqlCon;
        static string EmployeeName, UserJobposition;
        int totalrowsAuth = 0, totalrowsAffadivit = 0, totalRowDuration = 0;
        int startofNextWeek;
        string FilespathIn, FilespathOut, ArchFile;
        int V = 0, A = 0, M = 0;
        int[,] report1 = new int[15, 32];
        int[,] DeepReport = new int[100, 32];
        int[,] rep1 = new int[20, 20];
        int i = 0;
        Thread th;
        bool deleteEmptyRows = false;
        bool parrtialAll = true;
        DeviceInfo AvailableScanner = null;
        string PathImage;
        int rCnt;
        bool uploadDocx = true;
        bool textNumber = false;
        int cCnt;
        PictureBox org;
        int handIndex = 0;
        int rw = 0;
        bool VCIndexLoad = false;
        string authNo = "";
        bool setupRequres = false;
        string NewFileNamePic;
        int Messid = 1;
        int MessageDocNo = 0;
        string ArchfilePath = "";
        string MessageNo = "ق س ج/160/01/";
        private static bool NewSettings = false, FirstDate = false, LastDate = false;
        private static int TableIndex = -1, IDNo = -1, IDVANo = -1;
        static string[] query = new string[15];
        static string[] queryUpdate = new string[15];

        string strHijriDate;
        string strViseConsul;
        string bolApplicantSex;
        string filePath = "";
        int intMessageType;
        string strMessageType = "", strEmbassySource = "";
        static string[] AppNameA = new string[2000];
        static string[] oldNewA= new string[2000];
        static string[] oldNewM = new string[2000];
        static string[] DocIDA = new string[2000];
        static string[] GriDateA = new string[2000];


        static string[] MandoubM = new string[2000];
        static string[] AppNameM = new string[2000];

        static string[] DocIDM = new string[2000];
        static string[] GriDateM = new string[2000];

        static int[] IDA = new int[2000];
        static int[] IDV = new int[2000];
        static string[] DocA = new string[2000];
        static string[] DocV = new string [2000];
        static int[] IDM = new int[2000];
        static string[] DocM = new string[2000];
        string[] qureyDataUpdate = new string[7];
        static string[] queryVA = new string[15];
        static string[] queryuPDATE = new string[15];
        static string[] queryTable = new string[15];

        static string[] TableList = new string[15];
        static string[] TableArch = new string[15];
        static string[] TableDocID = new string[15];
        static string[] DuratioListquery = new string[15];
        bool ArchType = false;
        static string[] queryDateList = new string[20];
        static string[] reportItems = new string[20];

        static string[] querydatabase = new string[20];
        static string[] RetrievedNameAffadivit = new string[1000];
        static string[] RetrievedNoAffadivit = new string[1000];
        static string[] RetrievedTypeAffadivit = new string[1000];

        static string[] RetrievedNameAuth = new string[1000];
        static string[] RetrievedTypeAuth = new string[1000];
        static string[] RetrievedAuthPers = new string[1000];
        static string[] RetrievedNoAuth = new string[1000];
        static int[] Months = new int[12];
        static int[,] duartionReport = new int[11, 3000];
        static int[,] duarReport = new int[3, 12];
        static int[,] duartionReportAuth = new int[11, 3000];

        static string[] DocType = new string[6];
        static string[] DocTypeVA = new string[1000];
        static int[] IDNoList = new int[1000];
        static string[] AppNames = new string[1000];
        static string route;
        int[] monthSumV = new int[15];
        int[] monthSumH = new int[15];

        string Model, Output, ServerIP, Login, Pass, Database, FileArch;
        int Hiday, Himonth;
        string[] travType = new string[2];
        string[] IDList= new string[100];
        string[] travType1 = new string[5];
        string[] travType2 = new string[10];
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        string FormDataFile;
        Excel.Range range;
        int iqrarType = 0;
        string fileVersio;
        string primeryLink = "";
        string PrimariFiles = @"D:\PrimariFiles\";
        string Server = "57";
        bool Pers_Peope = true;
        int cl = 0;
        string[] items;
        string[] values;
        string[] colIDs = new string[100];
        bool Showed = false;
        string Career = "";
        bool MessageShowed = false;
        int maxRange = 10;
        bool autoCompleteMode = false;
        string mandoubInfo = "";
        string archState = "";
        bool nameNo = true;
        int ProcReqID = 0;
        string ServerModelFiles, ServerModelForms;
        string[] rightColNames;
        string LocalModelFiles = "";
        string LocalModelForms = "";
        public MainForm(string career,int id, string server, string Employee, string jobposition, string dataSource56, string dataSource57, string filepathIn, string filepathOut, string archFile, string formDataFile, bool pers_Peope, string gregorianDate, string hijriDate, string modelFiles,string modelForms)
        {
            InitializeComponent();
            userId = id;
            Server = server;
            DataSource = dataSource57;
            DataSource56 = dataSource56;
            DataSource57 = dataSource57;
            //checkColumnNames("تقاضي_4");
            GregorianDate = gregorianDate;
            HijriDate = hijriDate;
            Console.WriteLine(1);
            //MessageBox.Show(HijriDate);
            // BackupDataBase(DataSource57, "AhwalDataBase");

            //DeleteEmptyFiles("TableAddContext");
            if (Server == "57")
            {


                panel2.BackColor = System.Drawing.SystemColors.ButtonShadow;

            }
            else if (Server == "56")
            {
                DataSource = dataSource56;
                this.Name = "القائمة الرئيسة نافذة قسم شؤون الرعايا";
                
            }

            if (Directory.Exists(@"D:\"))
            {
                primeryLink = @"D:\PrimariFiles\";

            }
            else
            {
                string appFileName = Environment.GetCommandLineArgs()[0];
                string directory = Path.GetDirectoryName(appFileName);
                directory = directory + @"\";
                primeryLink = directory + @"PrimariFiles\";
                //if (Server == "57")
                //    primeryLink = directory + @"PrimariFiles\Personnel\";
                //else if (Server == "56")
                //    primeryLink = directory + @"PrimariFiles\SuddaneseAffairs\";
            }

            LocalModelFiles = primeryLink + @"ModelFiles\";
            LocalModelForms = primeryLink + @"FormData\";

            Console.WriteLine(2);
            sqlCon = new SqlConnection(DataSource);
            Console.WriteLine(3);
            columnNames();
            Console.WriteLine(4);
            GregorianDate = gregorianDate;
            Affbtn0.Visible = !pers_Peope;
            //com247, 37 aff
            //com 429, 459 pers
            if (pers_Peope)
            {
                //MessageBox.Show("الاحوال الشخضية");
                foreach (Control control in this.Controls)
                {
                    if (control.Name.Contains("persbtn"))
                    {
                        control.Visible = true;
                        control.BringToFront();                        
                    } 
                    else if (control.Name.Contains("Affbtn"))
                    {
                        control.Visible = false; ;
                        control.SendToBack();
                    }
                }
                Affbtn3.Visible = Affbtn0.Visible = Affbtn5.Visible = false;


            }
            else {
                //MessageBox.Show("شؤون الرعايا");
                foreach (Control control in this.Controls)
                {
                    if (control.Name.Contains("persbtn"))
                    {
                        control.Visible = false;
                        control.SendToBack();
                    }
                    else if (control.Name.Contains("Affbtn"))
                    {
                        control.Visible = true;
                        control.BringToFront();
                    }
                }
                persbtn6.Visible = false;
                Combtn0.Location = new System.Drawing.Point(427, 402);
                Combtn1.Location = new System.Drawing.Point(427, 402 + 39);
                Combtn2.Location = new System.Drawing.Point(427, 402 +(39*2));
                Combtn3.Location = new System.Drawing.Point(427, 402 +(39*3));
                Combtn4.Location = new System.Drawing.Point(427, 402 +(39*4));
            }
            perbtn1.Visible = false;
            Pers_Peope = pers_Peope;
            Affbtn6.Visible = !pers_Peope;
            uploadDocx = true;
            
            EmployeeName = Employee;
            Career = career;
            //if (Career == "موظف ارشفة") timer5.Enabled = true;
            ServerModelFiles = modelFiles;
            ServerModelForms = modelForms;
            FormDataFile = formDataFile;
            FilespathIn = filepathIn;
            ArchFile = archFile;
            FilespathOut = filepathOut;
            
            UserJobposition = jobposition;
            //persbtn2MessageBox.Show(UserJobposition);
            ConsulateEmployee.Text = EmployeeName;
            TablesList();
            fileVersio = primeryLink + @"\SuddaneseAffairs\getVersio.txt";
            Console.WriteLine(5);
            CurrentVersion = File.ReadAllText(fileVersio);
            //MessageBox.Show(CurrentVersion);
            Console.WriteLine("Main fileVersio " + Server);
            ClearFileds();
            //MessageBox.Show(DataSource);
            loadSettings(DataSource, false, false, false, false);
            ReportNo.Text = "ق س ج/" + DateTime.Now.ToString("dd") + "/11" + "/160";

            Console.WriteLine(6);
            perbtn1.Visible = false;
            persbtn2.Visible = false;
            persbtn2.SendToBack();
            if (UserJobposition.Contains("قنصل"))
            {
                picSettings.Visible = Affbtn0.Visible = true;
                empUpdate.Visible = false;
                picVersio.BringToFront();
            }
            else
            {
                picSettings.Visible = false;
                Aprove.Text = "تعديل بيانات الدخول";
                empUpdate.BringToFront();
                    empUpdate.Visible = true;
            }

                if (!Directory.Exists(PrimariFiles))
            {
                string appFileName = Environment.GetCommandLineArgs()[0];
                string directory = Path.GetDirectoryName(appFileName);
                directory = directory + @"\";
                PrimariFiles = directory + @"PrimariFiles\";
            }

            if (!File.Exists(primeryLink + "fileUpdate.txt"))
            {

                dataSourceWrite(primeryLink + "fileUpdate.txt", "files are fully update");
            }
            Console.WriteLine(7);
            //backgroundWorker1.RunWorkerAsync();
            //backgroundWorker2.RunWorkerAsync();


            //MessageBox.Show("FilespathIn " + FilespathIn);
            //MessageBox.Show("ServerModelFiles " + ServerModelFiles);
            //MessageBox.Show("FormDataFile " + FormDataFile);
            //MessageBox.Show("ServerModelForms " + ServerModelForms);

            quorterS[0] = "-01-01-";
            quorterE[0] = "-01-31-";
            quorterS[1] = "-02-01-";
            quorterE[1] = "-02-29-";
            quorterS[2] = "-011-01";
            quorterE[2] = "-12-31";

            //string from = "2022" + quorterS[2];
            //string to = "2022" + quorterE[2];
            //DailyListcustm(from, to, 11);
            //PrintReport.PerformClick();

        }

        private void backgroundWorker1_DoWork_1(object sender, DoWorkEventArgs e)
        {

            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            string date = DateTime.Now.ToString("dd/MM/yyyy");

            string[] archFiles = Directory.GetFiles(FilespathOut);
            foreach (string str in archFiles) {
                FileInfo fileInfo = new FileInfo(str);
                DateTime dt = fileInfo.LastWriteTime.Date;
                
                    Console.WriteLine (date + " - " + str +" date is "+dt.ToString().Split(' ')[0]);
                try
                {
                    if (date != dt.ToString().Split(' ')[0]) File.Delete(str);
                }
                catch (Exception ex) { }
            }
            //string DFile = @"D:\";
            //if (Directory.Exists(DFile))
            //{
            //    archFiles = Directory.GetFiles(DFile);
            //    foreach (string str in archFiles)
            //    {
            //        FileInfo fileInfo = new FileInfo(str);
            //        DateTime dt = fileInfo.LastWriteTime.Date;

            //        Console.WriteLine(date + " - " + str + " date is " + dt.ToString().Split(' ')[0]);
            //        try
            //        {
            //            if (fileInfo.Name.Contains("ArchiveFiles") &&  date != dt.ToString().Split(' ')[0]) File.Delete(str);
            //        }
            //        catch (Exception ex) { }
            //    }
            //}
            //string DFile = @"D:\PrimariFiles\";
            //if (Directory.Exists(DFile))
            //{
            //    archFiles = Directory.GetFiles(DFile);
            //    foreach (string str in archFiles)
            //    {
            //        if (!str.Contains(".txt"))
            //        {
            //            try
            //            {
            //                File.Delete(str);
            //            }
            //            catch (Exception ex) { }
            //        }
            //    }
            //}
            
            //DFile = @"D:\PrimariFiles\ModelFiles";
            ////DFile = @"\\192.168.100.100\Users\Public\Documents\ModelFiles";
            //if (Directory.Exists(DFile))
            //{
            //    archFiles = Directory.GetFiles(DFile);
            //    foreach (string str in archFiles)
            //    {
            //        FileInfo fileInfo = new FileInfo(str);
            //        if (fileInfo.Name.Contains("Docx")||str.Contains(".pdf") ||str.Contains(".odt") || str.Contains(".xlsx") || str.Contains(".txt")|| str.Contains(".jpg")|| str.Contains(".jpg"))
            //        {
            //            try {
            //                File.Delete(str); 
            //            }
            //            catch (Exception ex) { }
            //        }
            //    }
            //}
            
            //DFile = @"D:\PrimariFiles\FormData";
            ////DFile = @"\\192.168.100.100\Users\Public\Documents\FormData";
            //if (Directory.Exists(DFile))
            //{
            //    archFiles = Directory.GetFiles(DFile);
            //    foreach (string str in archFiles)
            //    {
            //        FileInfo fileInfo = new FileInfo(str);
                    
            //        if (str.Contains(".db") ||str.Contains(".pdf") ||str.Contains(".odt") || str.Contains(".xlsx") || str.Contains(".txt")|| str.Contains(".jpg")|| str.Contains(".jpg"))
            //        {
            //            try {
            //                File.Delete(str); 
            //            }
            //            catch (Exception ex) { }
            //        }
            //    }
            //}

        }
        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
        //private string[] getColList(string table)
        //{
        //    SqlConnection sqlCon = new SqlConnection(DataSource57);
        //    if (sqlCon.State == ConnectionState.Closed)

        //        sqlCon.Open();
        //    SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('" + table + "')", sqlCon);
        //    sqlDa.SelectCommand.CommandType = CommandType.Text;
        //    DataTable dtbl = new DataTable();
        //    sqlDa.Fill(dtbl);
        //    sqlCon.Close();
        //    string[] allList = new string[dtbl.Rows.Count];
        //    //MessageBox.Show(dtbl.Rows.Count.ToString());
        //    int i = 0;
        //    foreach (DataRow row in dtbl.Rows)
        //    {
        //        allList[i] = row["name"].ToString();
        //        i++;
        //    }
        //    return allList;

        //}
        private bool doubleCheckArch(string v1)
        {
            int index = Convert.ToInt32(v1.Split('/')[3]) -1;
            string table = TableList[index];
            string arch = TableArch[index];
            string docID = TableDocID[index];
            //MessageBox.Show(table + " - "+arch);
            string query;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "select "+arch+" FROM "+table+" where "+ docID+ " = @" + docID;
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, Con);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@" + docID, v1);
            DataTable dtbl2 = new DataTable();
            sqlDa.Fill(dtbl2);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl2.Rows)
            {
                if (dataRow[arch].ToString().Contains("مؤرشف نهائي"))
                return true;
            }
            return false;
        }

        private string getWafidTable(int index)
        {
            string table = "";
            switch (index)
            {
                case 0:
                    table = "TableWafid";
                    break;
                case 1:
                    table = "TableWafidJed";
                    break;
                case 2:
                    table = "TableWafidMekkah";
                    break;
                case 3:
                    table = "TableTarheel";
                    break;
                case 4:
                    table = "TableTransfer";
                    break;
                case 5:
                    table = "TableCommity";
                    break;
            }
            return table;
        }
        private void deleteRowsData(string v1)
        {
            string query;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "DELETE FROM archives where docID = @docID";
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@docID", v1);
            sqlCmd.ExecuteNonQuery();
            Con.Close();
        }

        void FillDataGridAdd()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }

            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT ArabCountries from TableListCombo", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl2 = new DataTable();
            sqlDa.Fill(dtbl2);
            sqlCon.Close();
            int id = 1;
            foreach (DataRow dataRow in dtbl2.Rows)
            {
                upDateCountry(id, dataRow["ArabCountries"].ToString());
                id++;
            }
        }

        private void upDateCountry(int id, string text)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("update TableListCombo set ArabCountries=@ArabCountries where ID=@id", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@ArabCountries", text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void BackupDataBase(string source, string dataBase)
        {
            //OpenFileDialog dlg = new OpenFileDialog();
            //dlg.ShowDialog();

            string file = "D:";
            //dlg.FileName = ;
            string query = "BACKUP DATABASE " + dataBase + " TO  DISK = '" + file + "\\" + dataBase + "-" + DateTime.Now.Ticks.ToString() + ".bak'";
            string query1 = "BACKUP DATABASE [AhwalDataBase] TO  DISK = N'D:\\SudanAffairs452145' WITH NOFORMAT, NOINIT,  NAME = N'AhwalDataBase-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10GO";
            try
            {
                SqlConnection sqlCon = new SqlConnection(source);
                if (sqlCon.State == ConnectionState.Closed)
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return; }
                SqlCommand cmd = new SqlCommand(query, sqlCon);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Backup is done !!");
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
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

        private string getPreivilage()
        {
            string prev = "0_0_0_0_0_0_0_0_0_0";
            //MessageBox.Show(DataSource);
            using (SqlConnection saConn = new SqlConnection(DataSource))
            {
                saConn.Open();

                string query = "select previliage from TableUser where EmployeeName=@EmployeeName";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.AddWithValue("@EmployeeName", EmployeeName);
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                foreach (DataRow dataRow in table.Rows)
                {
                    if (!String.IsNullOrEmpty(dataRow["previliage"].ToString()))
                        prev = dataRow["previliage"].ToString();
                }
                saConn.Close();
            }
            return prev;
        }
        private string checkColumnName(string source1, string table)
        {
            SqlConnection sqlCon = new SqlConnection(source1);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return ""; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS " + table, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            int itemsCount = 0;
            string item = "", value = "";
            items = new string[dtbl.Rows.Count];
            values = new string[dtbl.Rows.Count];
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(dataRow["COLUMN_NAME"].ToString()) && dataRow["COLUMN_NAME"].ToString() != "ID" && dataRow["COLUMN_NAME"].ToString() != "Data1" && dataRow["COLUMN_NAME"].ToString() != "Data2" && dataRow["COLUMN_NAME"].ToString() != "Data3")
                {
                    if (itemsCount == 0)
                    {
                        items[0] = dataRow["COLUMN_NAME"].ToString();
                        values[0] = "@" + dataRow["COLUMN_NAME"].ToString();
                        item = dataRow["COLUMN_NAME"].ToString();
                        value = "@" + dataRow["COLUMN_NAME"].ToString();
                    }
                    else {
                        items[itemsCount] = dataRow["COLUMN_NAME"].ToString();
                        values[itemsCount] = "@" + dataRow["COLUMN_NAME"].ToString();
                        item = item + "," + dataRow["COLUMN_NAME"].ToString();
                        value = value + ",@" + dataRow["COLUMN_NAME"].ToString();
                    }
                    itemsCount++;

                }
            }
            string query = "INSERT INTO " + table + " (" + item + ") values (" + value + ")";


            return query + "_" + itemsCount.ToString();

        }

        private void addData(string source1, string source2, string table)
        {
            string query = checkColumnName(source1, table);
            //MessageBox.Show(query);
            SqlConnection sqlCon = new SqlConnection(source1);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from " + table, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            MessageBox.Show("select * from " + table);
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            //MessageBox.Show("length " + query.Split('_')[1]);
            sqlCon = new SqlConnection(source2);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand(query.Split('_')[0], sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                //Console.WriteLine("id " + dataRow["ID"].ToString());
                if (Convert.ToInt32(dataRow["ID"].ToString()) > 5611)
                {
                    for (int idcol = 0; idcol < Convert.ToInt32(query.Split('_')[1]); idcol++)
                    {
                        //if (!values[idcol].Contains("Data"))
                        //Console.WriteLine("idcol " + values[idcol]);
                        //Console.WriteLine("items " + items[idcol]);
                        if (items[idcol].Contains("Data"))
                            sqlCmd.Parameters.AddWithValue(values[idcol], (byte[])dataRow[items[idcol]]);
                        else
                            sqlCmd.Parameters.AddWithValue(values[idcol], dataRow[items[idcol]].ToString());
                    }
                    sqlCmd.ExecuteNonQuery();

                    sqlCon.Close();
                    return;
                }
            }


        }
        private void columnNames() {
            queryNewYear[0] = "INSERT INTO TableDocIqrar (DocID,GriDate) values (@DocID,@GriDate)";
            queryNewYear[1] = "INSERT INTO TableTravIqrar (DocID,GriDate) values (@DocID,@GriDate)";
            queryNewYear[2] = "INSERT INTO TableMultiIqrar (DocID,GriDate) values (@DocID,@GriDate)";
            queryNewYear[3] = "INSERT INTO TableVisaApp (DocID,GriDate) values (@DocID,@GriDate)";
            queryNewYear[4] = "INSERT INTO TableFamilySponApp (DocID,GriDate) values (@DocID,@GriDate)";
            queryNewYear[5] = "INSERT INTO TableForensicApp (DocID,GriDate) values (@DocID,@GriDate)";
            queryNewYear[6] = "INSERT INTO TableTRName (DocID,GriDate) values (@DocID,@GriDate)";
            queryNewYear[7] = "INSERT INTO TableStudent (DocID,GriDate) values (@DocID,@GriDate)";
            queryNewYear[8] = "INSERT INTO TableMarriage (DocID,GriDate) values (@DocID,@GriDate)";
            queryNewYear[9] = "INSERT INTO TableCollection (رقم_المعاملة,التاريخ_الميلادي) values (@رقم_المعاملة,@التاريخ_الميلادي)";
            queryNewYear[11] = "INSERT INTO TableAuth (التاريخ_الميلادي,رقم_التوكيل) values (@رقم_التوكيل,@التاريخ_الميلادي)";
            queryNewYear[12] = " update TableSettings set SudAffNo=@SudAffNo where ID = @id";



        }
        private void NewYearEntry(int FormType, string year, string Gredate)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand(queryNewYear[FormType - 1], sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            if (FormType == 12)
            {
                sqlCmd.Parameters.AddWithValue("@رقم_التوكيل", "ق س ج/80/" + year + "/" + (FormType + 1).ToString() + "/0");
                sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", Gredate);
            }
            else if (FormType == 13)
            {                
                sqlCmd.Parameters.AddWithValue("@رقم_المعاملة", "ق س ج/80/22/13/0");
                sqlCmd.Parameters.AddWithValue("@id", 1);
            }
            else
            {
                sqlCmd.Parameters.AddWithValue("@DocID", "ق س ج/80/" + year + "/" + FormType.ToString() + "/0");
                sqlCmd.Parameters.AddWithValue("@GriDate", Gredate);
            }
            sqlCmd.ExecuteNonQuery();

            sqlCon.Close();
        }

        private void UserLogOut()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return; }
                SqlCommand sqlCmd = new SqlCommand("update TableUserLog set timeDateOut=@timeDateOut where ID=@id", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@id", userId);
                sqlCmd.Parameters.AddWithValue("@timeDateOut", DateTime.Now.ToString("G"));
                sqlCmd.ExecuteNonQuery();
                sqlCon.Close();
            }
            catch (Exception ex) { 
            }
        }

        private void closeToUpdate(string text)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("update TableSettings set closeToUpdate=@closeToUpdate where ID=@id", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", 1);
            sqlCmd.Parameters.AddWithValue("@closeToUpdate", text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void VersionUpdate(string version)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource56);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("update TableSettings set Version=@Version where ID='1'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@Version", version);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private string getIqrarType()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return ""; }
            string settingData = "select نوع_الإجراء from TableCollection where ID='1'";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            string ver = "1.0.0.0";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                ver = dataRow["Version"].ToString();
            }
            return ver;
        }


        private string getVersio()
        {
            //return "";
            string ver = "1.0.0.0";
            SqlConnection sqlCon = new SqlConnection(DataSource56);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return ""; }
                string settingData = "select Version from TableSettings where ID='1'";
                SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);


                foreach (DataRow dataRow in dtbl.Rows)
                {
                    ver = dataRow["Version"].ToString();

                }
            }
            catch (Exception ex) { }
            return ver;
        }


        private string getAppFolder()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return ""; }
            string settingData = "select FolderApp from TableSettings where ID='1'";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            string ver = "";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                ver = dataRow["FolderApp"].ToString();

            }
            return ver;
        }

        private void ReqestData()
        {
            //MessageBox.Show(rw.ToString());
            for (rCnt = 1; rCnt < rw; rCnt++)
            {
                //string Suitcase = (string)(range.Cells[rCnt, 6] as Excel.Range).Value2;
                ////MessageBox.Show("Suitcase " + Suitcase);
                //string finishDate = Convert.ToString((range.Cells[rCnt, 5] as Excel.Range).Value2);
                ////MessageBox.Show("finishDate " + finishDate);
                //string receiveDate = Convert.ToString((range.Cells[rCnt, 4] as Excel.Range).Value2);
                ////MessageBox.Show("receiveDate " + receiveDate);
                //string messDate = Convert.ToString((range.Cells[rCnt, 3] as Excel.Range).Value2);
                ////MessageBox.Show("messDate " + messDate);
                //string messID = Convert.ToString((range.Cells[rCnt, 2] as Excel.Range).Value2);
                ////MessageBox.Show("messID " + messID);
                //string messName = Convert.ToString((range.Cells[rCnt, 1] as Excel.Range).Value2);
                ////MessageBox.Show("messName " + messName);
                //string messComment = "لا تعليق";// Convert.ToString((range.Cells[rCnt, 7] as Excel.Range).Value2);
                ////MessageBox.Show("messComment " + messComment);
                //NewMandoubData(Suitcase, finishDate, receiveDate, messDate, messID, messName, messComment);
                string jobs = (string)(range.Cells[rCnt, 2] as Excel.Range).Value2;
                NewMandoubData(jobs);
            }
        }

        private void NewMandoubData(string Suitcase, string finishDate, string receiveDate, string messDate, string messID, string messName, string messComment)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource56);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }

            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableSuitCase (رقم_لبرقية, تاريخ_لبرقية, مقدم_الطلب, القضية, تاريخ_الاستلام, تاريخ_الرفع, التاريخ_الميلادي, التاريخ_الهجري, مدير_القسم, اسم_الموظف, تعليق)  values (@رقم_لبرقية, @تاريخ_لبرقية, @مقدم_الطلب, @القضية, @تاريخ_الاستلام, @تاريخ_الرفع, @التاريخ_الميلادي, @التاريخ_الهجري, @مدير_القسم, @اسم_الموظف, @تعليق) ", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_لبرقية", messID);
            sqlCmd.Parameters.AddWithValue("@تاريخ_لبرقية", messDate);
            sqlCmd.Parameters.AddWithValue("@مقدم_الطلب", messName);
            sqlCmd.Parameters.AddWithValue("@القضية", Suitcase);
            sqlCmd.Parameters.AddWithValue("@تاريخ_الاستلام", receiveDate);
            sqlCmd.Parameters.AddWithValue("@تاريخ_الرفع", finishDate);
            sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", GregorianDate.Trim());
            sqlCmd.Parameters.AddWithValue("@التاريخ_الهجري", HijriDate.Trim());
            sqlCmd.Parameters.AddWithValue("@مدير_القسم", attendedVC.Text.Trim());
            sqlCmd.Parameters.AddWithValue("@اسم_الموظف", ConsulateEmployee.Text.Trim());
            sqlCmd.Parameters.AddWithValue("@تعليق", messComment);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void NewMandoubData(string jobs)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource56);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }

            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableListCombo (jobs)  values (@jobs) ", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@jobs", jobs);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void loadExcel()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
        }



        private string[] qureyFunction(string tableName, bool state)
        {
            qureyDataUpdate[0] = "UPDATE " + tableName + " SET Data1=@Data1 WHERE ID=@ID";
            qureyDataUpdate[1] = "UPDATE " + tableName + " SET Extension1=@Extension1 WHERE ID=@ID";
            if (state) qureyDataUpdate[2] = "UPDATE " + tableName + " SET FileName1=@FileName1 WHERE ID=@ID";
            else qureyDataUpdate[2] = "UPDATE " + tableName + " SET ارشفة_المستندات=@ارشفة_المستندات WHERE ID=@ID";

            qureyDataUpdate[3] = "UPDATE " + tableName + " SET Data2=@Data2 WHERE ID=@ID";
            qureyDataUpdate[4] = "UPDATE " + tableName + " SET Extension2=@Extension2 WHERE ID=@ID";
            if (state) qureyDataUpdate[5] = "UPDATE " + tableName + " SET FileName2=@FileName2 WHERE ID=@ID";
            else qureyDataUpdate[5] = "UPDATE " + tableName + " SET المكاتبة_النهائية=@المكاتبة_النهائية WHERE ID=@ID";

            if (state) qureyDataUpdate[6] = "UPDATE " + tableName + " SET ArchivedState=@ArchivedState WHERE ID=@ID";
            else qureyDataUpdate[6] = "UPDATE " + tableName + " SET حالة_الارشفة=@حالة_الارشفة WHERE ID=@ID";



            return qureyDataUpdate;
        }

        private void TablesList()
        {
            travType2[0] = "إثبات حياة";
            travType2[1] = "إثبات حالة إجتماعية(متزوج)";
            travType2[1] = "إثبات حالة إجتماعية(غير متزوج)";
            travType2[1] = "إثبات حالة إجتماعية(أرملة)";
            travType2[1] = "إعفاء خروج جزئي";
            travType2[1] = "بلوغ سن الرشد";
            travType2[1] = "خطة إسكانية";
            travType2[1] = "إعالة أسرية";

            query[1] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender,Embassy,ProType  from TableTravIqrar where DocID=@DocID";
            query[2] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender,IqrarPurpose from TableMultiIqrar where DocID=@DocID";
            query[3] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableVisaApp where DocID=@DocID";
            query[4] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender  from TableFamilySponApp where DocID=@DocID";
            query[5] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableForensicApp where DocID=@DocID";
            query[6] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender,IqrarType  from TableTRName where DocID=@DocID";
            query[7] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender  from TableStudent where DocID=@DocID";
            query[8] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender  from TableMarriage where DocID=@DocID";
            query[9] = "select ID,مقدم_الطلب,التاريخ_الميلادي,مقدم_الطلب,اسم_المندوب,حالة_الارشفة,مقدم_الطلب,مقدم_الطلب,التاريخ_الهجري,موقع_المعاملة,النوع,نوع_الإجراء  from TableCollection where رقم_المعاملة=@رقم_المعاملة";
            query[12] = "select ID,مقدم_الطلب,التاريخ_الميلادي,المعالجة,اسم_المندوب,حالة_الارشفة,ارشفة_المستندات,المكاتبة_النهائية,التاريخ_الهجري,موقع_التوكيل,النوع,وجهة_التوكيل,نوع_التوكيل  from TableAuth where رقم_التوكيل=@رقم_التوكيل";
            query[10] = "select ID,AppName,AuthName,AuthNo,Gender,Institute,GriDate,Viewed,FileName1,Comment,ArchivedState  from TableReceMess where DocID=@DocID";
            query[11] = "select ID,AppName,Gender,Institute,GriDate,Viewed,FileName1,Comment,ArchivedState,HandTime  from TableHandAuth where DocID=@DocID";





            travType1[0] = "";
            travType1[1] = "إقرار بطلب نقل كفالة";
            travType1[2] = "إقرار بالموافقة بنقل كفالة";
            travType1[3] = "إقرار بالموافقة بنقل كفالة";
            travType1[4] = "إقرار بموافقة استقدام";


            queryVA[0] = "select ID,AppName,Viewed,ArchivedState,DocID,GriDate,DataInterType,FileName2,DataMandoubName,DataInterName  from TableDocIqrar";
            queryVA[1] = "select ID,AppName,Viewed,ArchivedState,DocID,GriDate,DataInterType,FileName2,DataMandoubName,DataInterName  from TableTravIqrar";
            queryVA[2] = "select ID,AppName,Viewed,ArchivedState,DocID,GriDate,DataInterType,FileName2,DataMandoubName,DataInterName  from TableMultiIqrar";
            queryVA[3] = "select ID,AppName,Viewed,ArchivedState,DocID,GriDate,DataInterType,FileName2,DataMandoubName,DataInterName  from TableVisaApp";
            queryVA[4] = "select ID,AppName,Viewed,ArchivedState,DocID,GriDate,DataInterType,FileName2,DataMandoubName,DataInterName  from TableFamilySponApp";
            queryVA[5] = "select ID,AppName,Viewed,ArchivedState,DocID,GriDate,DataInterType,FileName2,DataMandoubName,DataInterName  from TableForensicApp";
            queryVA[6] = "select ID,AppName,Viewed,ArchivedState,DocID,GriDate,DataInterType,FileName2,DataMandoubName,DataInterName  from TableTRName";
            queryVA[7] = "select ID,AppName,Viewed,ArchivedState,DocID,GriDate,DataInterType,FileName2,DataMandoubName,DataInterName  from TableStudent";
            queryVA[8] = "select ID,AppName,Viewed,ArchivedState,DocID,GriDate,DataInterType,FileName2,DataMandoubName,DataInterName  from TableMarriage";
            queryVA[9] = "select ID,مقدم_الطلب,حالة_الارشفة,حالة_الارشفة,رقم_المعاملة,التاريخ_الميلادي,طريقة_الطلب,مقدم_الطلب,اسم_المندوب,نوع_الإجراء  from TableCollection";
            queryVA[10] = "";
            queryVA[11] = "select ID,مقدم_الطلب,المعالجة,حالة_الارشفة,رقم_التوكيل,التاريخ_الميلادي,DocxData,Extension3,طريقة_الطلب,المكاتبة_النهائية,اسم_المندوب,اسم_الموظف,fileUpload from TableAuth";

            
            queryuPDATE[0] = "update TableDocIqrar set ArchivedState=@ArchivedState where ID=id";
            queryuPDATE[1] = "update TableTravIqrar set ArchivedState=@ArchivedState where ID=id";
            queryuPDATE[2] = "update TableMultiIqrar set ArchivedState=@ArchivedState where ID=id";
            queryuPDATE[3] = "update TableVisaApp set ArchivedState=@ArchivedState where ID=id";
            queryuPDATE[4] = "update TableFamilySponApp set ArchivedState=@ArchivedState where ID=id";
            queryuPDATE[5] = "update TableForensicApp set ArchivedState=@ArchivedState where ID=id";
            queryuPDATE[6] = "update TableTRName set ArchivedState=@ArchivedState where ID=id";
            queryuPDATE[7] = "update TableStudent set ArchivedState=@ArchivedState where ID=id";
            queryuPDATE[8] = "update TableMarriage set ArchivedState=@ArchivedState where ID=id";
            queryuPDATE[9] = "update TableCollection set حالة_الارشفة=@حالة_الارشفة where ID=id";
            queryuPDATE[10] = "";
            queryuPDATE[10] = "update TableAuth set حالة_الارشفة=@حالة_الارشفة where ID=id";

            TableList[0] = "TableDocIqrar";
                TableList[1] = "TableTravIqrar";
                TableList[2] = "TableMultiIqrar";
                TableList[3] = "TableVisaApp";
                TableList[4] = "TableFamilySponApp";
                TableList[5] = "TableForensicApp";
                TableList[6] = "TableTRName";
                TableList[7] = "TableStudent";
                TableList[8] = "TableMarriage";
                TableList[9] = "TableCollection";
                TableList[10] = "";
                TableList[11] = "TableAuth";
                TableList[12] = "TableWafid";
                TableList[13] = "TableSuitCase";

            TableArch[0] = "ArchivedState";
            TableArch[1] = "ArchivedState";
                TableArch[2] = "ArchivedState";
                TableArch[3] = "ArchivedState";
                TableArch[4] = "ArchivedState";
                TableArch[5] = "ArchivedState";
                TableArch[6] = "ArchivedState";
                TableArch[7] = "ArchivedState";
                TableArch[8] = "ArchivedState";
                TableArch[9] = "ArchivedState";
                TableArch[10] = "ArchivedState";
                TableArch[11] = "حالة_الارشفة";
                TableArch[12] = "ArchivedState";
                TableArch[13] = "ArchivedState";

            TableDocID[0] = "docID";
            TableDocID[1] = "docID";
                TableDocID[2] = "docID";
                TableDocID[3] = "docID";
                TableDocID[4] = "docID";
                TableDocID[5] = "docID";
                TableDocID[6] = "docID";
                TableDocID[7] = "docID";
                TableDocID[8] = "docID";
                TableDocID[9] = "docID";
                TableDocID[10] = "docID";
                TableDocID[11] = "رقم_التوكيل";
                TableDocID[12] = "رقم_المعاملة";
                TableDocID[13] = "docID";

                queryTable[0] = "TableDocIqrar";
            queryTable[1] = "TableTravIqrar";
            queryTable[2] = "TableMultiIqrar";
            queryTable[3] = "TableVisaApp";
            queryTable[4] = "TableFamilySponApp";
            queryTable[5] = "TableForensicApp";
            queryTable[6] = "TableTRName";
            queryTable[7] = "TableStudent";
            queryTable[8] = "TableMarriage";
            queryTable[9] = "TableCollection";
            queryTable[10] = "TableAuth";

            DuratioListquery[0] = "select ID  from TableDocIqrar where GriDate=@GriDate";
            DuratioListquery[1] = "select ID  from TableTravIqrar where GriDate=@GriDate";
            DuratioListquery[2] = "select ID  from TableMultiIqrar where GriDate=@GriDate";
            DuratioListquery[3] = "select ID  from TableVisaApp where GriDate=@GriDate";
            DuratioListquery[4] = "select ID  from TableFamilySponApp where GriDate=@GriDate";
            DuratioListquery[5] = "select ID  from TableForensicApp where GriDate=@GriDate";
            DuratioListquery[6] = "select ID  from TableTRName where GriDate=@GriDate";
            DuratioListquery[7] = "select ID  from TableStudent where GriDate=@GriDate";
            DuratioListquery[8] = "select ID  from TableMarriage where GriDate=@GriDate";
            DuratioListquery[9] = "select ID  from TableCollection where التاريخ_الميلادي=@التاريخ_الميلادي";
            DuratioListquery[10] = "select ID,إجراء_التوكيل  from TableAuth where التاريخ_الميلادي=@التاريخ_الميلادي";

            reportItems[0] = "اقرار استخراج أوراق ثبوتية موافقة بالسفر";
            reportItems[1] = "إقرار لاغراض مختلفة";
            reportItems[2] = "إقرار لاثبات صحة إسمين";
            reportItems[3] = "إقرارات عامة";
            reportItems[4] = "إقرار كفالة أفراد أسرة";
            reportItems[5] = "توكيل";
            reportItems[6] = "إفادة للادلة الجنائية";
            reportItems[7] = "إفادة عدم ممانعة زواج";
            reportItems[8] = "إفادة تسجيل ببرنامج دراسي";
            reportItems[9] = "مذكرة لمنح تأشيرة";




            queryDateList[0] = "select AppName from TableDocIqrar,GriDate where GriDate=@GriDate";
            queryDateList[1] = "select AppName,ProType,DocID,ArchivedState,DataInterType,GriDate from TableTravIqrar where GriDate=@GriDate";
            queryDateList[2] = "select AppName,IqrarPurpose,DocID,ArchivedState,DataInterType,GriDate from TableMultiIqrar where GriDate=@GriDate";
            queryDateList[3] = "select AppName,IqrarType,DocID,ArchivedState,DataInterType,GriDate from TableTRName where GriDate=@GriDate";
            queryDateList[4] = "select مقدم_الطلب,نوع_الإجراء,رقم_المعاملة,حالة_الارشفة,طريقة_الطلب,التاريخ_الميلادي from TableCollection where التاريخ_الميلادي=@التاريخ_الميلادي";
            queryDateList[5] = "select AppName,ProCase,DocID,ArchivedState,DataInterType,GriDate from TableFamilySponApp where GriDate=@GriDate";
            queryDateList[6] = "select مقدم_الطلب,إجراء_التوكيل,نوع_التوكيل,رقم_التوكيل,الموكَّل,حالة_الارشفة,طريقة_الطلب,نوع_التوكيل ,التاريخ_الميلادي from TableAuth where التاريخ_الميلادي=@التاريخ_الميلادي";
            queryDateList[7] = "select AppName,DocID,ArchivedState,DataInterType,GriDate from TableForensicApp where GriDate=@GriDate";
            queryDateList[8] = "select AppName,DocID,ArchivedState,DataInterType,GriDate from TableMarriage where GriDate=@GriDate";
            queryDateList[9] = "select AppName,DocID,ArchivedState,DataInterType,GriDate from TableStudent where GriDate=@GriDate";
            queryDateList[10] = "select AppName,DocID,ArchivedState,DataInterType,GriDate from TableVisaApp where GriDate=@GriDate";
            queryDateList[11] = "select DocNo,GriDate from TableHandAuth where GriDate=@GriDate";
            queryDateList[12] = "select رقم_اذن_الدفن,التاريخ_الميلادي from TablePassAway where التاريخ_الميلادي=@التاريخ_الميلادي";
            queryDateList[13] = "select رقم_المعاملة,التاريخ_الميلادي from TableMerrageDoc where التاريخ_الميلادي=@التاريخ_الميلادي";
            queryDateList[14] = "select رقم_المعاملة,التاريخ_الميلادي from TableDivorce  where التاريخ_الميلادي=@التاريخ_الميلادي";


            querydatabase[0] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableDocIqrar where ID=@id";
            querydatabase[1] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableTravIqrar where ID=@id";
            querydatabase[2] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableMultiIqrar where ID=@id";
            querydatabase[3] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableVisaApp where ID=@id";
            querydatabase[4] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableFamilySponApp where ID=@id";
            querydatabase[5] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableForensicApp where ID=@id";
            querydatabase[6] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableTRName where ID=@id";
            querydatabase[7] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableStudent where ID=@id";
            querydatabase[8] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableMarriage where ID=@id";
            querydatabase[9] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableCollection where ID=@id";


            DocTypeVA[0] = "إقرار باستخراج أوراق ثبوتية";
            DocTypeVA[1] = "إقرار سفر اسرة";
            DocTypeVA[2] = "إقرار";
            DocTypeVA[3] = "تأشيرة سفر";
            DocTypeVA[4] = "إقرار كفالة عائلية";
            DocTypeVA[5] = "افادة لمن يهمه الامر";
            DocTypeVA[6] = "إقرار بمطابقة اسمين";
            DocTypeVA[7] = "شهادة لمن يهمه الامر";
            DocTypeVA[8] = "شهادة لمن يهمه الامر";
            DocTypeVA[9] = "إقرار";
            DocTypeVA[10] = "توكيل";

            travType[0] = "استخراج وثائق";
            travType[1] = "عدم ممانعة سفر";




        }

        private int daysOfMonth(int month, int year)
        {
            Months[0] = 31;

            if (year % 400 == 0 || (year % 4 == 0 && year % 100 != 0))
                Months[1] = 29;
            else Months[1] = 28;

            Months[2] = 31;

            Months[3] = 30;
            Months[4] = 31;
            Months[5] = 30;

            Months[6] = 31;
            Months[7] = 31;
            Months[8] = 30;

            Months[9] = 31;
            Months[10] = 30;
            Months[11] = 31;

            return Months[month];
        }

        private void ClearFileds()
        {

            TableIndex = -1;
            IDNo = -1;
            ReportType.SelectedIndex = 0;
            SearchPanel.Visible = false;
            ReportPanel.Visible = false;
            attendedVC.SelectedIndex = 2;
            if (UserJobposition.Contains("قنصل"))
            {
                picVersio.Visible = labversio.Visible = true;
                picadd.Visible = true;
                labelmonth.Visible = true;
                picremove.Visible = true;
                labeldate.Visible = true;
                picaddmonth.Visible = true;
                pictremovemonth.Visible = true;
                btnMessSave.Text = "حفظ وتأكيد";
                // btnDeleteHand.Visible = true;                
            }
            txtEmbassey.SelectedIndex = 26;
        }
        private void StaredColumns()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select * from TableAuthRights", sqlCon);
            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa1.Fill(dtbl);
            dataGridView7.DataSource = dtbl;

            sqlCon.Close();


            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var fileinfo = new FileInfo(sfd.FileName);
                        using (var package = new ExcelPackage(fileinfo))
                        {
                            ExcelWorksheet excelsheet = package.Workbook.Worksheets.Add("Rights");
                            excelsheet.Cells.LoadFromDataTable(dtbl);

                            //excelsheet.Cells.LoadFromCollection(dataGridView7.DataSource); 
                            package.Save();

                        }
                    }
                    catch (Exception ex)
                    {
                    }


                }
            }
        }


        private void StaredColumns1()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select * from TableAddContext", sqlCon);
            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa1.Fill(dtbl);
            dataGridView7.DataSource = dtbl;

            sqlCon.Close();


            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var fileinfo = new FileInfo(sfd.FileName);
                        using (var package = new ExcelPackage(fileinfo))
                        {
                            ExcelWorksheet excelsheet = package.Workbook.Worksheets.Add("Context");
                            excelsheet.Cells.LoadFromDataTable(dtbl);

                            //excelsheet.Cells.LoadFromCollection(dataGridView7.DataSource); 
                            package.Save();

                        }
                    }
                    catch (Exception ex)
                    {
                    }


                }
            }
        }

        private void StatisticInfo()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select نوع_التوكيل,إجراء_التوكيل,التاريخ_الميلادي, itext1,itext2,itext3,itext4,itext5,icheck1,itxtDate1,icombo1,icombo2,ibtnAdd1  from TableAuth", sqlCon);
            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa1.Fill(dtbl);
            dataGridView7.DataSource = dtbl;

            sqlCon.Close();


            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var fileinfo = new FileInfo(sfd.FileName);
                        using (var package = new ExcelPackage(fileinfo))
                        {
                            ExcelWorksheet excelsheet = package.Workbook.Worksheets.Add("Statistics");
                            excelsheet.Cells.LoadFromDataTable(dtbl);

                            //excelsheet.Cells.LoadFromCollection(dataGridView7.DataSource); 
                            package.Save();

                        }
                    }
                    catch (Exception ex)
                    {
                    }


                }
            }
        }
        void DailyList(string dateFrom)
        {
            totalrowsAuth = 0;

            totalrowsAffadivit = 0;
            int y = 0;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            DataTable dtbl = new DataTable();
            DataTable dtbl1 = new DataTable();
            if (sqlCon.State == ConnectionState.Closed)
            {
                //label1.Text = "";
                if (sqlCon.State == ConnectionState.Closed)
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return; }
                for (TableIndex = 1; TableIndex <= 6; TableIndex++)
                {
                    int x = 0;

                    if (TableIndex == 6 )
                    {
                        SqlDataAdapter sqlDa1 = new SqlDataAdapter(queryDateList[TableIndex], sqlCon);
                        sqlDa1.SelectCommand.CommandType = CommandType.Text;
                        //MessageBox.Show(dateFrom.Split('-')[1] +"-"+ dateFrom.Split('-')[0]+"-"+ dateFrom.Split('-')[2]);
                        sqlDa1.SelectCommand.Parameters.AddWithValue("@التاريخ_الميلادي", dateFrom);
                        sqlDa1.Fill(dtbl1);
                        dataGridView2.DataSource = dtbl1;

                        //for (totalrowsAuth = 0; totalrowsAuth < dtbl1.Rows.Count; totalrowsAuth++)
                        //{
                        //    if (!(dataGridView2.Rows[x].Cells[5].Value.ToString().Contains("-") && dataGridView2.Rows[x].Cells[6].Value.ToString() != "حضور مباشرة إلى القنصلية"))
                        //    { 
                        //        RetrievedAuthPers[totalrowsAuth] = dataGridView2.Rows[x].Cells[4].Value.ToString();
                        //    RetrievedNoAuth[totalrowsAuth] = dataGridView2.Rows[x].Cells[3].Value.ToString();
                        //    RetrievedNameAuth[totalrowsAuth] = dataGridView2.Rows[x].Cells[0].Value.ToString().Replace("_"," و");

                        //        x++;
                        //}
                        totalrowsAuth = dtbl1.Rows.Count;
                        //}

                    }
                    else
                    {
                        SqlDataAdapter sqlDa = new SqlDataAdapter(queryDateList[TableIndex], sqlCon);
                        sqlDa.SelectCommand.CommandType = CommandType.Text;
                        if(TableIndex == 4) 
                            sqlDa.SelectCommand.Parameters.AddWithValue("@التاريخ_الميلادي", dateFrom);
                        else
                            sqlDa.SelectCommand.Parameters.AddWithValue("@GriDate", dateFrom);

                        sqlDa.Fill(dtbl);
                        string[] arrangeData = new string[5];
                        foreach (DataRow row in dtbl.Rows)
                        {

                            if (!(row["ArchivedState"].ToString().Contains("-") && row["DataInterType"].ToString() != "حضور مباشرة إلى القنصلية"))
                            {

                                switch (TableIndex)
                                {
                                    case 1:
                                        RetrievedNameAffadivit[totalrowsAffadivit] = row["AppName"].ToString();
                                        switch (row["ProType"].ToString()) {
                                            case "0":
                                                RetrievedTypeAffadivit[totalrowsAffadivit] = "استخراج وثائق للابناء";
                                                break;
                                            case "1":
                                                RetrievedTypeAffadivit[totalrowsAffadivit] = "عدم ممانعة سفر الابناء";
                                                break;
                                            case "2":
                                                RetrievedTypeAffadivit[totalrowsAffadivit] = "استخراج وثائق وعدم ممانعة سفر الابناء";
                                                break;
                                            case "3":
                                                RetrievedTypeAffadivit[totalrowsAffadivit] = "استخراج وثائق وعدم ممانعة سفر الابناء والزوجة";
                                                break;
                                            case "4":
                                                RetrievedTypeAffadivit[totalrowsAffadivit] = "استخراج وثائق وعدم ممانعة سفر الابناء بصحبة مرافق غير الزوجة";
                                                break;
                                            case "5":
                                                RetrievedTypeAffadivit[totalrowsAffadivit] = "عدم ممانعة سفر الابناء والزوجة";
                                                break;
                                            case "6":
                                                RetrievedTypeAffadivit[totalrowsAffadivit] = "عدم ممانعة سفر الزوجة";
                                                break;
                                        }

                                        arrangeData = row["DocID"].ToString().Split('/');
                                        if (arrangeData.Length != 4)
                                            //MessageBox.Show(TableIndex.ToString());
                                            if (arrangeData.Length == 4)
                                                RetrievedNoAffadivit[totalrowsAffadivit] = arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];
                                            else if (arrangeData.Length == 5)
                                                RetrievedNoAffadivit[totalrowsAffadivit] = arrangeData[4] + "/" + arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];
                                        if (!string.IsNullOrEmpty(row["ProType"].ToString()))
                                        {
                                            iqrarType = Convert.ToInt16(row["ProType"].ToString());
                                            switch (iqrarType)
                                            {
                                                case 0:
                                                    RetrievedTypeAffadivit[totalrowsAffadivit] = travType[0];
                                                    break;
                                                case 1:
                                                    RetrievedTypeAffadivit[totalrowsAffadivit] = travType[1];
                                                    break;
                                                case 2:
                                                    RetrievedTypeAffadivit[totalrowsAffadivit] = travType[0] + " و" + travType[1];
                                                    arrangeData = row["DocID"].ToString().Split('/');
                                                    if (arrangeData.Length == 4)
                                                        RetrievedNoAffadivit[totalrowsAffadivit] = arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];
                                                    else if (arrangeData.Length == 5)
                                                        RetrievedNoAffadivit[totalrowsAffadivit] = arrangeData[4] + "/" + arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];
                                                    break;
                                            }
                                        }
                                        break;
                                    case 2:


                                        RetrievedNameAffadivit[totalrowsAffadivit] = row["AppName"].ToString();
                                        RetrievedTypeAffadivit[totalrowsAffadivit] = "إقرار مشفوع باليمين";

                                        arrangeData = row["DocID"].ToString().Split('/');
                                        if (arrangeData.Length == 4)
                                            RetrievedNoAffadivit[totalrowsAffadivit] = arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];
                                        else if (arrangeData.Length == 5)
                                            RetrievedNoAffadivit[totalrowsAffadivit] = arrangeData[4] + "/" + arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];

                                        break;
                                    case 3:
                                        RetrievedNameAffadivit[totalrowsAffadivit] = row["AppName"].ToString();
                                        RetrievedTypeAffadivit[totalrowsAffadivit] = row["IqrarType"].ToString();
                                        arrangeData = row["DocID"].ToString().Split('/');
                                        if (arrangeData.Length == 4)
                                            RetrievedNoAffadivit[totalrowsAffadivit] = arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];
                                        else if (arrangeData.Length == 5)
                                            RetrievedNoAffadivit[totalrowsAffadivit] = arrangeData[4] + "/" + arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];

                                        break;
                                    case 4:
                                        RetrievedNameAffadivit[totalrowsAffadivit] = row["مقدم_الطلب"].ToString();
                                        RetrievedTypeAffadivit[totalrowsAffadivit] = row["نوع_الإجراء"].ToString();
                                        arrangeData = row["رقم_المعاملة"].ToString().Split('/');
                                        if (arrangeData.Length == 4)
                                            RetrievedNoAffadivit[totalrowsAffadivit] = arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];
                                        else if (arrangeData.Length == 5)
                                            RetrievedNoAffadivit[totalrowsAffadivit] = arrangeData[4] + "/" + arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];

                                        break;
                                    case 5:
                                        RetrievedNameAffadivit[totalrowsAffadivit] = row["AppName"].ToString();
                                        if (!string.IsNullOrEmpty(row["ProCase"].ToString()))
                                        {
                                            iqrarType = Convert.ToInt32(row["ProCase"].ToString());
                                            switch (iqrarType)
                                            {
                                                case 1:
                                                    RetrievedTypeAffadivit[totalrowsAffadivit] = travType1[1];
                                                    break;
                                                case 2:
                                                    RetrievedTypeAffadivit[totalrowsAffadivit] = travType1[2];
                                                    break;
                                                case 3:
                                                    RetrievedTypeAffadivit[totalrowsAffadivit] = travType1[3];
                                                    break;
                                                case 4:
                                                    RetrievedTypeAffadivit[totalrowsAffadivit] = travType1[4];
                                                    break;
                                            }
                                        }
                                        arrangeData = row["DocID"].ToString().Split('/');
                                        if (arrangeData.Length == 4)
                                            RetrievedNoAffadivit[totalrowsAffadivit] = arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];
                                        else if (arrangeData.Length == 5)
                                            RetrievedNoAffadivit[totalrowsAffadivit] = arrangeData[4] + "/" + arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];
                                        break;
                                }
                                totalrowsAffadivit++;
                            }
                        }

                        dtbl.Clear();
                    }






                }
            }
            sqlCon.Close();
        }

        int DailyList(string dateFrom, string dateTo)
        {
            totalRowDuration = 0;
            int w = 0;
            string Currentmonth = "0", CurrentDay = "0", CurrentDate = "0";
            DateTime datetimeS = dateTimeFrom.Value;
            DateTime datetimeE = dateTimeTo.Value;
            if (datetimeS > datetimeE)
            {
                string datetime = dateFrom;
                dateFrom = dateTo;
                dateTo = datetime;

            }

            string[] YearMonthDayS = dateFrom.Split('-');
            int yearS, monthS, dateS;
            yearS = Convert.ToInt16(YearMonthDayS[0]);
            monthS = Convert.ToInt16(YearMonthDayS[1]);
            dateS = Convert.ToInt16(YearMonthDayS[2]);
            DateTime dateValue = new DateTime(yearS, monthS, dateS);


            int dayeofWeek = ((int)dateValue.DayOfWeek);

            if (dayeofWeek == 0) { startofNextWeek = dateS + 7; }
            else if (dayeofWeek == 1) { startofNextWeek = dateS + 6; }
            else if (dayeofWeek == 2) { startofNextWeek = dateS + 5; }
            else if (dayeofWeek == 3) { startofNextWeek = dateS + 4; }
            else if (dayeofWeek == 4) { startofNextWeek = dateS + 3; }
            else if (dayeofWeek == 5) { startofNextWeek = dateS + 2; }
            else if (dayeofWeek == 6) { startofNextWeek = dateS + 1; }


            string[] YearMonthDayE = dateTo.Split('-');
            int yearE, monthE, dateE;
            yearE = Convert.ToInt16(YearMonthDayE[0]);
            monthE = Convert.ToInt16(YearMonthDayE[1]);
            dateE = Convert.ToInt16(YearMonthDayE[2]);

            SqlConnection sqlCon = new SqlConnection(DataSource);
            DataTable dtbl = new DataTable();
            DataTable dtbl1 = new DataTable();
            if (sqlCon.State == ConnectionState.Closed)
            {
                if (sqlCon.State == ConnectionState.Closed)
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return 0; }

                for (int y = yearS; y <= yearE; y++)
                {
                    for (int m = monthS; m <= monthE && m <= 12; m++)
                    {
                        int d;
                        for (d = dateS; d <= dateE && d <= daysOfMonth(m - 1, y); d++)
                        {
                            if (m < 10) Currentmonth = "0" + m.ToString();
                            else Currentmonth = m.ToString();
                            if (d < 10) CurrentDay = "0" + d.ToString();
                            else CurrentDay = d.ToString();
                            CurrentDate = CurrentDay + "-" + Currentmonth + "-" + y.ToString();

                            for (TableIndex = 0; TableIndex < 6; TableIndex++)
                            {
                                if (TableIndex == 6)
                                {
                                    SqlDataAdapter sqlDa1 = new SqlDataAdapter(queryDateList[TableIndex], sqlCon);
                                    sqlDa1.SelectCommand.CommandType = CommandType.Text;
                                    sqlDa1.SelectCommand.Parameters.AddWithValue("@التاريخ_الميلادي", dateFrom);
                                    sqlDa1.Fill(dtbl1);
                                    dataGridView5.DataSource = dtbl1;
                                }
                                else
                                {
                                    SqlDataAdapter sqlDa = new SqlDataAdapter(queryDateList[TableIndex], sqlCon);
                                    sqlDa.SelectCommand.CommandType = CommandType.Text;
                                    sqlDa.SelectCommand.Parameters.AddWithValue("@GriDate", dateFrom);
                                    sqlDa.Fill(dtbl);
                                    dataGridView2.DataSource = dtbl;
                                }
                                duartionReport[w, TableIndex] = dtbl.Rows.Count;
                                duartionReportAuth[w, TableIndex] = dtbl1.Rows.Count;
                                totalRowDuration = 1;
                            }
                        }
                        if (d == startofNextWeek)
                        {
                            w = 0;
                            startofNextWeek += 7;
                        }
                    }
                }
            }
            sqlCon.Close();

            return totalRowDuration;
        }

        private int AllProType()
        {
            for (int i = 0; i < 100; i++) ProType[i] = ""; ;

            DataTable dtbl = new DataTable();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return 0; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("select نوع_التوكيل from TableAuth", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.Fill(dtbl);
            int z = 0;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                bool found = false;

                for (int a = 0; a < comboBox3.Items.Count; a++)
                {

                    if (dataRow["نوع_التوكيل"].ToString() == comboBox3.Items[a].ToString())
                        found = true;

                }
                if (!found)
                {
                    Console.WriteLine(z.ToString() + ". نوع_التوكيل " + dataRow["نوع_التوكيل"].ToString());
                    if (dataRow["نوع_التوكيل"].ToString() != "")
                        comboBox3.Items.Add(dataRow["نوع_التوكيل"].ToString());

                }

            }
            string ReportName = "Report" + DateTime.Now.ToString("mmss") + ".docx";
            AuthTypes(comboBox3.Items.Count, ReportName, ProType);
            return comboBox3.Items.Count;
        }
        bool DeepStatics(string dateFrom, string dateTo, int month)
        {
            int proTypeCount = AllProType();

            bool foundData = false;
            string CurrentDay = "", Currentmonth = "", CurrentDate = "";
            string[] YearMonthDayS = dateFrom.Split('-');
            int yearS, monthS, dateS;
            yearS = Convert.ToInt16(YearMonthDayS[0]);
            monthS = Convert.ToInt16(YearMonthDayS[1]);
            dateS = Convert.ToInt16(YearMonthDayS[2]);

            string[] YearMonthDayE = dateTo.Split('-');
            int dateE = Convert.ToInt16(YearMonthDayE[2]);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            DataTable dtbl = new DataTable();
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return false; }
            for (int T = 0; T < 100; T++)
                    for (int yy = 0; yy < 31; yy++)
                        DeepReport[T, yy] = 0;
                int d;
                int m = monthS;
                int y = yearS;

                //Console.WriteLine("*********************" + month.ToString() + "*********************");
                for (d = dateS; d <= dateE && d <= daysOfMonth(m - 1, y); d++)
                {
                    int type = 0;
                    if (m < 10) Currentmonth = "0" + m.ToString();
                    else Currentmonth = m.ToString();
                    if (d < 10) CurrentDay = "0" + d.ToString();
                    else CurrentDay = d.ToString();
                    CurrentDate = CurrentDay + "-" + Currentmonth + "-" + y.ToString();
                    SqlDataAdapter sqlDa1 = new SqlDataAdapter(queryDateList[6], sqlCon);
                    sqlDa1.SelectCommand.CommandType = CommandType.Text;
                    sqlDa1.SelectCommand.Parameters.AddWithValue("@التاريخ_الميلادي", CurrentDate);
                    sqlDa1.Fill(dtbl);
                    dataGridView1.DataSource = dtbl;
                    foreach (DataRow dataRow in dtbl.Rows)
                    {
                        for (int x = 0; x < proTypeCount; x++)
                        {
                            BindingSource bs = new BindingSource();
                            bs.DataSource = dataGridView1.DataSource;
                            bs.Filter = dataGridView1.Columns[7].HeaderText.ToString() + " LIKE '%" + comboBox3.Items[x] + "%'";
                            dataGridView1.DataSource = bs;
                            //DeepReport[type, d] = dtbl.Rows.Count;
                            //MessageBox.Show(ProType[x]  +" -- " + dataGridView1.RowCount.ToString());
                        }
                        type++;
                    }
                    dtbl.Rows.Clear();
                

                //rep1[month, 0] = monthS;


                //int[] tempX = new int[12];
                //for (int x = 1; x < 12; x++)
                //{
                //    int tempdatat = 0;
                //    for (int yy = 0; yy <= 31; yy++)
                //    {
                //        tempdatat = tempdatat + report1[x, yy];
                //        report1[x, yy] = 0;
                //        if (tempdatat != 0)
                //            foundData = true;
                //    }
                //    rep1[month, x] = tempdatat;
                //}
            }


            if (foundData)
                totalRowDuration = 1;
            else totalRowDuration = 0;

            sqlCon.Close();
            return foundData;
        }

        bool DailyListcustm(string dateFrom, string dateTo, int month)
        {
            totalRowDuration = 0;
            bool foundData = false;
            int w = 0;
            string Currentmonth = "0", CurrentDay = "0", CurrentDate = "0";
            int[,] fileTable = new int[15, 16];

            string[] YearMonthDayS = dateFrom.Split('-');
            int yearS, monthS, dateS;
            yearS = Convert.ToInt16(YearMonthDayS[0]);
            monthS = Convert.ToInt16(YearMonthDayS[1]);
            dateS = Convert.ToInt16(YearMonthDayS[2]);
            DateTime dateValue = new DateTime(yearS, monthS, dateS);

            //MessageBox.Show("dateFrom=" + dateFrom + " dateTo=" + dateTo.ToString() );
            int dayeofWeek = ((int)dateValue.DayOfWeek);

            if (dayeofWeek == 0) { startofNextWeek = dateS + 7; }
            else if (dayeofWeek == 1) { startofNextWeek = dateS + 6; }
            else if (dayeofWeek == 2) { startofNextWeek = dateS + 5; }
            else if (dayeofWeek == 3) { startofNextWeek = dateS + 4; }
            else if (dayeofWeek == 4) { startofNextWeek = dateS + 3; }
            else if (dayeofWeek == 5) { startofNextWeek = dateS + 2; }
            else if (dayeofWeek == 6) { startofNextWeek = dateS + 1; }
            //MessageBox.Show(dayeofWeek.ToString());
            //MessageBox.Show(dateTo);
            string[] YearMonthDayE = dateTo.Split('-');
            int yearE, monthE, dateE;

            yearE = Convert.ToInt16(YearMonthDayE[0]);
            monthE = Convert.ToInt16(YearMonthDayE[1]);
            dateE = Convert.ToInt16(YearMonthDayE[2]);
            //MessageBox.Show("Y=" + yearE.ToString() + "M=" + monthE.ToString() + "D=" + dateE.ToString());
            SqlConnection sqlCon = new SqlConnection(DataSource);
            DataTable dtbl = new DataTable();
            DataTable dtbl1 = new DataTable();
            DataTable dtbl2 = new DataTable();
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return false; }
            //for (int y = yearS; y <= yearE; y++)
            //{                    
            //    for (int m = monthS; m <= monthE && m <= 12; m++)
            //    {
            for (int x = 0; x < 15; x++)
            {
                for (int yy = 0; yy < 31; yy++)
                { report1[x, yy] = 0; }
            }
                int d;
                int m = monthS;
                int y = yearS;
                //Console.WriteLine("*********************" + month.ToString() + "*********************");
                for (d = dateS; d <= dateE && d <= daysOfMonth(m - 1, y); d++)
                {
                    if (m < 10) Currentmonth = "0" + m.ToString();
                    else Currentmonth = m.ToString();
                    if (d < 10) CurrentDay = "0" + d.ToString();
                    else CurrentDay = d.ToString();
                    CurrentDate = Currentmonth+ "-" + CurrentDay + "-" + y.ToString();
                    if (Server == "57")
                    {
                        for (TableIndex = 1; TableIndex < 15; TableIndex++)
                        {
                            SqlDataAdapter sqlDa1 = new SqlDataAdapter(queryDateList[TableIndex], sqlCon);
                            //Console.WriteLine(TableIndex);
                            if (TableIndex == 4 ||TableIndex == 6 || TableIndex > 11)
                            {
                                sqlDa1.SelectCommand.CommandType = CommandType.Text;
                                sqlDa1.SelectCommand.Parameters.AddWithValue("@التاريخ_الميلادي", CurrentDate);
                                sqlDa1.Fill(dtbl1);
                                report1[TableIndex, d] = dtbl1.Rows.Count;
                                //Console.WriteLine(CurrentDate.ToString() + "-" + queryDateList[TableIndex].Split(' ')[3] + "-" + report1[TableIndex, d].ToString());
                            }
                            else if (TableIndex == 11)
                            {
                                sqlDa1.SelectCommand.CommandType = CommandType.Text;
                                sqlDa1.SelectCommand.Parameters.AddWithValue("@GriDate", CurrentDate);
                                sqlDa1.Fill(dtbl2);
                                int AuthCount = 0;
                                foreach (DataRow row in dtbl2.Rows)
                                {

                                    if (row["DocNo"].ToString().All(char.IsDigit))
                                    {
                                        AuthCount = AuthCount + Convert.ToInt32(row["DocNo"].ToString());
                                    }
                                }
                                report1[TableIndex, d] = AuthCount;
                            //Console.WriteLine(CurrentDate.ToString() + "-" + queryDateList[TableIndex].Split(' ')[3] + "-" + report1[TableIndex, d].ToString());
                            //Console.WriteLine(d.ToString() + "-" + dtbl2.Rows.Count.ToString());
                        }
                            else
                            {
                                sqlDa1.SelectCommand.CommandType = CommandType.Text;
                                sqlDa1.SelectCommand.Parameters.AddWithValue("@GriDate", CurrentDate);
                                sqlDa1.Fill(dtbl);
                                report1[TableIndex, d] = dtbl.Rows.Count;
                            //Console.WriteLine(CurrentDate.ToString() + "-" + queryDateList[TableIndex].Split(' ')[3] + "-" + report1[TableIndex, d].ToString());

                        }
                            dtbl.Rows.Clear();
                            dtbl1.Rows.Clear();
                            dtbl2.Rows.Clear();
                        }
                    }
                    else if (Server == "56")

                    {
                        for (TableIndex = 1; TableIndex < 7; TableIndex++)
                        {
                            string query = "select مقدم_الطلب  from "+ getFileTable(TableIndex-1) + " where التاريخ_الميلادي=@التاريخ_الميلادي";

                            SqlDataAdapter sqlDa1 = new SqlDataAdapter(query, sqlCon);
                            sqlDa1.SelectCommand.CommandType = CommandType.Text;
                            sqlDa1.SelectCommand.Parameters.AddWithValue("@التاريخ_الميلادي", CurrentDate);
                            sqlDa1.Fill(dtbl1);
                            report1[TableIndex, d] = dtbl1.Rows.Count;
                            dtbl1.Rows.Clear();
                        }
                    }   
                

                

            }
rep1[month, 0] = monthS;
                Console.WriteLine("--------------------------------"+ month.ToString()+"---------------------------------");
                int[] tempX = new int[15];
                for (int x = 1; x < 15; x++)
                {
                    int tempdatat = 0;
                    for (int yy = 0; yy <= 31; yy++)
                    {
                        tempdatat = tempdatat + report1[x, yy];
                        Console.WriteLine(queryDateList[x].ToString().Split(' ')[3] + "-" + yy.ToString() + "-" + report1[x, yy].ToString());
                        report1[x, yy] = 0;
                        if (tempdatat != 0)
                            foundData = true;
                    }
                    rep1[month, x] = tempdatat;
                }

            if (foundData)
                totalRowDuration = 1;
            else totalRowDuration = 0;

            sqlCon.Close();
            return foundData;
        }
        

        private void correctData() {
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select ID,GriDate from TableVisaApp", sqlCon);
            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl1 = new DataTable();
            sqlDa1.Fill(dtbl1);
            foreach (DataRow dataRow in dtbl1.Rows)
            {

                if (dataRow["GriDate"].ToString().Contains("-"))
                {
                    string[] str = dataRow["GriDate"].ToString().Split('-');

                    UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), str[2] + "-" + str[1] + "-" + str[0], "TableVisaApp");
                }

            }
        }
        private void CreateNotArchivedFiles(int rows, string reportName, string[] GriDate, string[] DocID, string[] AppName, string proType1, string proType2)
        {
            route = FilespathIn + @"\NonArchivedFiles.docx";
            string ActiveCopy = FilespathOut + reportName;
            if (!File.Exists(ActiveCopy))
            {
                System.IO.File.Copy(route, ActiveCopy);
                using (var document = DocX.Load(ActiveCopy))
                {
                    System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                    InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);

                    string strHeader = "الرقم : " + ReportNo.Text + "     " + "التاريخ :" + GregorianDate + " م" + "     " + "الموافق : " + HijriDate + "هـ" + Environment.NewLine;
                    document.InsertParagraph(strHeader)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(16d)
                    .Alignment = Alignment.center;
                    string MessageDir = "مرفق أدناه قائمة الملفات " + proType1 + Environment.NewLine +
                        Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ"
                        + Environment.NewLine;
                    document.InsertParagraph(MessageDir)
                        .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                        .FontSize(18d)
                        .Direction = Direction.RightToLeft;

                    var t = document.AddTable(rows + 1, 4);
                    t.Design = TableDesign.TableGrid;
                    t.Alignment = Alignment.center;

                    t.SetColumnWidth(3, 40);
                    t.SetColumnWidth(2, 180);
                    t.SetColumnWidth(1, 100);
                    t.SetColumnWidth(0, 80);


                    t.Rows[0].Cells[0].Paragraphs[0].Append("رقم المكاتبة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[1].Paragraphs[0].Append(proType2).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[2].Paragraphs[0].Append("اسم مقدم الطلب").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[3].Paragraphs[0].Append("الرقم").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;

                    for (int x = 1; x <= rows; x++)
                    {
                        t.Rows[x].Cells[0].Paragraphs[0].Append(GriDate[x - 1]).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                         if (Pers_Peope)
                            t.Rows[x].Cells[1].Paragraphs[0].Append(DocID[x - 1]).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                         else
                            t.Rows[x].Cells[1].Paragraphs[0].Append(DocID[x - 1] ).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                        t.Rows[x].Cells[2].Paragraphs[0].Append(AppName[x - 1]).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                        t.Rows[x].Cells[3].Paragraphs[0].Append(x.ToString() + ".").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                    }



                    var p = document.InsertParagraph(Environment.NewLine);
                    p.InsertTableAfterSelf(t);

                    string strAttvCo = Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ" + Environment.NewLine + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + attendedVC.Text + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + "ع/ القنصل العام بالإنابة";
                    var AttvCo = document.InsertParagraph(strAttvCo)
                        .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                        .FontSize(20d)
                        .Bold()
                        .Alignment = Alignment.center;


                    document.Save();
                    Process.Start("WINWORD.EXE", ActiveCopy);

                }
            }
            else
            {
                MessageBox.Show("يرجى حذف الملف الموجودأولاً");
                PrintReport.Enabled = true;

            }

        }


        private void CreateMandounbFiles(int rows, string reportName, string[] GriDate, string[] DocID, string[] AppName, string[] mandoubName)
        {
            route = FilespathIn + @"\NonArchivedFiles.docx";
            string ActiveCopy = FilespathOut + reportName;
            if (!File.Exists(ActiveCopy))
            {
                System.IO.File.Copy(route, ActiveCopy);
                using (var document = DocX.Load(ActiveCopy))
                {
                    System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                    InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);

                    string strHeader = "الرقم : " + ReportNo.Text + "     " + "التاريخ :" + GregorianDate + " م" + "     " + "الموافق : " + HijriDate + "هـ" + Environment.NewLine;
                    document.InsertParagraph(strHeader)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(16d)
                    .Alignment = Alignment.center;
                    string MessageDir = "مرفق أدناه قائمة بالمعاملات غير المكتملة لمناديب القنصلية العامة " + Environment.NewLine +
                        Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ"
                        + Environment.NewLine;
                    document.InsertParagraph(MessageDir)
                        .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                        .FontSize(18d)
                        .Direction = Direction.RightToLeft;

                    var t = document.AddTable(rows + 1, 5);
                    t.Design = TableDesign.TableGrid;
                    t.Alignment = Alignment.center;


                    t.SetColumnWidth(4, 40);
                    t.SetColumnWidth(3, 150);
                    t.SetColumnWidth(2, 130);
                    t.SetColumnWidth(1, 120);
                    t.SetColumnWidth(0, 80);

                    t.Rows[0].Cells[0].Paragraphs[0].Append("تاريخ المكاتبة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[1].Paragraphs[0].Append("رقم المكاتبة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[2].Paragraphs[0].Append("اسم المندوب").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[3].Paragraphs[0].Append("اسم مقدم الطلب").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[4].Paragraphs[0].Append("الرقم").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;

                    for (int x = 1; x <= rows; x++)
                    {

                        t.Rows[x].Cells[0].Paragraphs[0].Append(GriDate[x - 1]).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                        t.Rows[x].Cells[1].Paragraphs[0].Append(DocID[x - 1]).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                        t.Rows[x].Cells[2].Paragraphs[0].Append(mandoubName[x - 1]).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                        t.Rows[x].Cells[3].Paragraphs[0].Append(AppName[x - 1]).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                        t.Rows[x].Cells[4].Paragraphs[0].Append(x.ToString() + ".").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                    }



                    var p = document.InsertParagraph(Environment.NewLine);
                    p.InsertTableAfterSelf(t);

                    string strAttvCo = Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ" + Environment.NewLine + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + attendedVC.Text + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + "ع/ القنصل العام بالإنابة";
                    var AttvCo = document.InsertParagraph(strAttvCo)
                        .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                        .FontSize(20d)
                        .Bold()
                        .Alignment = Alignment.center;


                    document.Save();
                    Process.Start("WINWORD.EXE", ActiveCopy);

                }
            }
            else
            {
                MessageBox.Show("يرجى حذف الملف الموجودأولاً");
                PrintReport.Enabled = true;

            }

        }

        private void AuthTypes(int rows, string reportName, string[] AuthType)
        {
            route = FilespathIn + @"\نوع_التواكيل.docx";
            string ActiveCopy = FilespathOut + reportName;
            System.IO.File.Copy(route, ActiveCopy);
            using (var document = DocX.Load(ActiveCopy))
            {
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);

                string strHeader = "الرقم : " + ReportNo.Text + "     " + "التاريخ :" + GregorianDate + " م" + "     " + "الموافق : " + HijriDate + "هـ" + Environment.NewLine;
                document.InsertParagraph(strHeader)
                .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                .FontSize(16d)
                .Alignment = Alignment.center;
                string MessageDir = "الى : خارجية - الخرطوم" + Environment.NewLine + "من سوداني - جدة" + Environment.NewLine + "لعناية السيد/ مدير إدارة التوثيق"
                    + Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ";
                document.InsertParagraph(MessageDir)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(18d)
                    .Direction = Direction.RightToLeft;

                var t = document.AddTable(1, 2);
                t.Design = TableDesign.TableGrid;
                t.Alignment = Alignment.center;
                t.SetColumnWidth(1, 40);
                t.SetColumnWidth(0, 140);

                t.Rows[0].Cells[0].Paragraphs[0].Append("نوع التوكيل").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[1].Paragraphs[0].Append("الرقم").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                int x = 1;

                for (int count = 1; count <= rows; count++)

                {
                    t.InsertRow();
                    t.Rows[x].Cells[0].Paragraphs[0].Append(comboBox3.Items[count - 1].ToString()).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                    t.Rows[x].Cells[1].Paragraphs[0].Append(x.ToString() + ".").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                    x++;
                }



                var p = document.InsertParagraph(Environment.NewLine);
                p.InsertTableAfterSelf(t);

                string strAttvCo = Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ" + Environment.NewLine + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + attendedVC.Text + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + "ع/ القنصل العام بالإنابة";
                var AttvCo = document.InsertParagraph(strAttvCo)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(20d)
                    .Bold()
                    .Alignment = Alignment.center;


                document.Save();
                Process.Start("WINWORD.EXE", ActiveCopy);

            }



        }

        private void CreateDailyReportIqrar(int rows, string reportName, string DocumentType, bool AffadaivtAuth)
        {
            route = FilespathIn + @"\DailyReport.docx";
            string ActiveCopy = FilespathOut + reportName;
            System.IO.File.Copy(route, ActiveCopy);
            using (var document = DocX.Load(ActiveCopy))
            {
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);

                string strHeader = "الرقم : " + ReportNo.Text + "     " + "التاريخ :" + GregorianDate + " م" + "     " + "الموافق : " + HijriDate + "هـ" + Environment.NewLine;
                document.InsertParagraph(strHeader)
                .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                .FontSize(16d)
                .Alignment = Alignment.center;
                string MessageDir = "الى : خارجية - الخرطوم" + Environment.NewLine + "من سوداني - جدة" + Environment.NewLine + "لعناية السيد/ مدير إدارة التوثيق"
                    + Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ"
                    + Environment.NewLine + "بالإشارة إلى برقيتكم بالرقم: و خ/توثيق/97 بتاريخ 23/04/2014 م بشأن إصدار راجعة " + DocumentType + "، نفيدكم باعتماد القنصلية العامة للمعاملات الصادرة طرفها للمذكورين بالجدول أدناه" + " بتاريخ " + dateTimeFrom.Text;
                document.InsertParagraph(MessageDir)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(18d)
                    .Direction = Direction.RightToLeft;

                var t = document.AddTable(1, 4);
                t.Design = TableDesign.TableGrid;
                t.Alignment = Alignment.center;
                t.SetColumnWidth(3, 40);
                t.SetColumnWidth(2, 150);
                t.SetColumnWidth(1, 170);
                t.SetColumnWidth(0, 140);

                t.Rows[0].Cells[0].Paragraphs[0].Append("الرقم المرجعي للمعاملة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[1].Paragraphs[0].Append("نوع المعاملة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[2].Paragraphs[0].Append("اسم مقدم الطلب").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[3].Paragraphs[0].Append("الرقم").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                int x = 1;

                for (int count = 1; count <= rows; count++)

                {
                    if (AffadaivtAuth)
                    {
                        if (RetrievedNameAffadivit[count - 1] != "")

                        {
                            t.InsertRow();
                            t.Rows[x].Cells[0].Paragraphs[0].Append(RetrievedNoAffadivit[count - 1]).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                            t.Rows[x].Cells[1].Paragraphs[0].Append(RetrievedTypeAffadivit[count - 1]).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                            t.Rows[x].Cells[2].Paragraphs[0].Append(RetrievedNameAffadivit[count - 1]).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                            t.Rows[x].Cells[3].Paragraphs[0].Append(x.ToString() + ".").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                            x++;
                        }
                    }
                }



                var p = document.InsertParagraph(Environment.NewLine);
                p.InsertTableAfterSelf(t);

                string strAttvCo = Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ" + Environment.NewLine + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + attendedVC.Text + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + "ع/ القنصل العام بالإنابة";
                var AttvCo = document.InsertParagraph(strAttvCo)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(20d)
                    .Bold()
                    .Alignment = Alignment.center;


                document.Save();
                Process.Start("WINWORD.EXE", ActiveCopy);

            }



        }


        private void CreateDailyReportAuth(int rows, string reportName, string DocumentType, bool AffadaivtAuth)
        {
            route = FilespathIn + @"\DailyReport.docx";
            string ActiveCopy = FilespathOut + reportName;
            System.IO.File.Copy(route, ActiveCopy);
            using (var document = DocX.Load(ActiveCopy))
            {
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);

                string strHeader = "الرقم : " + ReportNo.Text + "     " + "التاريخ :" + GregorianDate + " م" + "     " + "الموافق : " + HijriDate + "هـ" + Environment.NewLine;
                document.InsertParagraph(strHeader)
                .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                .FontSize(16d)
                .Alignment = Alignment.center;
                string MessageDir = "الى : خارجية - الخرطوم" + Environment.NewLine + "من سوداني - جدة" + Environment.NewLine + "لعناية السيد/ مدير إدارة التوثيق"
                    + Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ"
                    + Environment.NewLine + "بالإشارة إلى برقيتكم بالرقم: و خ/توثيق/97 بتاريخ 23/04/2014 م بشأن إصدار راجعة " + DocumentType + "، نفيدكم باعتماد القنصلية العامة للمعاملات الصادرة طرفها للمذكورين بالجدول أدناه" + " بتاريخ " + dateTimeFrom.Text;
                document.InsertParagraph(MessageDir)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(18d)
                    .Direction = Direction.RightToLeft;

                var t = document.AddTable(1, 4);
                t.Design = TableDesign.TableGrid;
                t.Alignment = Alignment.center;
                t.SetColumnWidth(3, 40);
                t.SetColumnWidth(2, 150);
                t.SetColumnWidth(1, 170);
                t.SetColumnWidth(0, 140);

                t.Rows[0].Cells[0].Paragraphs[0].Append("الرقم المرجعي للمعاملة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[1].Paragraphs[0].Append("اسم الوكيل").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[2].Paragraphs[0].Append("اسم مقدم الطلب").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[3].Paragraphs[0].Append("الرقم").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;

                int CurrentRows = 1;

                for (int x = 1; x <= rows; x++)
                {
                    if (dataGridView2.Rows[x - 1].Cells[0].Value.ToString() != "" && dataGridView2.Rows[x - 1].Cells[5].Value.ToString() == "مؤرشف نهائي")
                    {
                        if (!AffadaivtAuth)
                        {
                            t.InsertRow();
                            string str = "";
                            string[] arrangeData = dataGridView2.Rows[x - 1].Cells[3].Value.ToString().Split('/');
                            if (arrangeData.Length == 4)
                                str = arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];
                            else if (arrangeData.Length == 5)
                                str = arrangeData[4] + "/" + arrangeData[3] + "/" + arrangeData[2] + "/" + arrangeData[1] + "/" + arrangeData[0];

                            t.Rows[CurrentRows].Cells[1].Paragraphs[0].Append(dataGridView2.Rows[x - 1].Cells[4].Value.ToString().Replace("_", " و")).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                            t.Rows[CurrentRows].Cells[0].Paragraphs[0].Append(str).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                            t.Rows[CurrentRows].Cells[2].Paragraphs[0].Append(dataGridView2.Rows[x - 1].Cells[0].Value.ToString().Replace("_", " و")).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                            t.Rows[CurrentRows].Cells[3].Paragraphs[0].Append(CurrentRows.ToString() + ".").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                            CurrentRows++;
                        }

                    }
                }



                var p = document.InsertParagraph(Environment.NewLine);
                p.InsertTableAfterSelf(t);

                string strAttvCo = Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ" + Environment.NewLine + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + attendedVC.Text + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + "ع/ القنصل العام بالإنابة";
                var AttvCo = document.InsertParagraph(strAttvCo)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(20d)
                    .Bold()
                    .Alignment = Alignment.center;


                document.Save();
                Process.Start("WINWORD.EXE", ActiveCopy);

            }



        }

        private void fillDate(string queryInfo, string items)
        {
            
            string column = "@" + items;
            dataRowTable = new DataTable();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(queryInfo, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue(column, "");
            sqlDa.Fill(dataRowTable);
            sqlCon.Close();
            
            
            foreach (DataRow dataRow in dataRowTable.Rows)
            {
                string name2 = dataRow[items].ToString();
                bool found2 = false;
                for (int a = 0; a < yearReport.Items.Count; a++)
                {
                    if (name2.Split('-').Length != 3 || name2.Split('-')[2] == yearReport.Items[a].ToString())
                    {
                        found2 = true; 
                        break;
                    }
                    //else found2 = false;

                }
                if (!found2)
                {
                    if (dataRow[items].ToString().Split('-').Length == 3)
                    {
                        if (dataRow[items].ToString().Split('-')[2].Contains("20"))
                            if (name2.Split('-')[2] != "")
                            {
                                Console.WriteLine(name2.Split('-')[2]);
                                yearReport.Items.Add(name2.Split('-')[2]);
                            }
                    }
                }

            }
            
        }


        private void CreateDurationReport(int[,] Report, string reportName)
        {
            int lengthRow = 3;
            string title = "تقرير المعاملات لكل من شهر " + Monthorder(rep1[0, 0]) + " و" + Monthorder(rep1[1, 0]) + " و" + Monthorder(rep1[2, 0]) + " للعام " + yearReport.Text + "م" ;
            if (ReportType.SelectedIndex == 8)
            {
                lengthRow = 16;
                title = "تقرير المعاملات للعام " + yearReport.Text + "م";
            }
            route = FilespathIn + @"\DailyReportCopy.docx";
            string ActiveCopy = FilespathOut + reportName;
            System.IO.File.Copy(route, ActiveCopy);
            using (DocX document = DocX.Load(ActiveCopy))
            {
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
                string strHeader = "الرقم: " + ReportNo.Text + "     " + "التاريخ:" + GregorianDate + " م" + "     " + "الموافق: " + HijriDate + "هـ" +
                    Environment.NewLine + title;
                document.InsertParagraph(strHeader)
                .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                .FontSize(18d)
                .Alignment = Alignment.center;
                int col = 8;
                if (Server == "57") col = 16;
                    var t = document.AddTable(2 + lengthRow, col);

                if (Server == "57") {
                    
                    t.Design = TableDesign.TableGrid;
                    t.Alignment = Alignment.center;
                    t.SetColumnWidth(15, 50);
                    t.SetColumnWidth(14, 50);
                    t.SetColumnWidth(13, 50);
                    t.SetColumnWidth(12, 50);
                    t.SetColumnWidth(11, 50);
                    t.SetColumnWidth(10, 50);
                    t.SetColumnWidth(9, 50);
                    t.SetColumnWidth(8, 50);
                    t.SetColumnWidth(7, 50);
                    t.SetColumnWidth(6, 50);
                    t.SetColumnWidth(5, 50);
                    t.SetColumnWidth(4, 50);
                    t.SetColumnWidth(3, 50);
                    t.SetColumnWidth(2, 50);
                    t.SetColumnWidth(1, 50);
                    t.SetColumnWidth(0, 50);

                    reportItems[0] = "اقرار استخراج أوراق ثبوتية موافقة بالسفر";//0+1
                    reportItems[1] = "إقرار لاغراض مختلفة";//2+4
                    reportItems[2] = "إقرار لاثبات صحة إسمين";
                    reportItems[3] = "إقرارات عامة";//2+4
                    reportItems[4] = "إقرار كفالة أفراد أسرة";
                    reportItems[5] = "توكيل";
                    reportItems[6] = "إفادة للادلة الجنائية";
                    reportItems[7] = "إفادة عدم ممانعة زواج";
                    reportItems[8] = "إفادة تسجيل ببرنامج دراسي";
                    reportItems[9] = "مذكرة لمنح تأشيرة";
                    reportItems[10] = "التوثيق";
                    reportItems[11] = "إذن دفن";
                    reportItems[12] = "وثيقة زواج";
                    reportItems[13] = "وثيقة طلاق";
                    reportItems[14] = "مجموع المعاملات";

                    t.Rows[0].Cells[0].Paragraphs[0].Append(reportItems[14]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[1].Paragraphs[0].Append(reportItems[13]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[2].Paragraphs[0].Append(reportItems[12]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[3].Paragraphs[0].Append(reportItems[11]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[4].Paragraphs[0].Append(reportItems[10]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[5].Paragraphs[0].Append(reportItems[9]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[6].Paragraphs[0].Append(reportItems[8]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[7].Paragraphs[0].Append(reportItems[7]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[8].Paragraphs[0].Append(reportItems[6]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[9].Paragraphs[0].Append(reportItems[5]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[10].Paragraphs[0].Append(reportItems[4]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[11].Paragraphs[0].Append(reportItems[3]).FontSize(12d).Bold().Alignment = Alignment.center; //7 8
                    t.Rows[0].Cells[12].Paragraphs[0].Append(reportItems[2]).FontSize(12d).Bold().Alignment = Alignment.center;//5
                    t.Rows[0].Cells[13].Paragraphs[0].Append(reportItems[1]).FontSize(12d).Bold().Alignment = Alignment.center; //3
                    t.Rows[0].Cells[14].Paragraphs[0].Append(reportItems[0]).FontSize(12d).Bold().Alignment = Alignment.center;//0 1 2 4 6 9
                    t.Rows[0].Cells[15].Paragraphs[0].Append("الشهر").FontSize(12d).Bold().Alignment = Alignment.center;

                    int AllSum = 0;
                    for (int c = 1; c < 15; c++)
                    {
                        AllSum = 0;
                        for (int r = 0; r < lengthRow; r++)
                        {
                            AllSum = AllSum + rep1[r, c];
                        }

                        monthSumH[c] = AllSum;
                    }


                    for (int r = 0; r < lengthRow; r++)
                    {
                        AllSum = 0;
                        for (int c = 1; c < 15; c++)
                        {
                            AllSum = AllSum + rep1[r, c];
                        }

                        monthSumV[r] = AllSum;

                    }

                    AllSum = 0;
                    for (int c = 1; c < 15; c++)
                    {
                        AllSum = AllSum + monthSumH[c];
                    }


                    int x = 0;
                    //for (int x = 0; x < 7; x++)
                    for (int w = 0; w < lengthRow; w++)
                    {
                        //Console.WriteLine(w.ToString() + "---- monthSumV " + monthSumV[w].ToString());
                        t.Rows[w + 1].Cells[0].Paragraphs[0].Append(monthSumV[w].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[1].Paragraphs[0].Append(rep1[w, 14].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[2].Paragraphs[0].Append(rep1[w, 13].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[3].Paragraphs[0].Append(rep1[w, 12].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[4].Paragraphs[0].Append(rep1[w, 11].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[5].Paragraphs[0].Append(rep1[w, 10].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[6].Paragraphs[0].Append(rep1[w, 9].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[7].Paragraphs[0].Append(rep1[w, 8].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[8].Paragraphs[0].Append(rep1[w, 7].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[9].Paragraphs[0].Append(rep1[w, 6].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[10].Paragraphs[0].Append(rep1[w, 5].ToString()).FontSize(12d).Bold().Alignment = Alignment.center; //7 8
                        t.Rows[w + 1].Cells[11].Paragraphs[0].Append(rep1[w, 4].ToString()).FontSize(12d).Bold().Alignment = Alignment.center; //7 8
                        t.Rows[w + 1].Cells[12].Paragraphs[0].Append(rep1[w, 3].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;//5
                        t.Rows[w + 1].Cells[13].Paragraphs[0].Append(rep1[w, 2].ToString()).FontSize(12d).Bold().Alignment = Alignment.center; //3
                        t.Rows[w + 1].Cells[14].Paragraphs[0].Append(rep1[w, 1].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;//0 1 2 4 6 9
                        t.Rows[w + 1].Cells[15].Paragraphs[0].Append(Monthorder(rep1[w, 0]).ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    }
                    //int w = 4;
                    t.Rows[lengthRow + 1].Cells[0].Paragraphs[0].Append(AllSum.ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[1].Paragraphs[0].Append(monthSumH[14].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[2].Paragraphs[0].Append(monthSumH[13].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[3].Paragraphs[0].Append(monthSumH[12].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[4].Paragraphs[0].Append(monthSumH[11].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[5].Paragraphs[0].Append(monthSumH[10].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[6].Paragraphs[0].Append(monthSumH[9].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[7].Paragraphs[0].Append(monthSumH[8].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[8].Paragraphs[0].Append(monthSumH[7].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[9].Paragraphs[0].Append(monthSumH[6].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[10].Paragraphs[0].Append(monthSumH[5].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[11].Paragraphs[0].Append(monthSumH[4].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[12].Paragraphs[0].Append(monthSumH[3].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[13].Paragraphs[0].Append(monthSumH[2].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[14].Paragraphs[0].Append(monthSumH[1].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[15].Paragraphs[0].Append("إجمالي الشهور").FontSize(12d).Bold().Alignment = Alignment.center;
                }
                else if (Server == "56")
                {

                    t.Design = TableDesign.TableGrid;
                    t.Alignment = Alignment.center;
                    t.SetColumnWidth(7, 70);
                    t.SetColumnWidth(6, 70);
                    t.SetColumnWidth(5, 75);
                    t.SetColumnWidth(4, 75);
                    t.SetColumnWidth(3, 70);
                    t.SetColumnWidth(2, 70);
                    t.SetColumnWidth(1, 75);
                    t.SetColumnWidth(0, 55);

                    reportItems[0] = "إجراء خروج نهائي عام";//0+1
                    reportItems[1] = "إجراء خروج نهائي لمنطقة جدة";//2+4
                    reportItems[2] = "إجراء خروج نهائي لمنطقة مكة";
                    reportItems[3] = "إجراء خروج نهائي بالترحيل";//2+4
                    reportItems[4] = "إجراء تحويل مقابل مالي";
                    reportItems[5] = "إجراء خروج نهائي المحكمة العمالية";
                    reportItems[6] = "مجموع المعاملات";

                    t.Rows[0].Cells[0].Paragraphs[0].Append(reportItems[6]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[1].Paragraphs[0].Append(reportItems[5]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[2].Paragraphs[0].Append(reportItems[4]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[3].Paragraphs[0].Append(reportItems[3]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[4].Paragraphs[0].Append(reportItems[2]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[5].Paragraphs[0].Append(reportItems[1]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[6].Paragraphs[0].Append(reportItems[0]).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[7].Paragraphs[0].Append("الشهر").FontSize(12d).Bold().Alignment = Alignment.center;

                    int AllSum = 0;
                    for (int c = 1; c < 7; c++)
                    {
                        AllSum = 0;
                        for (int r = 0; r < lengthRow; r++)
                        {
                            AllSum = AllSum + rep1[r, c];
                        }

                        monthSumH[c] = AllSum;
                    }


                    for (int r = 0; r < lengthRow; r++)
                    {
                        AllSum = 0;
                        for (int c = 1; c < 7; c++)
                        {
                            AllSum = AllSum + rep1[r, c];
                        }

                        monthSumV[r] = AllSum;

                    }

                    AllSum = 0;
                    for (int c = 1; c < 7; c++)
                    {
                        AllSum = AllSum + monthSumH[c];
                    }
                    int x = 0;
                    for (int w = 0; w < lengthRow; w++)
                    {
                        t.Rows[w + 1].Cells[0].Paragraphs[0].Append(monthSumV[w].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[1].Paragraphs[0].Append(rep1[w, 6].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[2].Paragraphs[0].Append(rep1[w, 5].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[3].Paragraphs[0].Append(rep1[w,4].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[4].Paragraphs[0].Append(rep1[w, 3].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[5].Paragraphs[0].Append(rep1[w, 2].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[6].Paragraphs[0].Append(rep1[w, 1].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[7].Paragraphs[0].Append(Monthorder(rep1[w, 0]).ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    }

                    t.Rows[lengthRow + 1].Cells[0].Paragraphs[0].Append(AllSum.ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[1].Paragraphs[0].Append(monthSumH[6].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[2].Paragraphs[0].Append(monthSumH[5].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[3].Paragraphs[0].Append(monthSumH[4].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[4].Paragraphs[0].Append(monthSumH[3].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[5].Paragraphs[0].Append(monthSumH[2].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[6].Paragraphs[0].Append(monthSumH[1].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                    t.Rows[lengthRow + 1].Cells[7].Paragraphs[0].Append("إجمالي الشهور").FontSize(12d).Bold().Alignment = Alignment.center;
                }



                var p = document.InsertParagraph(Environment.NewLine);
                p.InsertTableAfterSelf(t);
                string strAttvCo = Environment.NewLine + Environment.NewLine + attendedVC.Text + Environment.NewLine + "ع/ القنصل العام بالإنابة";
                var AttvCo = document.InsertParagraph(strAttvCo)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(20d)
                    .Bold()
                    .Alignment = Alignment.center;

                document.Save();
                Process.Start("WINWORD.EXE", ActiveCopy);
            }

        }

        private string getFileTable(int index)
        {
            string table = "";
            switch (index)
            {
                case 0:
                    table = "TableWafid";
                    break;
                case 1:
                    table = "TableWafidJed";
                    break;
                case 2:
                    table = "TableWafidMekkah";
                    break;
                case 3:
                    table = "TableTarheel";
                    break;
                case 4:
                    table = "TableTransfer";
                    break;
                case 5:
                    table = "TableCommity";
                    break;
            }
            return table;
        }

        private string getNameTable(int index)
        {
            string table = "";
            switch (index)
            {
                case 0:
                    table = "وافدين عام";
                    break;
                case 1:
                    table = "وافدين جدة";
                    break;
                case 2:
                    table = "وافدين مكة";
                    break;
                case 3:
                    table = "الترحيل عام";
                    break;
                case 4:
                    table = "تحويل المقابل المالي";
                    break;
                case 5:
                    table = "المحكمة العمالية";
                    break;
            }
            return table;
        }
        private string weekorder(int week)
        {
            switch (week)
            {
                case 0:
                    return "الأول";

                case 1:
                    return "الثاني";

                case 2:
                    return "الثالث";

                case 3:
                    return "الرابع";

                case 4:
                    return "الخامس";

                default:
                    return "";

            }
        }


        private string Monthorder(int month)
        {
            switch (month)
            {
                case 1:
                    return "يناير";


                case 2:
                    return "فبراير";


                case 3:
                    return "مارس";


                case 4:
                    return "ابريل";

                case 5:
                    return "مايو";


                case 6:
                    return "يونيو";


                case 7:
                    return "يوليو";


                case 8:
                    return "أغسطس";

                case 9:
                    return "سبتمبر";


                case 10:
                    return "اكتوبر";


                case 11:
                    return "نوفمبر";


                case 12:
                    return "ديسمبر";
                default:
                    return "يناير";

            }
        }

        void SearchByName(string search)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);

            if (sqlCon.State == ConnectionState.Closed)

                if (txtSearch.Text != "")
                {
                    for (TableIndex = 1; TableIndex < 13; TableIndex++)
                    {
                        if (sqlCon.State == ConnectionState.Closed)
                            try
                            {
                                sqlCon.Open();
                            }
                            catch (Exception ex) { return; }
                        SqlCommand sqlCmd1 = new SqlCommand(query[TableIndex], sqlCon);
                        if (TableIndex < 12)
                        {

                            sqlCmd1.Parameters.Add("@AppName", SqlDbType.NVarChar).Value = search;
                            var reader = sqlCmd1.ExecuteReader();

                            if (reader.Read())
                            {
                                if (TableIndex != 10 && TableIndex != 11)
                                {
                                    IDNo = Convert.ToInt32(reader["ID"].ToString());
                                    txtSearch.Text = reader["DocID"].ToString();
                                    date.Text = reader["GriDate"].ToString();

                                    string viewSt = reader["Viewed"].ToString();
                                    string filename1 = reader["FileName1"].ToString();
                                    string filename2 = reader["FileName2"].ToString();
                                    if (filename1 == "text1.txt") Arch1.Visible = false;
                                    if (filename2 == "text2.txt") Arch2.Visible = false;

                                  

                                    string mandoub = reader["DataMandoubName"].ToString();
                                    if (mandoub != "")
                                        Apptype.Text = "بواسطة مندوب القنصلية " + mandoub;
                                    else Apptype.Text = "حضور مباشرة إلى القنصلية";


                                    if (reader["ArchivedState"].ToString() == "مؤرشف نهائي")
                                    {
                                        ArchiveSt.CheckState = CheckState.Checked;
                                        ArchiveSt.Text = "المكاتبة مؤرشفة";
                                        ArchiveSt.BackColor = Color.Green;
                                    }
                                    else if (reader["ArchivedState"].ToString().Contains("ملغي"))
                                    {
                                        ArchiveSt.CheckState = CheckState.Unchecked;
                                        ArchiveSt.Text = "المكاتبة ملغية";
                                        ArchiveSt.BackColor = Color.Red;
                                    }
                                    else
                                    {
                                        ArchiveSt.CheckState = CheckState.Unchecked;
                                        ArchiveSt.Text = "المكاتبة غير مؤرشفة";
                                        ArchiveSt.BackColor = Color.Red;
                                    }

                                    SearchPanel.Height = 296;
                                }
                                switch (TableIndex)
                                {
                                    case 1:
                                        strHijriDate = reader["Hijri"].ToString();
                                        strViseConsul = reader["AtteVicCo"].ToString();
                                        bolApplicantSex = reader["Gender"].ToString();
                                        if (bolApplicantSex == "ذكر") bolApplicantSex = "المواطن"; else bolApplicantSex = "المواطنة";
                                        intMessageType = Convert.ToInt32(reader["ProType"].ToString());
                                        strMessageType = "";
                                        switch (intMessageType)
                                        {
                                            case 0:
                                                strMessageType = travType[intMessageType];
                                                break;
                                            case 1:
                                                strMessageType = travType[intMessageType];
                                                break;
                                            case 2:
                                                strMessageType = travType[0] + " و" + travType[1];
                                                break;
                                        }
                                        txtEmbassey.Text = strEmbassySource = reader["Embassy"].ToString();

                                        PrintMessage.Visible = true;
                                        DetecedForm.Width = 184;
                                        break;
                                    case 2:
                                        strHijriDate = reader["Hijri"].ToString();
                                        strViseConsul = reader["AtteVicCo"].ToString();
                                        bolApplicantSex = reader["Gender"].ToString();
                                        if (bolApplicantSex == "ذكر") bolApplicantSex = "المواطن"; else bolApplicantSex = "المواطنة";
                                        strMessageType = reader["IqrarPurpose"].ToString();
                                        if (txtEmbassey.Text == "") strEmbassySource = "الخرطوم";
                                        else
                                            strEmbassySource = txtEmbassey.Text;
                                        PrintMessage.Visible = true;
                                        DetecedForm.Width = 184;
                                        break;
                                    case 4:
                                        strHijriDate = reader["Hijri"].ToString();
                                        strViseConsul = reader["AtteVicCo"].ToString();
                                        bolApplicantSex = reader["Gender"].ToString();
                                        if (bolApplicantSex == "ذكر") bolApplicantSex = "المواطن"; else bolApplicantSex = "المواطنة";
                                        intMessageType = Convert.ToInt32(reader["ProCase"].ToString());
                                        strMessageType = "";
                                        switch (intMessageType)
                                        {
                                            case 0:
                                                strMessageType = travType[intMessageType];
                                                break;
                                            case 1:
                                                strMessageType = travType[intMessageType];
                                                break;
                                            case 2:
                                                strMessageType = travType[0] + " و" + travType[1];
                                                break;
                                        }
                                        if (txtEmbassey.Text == "") strEmbassySource = "الخرطوم";
                                        else
                                            strEmbassySource = txtEmbassey.Text;
                                        PrintMessage.Visible = true;
                                        DetecedForm.Width = 184;
                                        break;


                                    case 10:

                                        Messid = Convert.ToInt32(reader["ID"].ToString());
                                        txtMessApp.Text = reader["DocID"].ToString();
                                        txtMessAuth.Text = reader["AuthName"].ToString();
                                        txtAuthNo.Text = reader["AuthNo"].ToString();
                                        if (reader["Gender"].ToString() == "ذكر")
                                        {
                                            checkMessAppSex.Text = "ذكر";
                                            checkMessAppSex.CheckState = CheckState.Unchecked;
                                        }
                                        else
                                        {
                                            checkMessAppSex.Text = "أنثى";
                                            checkMessAppSex.CheckState = CheckState.Checked;
                                        }
                                        txtMessSource.Text = reader["Institute"].ToString();
                                        txtMessGreDate.Text = reader["GriDate"].ToString();
                                        processed.Text = reader["Viewed"].ToString();
                                        if (btnMessSave.Text == "حفظ")
                                        {
                                            btnMessSave.Text = "تعديل";
                                        }
                                        else if (btnMessSave.Text == "حفظ وتأكيد")
                                            btnMessSave.Text = "تعديل وتأكيد";

                                        if (reader["ArchivedState"].ToString() == "مؤرشف نهائي")
                                        {
                                            checkMessArch.CheckState = CheckState.Checked;
                                            checkMessArch.Text = "المكاتبة مؤرشفة";
                                            checkMessArch.BackColor = Color.Green;
                                            btnMessArch.Text = "معاينة المكاتبة";
                                            btnMessArch.Width = 93;
                                            processed.Visible = button13.Visible = true;
                                        }
                                        else if (reader["ArchivedState"].ToString().Contains("ملغي"))
                                        {
                                            checkMessArch.CheckState = CheckState.Checked;
                                            checkMessArch.Text = "المكاتبة تم إلغاؤها";
                                            checkMessArch.BackColor = Color.Green;
                                            btnMessArch.Text = "معاينة المكاتبة";
                                            btnMessArch.Width = 93;
                                            processed.Visible = button13.Visible = true;
                                        }
                                        else
                                        {
                                            checkMessArch.CheckState = CheckState.Unchecked;
                                            checkMessArch.Text = "غير مؤرشف";
                                            checkMessArch.BackColor = Color.Red;
                                            btnMessArch.Text = "تحميل ملف ارشفة المعاملة";
                                            btnMessArch.Width = 184;
                                            processed.Visible = button13.Visible = false;
                                        }


                                        break;

                                    case 11:

                                        Messid = Convert.ToInt32(reader["ID"].ToString());
                                        txtHAAuthentic.Text = reader["AppName"].ToString();
                                        if (reader["Gender"].ToString() == "ذكر")
                                        {
                                            checkHASex.Text = "ذكر";
                                            checkHASex.CheckState = CheckState.Unchecked;
                                        }
                                        else
                                        {
                                            checkHASex.Text = "أنثى";
                                            checkHASex.CheckState = CheckState.Checked;
                                        }
                                        comboBox2.Text = reader["Institute"].ToString();
                                        txtHAGredate.Text = reader["HandTime"].ToString();
                                        HandProcess.Text = reader["Viewed"].ToString();
                                        if (btnSaveArch.Text == "حفظ")
                                        {
                                            btnSaveArch.Text = "تعديل";
                                        }
                                        else if (btnSaveArch.Text == "حفظ وتأكيد")
                                            btnSaveArch.Text = "تعديل وتأكيد";
                                        break;
                                }
                            }
                        }
                        else
                        {
                            SearchPanel.Height = 40;
                            sqlCmd1.Parameters.Add("@رقم_التوكيل", SqlDbType.NVarChar).Value = txtSearch.Text;
                            var reader = sqlCmd1.ExecuteReader();
                            if (reader.Read())
                            {
                                authNo = txtSearch.Text;
                                IDNo = Convert.ToInt32(reader["ID"].ToString());
                                applicant.Text = reader["مقدم_الطلب"].ToString();
                                date.Text = reader["التاريخ_الميلادي"].ToString();
                                string viewSt = reader["المعالجة"].ToString();
                                string filename1 = reader["ارشفة_المستندات"].ToString();
                                string filename2 = reader["المكاتبة_النهائية"].ToString();
                                if (filename1 == "text1.txt") Arch1.Visible = false;
                                if (filename2 == "text2.txt") Arch2.Visible = false;

                                
                                string mandoub = reader["اسم_المندوب"].ToString();
                                if (mandoub != "")
                                    Apptype.Text = "بواسطة مندوب القنصلية " + mandoub;
                                else Apptype.Text = "حضور مباشرة إلى القنصلية";


                                if (reader["حالة_الارشفة"].ToString() != "غير مؤرشف")
                                {
                                    ArchiveSt.CheckState = CheckState.Checked;
                                    ArchiveSt.Text = "مؤرشف";
                                    ArchiveSt.BackColor = Color.Green;
                                }
                                else
                                {
                                    ArchiveSt.CheckState = CheckState.Unchecked;
                                    ArchiveSt.Text = "غير مؤرشف";
                                    ArchiveSt.BackColor = Color.Red;
                                }

                                SearchPanel.Height = 296;
                                break;
                            }
                        }
                        sqlCon.Close();
                    }
                }

        }

        void FillDataGridView(string search)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);

            if (sqlCon.State == ConnectionState.Closed)

                if (txtSearch.Text != "")
                {
                    for (TableIndex = 1; TableIndex < 13; TableIndex++)
                    {
                        if (sqlCon.State == ConnectionState.Closed)
                            try
                            {
                                sqlCon.Open();
                            }
                            catch (Exception ex) { return; }
                        SqlCommand sqlCmd1 = new SqlCommand(query[TableIndex], sqlCon);
                        if (TableIndex < 12)
                        {

                            sqlCmd1.Parameters.Add("@DocID", SqlDbType.NVarChar).Value = search;
                            var reader = sqlCmd1.ExecuteReader();

                            if (reader.Read())
                            {
                                if (TableIndex != 10 && TableIndex != 11)
                                {
                                    IDNo = Convert.ToInt32(reader["ID"].ToString());
                                    applicant.Text = reader["AppName"].ToString();
                                    date.Text = reader["GriDate"].ToString();

                                    string viewSt = reader["Viewed"].ToString();
                                    string filename1 = reader["FileName1"].ToString();
                                    string filename2 = reader["FileName2"].ToString();
                                    if (filename1 == "text1.txt") Arch1.Visible = false;
                                    if (filename2 == "text2.txt") Arch2.Visible = false;


                                    string mandoub = reader["DataMandoubName"].ToString();
                                    if (mandoub != "")
                                        Apptype.Text = "بواسطة مندوب القنصلية " + mandoub;
                                    else Apptype.Text = "حضور مباشرة إلى القنصلية";


                                    if (reader["ArchivedState"].ToString() == "مؤرشف نهائي")
                                    {
                                        ArchiveSt.CheckState = CheckState.Checked;
                                        ArchiveSt.Text = "المكاتبة مؤرشفة";
                                        ArchiveSt.BackColor = Color.Green;
                                    }
                                    else if (reader["ArchivedState"].ToString().Contains("ملغي"))
                                    {
                                        ArchiveSt.CheckState = CheckState.Unchecked;
                                        ArchiveSt.Text = "المكاتبة ملغية";
                                        ArchiveSt.BackColor = Color.Red;
                                    } else
                                    {
                                        ArchiveSt.CheckState = CheckState.Unchecked;
                                        ArchiveSt.Text = "المكاتبة غير مؤرشفة";
                                        ArchiveSt.BackColor = Color.Red;
                                    }

                                    SearchPanel.Height = 296;
                                }
                                switch (TableIndex)
                                {
                                    case 1:
                                        //query[1] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender,Embassy,ProType  from TableTravIqrar where DocID=@DocID";
                                        //query[2] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender,IqrarPurpose from TableMultiIqrar where DocID=@DocID";
                                        //query[3] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableVisaApp where DocID=@DocID";
                                        //query[4] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender  from TableFamilySponApp where DocID=@DocID";
                                        //query[5] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableForensicApp where DocID=@DocID";
                                        //query[6] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender  from TableTRName where DocID=@DocID";
                                        //query[7] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender,IqrarPurpose  from TableStudent where DocID=@DocID";
                                        //query[8] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender  from TableMarriage where DocID=@DocID";
                                        //query[9] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2,Hijri,AtteVicCo,Gender,SpecType  from TableCollection where DocID=@DocID";
                                        //query[10] = "select ID,مقدم_الطلب,التاريخ_الميلادي,المعالجة,اسم_المندوب,حالة_الارشفة,ارشفة_المستندات,المكاتبة_النهائية,التاريخ_الهجري,موقع_التوكيل,النوع,وجهة_التوكيل,نوع_التوكيل  from TableAuth where رقم_التوكيل=@رقم_التوكيل";


                                        strHijriDate = reader["Hijri"].ToString();
                                        strViseConsul = reader["AtteVicCo"].ToString();
                                        bolApplicantSex = reader["Gender"].ToString();
                                        if (bolApplicantSex == "ذكر") bolApplicantSex = "المواطن"; else bolApplicantSex = "المواطنة";
                                        intMessageType = Convert.ToInt32(reader["ProType"].ToString());
                                        strMessageType = "";
                                        switch (intMessageType)
                                        {
                                            case 0:
                                                strMessageType = travType[intMessageType];
                                                break;
                                            case 1:
                                                strMessageType = travType[intMessageType];
                                                break;
                                            case 2:
                                                strMessageType = travType[0] + " و" + travType[1];
                                                break;
                                        }
                                        txtEmbassey.Text = strEmbassySource = reader["Embassy"].ToString();

                                        PrintMessage.Visible = true;
                                        DetecedForm.Width = 184;
                                        break;
                                    case 2:
                                        strHijriDate = reader["Hijri"].ToString();
                                        strViseConsul = reader["AtteVicCo"].ToString();
                                        bolApplicantSex = reader["Gender"].ToString();
                                        if (bolApplicantSex == "ذكر") bolApplicantSex = "المواطن"; else bolApplicantSex = "المواطنة";
                                        strMessageType = reader["IqrarPurpose"].ToString();
                                        if (txtEmbassey.Text == "") strEmbassySource = "الخرطوم";
                                        else
                                            strEmbassySource = txtEmbassey.Text;
                                        PrintMessage.Visible = true;
                                        DetecedForm.Width = 184;
                                        break;
                                    case 4:
                                        strHijriDate = reader["Hijri"].ToString();
                                        strViseConsul = reader["AtteVicCo"].ToString();
                                        bolApplicantSex = reader["Gender"].ToString();
                                        if (bolApplicantSex == "ذكر") bolApplicantSex = "المواطن"; else bolApplicantSex = "المواطنة";
                                        intMessageType = Convert.ToInt32(reader["ProCase"].ToString());
                                        strMessageType = "";
                                        switch (intMessageType)
                                        {
                                            case 0:
                                                strMessageType = travType[intMessageType];
                                                break;
                                            case 1:
                                                strMessageType = travType[intMessageType];
                                                break;
                                            case 2:
                                                strMessageType = travType[0] + " و" + travType[1];
                                                break;
                                        }
                                        if (txtEmbassey.Text == "") strEmbassySource = "الخرطوم";
                                        else
                                            strEmbassySource = txtEmbassey.Text;
                                        PrintMessage.Visible = true;
                                        DetecedForm.Width = 184;
                                        break;


                                    case 10:

                                        Messid = Convert.ToInt32(reader["ID"].ToString());
                                        txtMessApp.Text = reader["AppName"].ToString();
                                        txtMessAuth.Text = reader["AuthName"].ToString();
                                        txtAuthNo.Text = reader["AuthNo"].ToString();
                                        if (reader["Gender"].ToString() == "ذكر")
                                        {
                                            checkMessAppSex.Text = "ذكر";
                                            checkMessAppSex.CheckState = CheckState.Unchecked;
                                        }
                                        else
                                        {
                                            checkMessAppSex.Text = "أنثى";
                                            checkMessAppSex.CheckState = CheckState.Checked;
                                        }
                                        txtMessSource.Text = reader["Institute"].ToString();
                                        txtMessGreDate.Text = reader["GriDate"].ToString();
                                        processed.Text = reader["Viewed"].ToString();
                                        if (btnMessSave.Text == "حفظ")
                                        {
                                            btnMessSave.Text = "تعديل";
                                        }
                                        else if (btnMessSave.Text == "حفظ وتأكيد")
                                            btnMessSave.Text = "تعديل وتأكيد";

                                        if (reader["ArchivedState"].ToString() == "مؤرشف نهائي")
                                        {
                                            checkMessArch.CheckState = CheckState.Checked;
                                            checkMessArch.Text = "المكاتبة مؤرشفة";
                                            checkMessArch.BackColor = Color.Green;
                                            btnMessArch.Text = "معاينة المكاتبة";
                                            btnMessArch.Width = 93;
                                            processed.Visible = button13.Visible = true;
                                        }
                                        else if (reader["ArchivedState"].ToString().Contains("ملغي"))
                                        {
                                            checkMessArch.CheckState = CheckState.Checked;
                                            checkMessArch.Text = "المكاتبة تم إلغاؤها";
                                            checkMessArch.BackColor = Color.Green;
                                            btnMessArch.Text = "معاينة المكاتبة";
                                            btnMessArch.Width = 93;
                                            processed.Visible = button13.Visible = true;
                                        }
                                        else
                                        {
                                            checkMessArch.CheckState = CheckState.Unchecked;
                                            checkMessArch.Text = "غير مؤرشف";
                                            checkMessArch.BackColor = Color.Red;
                                            btnMessArch.Text = "تحميل ملف ارشفة المعاملة";
                                            btnMessArch.Width = 184;
                                            processed.Visible = button13.Visible = false;
                                        }


                                        break;

                                    case 11:

                                        Messid = Convert.ToInt32(reader["ID"].ToString());
                                        txtHAAuthentic.Text = reader["AppName"].ToString();
                                        if (reader["Gender"].ToString() == "ذكر")
                                        {
                                            checkHASex.Text = "ذكر";
                                            checkHASex.CheckState = CheckState.Unchecked;
                                        }
                                        else
                                        {
                                            checkHASex.Text = "أنثى";
                                            checkHASex.CheckState = CheckState.Checked;
                                        }
                                        comboBox2.Text = reader["Institute"].ToString();
                                        txtHAGredate.Text = reader["HandTime"].ToString();
                                        HandProcess.Text = reader["Viewed"].ToString();
                                        if (btnSaveArch.Text == "حفظ")
                                        {
                                            btnSaveArch.Text = "تعديل";
                                        }
                                        else if (btnSaveArch.Text == "حفظ وتأكيد")
                                            btnSaveArch.Text = "تعديل وتأكيد";
                                        break;
                                }
                            }
                        }
                        else
                        {
                            SearchPanel.Height = 40;
                            sqlCmd1.Parameters.Add("@رقم_التوكيل", SqlDbType.NVarChar).Value = txtSearch.Text;
                            var reader = sqlCmd1.ExecuteReader();
                            if (reader.Read())
                            {
                                authNo = txtSearch.Text;
                                IDNo = Convert.ToInt32(reader["ID"].ToString());
                                applicant.Text = reader["مقدم_الطلب"].ToString();
                                date.Text = reader["التاريخ_الميلادي"].ToString();
                                string viewSt = reader["المعالجة"].ToString();
                                string filename1 = reader["ارشفة_المستندات"].ToString();
                                string filename2 = reader["المكاتبة_النهائية"].ToString();
                                if (filename1 == "text1.txt") Arch1.Visible = false;
                                if (filename2 == "text2.txt") Arch2.Visible = false;
                                
                                string mandoub = reader["اسم_المندوب"].ToString();
                                if (mandoub != "")
                                    Apptype.Text = "بواسطة مندوب القنصلية " + mandoub;
                                else Apptype.Text = "حضور مباشرة إلى القنصلية";


                                if (reader["حالة_الارشفة"].ToString() != "غير مؤرشف")
                                {
                                    ArchiveSt.CheckState = CheckState.Checked;
                                    ArchiveSt.Text = "مؤرشف";
                                    ArchiveSt.BackColor = Color.Green;
                                }
                                else
                                {
                                    ArchiveSt.CheckState = CheckState.Unchecked;
                                    ArchiveSt.Text = "غير مؤرشف";
                                    ArchiveSt.BackColor = Color.Red;
                                }

                                SearchPanel.Height = 296;
                                break;
                            }
                        }
                        sqlCon.Close();
                    }
                }

        }

        private void selectArchData(string docid)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from  archives where docID=@docID", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@docID", docid);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            archState = "مؤرشف نهائي";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                mandoubInfo = dataRow["mandoubName"].ToString() + "_" + dataRow["appType"].ToString();
                if (dataRow["appOldNew"].ToString() == "new" )
                    archState = "غير مؤرشف";
                if (dataRow["appOldNew"].ToString() == "old" && dataRow["appType"].ToString() == "عن طريق مندوب للقنصلية")
                    archState = "لم يتم أرشفة التعديل";
                if (dataRow["appOldNew"].ToString() == "في انتظار نسخة المواطن")
                    archState = "في انتظار نسخة المواطن";
                
                if (dataRow["appType"].ToString() == "عن طريق مندوب للقنصلية") Apptype.Text = "عن طريق المندوب " + dataRow["mandoubName"].ToString();
                else Apptype.Text = dataRow["appType"].ToString();
            }
            
        }

        void FillDatafromGenArch(string search, string column)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);

            if (sqlCon.State == ConnectionState.Closed)

                if (txtSearch.Text != "")
                {

                    if (sqlCon.State == ConnectionState.Closed)
                        try
                        {
                            sqlCon.Open();
                        }
                        catch (Exception ex) { return; }
                    SqlCommand sqlCmd1 = new SqlCommand("select * from TableGeneralArch where " + column + "=@col", sqlCon);
                        

                            sqlCmd1.Parameters.Add("@col", SqlDbType.NVarChar).Value = search;
                            var reader = sqlCmd1.ExecuteReader();

                    if (reader.Read())
                    {
                        
                        IDNo = Convert.ToInt32(reader["رقم_المرجع"].ToString());
                        if (column == "رقم_معاملة_القسم")
                            applicant.Text = reader["الاسم"].ToString();
                        else if (column == "الاسم")
                            txtSearch.Text = reader["رقم_معاملة_القسم"].ToString();
                        date.Text = reader["التاريخ"].ToString();


                        //Apptype.Text = 

                        selectArchData(search);
                        if (archState == "مؤرشف نهائي")
                        {
                            ArchiveSt.CheckState = CheckState.Checked;
                            ArchiveSt.Text = "المكاتبة مؤرشفة";
                            ArchiveSt.BackColor = Color.Green;
                            
                        }
                        else if (archState.Contains("ملغي"))
                        {
                            ArchiveSt.CheckState = CheckState.Unchecked;
                            ArchiveSt.Text = "المكاتبة ملغية";
                            ArchiveSt.BackColor = Color.Red;
                        }
                        else if (archState == "لم يتم أرشفة التعديل")
                        {
                            ArchiveSt.CheckState = CheckState.Checked;
                            ArchiveSt.Text = "لم يتم أرشفة التعديل";
                            ArchiveSt.BackColor = Color.Green;
                        }
                        else if (archState == "في انتظار نسخة المواطن")
                        {
                            ArchiveSt.CheckState = CheckState.Unchecked;
                            ArchiveSt.Text = "في انتظار نسخة المواطن";
                            ArchiveSt.BackColor = Color.Red;
                        }
                        else
                        {
                            ArchiveSt.CheckState = CheckState.Unchecked;
                            ArchiveSt.Text = "المكاتبة غير مؤرشفة";
                            ArchiveSt.BackColor = Color.Red;
                            Arch2.Enabled = false;
                        }

                        SearchPanel.Height = 296;
                    }


                    sqlCon.Close();
                    
                }

        }
        void FillDatafromGenArch(string search, string doc, Button button)
        {

           
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableGeneralArch where  رقم_معاملة_القسم=N'" + txtSearch.Text + "' and نوع_المستند='" + doc + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow reader in dtbl.Rows)
            {
                var name = reader["المستند"].ToString();
                if (name == "") return;
                try
                {
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                    button.Enabled = false;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                    button.Enabled = true;
                }
                catch (Exception ex) { }
            }

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

        private void CreateMessageWord(string ApplicantName, string EmbassySource, string IqrarNo, string MessageType, string ApplicantSex, string GregorianDate, string HijriDate, string ViseConsul)
        {
            string ActiveCopy;
            string ReportName = DateTime.Now.ToString("mmss");
            string routeDoc = FilespathIn + @"\MessageCap.docx";
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
                BookDateGre.Text = GregorianDate;
                BookGregorDate2.Text = GregorianDate;
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
        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                FillDataGridView(txtSearch.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
            }
        }


        private void btnSearch_Click(object sender, EventArgs e)
        {

        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            //MessageBox.Show("main formclosing");
            //FormDataBase formDataBase = new FormDataBase(DataSource, FilespathIn, FilespathOut, ArchFile, FormDataFile);
            //formDataBase.ShowDialog();
            //this.Close();
        }

        private void txtSearch_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13) DetecedForm.PerformClick();

        }




        private void txtSearch_TextChanged_2(object sender, EventArgs e)
        {
            if (!nameNo) return;
            //MessageBox.Show("رقم_معاملة_القسم");
            FillDatafromGenArch(txtSearch.Text, "رقم_معاملة_القسم");
        }

        private void txtSearch_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13) DetecedForm.PerformClick();

        }



        private void Arch1_Click_1(object sender, EventArgs e)
        {
            //OpenFile(IDNo, 1);
            FillDatafromGenArch(txtSearch.Text, "data1",Arch1);
        }

        private void OpenMessFile(int id, string tableName)
        {
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select Data1, Extension1,FileName1  from " + tableName + " where ID=@id", Con);
            sqlCmd1.Parameters.Add("@ID", SqlDbType.Int).Value = id;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                var name = reader["FileName1"].ToString();
                var Data = (byte[])reader["Data1"];
                var ext = reader["Extension1"].ToString();
                var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                //button.Enabled = false;
                File.WriteAllBytes(NewFileName, Data);
                System.Diagnostics.Process.Start(NewFileName);
                //button.Enabled = true;

            }
            Con.Close();

        }
        private void OpenFil1e(int id, int fileNo)
        {
            string query = "select Data3, Extension3,DocxData from TableAuth where ID=@id";
            SqlConnection Con = new SqlConnection(DataSource);
            if (fileNo == 1)
            {
                query = "select Data1, Extension1,ارشفة_المستندات from TableAuth where ID=@id";
            }
            else if (fileNo == 2)
            {
                query = "select Data2, Extension2,المكاتبة_النهائية from TableAuth where ID=@id";
            }
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = id;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();

            if (reader.Read())
            {
                if (fileNo == 1)
                {
                    var name = reader["ارشفة_المستندات"].ToString();
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }
                else if (fileNo == 2)
                {
                    var name = reader["المكاتبة_النهائية"].ToString();
                    var Data = (byte[])reader["Data2"];
                    var ext = reader["Extension2"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);

                }
                else
                {

                    var name = reader["DocxData"].ToString();
                    var Data = (byte[])reader["Data3"];
                    var ext = reader["Extension3"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }


            }
            Con.Close();

        }
        private void OpenFile(int id, int fileNo)
        {
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand(querydatabase[TableIndex], Con);
            sqlCmd1.Parameters.Add("@ID", SqlDbType.Int).Value = id;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                if (fileNo == 1)
                {
                    var name = reader["FileName1"].ToString();
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }
                else
                {
                    var name = reader["FileName2"].ToString();
                    var Data = (byte[])reader["Data2"];
                    var ext = reader["Extension2"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }

            }
            Con.Close();

        }

        private void Arch2_Click(object sender, EventArgs e)
        {

        }


        private int HijriDateDifferment(string source)
        {
            int differment = 0;
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select hijriModification from TableSettings";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    differment = Convert.ToInt32(reader["hijriModification"].ToString());

                    labeldate.Text = differment.ToString();
                }

                saConn.Close();
            }
            return differment;
        }


        private void ReportType_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            switch (ReportType.SelectedIndex)
            {
                case 0:
                    ReportNo.Enabled = true;
                    attendedVC.Visible = true;
                    btnattendedVC.Enabled = true;
                    btnReportNo.Enabled = true;
                    ReportPanel.Height = 205;
                    break;
                case 1:
                    yearReport.Visible = false;
                    button24.Visible = false;
                    button24.Enabled = false;
                    button28.Visible = false;
                    dateTimeFrom.Visible = false;
                    dateTimeTo.Visible = false;
                    DailyList(GregorianDate);
                    if (totalrowsAffadivit > 0 || totalrowsAuth > 0)
                    {

                        PrintReport.Enabled = true;
                        PrintReport.Visible = true;
                        ReportPanel.Height = 205;


                    }
                    else
                    {
                        PrintReport.Enabled = false;
                        PrintReport.Visible = false;
                        ReportNo.Enabled = true;
                        attendedVC.Visible = true;
                        btnattendedVC.Enabled = true;
                        btnReportNo.Enabled = true;
                        ReportPanel.Height = 42;
                        MessageBox.Show("لا يوجد قائمة بالتاريخ المحدد");
                    }
                    var selectedOption = MessageBox.Show("", "إزالة الملفات غير المكتملة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (selectedOption == DialogResult.Yes)
                    {
                        deleteEmptyRows = true;
                        var partAll = MessageBox.Show("", "إستنثاء الملفات المؤرشفة بتاريخ اليوم؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (selectedOption == DialogResult.Yes)
                        {
                            parrtialAll = true;
                        }
                        else if (selectedOption == DialogResult.No)
                            parrtialAll = false;
                    }
                    break;
                case 2:
                    yearReport.Visible = false;
                    button24.Text = "يوم:";
                    button24.Visible = true;
                    button24.Enabled = true;
                    button28.Visible = false;
                    dateTimeFrom.Visible = true;
                    dateTimeTo.Visible = false;
                    dateTimeFrom.Width = 288;
                    btnattendedVC.Enabled = true;
                    btnReportNo.Enabled = true;
                    ReportPanel.Height = 205;
                    break;
                case 3:
                    yearReport.Visible = false;
                    button24.Text = "من:";
                    dateTimeFrom.Width = 113;
                    button24.Visible = true;
                    button28.Visible = true;
                    dateTimeFrom.Visible = true;
                    dateTimeTo.Visible = true;
                    ReportPanel.Height = 205;
                    break;
                case 4:
                    button24.Text = "السنة:";
                    quorterS[0] = "-01-01-";
                    quorterE[0] = "-01-31-";
                    quorterS[1] = "-02-01-";
                    quorterE[1] = "-02-29-";
                    quorterS[2] = "-03-01-";
                    quorterE[2] = "-03-31-";
                    button24.Enabled = true;
                    yearReport.Visible = false;
                    ReportPanel.Height = 205;
                    button24.Visible = true;
                    yearReport.Visible = true;
                    button28.Visible = false;
                    dateTimeTo.Visible = false;
                    break;
                case 5:
                    button24.Text = "السنة:";
                    quorterS[0] = "-04-01-";
                    quorterE[0] = "-04-30-";
                    quorterS[1] = "-05-01-";
                    quorterE[1] = "-05-31-";
                    quorterS[2] = "-06-01-";
                    quorterE[2] = "-06-30-";
                    yearReport.Visible = false;
                    button24.Enabled = true;
                    button24.Visible = true;
                    yearReport.Visible = true;
                    button28.Visible = false;
                    dateTimeTo.Visible = false;
                    ReportPanel.Height = 205;
                    break;
                case 6:
                    button24.Text = "السنة:";
                    quorterS[0] = "-07-01-";
                    quorterE[0] = "-07-31-";
                    quorterS[1] = "-08-01-";
                    quorterE[1] = "-08-31-";
                    quorterS[2] = "-09-01-";
                    quorterE[2] = "-09-30-";
                    yearReport.Visible = false;
                    button24.Enabled = true;
                    button24.Visible = true;
                    yearReport.Visible = true;
                    button28.Visible = false;
                    dateTimeTo.Visible = false;
                    ReportPanel.Height = 205;
                    break;
                case 7:
                    button24.Text = "السنة:";
                    quorterS[0] = "-10-01-";
                    quorterE[0] = "-10-30-";
                    quorterS[1] = "-11-01-";
                    quorterE[1] = "-11-30-";
                    quorterS[2] = "-12-01-";
                    quorterE[2] = "-12-31-";
                    yearReport.Visible = false;
                    button24.Enabled = true;
                    button24.Visible = true;
                    yearReport.Visible = true;
                    button28.Visible = false;
                    dateTimeTo.Visible = false;
                    ReportPanel.Height = 205;
                    break;
                case 8:
                    button24.Text = "السنة:";
                    quorterS[0] = "-01-01-";
                    quorterE[0] = "-01-31-";
                    quorterS[1] = "-02-01-";
                    quorterE[1] = "-02-29-";
                    quorterS[2] = "-03-01-";
                    quorterE[2] = "-03-31-";
                    quorterS[3] = "-04-01-";
                    quorterE[3] = "-04-30-";
                    quorterS[4] = "-05-01-";
                    quorterE[4] = "-05-31-";
                    quorterS[5] = "-06-01-";
                    quorterE[5] = "-06-30-";
                    quorterS[6] = "-07-01-";
                    quorterE[6] = "-07-31-";
                    quorterS[7] = "-08-01-";
                    quorterE[7] = "-08-31-";
                    quorterS[8] = "-09-01-";
                    quorterE[8] = "-09-30-";
                    quorterS[9] = "-10-01-";
                    quorterE[9] = "-10-30-";
                    quorterS[10] = "-11-01-";
                    quorterE[10] = "-11-30-";
                    quorterS[11] = "-12-01-";
                    quorterE[11] = "-12-31-";
                    yearReport.Visible = false;
                    button24.Enabled = true;
                    button24.Visible = true;
                    yearReport.Visible = true;
                    button28.Visible = false;
                    dateTimeTo.Visible = false;
                    ReportPanel.Height = 205;
                    break;
                case 9:

                    reportpass.Visible = true;

                    break;
                case 10:
                    fillDataGridReports();
                    btnReportSub.Visible = PrintReport.Visible = txtReportSub.Visible = true;
                    PrintReport.Text = "إضافة";

                    break;
            }
            //yearReport.Items.Clear();
            if (ReportType.SelectedIndex >= 4 && ReportType.SelectedIndex != 10)
            {
                //MessageBox.Show(ReportType.SelectedIndex.ToString());
                getDate();
            }
        }

        private void insertDoc(string dataSource, string extn1, string DocName1, string messNo, string docType, byte[] buffer1)
        {

            string query = "INSERT INTO TableGeneralArch (Data1,Extension1,نوع_المستند,رقم_معاملة_القسم,المستند,الموظف,التاريخ) values (@Data1,@Extension1,@نوع_المستند,@رقم_معاملة_القسم,@المستند,@الموظف,@التاريخ)";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", messNo);
            sqlCmd.Parameters.AddWithValue("@نوع_المستند", docType);
            sqlCmd.Parameters.AddWithValue("@الموظف", ConsulateEmployee.Text);
            sqlCmd.Parameters.AddWithValue("@التاريخ", GregorianDate);

            sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
            sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
            sqlCmd.Parameters.Add("@المستند", SqlDbType.NVarChar).Value = DocName1;


            //if (filePath != "")
            //{

            //    using (Stream stream = File.OpenRead(filePath))
            //    {
            //        byte[] buffer1 = new byte[stream.Length];
            //        stream.Read(buffer1, 0, buffer1.Length);
            //        var fileinfo1 = new FileInfo(filePath);
            //        string extn1 = fileinfo1.Extension;
            //        string DocName1 = fileinfo1.Name;
            //        sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
            //        sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
            //        sqlCmd.Parameters.Add("@المستند", SqlDbType.NVarChar).Value = DocName1;

            //    }
            //}
            sqlCmd.ExecuteNonQuery();

            sqlCon.Close();
        }

        private void getDate()
        {
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            if (Server == "57")
            {
                for (TableIndex = 1; TableIndex < 12; TableIndex++)
                {
                    Console.WriteLine("TableIndex " + queryDateList[TableIndex]);
                    if (TableIndex == 6||TableIndex == 4)
                    {
                        fillDate(queryDateList[TableIndex], "التاريخ_الميلادي");                         
                    }
                    else 
                    {
                        fillDate(queryDateList[TableIndex], "GriDate");                        
                    }                    
                }
            }
            else if (Server == "56")
            {
                for (TableIndex = 1; TableIndex < 7; TableIndex++)
                {
                    //string query = "select مقدم_الطلب  from " + getFileTable(TableIndex - 1) + " where التاريخ_الميلادي=@التاريخ_الميلادي";
                    string query = "select التاريخ_الميلادي  from " + getFileTable(TableIndex - 1);
                    fillDate(query, "التاريخ_الميلادي");                    
                }
            }
            sqlCon.Close();
        }


        private string ConvertTostring(string gregorianDate)
        {
            string[] strlist = new string[3];
            strlist = gregorianDate.Split('-');
            int monthInt = Convert.ToInt32("06");
            string strMonth = "";
            switch (monthInt)
            {
                case 1:
                    strMonth = "Jan";
                    break;
                case 2:
                    strMonth = "Feb";
                    break;
                case 3:
                    strMonth = "Mar";
                    break;
                case 4:
                    strMonth = "Apr";
                    break;
                case 5:
                    strMonth = "May";
                    break;
                case 6:
                    strMonth = "Jun";
                    break;
                case 7:
                    strMonth = "Jul";
                    break;
                case 8:
                    strMonth = "Aug";
                    break;
                case 9:
                    strMonth = "Sep";
                    break;
                case 10:
                    strMonth = "Oct";
                    break;
                case 11:
                    strMonth = "Nov";
                    break;
                case 12:
                    strMonth = "Dec";
                    break;
            }
            //MessageBox.Show(gregorianDate + Environment.NewLine + strlist[0] + "-" + strMonth + "-" + strlist[2]);
            return strlist[0] + "-" + strMonth + "-" + strlist[2];

        }

        //private string ConvertoString(string gregorianDate)
        //{
        //    

        //}

        private void PrintReport_Click(object sender, EventArgs e)
        {
            string ReportName1 = "Report1" + DateTime.Now.ToString("mmss") + ".docx";
            string ReportName2 = "Report2" + DateTime.Now.ToString("mmss") + ".docx";
            PrintReport.Enabled = false;
            PrintReport.Text = "تجري عملية الطباعة";
            if (ReportType.SelectedIndex != 10) {
                if (totalrowsAuth > 0)
                {
                    CreateDailyReportAuth(totalrowsAuth, ReportName1, " التواكيل", false);

                }
                if (totalrowsAffadivit > 0)
                {
                    CreateDailyReportIqrar(totalrowsAffadivit, ReportName2, " الإقرارات", true);

                }
                if (rowFound)
                {
                    while (File.Exists(ReportName2))
                        ReportName2 = "Report2" + DateTime.Now.ToString("mmss") + ".docx";
                    CreateDurationReport(rep1, ReportName2);

                }
                totalRowDuration = 0;
                totalrowsAffadivit = 0;
                totalrowsAuth = 0;
                PrintReport.Text = "طباعة التقرير";
                PrintReport.Enabled = false;
                PrintReport.Visible = false;
            } else {
                OpenFileDialog dlg = new OpenFileDialog();
                //dlg.ShowDialog();
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string filePath = dlg.FileName;
                    using (Stream stream = File.OpenRead(filePath))
                    {
                        byte[] buffer1 = new byte[stream.Length];
                        stream.Read(buffer1, 0, buffer1.Length);
                        var fileinfo1 = new FileInfo(filePath);
                        string extn1 = fileinfo1.Extension;
                        string DocName1 = fileinfo1.Name;
                        insertDoc(DataSource, extn1, DocName1, txtReportSub.Text, "تقرير", buffer1);
                    }
                }
            }
            ReportPanel.Height = 36;
        }

        private void IqrarBox_SelectedIndexChanged_1(object sender, EventArgs e)
        {


            if (persbtn3.SelectedIndex >= 0 && persbtn3.SelectedIndex <= 7)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    string[] str = new string[persbtn3.Items.Count];
                    for (int x = 0; x < persbtn3.Items.Count; x++) { str[x] = persbtn3.Items[x].ToString(); }
                    string[] strSub = new string[1] { "" };
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn3.SelectedIndex, FormDataFile, FilespathOut, 3, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    Form3 form3 = new Form3(attendedVC.SelectedIndex, IDNo, persbtn3.SelectedIndex, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form3.ShowDialog();
                }
            }

            else if (persbtn3.SelectedIndex == 8)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn3.Items.Count];
                    for (int x = 0; x < persbtn3.Items.Count; x++) { str[x] = persbtn3.Items[x].ToString(); }
                    string[] strSub = new string[4] { "نقل كفالة مقدم الطلب إلى كفالة طرف ثاني", "نقل كفالة طرف ثاني إلى كفالة مقدم الطلب", "نقل كفالة أحد مكفولي مقدم الطلب إلى كفالة طرف ثاني", "استقدام على كفالة طرف ثاني" };
                    FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn3.SelectedIndex, FormDataFile, FilespathOut, 5, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    Form5 form5 = new Form5(attendedVC.SelectedIndex, IDNo, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form5.ShowDialog();
                }
            }
            else if (persbtn3.SelectedIndex == 9)
            {
                MessageBox.Show("Off");
                //Form1 form1 = new Form1(comboBox1.SelectedIndex,IDNo, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition);
                //form1.ShowDialog();
            }

            else if (persbtn3.SelectedIndex == 10)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn3.Items.Count];
                    for (int x = 0; x < persbtn3.Items.Count; x++) { str[x] = persbtn3.Items[x].ToString(); }
                    string[] strSub = new string[7] { "استخراج وثائق للابناء", "عدم ممانعة سفر الابناء", "استخراج وثائق وعدم ممانعة سفر الابناء", "استخراج وثائق وعدم ممانعة سفر الابناء والزوجة", "استخراج وثائق وعدم ممانعة سفر الابناء بصحبة مرافق غير الزوجة", "عدم ممانعة سفر الابناء والزوجة", "عدم ممانعة سفر الزوجة" };
                    FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn3.SelectedIndex, FormDataFile, FilespathOut, 2, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    Form2 form2 = new Form2(attendedVC.SelectedIndex, IDNo, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form2.ShowDialog();
                }
            }

            else if (persbtn3.SelectedIndex == 11)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn3.Items.Count];
                    for (int x = 0; x < persbtn3.Items.Count; x++) { str[x] = persbtn3.Items[x].ToString(); }
                    string[] strSub = new string[2] { "اثبات اسمان لذات واحدة", "اثبات صحة وثائق" };
                    FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn3.SelectedIndex, FormDataFile, FilespathOut, 7, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    Form7 form7 = new Form7(attendedVC.SelectedIndex, IDNo, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form7.ShowDialog();
                }
            }

            else if (persbtn3.SelectedIndex == 12)
            {

            }

        }

        private void IfadaBox_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (persbtn4.SelectedIndex == 0)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn4.Items.Count];
                    for (int x = 0; x < persbtn4.Items.Count; x++) { str[x] = persbtn4.Items[x].ToString(); }
                    string[] strSub = new string[4] { "إختر أو أكتب الغرض", "السفر للدراسة", "السفر للسياحة", "السفر للعلاج" };
                    FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn4.SelectedIndex, FormDataFile, FilespathOut, 6, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    if (mangerArch.CheckState == CheckState.Checked)
                    {
                        dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                        string[] str = new string[persbtn4.Items.Count];
                        for (int x = 0; x < persbtn4.Items.Count; x++) { str[x] = persbtn4.Items[x].ToString(); }
                        string[] strSub = new string[1] { "" };
                        FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn4.SelectedIndex, FormDataFile, FilespathOut, 6, str, strSub, true, MandoubM, GriDateM);
                        form2.ShowDialog();
                    }
                    else
                    {
                        dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                        Form6 form6 = new Form6(attendedVC.SelectedIndex, -1, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                        form6.ShowDialog();
                    }
                }
            }
            else if (persbtn4.SelectedIndex == 1)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn4.Items.Count];
                    for (int x = 0; x < persbtn4.Items.Count; x++) { str[x] = persbtn4.Items[x].ToString(); }
                    string[] strSub = new string[1] { "" };
                    FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn4.SelectedIndex, FormDataFile, FilespathOut, 8, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    Form8 form8 = new Form8(attendedVC.SelectedIndex, -1, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form8.ShowDialog();
                }
            }
            else if (persbtn4.SelectedIndex == 2)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn4.Items.Count];
                    for (int x = 0; x < persbtn4.Items.Count; x++) { str[x] = persbtn4.Items[x].ToString(); }
                    string[] strSub = new string[1] { "" };
                    FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn4.SelectedIndex, FormDataFile, FilespathOut, 10, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    Form10 form10 = new Form10(attendedVC.SelectedIndex, -1, 2, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form10.ShowDialog();
                }
            }
        }

        private void ShehadaBox_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (persbtn5.SelectedIndex == 0)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn5.Items.Count];
                    for (int x = 0; x < persbtn5.Items.Count; x++) { str[x] = persbtn5.Items[x].ToString(); }
                    string[] strSub = new string[2] { "عدم ممانعة زواج", "عدم ممانعة وشهادة كفاءة" };
                    FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn5.SelectedIndex, FormDataFile, FilespathOut, 9, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    Form9 form9 = new Form9(attendedVC.SelectedIndex, -1, 0, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form9.ShowDialog();
                }
            }
            if (persbtn5.SelectedIndex == 1)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn5.Items.Count];
                    for (int x = 0; x < persbtn5.Items.Count; x++) { str[x] = persbtn5.Items[x].ToString(); }
                    string[] strSub = new string[1] { "" };
                    FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn5.SelectedIndex, FormDataFile, FilespathOut, 9, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    if (mangerArch.CheckState == CheckState.Checked)
                    {
                        dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                        string[] str = new string[persbtn5.Items.Count];
                        for (int x = 0; x < persbtn5.Items.Count; x++) { str[x] = persbtn5.Items[x].ToString(); }
                        string[] strSub = new string[1] { "" };
                        FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn5.SelectedIndex, FormDataFile, FilespathOut, 9, str, strSub, true, MandoubM, GriDateM);
                        form2.ShowDialog();
                    }
                    else
                    {
                        dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                        Form9 form9 = new Form9(attendedVC.SelectedIndex, -1, 1, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                        form9.ShowDialog();
                    }
                }
            }
            else if (persbtn5.SelectedIndex == 2)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn5.Items.Count];
                    for (int x = 0; x < persbtn5.Items.Count; x++) { str[x] = persbtn5.Items[x].ToString(); }
                    string[] strSub = new string[1] { "" };
                    FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn5.SelectedIndex, FormDataFile, FilespathOut, 10, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    Form10 form10 = new Form10(attendedVC.SelectedIndex, -1, 3, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form10.ShowDialog();
                }
            }
            else if (persbtn5.SelectedIndex == 3)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn5.Items.Count];
                    for (int x = 0; x < persbtn5.Items.Count; x++) { str[x] = persbtn5.Items[x].ToString(); }
                    string[] strSub = new string[1] { "" };
                    FormPics form2 = new FormPics(Server, EmployeeName, attendedVC.Text, UserJobposition, DataSource, persbtn5.SelectedIndex, FormDataFile, FilespathOut, 15, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    MerriageDoc merriageDoc = new MerriageDoc(DataSource, false, EmployeeName, attendedVC.SelectedIndex, GregorianDate, HijriDate);
                    merriageDoc.ShowDialog();
                }
            }
            else if (persbtn5.SelectedIndex == 4)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn5.Items.Count];
                    for (int x = 0; x < persbtn5.Items.Count; x++) { str[x] = persbtn5.Items[x].ToString(); }
                    string[] strSub = new string[1] { "" };
                    FormPics form2 = new FormPics(Server, EmployeeName, attendedVC.Text, UserJobposition, DataSource, persbtn5.SelectedIndex, FormDataFile, FilespathOut, 16, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");                    
                    PassAway passAway = new PassAway(attendedVC.SelectedIndex, DataSource, FilespathIn, FilespathOut, UserJobposition, EmployeeName, GregorianDate, HijriDate);
                    passAway.ShowDialog();
                }
            }else if (persbtn5.SelectedIndex == 5)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn5.Items.Count];
                    for (int x = 0; x < persbtn5.Items.Count; x++) { str[x] = persbtn5.Items[x].ToString(); }
                    string[] strSub = new string[1] { "" };
                    FormPics form2 = new FormPics(Server, EmployeeName, attendedVC.Text, UserJobposition, DataSource, persbtn5.SelectedIndex, FormDataFile, FilespathOut, 17, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    FormDivorce formDivorce = new FormDivorce(DataSource, false, EmployeeName, attendedVC.SelectedIndex, GregorianDate, HijriDate);
                    formDivorce.ShowDialog();
                }
            }
        }

        private void VisaBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (persbtn6.SelectedIndex == 0)
            {
                if (mangerArch.CheckState == CheckState.Checked)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    string[] str = new string[persbtn6.Items.Count];
                    for (int x = 0; x < persbtn6.Items.Count; x++) { str[x] = persbtn6.Items[x].ToString(); }
                    string[] strSub = new string[1] { "" };
                    FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition, DataSource, persbtn6.SelectedIndex, FormDataFile, FilespathOut, 4, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    Form4 form4 = new Form4(attendedVC.SelectedIndex, -1, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form4.ShowDialog();
                }
            }
            if (persbtn6.SelectedIndex == 1)
            {
                MessageBox.Show("غير مدرجة حتى الآن");
            }

        }

        private void btnSearch_Click_3(object sender, EventArgs e)
        {
            FillDataGridView(txtSearch.Text);
        }



        private void button2_Click_1(object sender, EventArgs e)
        {
            if (SearchPanel.Visible == false)
            {
                flowLayoutPanel1.Visible = SearchPanel.Visible = true;
                panel4.Visible = false;
                PanelMandounb.Visible = fileManagePanel2.Visible = panelAuthAknow.Visible = panelReceMess.Visible = ReportPanel.Visible = false;

            }
            else SearchPanel.Visible = false;
        }



        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {

        }



        private void button3_Click_1(object sender, EventArgs e)
        {
            if (ReportPanel.Visible == false)
            {
                ReportPanel.BringToFront();
                flowLayoutPanel1.Visible = ReportPanel.Visible = true;
                panel4.Visible = false;
                PanelMandounb.Visible = SearchPanel.Visible = fileManagePanel2.Visible = panelReceMess.Visible = panelAuthAknow.Visible = SearchPanel.Visible = false;
            }
            else ReportPanel.Visible = false;
        }

        private void ViewArchShow(int Buttons, string Doc, string date, string AppName, string oldNew, int size)
        {
            
            Button btnArchieve = new Button();
            btnArchieve.Dock = System.Windows.Forms.DockStyle.Top;
            btnArchieve.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            btnArchieve.Location = new System.Drawing.Point(4, 125);
            btnArchieve.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            btnArchieve.Name = IDA[Buttons].ToString();
            btnArchieve.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            btnArchieve.Size = new System.Drawing.Size(568, 34 * size);
            btnArchieve.TabIndex = 512;
            btnArchieve.Click += new System.EventHandler(this.button_Click);
            if (size == 1)
                btnArchieve.Text = (Buttons + 1).ToString() + " - " + AppName + " - " + Doc+ " - " + date;
            else btnArchieve.Text = (Buttons + 1).ToString() + " - " + AppName + Environment.NewLine + Doc + " - " + date;

            btnArchieve.UseVisualStyleBackColor = true;
            flowLayoutPanel1.Controls.Add(btnArchieve);
        }

        private void button_Click(object sender, EventArgs e)
        {
            if (UserJobposition.Contains("قنصل"))
            {
                Button button = (Button)sender;
                var selectedOption = MessageBox.Show("", "حذف المكاتبة من سجل المكاتبات غير المؤرشفة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                {
                    deleteRowsData(Convert.ToInt32(button.Name), "archives", DataSource);
                }
            }
        }
        private void ViewMandoubShow(int Buttons, string Doc, int ID, string AppName, string strOldNew)
        {
            labelM.Text = labelM.Text + (Buttons + 1).ToString() + " - " + Doc + strOldNew + " باسم " + AppName + Environment.NewLine;
        }

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {
            fillNonArchInfo();

            
        }

        private void fillNonArchInfo()
        {
            foreach (Control control in flowLayoutPanel1.Controls)
            {
                if (control is Button) control.Visible = false;
            }

            string ReportName = "Report" + DateTime.Now.ToString("mmss") + ".docx";
            if (A <= 0) MessageBox.Show("لا توجد معاملات غير مؤرشفة");
            else if (A <= 100)
            {
                for (int x = 0; x < A; x++)
                {

                    ViewArchShow(x, DocA[x], GriDateA[x], AppNameA[x], oldNewA[x], 1);
                }
            }
            else if (A > 100)
            {
                labelVA.Text = "جاري طباعة الملخص";
                CreateNotArchivedFiles(A, ReportName, GriDateA, DocA, AppNameA, "رقم المعاملة المرجعي", "غير المؤرشفة");
            }
            else// if (UserJobposition.Contains("قنصل"))
            {
                var selectedOption = MessageBox.Show("", "طباعة الملفات غير المؤرشفة", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                {
                    CreateNotArchivedFiles(A, ReportName, GriDateA, DocA, AppNameA, "رقم المعاملة المرجعي", "غير المؤرشفة");
                }
            }
            //labelVA.Text = "";
            //for (int x = 1; x < V; x++)
            //{
            //    ViewArchShow(x, DocV[x], IDV[x], AppNameV[x], oldNewV[x]);
            //}
            //if (V <= 1) 
            //    MessageBox.Show("لا توجد معاملات غير معالجة");
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void btnAuth_Click_1(object sender, EventArgs e)
        {
            uploadDocx = false;
            dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
            //MessageBox.Show(HijriDate);
            Form11 form11 = new Form11(attendedVC.SelectedIndex, -1, "", DataSource, DataSource56, FilespathIn, FilespathOut, EmployeeName, UserJobposition, GregorianDate, HijriDate);
            form11.ShowDialog();
            //dataSourceWrite(primeryLink + @"\updatingSetup.txt", "Not Allowed");
            //this.Hide();

        }

        private void btnSearch_Click_2(object sender, EventArgs e)
        {

        }

        private void date_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click_3(object sender, EventArgs e)
        {
            if (panelReceMess.Visible == false)
            {
                flowLayoutPanel1.Visible = panelReceMess.Visible = true;
                panel4.Visible = false;
                PanelMandounb.Visible = fileManagePanel2.Visible = panelAuthAknow.Visible = SearchPanel.Visible = ReportPanel.Visible = false;


            }
            else panelReceMess.Visible = false;
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            fileComboBox(txtHAAuthentic, DataSource, "AppName", "TableHandAuth", true);
            //if(!autoCompleteMode) 
                autoCompleteTextBox(txtHandAuthNo, DataSource, "DocID", "TableHandAuth");
            fillDataGrid("");
            combCount.SelectedIndex = 0;
            imagecount = 0;
            ColorFulGrid9();
            ArchivePic.Visible = false;

            dataGridView8.Visible = true;
            dataGridView8.BringToFront();
            if (panelAuthAknow.Visible == false)
            {
                flowLayoutPanel1.Visible = false;
                panel4.Visible = true;


                panelAuthAknow.Visible = true;

                PanelMandounb.Visible = flowLayoutPanel1.Visible = fileManagePanel2.Visible = panelReceMess.Visible = SearchPanel.Visible = ReportPanel.Visible = false;

            }
            else
            {

                panel4.Visible = false;
                panelAuthAknow.Visible = false;
                flowLayoutPanel1.Visible = true;
            }
        }

        private void btnMessSave_Click_1(object sender, EventArgs e)
        {
            int iD = 1;
            if (btnSaveArch.Text.Contains("تعديل")) iD = Messid;
            SubNameData(iD, txtMessNo.Text, txtMessApp.Text, txtMessAuth.Text, txtAuthNo.Text, checkMessAppSex.Text, txtMessSource.Text, txtMessGreDate.Text, ArchfilePath, GregorianDate + "_تمت معالجة الملف بواسطة" + EmployeeName);
            btnMessSave.Text = "حفظ وتأكيد";
            Messid = 1;
            ArchfilePath = "";
            txtMessNo.Text = txtMessApp.Text = txtMessAuth.Text = txtAuthNo.Text = checkMessAppSex.Text = txtMessSource.Text = txtMessGreDate.Text = "";
        }

        private void SubNameData(int id, string DocID, string AppName, string AuthName, string AuthNo, string Gender, string Institute, string GriDate, string filePath, string viewed)
        {
            //MessageBox.Show("filePath  " + filePath);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableReceMess ( DocID,AppName,AuthName,AuthNo,Gender,Institute,GriDate,Data1,Extension1,FileName1,ArchivedState,Viewed) values (@DocID,@AppName,@AuthName,@AuthNo,@Gender,@Institute,@GriDate,@Data1,@Extension1,@FileName1,@ArchivedState,@Viewed)", sqlCon);
            if (id != 1) sqlCmd = new SqlCommand("UPDATE TableReceMess SET   DocID=@DocID,AppName=@AppName,AuthName=@AuthName,AuthNo=@AuthNoGender=@Gender,Institute=@Institute,GriDate=@GriDate,Data1=@Data1,Extension1=@Extension1,FileName1=@FileName1,ArchivedState=@ArchivedState,Viewed=@Viewed where ID=@ID", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.Parameters.AddWithValue("@DocID", DocID);
            sqlCmd.Parameters.AddWithValue("@AppName", AppName);
            sqlCmd.Parameters.AddWithValue("@AuthName", AuthName);
            sqlCmd.Parameters.AddWithValue("@AuthNo", AuthNo);
            sqlCmd.Parameters.AddWithValue("@Gender", Gender);
            sqlCmd.Parameters.AddWithValue("@Institute", Institute);
            sqlCmd.Parameters.AddWithValue("@GriDate", GriDate);
            if (filePath == "")
            {
                filePath = ArchFile + "text1.txt";
                sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");
            }
            else
                sqlCmd.Parameters.AddWithValue("@ArchivedState", "مؤرشف");
            using (Stream stream = File.OpenRead(filePath))
            {
                byte[] buffer1 = new byte[stream.Length];
                stream.Read(buffer1, 0, buffer1.Length);
                var fileinfo1 = new FileInfo(filePath);
                string extn1 = fileinfo1.Extension;
                string DocName1 = fileinfo1.Name;
                sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
                sqlCmd.Parameters.Add("@FileName1", SqlDbType.NVarChar).Value = DocName1;
            }
            sqlCmd.Parameters.AddWithValue("@Viewed", viewed);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void SubAuthData(int id, string DocID, string AppName, string Gender, string Institute, string GriDate, string filePath, string viewed, string docCount, string comment)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableHandAuth ( DocID,AppName,Gender,Institute,GriDate,Data1,Extension1,FileName1,ArchivedState,Viewed,HandTime,DocNo,Comment,AVConsule) values (@DocID,@AppName,@Gender,@Institute,@GriDate,@Data1,@Extension1,@FileName1,@ArchivedState,@Viewed,@HandTime,@DocNo,@Comment,@AVConsule)", sqlCon);
            if (id != 1) sqlCmd = new SqlCommand("UPDATE TableHandAuth SET  DocID=@DocID,AppName=@AppName,Gender=@Gender,Institute=@Institute,GriDate=@GriDate,Data1=@Data1,Extension1=@Extension1,FileName1=@FileName1,ArchivedState=@ArchivedState,Viewed=@Viewed,HandTime=@HandTime,DocNo=@DocNo,Comment=@Comment,AVConsule=@AVConsule where ID=@ID", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.Parameters.AddWithValue("@DocID", DocID);
            sqlCmd.Parameters.AddWithValue("@AppName", AppName);
            sqlCmd.Parameters.AddWithValue("@Gender", Gender);
            sqlCmd.Parameters.AddWithValue("@Institute", comboBox2.Text);
            sqlCmd.Parameters.AddWithValue("@HandTime", txtHAGredate.Text);
            sqlCmd.Parameters.AddWithValue("@DocNo", docCount);
            sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate);
            sqlCmd.Parameters.AddWithValue("@Comment", comment);
            sqlCmd.Parameters.AddWithValue("@AVConsule", attendedVC.Text);

            if (filePath == "")
            {
                filePath = ArchFile + "text1.txt";
                sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");
            }
            else
                sqlCmd.Parameters.AddWithValue("@ArchivedState", "مؤرشف");

            if (filePath == "") { MessageBox.Show("لا يوجد رابط بيانات صالح"); return; }
            using (Stream stream = File.OpenRead(filePath))
            {
                byte[] buffer1 = new byte[stream.Length];
                stream.Read(buffer1, 0, buffer1.Length);
                var fileinfo1 = new FileInfo(filePath);
                string extn1 = fileinfo1.Extension;
                string DocName1 = fileinfo1.Name;
                sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
                sqlCmd.Parameters.Add("@FileName1", SqlDbType.NVarChar).Value = DocName1;
            }
            sqlCmd.Parameters.AddWithValue("@Viewed", viewed);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }


        private void checkMessAppSex_CheckedChanged(object sender, EventArgs e)
        {
            if (checkMessAppSex.CheckState == CheckState.Unchecked)
            {
                checkMessAppSex.Text = "ذكر";
            }
            else if (checkMessAppSex.CheckState == CheckState.Checked)
            {
                checkMessAppSex.Text = "إنثى";
            }

        }

        private void btnMessArch_Click_1(object sender, EventArgs e)
        {
            if (btnMessArch.Text == "تحميل ملف ارشفة المعاملة")
            {
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.ShowDialog();
                ArchfilePath = dlg.FileName;
            }
            else
            {
                OpenMessFile(Messid, "TableReceMess");
            }
        }

        private void txtMessNo_TextChanged_1(object sender, EventArgs e)
        {
            FillDataGridView(txtMessNo.Text);
        }

        private void btnMessView_Click(object sender, EventArgs e)
        {

        }

        private void Arch2_Click_2(object sender, EventArgs e)
        {
            //OpenFile(IDNo, 2);
            FillDatafromGenArch(txtSearch.Text, "data2",Arch2);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            ArchfilePath = dlg.FileName;


        }

        private void ArchfilePathEdit(string archfile)
        {
            if (archfile != "")
            {
                SqlConnection sqlCon = new SqlConnection(DataSource);
                if (sqlCon.State == ConnectionState.Closed)
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return; }
                SqlCommand sqlCmd = new SqlCommand("UPDATE TableReceMess SET Data1=@Data1,Extension1=@Extension1,FileName1=@FileName1 WHERE ID=@ID", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@ID", Messid);
                sqlCmd.Parameters.AddWithValue("@ArchivedState", "مؤرشف بتاريخ " + GregorianDate);
                using (Stream stream = File.OpenRead(archfile))
                {
                    byte[] buffer1 = new byte[stream.Length];
                    stream.Read(buffer1, 0, buffer1.Length);
                    var fileinfo1 = new FileInfo(filePath);
                    string extn1 = fileinfo1.Extension;
                    string DocName1 = fileinfo1.Name;
                    sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                    sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
                    sqlCmd.Parameters.Add("@FileName1", SqlDbType.NVarChar).Value = DocName1;
                }

                if (btnMessSave.Text == "حفظ")
                {
                    sqlCmd.Parameters.AddWithValue("@Viewed", GregorianDate);

                }
                else if (btnMessSave.Text == "حفظ وتأكيد")
                { sqlCmd.Parameters.AddWithValue("@Viewed", GregorianDate + "_تمت معالجة الملف بواسطة" + EmployeeName); }



                sqlCmd.ExecuteNonQuery();
                sqlCon.Close();
            }
        }



        private void btnHAArch_Click(object sender, EventArgs e)
        {

        }

        private void btnHAView_Click(object sender, EventArgs e)
        {

        }

        private void txtHandAuthNo_TextChanged_1(object sender, EventArgs e)
        {

            OpenReportArch(Messid.ToString());
        }

        private void button22_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            ArchfilePath = dlg.FileName;
        }

        private void btnHandArch_Click_1(object sender, EventArgs e)
        {
            if (btnMessArch.Text == "تحميل ملف ارشفة المعاملة")
            {
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.ShowDialog();
                ArchfilePath = dlg.FileName;
            }
            else
            {
                OpenMessFile(Messid, "TableHandAuth");
            }
        }

        private void checkHASex_CheckedChanged(object sender, EventArgs e)
        {

        }


        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void processed_TextChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click_1(object sender, EventArgs e)
        {

        }


        private void PrintMessage_Click_1(object sender, EventArgs e)
        {

            CreateMessageWord(applicant.Text, txtEmbassey.Text, txtSearch.Text, strMessageType, bolApplicantSex, date.Text, HijriDate, attendedVC.Text);
            PrintMessage.Visible = false;
            DetecedForm.Width = 393;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            checkYear(DataSource);

            autoCompleteTextBox(applicant, DataSource, "الاسم", "TableGeneralArch");
            fileComboBox(AttendViceConsul, DataSource, "ArabicAttendVC", "TableListCombo", true);
            fileComboBox(perbtn1, DataSource, "AuthTypes", "TableListCombo", true);
            fileComboBox(attendedVC, DataSource, "ArabicAttendVC", "TableListCombo", true);
            if (attendedVC.Items.Count >= VCIndexData()) attendedVC.SelectedIndex = VCIndexData();
            if (AttendViceConsul.Items.Count >= VCIndexData()) AttendViceConsul.SelectedIndex = VCIndexData();
            fileComboBox(docCollectCombo, DataSource, "ArabicGenIgrar", "TableListCombo", true);
            fileComboBox(docCollectCombo, DataSource, "EnglishGenIgrar", "TableListCombo", false);
            
            VCIndexLoad = true; loadScanner();
            fillDataGrid("");
            updataArchData1();
            
        }
        private void autoCompleteTextBox(TextBox textbox, string source, string comlumnName, string tableName)
        {

            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                DataTable Textboxtable = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(Textboxtable);
                AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();
                bool newSrt = true;
                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    autoComplete.Add(dataRow[comlumnName].ToString());                    
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
                autoCompleteMode = true;
            }
        }
        private void loadScanner()
        {
            try
            {
                var deviceManager = new DeviceManager();



                for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++) // Loop Through the get List Of Devices.
                {
                    if (deviceManager.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType) // Skip device If it is not a scanner
                    {
                        continue;
                    }
                    AvailableScanner = deviceManager.DeviceInfos[i];
                    break;
                }


            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName, bool clear)
        {
            
            if(clear) combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;


                try
                {
                    cmd.ExecuteNonQuery();
                    DataTable table = new DataTable();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                    dataAdapter.Fill(table);

                    foreach (DataRow dataRow in table.Rows)
                    {
                        if (dataRow[comlumnName].ToString() != "")
                        {
                            bool found = false;
                            for (int x = 0; x < combbox.Items.Count; x++)
                            {
                                if (combbox.Items[x].ToString() == dataRow[comlumnName].ToString()) found = true;
                            }
                            if (!found) combbox.Items.Add(dataRow[comlumnName].ToString());
                        }
                    }
                }
                catch (Exception ex) { }
                saConn.Close();
            }
        }

        private void checkYear(string source)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            //MessageBox.Show(DateTime.Now.Year.ToString());
            string year = DateTime.Now.Year.ToString();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();
                string query = "select CurrentYear from TableSettings where ID = 1";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                foreach (DataRow dataRow in table.Rows)
                {
                    if (dataRow["CurrentYear"].ToString() != year)
                    {
                        btnNewYear.Enabled = true;
                        MessageBox.Show(" تشير الساعة إلى حلول عام ميلادي جديد، يرجى التواصل مع رئيس القسم لإعادة تصفير جميع المعاملات");
                        flowLayoutPanel3.BringToFront();
                    }
                }
                saConn.Close();
            }
        }

        private void comboBoxAuthValue_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] str = new string[perbtn1.Items.Count];
            for (int x = 0; x < perbtn1.Items.Count; x++) 
            {
                str[x] = perbtn1.Items[x].ToString(); 
            }
            string[] strSub = new string[1] { "" };
            dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
            FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition,DataSource, perbtn1.SelectedIndex, FormDataFile, FilespathOut, 12, str, strSub, true,MandoubM, GriDateM);
                form2.ShowDialog();
            
            
        }

        private void picremove_Click_1(object sender, EventArgs e)
        {

            loadSettings(DataSource, false, true, false, false);

        }

        private void picaddmonth_Click(object sender, EventArgs e)
        {

            loadSettings(DataSource, false, false, true, true);

        }

        private void mangerArch_CheckedChanged(object sender, EventArgs e)
        {
            if (mangerArch.CheckState == CheckState.Checked)
            {
                mangerArch.Text = "توزيع المهام";
                persbtn2.Visible = persbtn10.Visible = false;
                perbtn1.Visible = docCollectCombo.Visible = true;
                docCollectCombo.BringToFront();
            }
            else
            {
                docCollectCombo.SendToBack();
                mangerArch.Text = "ادخال البيانات";
                persbtn2.Visible = persbtn10.Visible = true;
                perbtn1.Visible = docCollectCombo.Visible = false;
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            StaredColumns();
            //StaredColumns1();
            //var selectedOption = MessageBox.Show("", "تحميل الاحصائيات؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            //if (selectedOption == DialogResult.Yes)
            //{
            //    StatisticInfo();
            //}
            var selectedOption = MessageBox.Show("", "قوائم المعاملات؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                StaredColumns1();
            }
        }

        private void merriageTable()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            button23.Enabled = false;

            string col = "Col0";


            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            for (rCnt = 1; rCnt < 67; rCnt++)
            {
                string strDate = (string)(range.Cells[rCnt, 1] as Excel.Range).Value2;
                int strMarriage = (int)(range.Cells[rCnt, 2] as Excel.Range).Value2;
                int strDivorce = (int)(range.Cells[rCnt, 3] as Excel.Range).Value2;
                //MessageBox.Show(strDate + " - " + strMarriage.ToString() + " - " + strDivorce.ToString());
                for (int x = 0; x < strMarriage; x++)
                {
                    UpdateMaririageColumn(DataSource, "زواج", strDate);

                }
                for (int x = 0; x < strDivorce; x++)
                {
                    UpdateMaririageColumn(DataSource, "طلاق", strDate);
                }
            }

            sqlCon.Close();
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            button23.Enabled = true;


        }
        private void CreateColumn(string Columnname)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("alter table TableAuthRight add " + Columnname.Replace(" ", "_") + " nvarchar(1000)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        
        private void CreateColumns(string Columnname)
        {
            string query = "alter table TableAuthRights add " + Columnname + " nvarchar(1000)";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { MessageBox.Show("query " + query + "DataSource " + DataSource); return; }
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            //MessageBox.Show(Columnname);
            try
            {
                sqlCmd.ExecuteNonQuery();
                //MessageBox.Show(Columnname);
            }
            catch (Exception ex) {
                // MessageBox.Show("query " + query + "DataSource " + DataSource);
            }
            sqlCon.Close();
        }
        private bool checkColumnName(string colNo)
        {
            //MessageBox.Show(dataSource);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return false ; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableAuthRight", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(dataRow["COLUMN_NAME"].ToString()))
                {
                    //MessageBox.Show(dataRow["COLUMN_NAME"].ToString());
                    if (dataRow["COLUMN_NAME"].ToString() == colNo.Replace(" ", "_"))
                    {
                        return true;
                    }
                }
            }
            //MessageBox.Show(colNo + "not found");
            return false;
        }
        private bool checkColumnNames(string colNo, string id)
        {
            
            string query = "select " + colNo + " from TableAuthRights";
            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(id.ToString() + " - " + colNo + "not found");
                return false;
            }
            
            
            
            sqlCon.Close();
            
            foreach (DataRow dataRow in dtbl.Rows)
            {
                try
                {
                    //Console.WriteLine("dataRow " + dataRow[colNo].ToString().TrimEnd().TrimStart() + " == colNo" + colNo);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(id.ToString() + " - "+colNo + "not found");
                    return false;
                }
            }
            //else MessageBox.Show(colNo + "found");
            return true;
        }
        
        private bool checkID(string id)
        {
            //MessageBox.Show(id);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("select ID from TableAuthRights", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["ID"].ToString() == id)
                {
                    //MessageBox.Show(dataRow["ID"].ToString());
                    return true;
                }
            }
            //MessageBox.Show(id + " not found");
            return false;
        }
        
        private void ColumnNamesLoad()
        {
            bool found = false;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableAuthRights", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            int colIndex = 0;
            rightColNames = new string[dtbl.Rows.Count - 1];
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["COLUMN_NAME"].ToString() != "" && dataRow["COLUMN_NAME"].ToString() != "ID")
                {
                    rightColNames[colIndex] = dataRow["COLUMN_NAME"].ToString();
                    colIndex++;
                }
            }
        }
        private void button23_Click(object sender, EventArgs e)
        {
            //oldFun();
            //

            string[] colName = getColName();
            for (int col = 0; col < colName.Length; col++)
            {
                //MessageBox.Show(colName[col]);
                if (!checkColumnNames(colName[col].Replace("-", "_"), IDList[col]))
                {
                    //MessageBox.Show("colName " +colName[col]);
                    CreateColumns(colName[col].Replace("-", "_"));
                    if (checkID("1"))
                        UpdateColumn(DataSource, colName[col].Replace("-","_"), 1, colName[col], "TableAuthRights");
                    else InsertColumn(DataSource, colName[col].Replace("-", "_"), 1, colName[col], "TableAuthRights");
                }
            }
            newFun();
        }

        private void newFun() {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            button23.Enabled = false;
            ColumnNamesLoad();

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            for (cCnt = 2; cCnt <= cl; cCnt++)
            {
                Console.WriteLine("rightColNames " + rightColNames.Length.ToString() + " cCnt " + cCnt.ToString());
                string cols = "";
                try
                {
                    string colname = (string)(range.Cells[1, cCnt] as Excel.Range).Value2;

                    if (string.IsNullOrEmpty(colname)) continue;

                     cols = colname.Replace(" ", "_").Replace("-", "_");
                    if (!checkColumnNames(cols, ""))
                    {
                        //MessageBox.Show(colname + "-"+colname);
                        CreateColumns(cols);
                        if (checkID("1"))
                            UpdateColumn(DataSource, cols, 1, cols, "TableAuthRights");
                        else InsertColumn(DataSource, cols, 1, cols, "TableAuthRights");
                    }
                    //MessageBox.Show(colname);
                    for (rCnt = 2; rCnt < rw; rCnt++)
                    {
                        try
                        {
                            string strData = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                            if (String.IsNullOrEmpty(strData)) strData = "";
                            //MessageBox.Show(rCnt +"_" +strData);
                            if (checkID(rCnt.ToString()))
                                UpdateColumn(DataSource, cols, rCnt, strData, "TableAuthRights");
                            else InsertColumn(DataSource, cols, rCnt, strData, "TableAuthRights");
                        }
                        catch (Exception ex)
                        {
                            //for (cCnt = 2; cCnt <= rightColNames.Length; cCnt++)
                            //{
                            //    col = rightColNames[cCnt - 2];

                            //    for (rCnt = 1; rCnt < 24; rCnt++)
                            //    {
                            //        try
                            //        {
                            //            string strData = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                            //            if (String.IsNullOrEmpty(strData)) strData = "";
                            //            UpdateColumn(DataSource, col, rCnt, strData, "TableAuthRights");
                            //        }
                            //        catch (Exception e1x)
                            //        {
                            //        }
                            //        //MessageBox.Show(rCnt.ToString());
                            //    }
                            //}
                        }
                        //MessageBox.Show(rCnt.ToString());
                    }
                }
                catch (Exception ex) { MessageBox.Show(cols); }
            }

            sqlCon.Close();
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            button23.Enabled = true;
        }
        private string[] getColName()
        {
            string[] colName = new string[1];
            SqlConnection sqlCon = new SqlConnection(DataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return colName; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("select ID, ColName from TableAddContext where ColRight <> '' and ColName is not null", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            colName = new string[dtbl.Rows.Count];
            int index = 0;
            foreach (DataRow row in dtbl.Rows)
            {
                colName[index] = row["ColName"].ToString().Replace("-", "_").TrimEnd().TrimStart();
                colName[index] = colName[index].Replace(" ", "_");
                IDList[index] = row["ID"].ToString();
                Console.WriteLine("colName["+ index.ToString()+"] " + colName[index]);
                index++;
            }
            return colName;
        }
        private void oldFun() {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            button23.Enabled = false;
            string col = "Col0";
            for (cCnt = 1; cCnt <= cl; cCnt++)
            {
                col = "Col" + (cCnt - 1).ToString();
                if (!checkColumnName(col))
                    CreateColumn(col);
            }



            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            for (cCnt = 1; cCnt <= cl; cCnt++)
            {
                col = "Col" + (cCnt - 1).ToString();


                for (rCnt = 1; rCnt < 24; rCnt++)
                {
                    string strData = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    if (String.IsNullOrEmpty(strData)) strData = "";
                    UpdateColumn(DataSource, col, rCnt, strData, "TableAuthRight");
                    //MessageBox.Show(rCnt.ToString());
                }
            }

            sqlCon.Close();
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            button23.Enabled = true;
        }
        private void pictremovemonth_Click(object sender, EventArgs e)
        {

            loadSettings(DataSource, false, false, false, true);

        }

        private void button29_Click(object sender, EventArgs e)
        {
            string[] str = new string[1] { "" };
            string[] strSub = new string[1] { "" };
            dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
            FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition,DataSource, perbtn1.SelectedIndex, FormDataFile, FilespathOut, 12, str, strSub, false,MandoubM, GriDateM);
            form2.ShowDialog();
            flowLayoutPanel1.Visible = true;
            panel4.Visible = false;
        }

        private void picadd_Click_1(object sender, EventArgs e)
        {

            loadSettings(DataSource, true, true, false, false);

        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (!VCIndexLoad) return;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableSettings SET VCIndesx = @VCIndesx WHERE ID = @ID", sqlCon);
            if (!Pers_Peope)
            {
                sqlCmd = new SqlCommand("UPDATE TableSettings SET AttendVCAffairs = @AttendVCAffairs WHERE ID = @ID", sqlCon);
            }
             
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", 1);
            if (Pers_Peope) 
                sqlCmd.Parameters.AddWithValue("@VCIndesx", attendedVC.SelectedIndex.ToString());
            else
                sqlCmd.Parameters.AddWithValue("@AttendVCAffairs", attendedVC.SelectedIndex.ToString());            
                sqlCmd.ExecuteNonQuery();
        }

        private int VCIndexData()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return 0; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT VCIndesx,AttendVCAffairs FROM TableSettings", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            int index = 0;
            foreach (DataRow dataRow in table.Rows)
            {
                if (Pers_Peope)
                {
                    if (!string.IsNullOrEmpty(dataRow["VCIndesx"].ToString()))
                    {
                        index = Convert.ToInt32(dataRow["VCIndesx"].ToString());
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(dataRow["AttendVCAffairs"].ToString()))
                    {
                        index = Convert.ToInt32(dataRow["AttendVCAffairs"].ToString());
                    }
                }
            }
            return index;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            
            try

            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.

                    var ScanerItem = device.Items[1]; // select the scanner.


                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);
                    PathImage = PrimariFiles + "ScanImg" + DateTime.Now.ToString("mmss") + ".jpg";
                    if (File.Exists(PathImage))
                    {
                        File.Delete(PathImage);
                    }
                    imgFile.SaveFile(PathImage);
                    ArchivePic.ImageLocation = PathImage;
                    ArchivePic.Visible = true;
                   // Console.WriteLine(PathImage);
                    org = new PictureBox();
                    org.Load(PathImage);

                    
                }
                else
                {

                    MessageBox.Show("لا يوجد جهاز ماسح متصل");
                }

            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void reTackPic_Click(object sender, EventArgs e)
        {
            
            try

            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.

                    var ScanerItem = device.Items[1]; // select the scanner.


                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);
                    PathImage = @"D:\ScanImg.jpg";


                    if (File.Exists(PathImage))
                    {
                        File.Delete(PathImage);
                    }
                    imgFile.SaveFile(PathImage);
                    ArchivePic.ImageLocation = PathImage;
                    
                    org = new PictureBox();
                    org.Load(PathImage);
                }
                else
                {

                    MessageBox.Show("لا يوجد جهاز ماسح متصل");
                }

            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void MainForm_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Visible = true;
            panel4.Visible = false;
        }

     

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void fillMandoubGrid()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource56);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT ID,MandoubNames,MandoubAreas,MandoubPhones FROM TableMandoudList", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            dataGridView10.DataSource = table;
            if (dataGridView10.Rows.Count > 1)
            {
                dataGridView10.Columns[0].Visible = false;
                dataGridView10.Columns[1].Width = 180;
                dataGridView10.Columns[2].Width = 80;
                dataGridView10.Columns[3].Width = 90;
            }
        }

        private void fillDataGrid(string text)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            
            try
            {if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT ID,AppName,DocID,Institute,GriDate,Viewed,HandTime,DocNo,Comment,AVConsule FROM TableHandAuth", sqlCon);
                if (text != "")
                    sqlDa = new SqlDataAdapter("SELECT ID,AppName,DocID,Institute,GriDate,Viewed,HandTime,DocNo,Comment,AVConsule  FROM TableHandAuth where Institute=@Institute", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                sqlDa.SelectCommand.Parameters.AddWithValue("@Institute", text);
                DataTable table = new DataTable();
                sqlDa.Fill(table);
                sqlCon.Close();
                dataGridView8.DataSource = table;
            }
            catch (Exception ex) { return; }
            
            if (dataGridView8.Rows.Count > 1) {
                handIndex = 0;
                
                dataGridView8.Columns[0].Visible = false;
                
                Messid = Convert.ToInt32(dataGridView8.Rows[handIndex].Cells[0].Value.ToString());                
            }
        }

        private void fillDataGridReports()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            
            try
            {if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT ID,التاريخ,الموظف,رقم_معاملة_القسم as موضوع_التقرير FROM TableGeneralArch where نوع_المستند=@نوع_المستند and الموظف=@الموظف", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                sqlDa.SelectCommand.Parameters.AddWithValue("@نوع_المستند", "تقرير");
                //if (UserJobposition.Contains("قنصل"))
                //    sqlDa.SelectCommand.Parameters.AddWithValue("@الموظف", "");
                //else 
                    sqlDa.SelectCommand.Parameters.AddWithValue("@الموظف", ConsulateEmployee.Text);
                DataTable table = new DataTable();
                sqlDa.Fill(table);
                sqlCon.Close();
                dataGridView11.DataSource = table;
            }
            
            catch (Exception ex) { return; }
            dataGridView11.BringToFront();
            dataGridView11.Visible = true;
            
        }

        private void ColorFulGrid9()
        {
            for (int i = 0; i < dataGridView8.Rows.Count - 1; i++)
            {
                dataGridView8.Rows[i].DefaultCellStyle.BackColor = Color.White;
                if (dataGridView8.Rows[i].Cells[3].Value.ToString().Contains("سودان"))
                {
                    dataGridView8.Rows[i].DefaultCellStyle.BackColor = Color.LightBlue;
                    countTimer++;
                    if (countTimer == 100) countTimer = 0;
                }

                if (dataGridView8.Rows[i].Cells[8].Value.ToString().Contains("مزورة"))
                {
                    dataGridView8.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                }
            }
        }


        private void OpenReportArch(string refNum)
        {
            string query;

            SqlConnection Con = new SqlConnection(DataSource);
            query = "select ID,AppName,DocID,Gender,Institute,GriDate,Data1,Extension1,FileName1,Viewed,HandTime,AVConsule from TableHandAuth where AppName=@ref";
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@ref", SqlDbType.NVarChar).Value = refNum;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                Messid = Convert.ToInt32(reader["ID"].ToString());
                string name = reader["FileName1"].ToString();
                var Data = (byte[])reader["Data1"];
                var ext = reader["Extension1"].ToString();
                NewFileNamePic = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                if(textNumber)
                txtHAAuthentic.Text = reader["AppName"].ToString();
                txtHandAuthNo.Text = reader["DocID"].ToString();
                
                if (reader["Gender"].ToString() == "ذكر")
                {
                    checkHASex.Text = "ذكر";
                    checkHASex.CheckState = CheckState.Unchecked;
                }
                else
                {
                    checkHASex.Text = "أنثى";
                    checkHASex.CheckState = CheckState.Checked;
                }
                comboBox2.Text = reader["Institute"].ToString();
                txtHAGredate.Text = reader["HandTime"].ToString();
                HandProcess.Text = reader["Viewed"].ToString();
                if (btnSaveArch.Text == "حفظ")
                {
                    btnSaveArch.Text = "تعديل";
                }
                else if (btnSaveArch.Text == "حفظ وتأكيد")
                    btnSaveArch.Text = "تعديل وتأكيد";
                
                File.WriteAllBytes(NewFileNamePic, Data);
                panel4.Visible = true;

                ArchivePic.ImageLocation = NewFileNamePic;
                org = new PictureBox();
                org.Load(NewFileNamePic);
                ArchivePic.SizeMode = PictureBoxSizeMode.Zoom;
                
                //ArchivePic.Image = ZoomPicture(ArchivePic.Image, new Size(trackBar1.Value, trackBar1.Value));
            }
            Con.Close();

        }

        private void OpenReportArchID(int refNum)
        {
            string query;

            SqlConnection Con = new SqlConnection(DataSource);
            query = "select AppName,DocID,Gender,Institute,GriDate,Data1,Extension1,FileName1,Viewed,HandTime,DocNo,AVConsule from TableHandAuth where ID=@ref";
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@ref", SqlDbType.Int).Value = refNum;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                
                string name = reader["FileName1"].ToString();
                var Data = (byte[])reader["Data1"];
                var ext = reader["Extension1"].ToString();
                PathImage = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                txtHAAuthentic.Text = reader["AppName"].ToString();
                txtHandAuthNo.Text = reader["DocID"].ToString();
                combCount.Text = reader["DocNo"].ToString();
                if (reader["Gender"].ToString() == "ذكر")
                {
                    checkHASex.Text = "ذكر";
                    checkHASex.CheckState = CheckState.Unchecked;
                }
                else
                {
                    checkHASex.Text = "أنثى";
                    checkHASex.CheckState = CheckState.Checked;
                }
                comboBox2.Text = reader["Institute"].ToString();
                txtHAGredate.Text = reader["HandTime"].ToString();
                HandProcess.Text = reader["Viewed"].ToString();
                if (btnSaveArch.Text == "حفظ")
                {
                    btnSaveArch.Text = "تعديل";
                }
                else if (btnSaveArch.Text == "حفظ وتأكيد")
                    btnSaveArch.Text = "تعديل وتأكيد";

                File.WriteAllBytes(PathImage, Data);
                panel4.Visible = true;

                if (ext != ".docx")
                {
                    Process.Start(PathImage);
                    //ArchivePic.ImageLocation = PathImage;
                    //org = new PictureBox();
                    //org.Load(PathImage);
                    //ArchivePic.SizeMode = PictureBoxSizeMode.Zoom;

                }
                else {
                    Process.Start("WINWORD.EXE", PathImage);
                }
                //ArchivePic.Image = ZoomPicture(ArchivePic.Image, new Size(trackBar1.Value, trackBar1.Value));
            }
            Con.Close();

        }

        Image ZoomPicture(Image img, Size size)
        {
            Bitmap bm = new Bitmap(img, Convert.ToInt32(img.Width + (img.Width * size.Width / 100)), Convert.ToInt32(img.Height + (img.Height * size.Height / 100)));
            Graphics gpu = Graphics.FromImage(bm);
            gpu.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            return bm;

        }

        private void ArchivePic_Click(object sender, EventArgs e)
        {
            if (ArchivePic.SizeMode == PictureBoxSizeMode.Zoom) {
                ArchivePic.SizeMode = PictureBoxSizeMode.AutoSize;
            }else ArchivePic.SizeMode = PictureBoxSizeMode.Zoom;
        }

       
        private void button16_Click(object sender, EventArgs e)
        {
            }

        private void date_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void DeleteEmptyFiles(string tableStr)
        {
            //TableAddContext
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * FROM " + tableStr, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();

            foreach (DataRow dataRow in table.Rows)
            {
                if(dataRow["label1"].ToString() == "" && dataRow["label2"].ToString() == "")
                deleteRowsData(Convert.ToInt32(dataRow["ID"].ToString()), tableStr, DataSource);
            }
        }

        private void btnDeleteHand_Click(object sender, EventArgs e)
        {
            var selectedOption = MessageBox.Show("", "حذف نموذج التوقيع", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                if (dataGridView8.Rows.Count > 1)                
                    deleteRowsData(Messid, "TableHandAuth", DataSource);
            }
            fillDataGrid("");
            ColorFulGrid9();
            
            ArchivePic.Visible = false;
            Messid = 1;
            btnSaveArch.Text = "حفظ وتأكيد";
        }

        private void deleteRowsData(int v1, string v2, string source)
        {
            string query;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "DELETE FROM " + v2 + " where ID = @ID";
            Console.WriteLine(query);
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", v1);
            try
            {
                sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex) { }
            Con.Close();
            fillDataGrid(comboBox2.Text);
        }


        
        private void deleteGenArch(string v1, string v2)
        {
            string query;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "DELETE FROM TableGeneralArch where رقم_المرجع = @رقم_المرجع and docTable=@docTable";
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_المرجع", v1);
            sqlCmd.Parameters.AddWithValue("@docTable", v2);
            sqlCmd.ExecuteNonQuery();
            Con.Close();
           Console.WriteLine("deleted files no " + v1 +" - "+v2);
        }

        private void txtHAGredate_TextChanged(object sender, EventArgs e)
        {
            if (txtHAGredate.Text.Length == 2) txtHAGredate.Text = "-" + txtHAGredate.Text;
            else if (txtHAGredate.Text.Length == 5) txtHAGredate.Text = "-" + txtHAGredate.Text;
        }

        private void txtHAAuthenticic_TextChanged(object sender, EventArgs e)
        {
            if (textNumber) return;
            OpenReportArch(Messid.ToString());
        }

        private void dataGridView8_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView8.Rows.Count > 1)
            {
                textNumber = true;
                Messid = Convert.ToInt32(dataGridView8.CurrentRow.Cells[0].Value.ToString());
                
                OpenReportArchID(Messid);
                btnSaveArch.Text = "تعديل وتأكيد";
                //dataGridView8.Visible = false;
                
            }
        }

        private string getFileNo(int id)
        {
            string str = "1";
            using (SqlConnection saConn = new SqlConnection(DataSource))
            {
                saConn.Open();

                string query = "select رقم_ملف_جدة,رقم_ملف_مكة,رقم_ملف_اللجنة,رقم_ملف_الوافدين,عدد_الأفراد,عدد_الأفراد_مكة,عدد_الأفراد_الوافدين,عدد_الأفراد_اللجنة,رقم_ملف_المقابل,عدد_الأفراد_المقابل from TableSettings";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;


                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                foreach (DataRow dataRow in table.Rows)
                {
                    if (!string.IsNullOrEmpty(dataRow["رقم_ملف_جدة"].ToString()))
                    {
                        switch (id)
                        {
                            case 0:
                                if (dataRow["رقم_ملف_جدة"].ToString() != "")
                                    str = dataRow["رقم_ملف_جدة"].ToString();
                                break;
                            case 1:
                                if (dataRow["رقم_ملف_مكة"].ToString() != "")
                                    str = dataRow["رقم_ملف_مكة"].ToString();
                                break;

                            case 2:
                                if (dataRow["رقم_ملف_الوافدين"].ToString() != "")
                                    str = dataRow["رقم_ملف_الوافدين"].ToString();
                                break;

                            case 3:
                                if (dataRow["رقم_ملف_اللجنة"].ToString() != "")
                                    str = dataRow["رقم_ملف_اللجنة"].ToString();
                                break;
                            case 4:
                                if (dataRow["رقم_ملف_المقابل"].ToString() != "")
                                    str = dataRow["رقم_ملف_المقابل"].ToString();
                                break;
                            case 5:
                                if (dataRow["عدد_الأفراد"].ToString() != "")
                                    str = dataRow["عدد_الأفراد"].ToString();
                                break;
                            case 6:
                                if (dataRow["عدد_الأفراد_مكة"].ToString() != "")
                                    str = dataRow["عدد_الأفراد_مكة"].ToString();
                                break;
                            case 7:
                                if (dataRow["عدد_الأفراد_الوافدين"].ToString() != "")
                                    str = dataRow["عدد_الأفراد_الوافدين"].ToString();
                                break;
                            case 8:
                                if (dataRow["عدد_الأفراد_اللجنة"].ToString() != "")
                                    str = dataRow["عدد_الأفراد_اللجنة"].ToString();
                                break;
                            
                            case 9:
                                if (dataRow["عدد_الأفراد_المقابل"].ToString() != "")
                                    str = dataRow["عدد_الأفراد_المقابل"].ToString();
                                break;
                        }
                    }
                }
                saConn.Close();
            }
            return str;
        }


        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            labelM.Text = "";
            foreach (Control control in flowLayoutPanel1.Controls)
            {
                if (control is Button) control.Visible = false;
            }
            for (int x = 0; x < M; x++)
            {
                string strOldNew = "";
                if (oldNewM[x] == "old") strOldNew = "-نسخة معدلة-";

                ViewArchShow(x, DocM[x], GriDateM[x], AppNameM[x] + " عن طريق " + MandoubM[x], strOldNew, 2);

            }
            if (M <= 0) 
                MessageBox.Show("لا توجد معاملات مناديب غير مؤرشفة");
            else
            {
                var selectedOption = MessageBox.Show("", "طباعة الملفات غير المؤرشفة", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                {
                    string ReportName = "Report" + DateTime.Now.ToString("mmss") + ".docx";
                    CreateMandounbFiles(M, ReportName, GriDateM, DocIDM, AppNameM, MandoubM);                 
                }
            }
        }

        private void dataGridView8_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DetecedForm_Click_1(object sender, EventArgs e)
        {
            
            GoToForm(TableIndex - 1, IDNo);
            ClearFileds();
        }

        private void picVersio_Click(object sender, EventArgs e)
        {
            string currentVersion = getVersio();
            //text 1.0.0.336.O------ server = 1.0.0.333.F
            string str = currentVersion.Split('.')[0] + "." + currentVersion.Split('.')[1] + "." + currentVersion.Split('.')[2] + "." + (Convert.ToInt32(currentVersion.Split('.')[3]) + 1).ToString();

            var selectedOption = MessageBox.Show("", "تحديث إجباري؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                VersionUpdate(str + ".F");
                
            }
            else
            {
                VersionUpdate(str + ".O");
            }
            timer4Update();
        }

        private void upDateClose()
        {
            string version = getVersio();
            try
            {
                File.Delete(primeryLink + "fileUpdate.txt");
                System.Diagnostics.Process.Start(getAppFolder() + @"\setup.exe");
                if(Server == "57")
                    dataSourceWrite(primeryLink + @"\Personnel\getVersio.txt", version);
                else if (Server == "56") 
                    dataSourceWrite(primeryLink + @"\SuddaneseAffairs\getVersio.txt", version);

                dataSourceWrite(primeryLink + @"\updatingSetup.txt", "updating");                
            }
            catch (Exception ex) {
                //MessageBox.Show("close");
            }
        }
        private void timer4Update() {
            int CV = 0;
            int cV = 0;
            string updateType = "O";
            //if (onUpdate) return;
            string currentVersion = getVersio();
            try
            {
                CV = Convert.ToInt32(CurrentVersion.Split('.')[3]);

                cV = Convert.ToInt32(currentVersion.Split('.')[3]);
                updateType = currentVersion.Split('.')[4];
                if (CV < cV && UserJobposition.Contains("قنصل"))
                {

                    empUpdate.Visible = false;
                }
                else
                {

                    empUpdate.Visible = true;
                }
            }

            catch (Exception ex) { return; }
            
            if (CV >= cV || updateType != "F")
                return;
            else if (CV < cV && updateType == "F" && !onUpdate)
            {
                Console.WriteLine(primeryLink + @"\updatingStatus.txt");
                string status = File.ReadAllText(primeryLink + @"\updatingStatus.txt");
                if (status == "Allowed")
                {
                    onUpdate = true;
                    upDateClose();
                }
            }
        }
        private void timer4_Tick(object sender, EventArgs e)
        {
            UserLogOut();
            if (deleteEmptyRows)
            {
                Console.WriteLine("deleteEmptyRows");
                {
                    DeleteEmptyFiles(parrtialAll);
                }
            }

        }

        private void DeleteEmptyFiles(bool partial)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string[] DocumentID = new string[40];
            string year = DateTime.Now.Year.ToString().Replace("20", "");

            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            for (TableIndex = 1; TableIndex <= 11; TableIndex++)
                {
                    if (TableIndex == 10) continue;
                    //ID,AppName,Viewed,ArchivedState,DocID,GriDate,DataInterType,FileName2,DataMandoubName,SpecType
                    SqlDataAdapter sqlDa = new SqlDataAdapter(queryVA[TableIndex], sqlCon);

                    sqlDa.SelectCommand.CommandType = CommandType.Text;
                    DataTable dtbl = new DataTable();
                    sqlDa.Fill(dtbl);

                    Console.WriteLine("deletefiles.....");
                    foreach (DataRow dataRow in dtbl.Rows)
                    {
                        //Console.WriteLine("TableList " + TableList[TableIndex] + " ---- GriDate " + dataRow["GriDate"].ToString());

                        if (partial)
                        {
                            if (TableIndex != 11 && TableIndex != 9)
                            {

                                if (dataRow["AppName"].ToString() == "" && dataRow["GriDate"].ToString() != GregorianDate)
                                {
                                    //MessageBox.Show(dataRow["ID"].ToString() + " - "+ TableList[TableIndex]);
                                    deleteRowsData(Convert.ToInt32(dataRow["ID"].ToString()), TableList[TableIndex], DataSource);
                                    deleteGenArch(dataRow["ID"].ToString(), TableList[TableIndex]);
                                    Console.WriteLine("ID " + dataRow["ID"].ToString() + " ---- TableList " + TableList[TableIndex]);
                                }
                            }
                            else
                            {
                                
                                if (dataRow["مقدم_الطلب"].ToString() == "" && dataRow["التاريخ_الميلادي"].ToString() != GregorianDate)
                                {
                                    //MessageBox.Show(dataRow["ID"].ToString() + " - " + TableList[TableIndex] + " - " + dataRow["مقدم_الطلب"].ToString());
                                    deleteRowsData(Convert.ToInt32(dataRow["ID"].ToString()), TableList[TableIndex], DataSource);
                                    deleteGenArch(dataRow["ID"].ToString(), TableList[TableIndex]);
                                }
                                Console.WriteLine("ID " + dataRow["ID"].ToString() + " ---- TableList " + TableList[TableIndex]);
                            }
                        }
                        else {
                            if (TableIndex != 11 && TableIndex != 9)
                            {
                                
                                if (dataRow["AppName"].ToString() == "" )
                                {
                                    //MessageBox.Show(dataRow["ID"].ToString() + " - " + TableList[TableIndex]);
                                    deleteRowsData(Convert.ToInt32(dataRow["ID"].ToString()), TableList[TableIndex], DataSource);
                                    deleteGenArch(dataRow["ID"].ToString(), TableList[TableIndex]);
                                    Console.WriteLine("ID " + dataRow["ID"].ToString() + " ---- TableList " + TableList[TableIndex]);
                                }
                            }
                            else
                            {

                                if (dataRow["مقدم_الطلب"].ToString() == "")
                                {
                                    //MessageBox.Show(dataRow["ID"].ToString() + " - " + TableList[TableIndex] + " - "+ dataRow["مقدم_الطلب"].ToString());
                                    deleteRowsData(Convert.ToInt32(dataRow["ID"].ToString()), TableList[TableIndex], DataSource);
                                    deleteGenArch(dataRow["ID"].ToString(), TableList[TableIndex]);
                                }
                                Console.WriteLine("ID " + dataRow["ID"].ToString() + " ---- TableList " + TableList[TableIndex]);
                            }
                        }
                        
                    }
                    Console.WriteLine("finish deleting files.....");

                }
            
            parrtialAll = true;
        }

       

        private void picUpdate_Click(object sender, EventArgs e)
        {            
            //upDateClose();

        }

        private void MainForm_MouseHover(object sender, EventArgs e)
        {
            uploadDocx = true;
            
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (mangerArch.CheckState == CheckState.Checked)
            {
                dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                string[] str = new string[persbtn3.Items.Count];
                for (int x = 0; x < persbtn3.Items.Count; x++) { str[x] = persbtn3.Items[x].ToString(); }
                string[] strSub = new string[4] { "إقرار بصيغة غير مدرجة", "اقرار بصيغة غير مدرجة مع الشهود", "إفادة لمن يهمه الأمر", "مذكرة لسفارة عربية" };
                FormPics form2 = new FormPics(Server,EmployeeName, attendedVC.Text,UserJobposition,DataSource, 12, FormDataFile, FilespathOut, 10, str, strSub, true,MandoubM, GriDateM);
                form2.ShowDialog();
                
            }
            else
            {
                dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                Form10 form10 = new Form10(attendedVC.SelectedIndex, IDNo, 0, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                form10.ShowDialog();
            }
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Suddanese_Affair_Click(object sender, EventArgs e)
        {
            
        }

        private void button30_Click(object sender, EventArgs e)
        {
            
        }

        private void yearReport_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool rows = false;
            
            int length = 3;
            if (ReportType.SelectedIndex == 8) length = 12;
            for (int s = 0; s < length; s++)
            {
                string from = yearReport.Text.Trim() + quorterS[s];
                string to = yearReport.Text.Trim() + quorterE[s];
                rows = DailyListcustm(from, to,s);
                if (rows) rowFound = true;
            }
            
            if (rowFound)
                {

                    PrintReport.Enabled = true;
                    PrintReport.Visible = true;
                    ReportPanel.Height = 205;

                }
                else
                {
                    PrintReport.Enabled = false;
                    PrintReport.Visible = false;
                    ReportPanel.Height = 42;
                    MessageBox.Show("لا يوجد قائمة بالتاريخ المحدد");
                }
            //}
            //MessageBox.Show(from + "-" + to + "----" + rows.ToString());
            
        }

        private void comboBox1_SelectedIndexChanged_2(object sender, EventArgs e)
        {
            //MessageBox.Show(Suddanese_Affair.SelectedIndex.ToString());
           
            getInSettings(Affbtn0.SelectedIndex);
           

        }

        private void getInSettings(int index)
        {
            bool modifyPermit = true;
            string[] prevStr = getPreivilage().Split('_');
            if (prevStr[index] != "1")
            {
                modifyPermit = false;
                MessageBox.Show("الملف ليس من صلاحيات حساب الموظف يرجى التواصل مع مدير القسم");
                //return;
            }

            if (mangerArch.CheckState == CheckState.Checked)
            {
                string[] str = new string[Affbtn0.Items.Count];
                for (int x = 0; x < Affbtn0.Items.Count; x++)
                {
                    str[x] = Affbtn0.Items[x].ToString();
                }

                string[] strSub = new string[Affairsbtn7.Items.Count];
                for (int x = 0; x < Affairsbtn7.Items.Count; x++)
                {
                    strSub[x] = Affairsbtn7.Items[x].ToString();
                }
                if (Affbtn0.SelectedIndex <= 5)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    FormPics form2 = new FormPics(Server, EmployeeName, attendedVC.Text, UserJobposition, DataSource, index, FormDataFile, FilespathOut, 13, str, strSub, true, MandoubM, GriDateM);
                    form2.ShowDialog();
                }
                else if (index == 11)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    NoteVerbal noteVerbal = new NoteVerbal(modifyPermit, attendedVC.SelectedIndex, GregorianDate, HijriDate, UserJobposition, DataSource, FilespathIn, FilespathOut, EmployeeName, 1, true);
                    noteVerbal.ShowDialog();
                }
            }

            else
            {
                //Affairsbtn7.Visible = true;
                if (index == 9 && UserJobposition.Contains("قنصل"))
                {
                    fileManagePanel2.Visible = true;
                    PanelMandounb.Visible = flowLayoutPanel1.Visible = panelAuthAknow.Visible = panelReceMess.Visible = SearchPanel.Visible = ReportPanel.Visible = false;
                    txtIndivNo.Text = txtFileNo.Text = "";
                }
                else if (index == 9 && !UserJobposition.Contains("قنصل"))
                {
                    MessageBox.Show("إدارة الملفات من صلاحيات مدير القسم فقط");
                }
                else if (index == 10)
                {
                    FormSuits formSuits = new FormSuits(DataSource, EmployeeName);
                    formSuits.ShowDialog();
                }
                else if (index == 11)
                {
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    
                    NoteVerbal noteVerbal = new NoteVerbal(modifyPermit, attendedVC.SelectedIndex, GregorianDate, HijriDate, UserJobposition, DataSource, FilespathIn, FilespathOut, EmployeeName, 1, false);
                    noteVerbal.ShowDialog();
                }
                else
                {
                    //Console.WriteLine("FormTimeLine");
                    dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                    FormSudAffairs formSudAffairs = new FormSudAffairs(modifyPermit, index, attendedVC.SelectedIndex, DataSource, FilespathIn, ArchFile + @"\", UserJobposition, EmployeeName, DataSource57);
                    formSudAffairs.ShowDialog(); Console.WriteLine(1);
                }
            }
        }
        private void loadSettings(string dataSource, bool day, bool daychange, bool month, bool monthchange)
        {
            SqlConnection Con = new SqlConnection(dataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select Modelfilespath,TempOutput,ServerName,Serverlogin,ServerPass,serverDatabase,hijriday,hijrimonth,FileArchive  from TableSettings where ID=@id", Con);
            sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
            try
            {
                if (Con.State == ConnectionState.Closed)
                    Con.Open();

                var reader = sqlCmd1.ExecuteReader();

                if (reader.Read())
                {
                    Model = reader["Modelfilespath"].ToString();
                    Output = reader["TempOutput"].ToString();
                    ServerIP = reader["ServerName"].ToString();
                    Login = reader["Serverlogin"].ToString();
                    Pass = reader["ServerPass"].ToString();
                    Database = reader["serverDatabase"].ToString();
                    Hiday = Convert.ToInt32(reader["hijriday"].ToString());
                    Himonth = Convert.ToInt32(reader["hijrimonth"].ToString());
                    FileArch = reader["FileArchive"].ToString();
                    if (daychange)
                    {
                        if (day) Hiday++;
                        else Hiday--;
                    }
                    if (monthchange)
                    {
                        if (month) Himonth++;
                        else Himonth--;
                    }
                }
            }
            catch (Exception ex)
            {
                if (Server == "57")
                {
                    MessageBox.Show("سيرفر قسم الاحوال الشخصية معطل أو غير متصل بالانترنت قم بإعادة تشغيل السيرفر أو التأكد من الاتصال بالانترنت");
                }
                else
                {
                    MessageBox.Show("سيرفر قسم شؤون الرعايا معطل أو غير متصل بالانترنت قم بإعادة تشغيل السيرفر أو التأكد من الاتصال بالانترنت");
                }
                Con.Close();
            }
            finally
            {
                Con.Close();
                if (Con.State == ConnectionState.Closed)
                {
                    try
                    {
                        Con.Open();
                        SqlCommand sqlCmd = new SqlCommand("SettingsAddorEdit", Con);
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@ID", 1);
                        sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                        sqlCmd.Parameters.AddWithValue("@Modelfilespath", Model);
                        sqlCmd.Parameters.AddWithValue("@TempOutput", Output);
                        sqlCmd.Parameters.AddWithValue("@ServerName", ServerIP);
                        sqlCmd.Parameters.AddWithValue("@Serverlogin", Login);
                        sqlCmd.Parameters.AddWithValue("@ServerPass", Pass);
                        sqlCmd.Parameters.AddWithValue("@serverDatabase", Database);
                        sqlCmd.Parameters.AddWithValue("@hijriday", Hiday);
                        sqlCmd.Parameters.AddWithValue("@hijrimonth", Himonth);
                        sqlCmd.Parameters.AddWithValue("@FileArchive", FileArch);
                        labeldate.Text = "فرق اليوم الهجري " + Hiday.ToString();
                        labelmonth.Text = "فرق الشهر الهجري " + Himonth.ToString();
                        sqlCmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("النظام متعطل يرجى التواصل مع مدير النظام");
                    }

                }
            }
        }

        private int getMaxRange(string dataSource)
        {
            int max = 0;
            SqlConnection Con = new SqlConnection(dataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select maxRange  from TableSettings where ID=1", Con);

            try
            {
                if (Con.State == ConnectionState.Closed)
                    Con.Open();
                sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
                var reader = sqlCmd1.ExecuteReader();

                if (reader.Read())
                {
                    max = Convert.ToInt32(reader["maxRange"].ToString());
                }
            }
            catch (Exception ex)
            {

                Con.Close();
            }
            return max;
        }


        private void loadMandInfo(string dataSource)
        {
            SqlConnection Con = new SqlConnection(dataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select MandoubNames,MandoubPhones from TableMandoudLis where MandoubAreas=@MandoubAreas", Con);
            sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();

            if (reader.Read())
            {
                Model = reader["MandoubNames"].ToString();
                Output = reader["MandoubPhones"].ToString();
                ServerIP = reader["MandoubAreas"].ToString();                
            }
            Con.Close();
            
            
        }

        private void button33_Click(object sender, EventArgs e)
        {
            
            ScanPic.Size = new System.Drawing.Size(228, 35);
            loadPic.Size = new System.Drawing.Size(228, 35);

            //228, 62
            loadPic.Enabled = button1.Visible = ScanPic.Enabled = false;
            
            try

            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.

                    var ScanerItem = device.Items[1]; // select the scanner.


                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);

                    PathImages[imagecount] = PrimariFiles + "ScanImg" + DateTime.Now.ToString("mmss") + (imagecount).ToString() + ".jpg";


                    if (File.Exists(PathImages[imagecount]))
                    {
                        File.Delete(PathImages[imagecount]);
                    }
                    imgFile.SaveFile(PathImages[imagecount]);
                    ArchivePic.ImageLocation = PathImages[imagecount];
                    dataGridView8.Visible = false;
                    ArchivePic.Visible = true;
                    imagecount++;
                }
                else
                {

                    MessageBox.Show("لا يوجد جهاز ماسح متصل");
                }

            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
            loadPic.Text = ScanPic.Text = "اضافة مستند آخر (" + imagecount.ToString() + ")";

            reLoadPic.Visible = RescanPic.Visible =  true;
            loadPic.Enabled = button1.Visible = ScanPic.Enabled = true;
        }

        private void button32_Click(object sender, EventArgs e)
        {
            
            try


            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.

                    var ScanerItem = device.Items[1]; // select the scanner.


                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);

                    PathImages[imagecount - 1] = PrimariFiles + "ScanImg" + DateTime.Now.ToString("mmss") + (imagecount - 1).ToString() + ".jpg";


                    if (File.Exists(PathImages[imagecount - 1]))
                    {
                        File.Delete(PathImages[imagecount - 1]);
                    }
                    imgFile.SaveFile(PathImages[imagecount - 1]);
                    ArchivePic.ImageLocation = PathImages[imagecount - 1];
                    
                    ArchivePic.Visible = true;
                    //panel1.Visible = false;

                }
                else
                {

                    MessageBox.Show("لا يوجد جهاز ماسح متصل");
                }

            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private string loadDocxFile()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {                
                return dlg.FileName;
            }
            return "";
        }

        private void loadPic_Click(object sender, EventArgs e)
        {
            ScanPic.Size = new System.Drawing.Size(228, 35);
            loadPic.Size = new System.Drawing.Size(228, 35);
            loadPic.Enabled = button1.Visible = ScanPic.Enabled = false;
            
            string fileName = loadDocxFile();
            if (fileName != "")
            {
                
                ArchivePic.ImageLocation = PathImages[imagecount];

                ArchivePic.Visible = true;
                imagecount++;
                persbtn2.BackColor = System.Drawing.Color.LightGreen;
                persbtn2.Text = "اضافة مستند آخر (" + (imagecount + 1).ToString() + ")";
                reLoadPic.Enabled = RescanPic.Visible = true;

            }
            dataGridView8.Visible = false;
            loadPic.Enabled = ScanPic.Enabled = true;
            reLoadPic.Visible = RescanPic.Visible = true;
        }

        private void reLoadPic_Click(object sender, EventArgs e)
        {
            
            string fileName = loadDocxFile();
            if (fileName != "")
            {
                ArchivePic.ImageLocation = PathImages[imagecount - 1] = fileName;
                ArchivePic.Visible = true;
            }
        }

        private string CreatePic( string[] location)
        {
            
            string ActiveCopy = PrimariFiles + "Docx" + DateTime.Now.ToString("mmss") + ".docx";
                using (DocX document = DocX.Create(ActiveCopy))
                {
                    Paragraph p1 = document.InsertParagraph();

                    // Append content to the Paragraph
                    for (int x = 0; x < imagecount; x++)
                    {
                        var image = document.AddImage(location[x]);
                        // Set Picture Height and Width.
                        var picture = image.CreatePicture(600, 500);

                        p1.AppendPicture(picture);
                    }
                    document.Save();

                    imagecount = 0;
                }
            return ActiveCopy;
        }
        private void button34_Click(object sender, EventArgs e)
        {
            string comment = "";
            var selectedOption = MessageBox.Show("", "هل الشهادة صحيحة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                comment = "";
            } else if (selectedOption == DialogResult.No) 
                comment = "مستند غير صحيح";

            if (PathImage == "")
            {
                MessageBox.Show("يرجى أرشفة نموذج التوقيع أولا");
                return;
            }

            if (!combCount.Text.All(char.IsDigit))
            {
                MessageBox.Show("يرجى تحديد عدد المستندات بصورة صحيحة");
                return;
            }
            if (txtHAGredate.Text.Length == 0)
            {
                MessageBox.Show("يرجى توضيح تاريخ توثيق الشهادة من آخر جهة");
                return;
            }
            if (txtHAGredate.Text.Length == 0)
            {
                MessageBox.Show("يرجى توضيح تاريخ توثيق الشهادة من آخر جهة");
                return;
            }
            if (txtHAAuthentic.Text.Length == 0)
            {
                MessageBox.Show("يرجى توضيح اسم موثق المكاتبة ");
                return;
            }

            int iD = 1;
            if (btnSaveArch.Text == "حفظ وتأكيد")
            {
                // Console.WriteLine("حفظ وتأكيد");
                SubAuthData(1, txtHandAuthNo.Text, txtHAAuthentic.Text, checkHASex.Text, comboBox2.Text, txtHAGredate.Text, CreatePic(PathImages), GregorianDate + "_تمت معالجة الملف بواسطة" + EmployeeName, combCount.Text, comment);

            }
            else if (btnSaveArch.Text == "تعديل وتأكيد")
            {
                iD = Messid;
                SubAuthData(iD, txtHandAuthNo.Text, txtHAAuthentic.Text, checkHASex.Text, comboBox2.Text, txtHAGredate.Text, CreatePic(PathImages), GregorianDate + "_تمت معالجة الملف بواسطة" + EmployeeName, combCount.Text, comment);
            }
            btnSaveArch.Text = "حفظ وتأكيد";
            txtHandAuthNo.Text = txtHAAuthentic.Text = checkHASex.Text = comboBox2.Text = txtHAGredate.Text = ArchfilePath = "";

            Messid = 1;
            fillDataGrid(comboBox2.Text);

            ColorFulGrid9();
            ArchfilePath = "";
            txtHandAuthNo.Text = txtHAAuthentic.Text = checkHASex.Text = comboBox2.Text = txtHAGredate.Text = ArchfilePath = "";
            panelAuthAknow.Visible = false;
            flowLayoutPanel1.Visible = true;
            panel4.Visible = false;
            imagecount = 0;
            ArchivePic.Image = null;
            ScanPic.Text = "بدء الارشفة";
            loadPic.Text = "تحميل من ملف";
            ScanPic.Size = new System.Drawing.Size(392, 59);
            loadPic.Size = new System.Drawing.Size(392, 59);
            reLoadPic.Visible = RescanPic.Visible = false;
        }

        private void comboBox1_SelectedIndexChanged_3(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView8.DataSource;
                bs.Filter = dataGridView8.Columns[4].HeaderText.ToString() + " LIKE '%" + GregorianDate + "%'";
                dataGridView8.DataSource = dataGridView8.DataSource = bs;
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                dateAuth.Visible = true;
                comboBox1.Visible = false;
            }
            else {
                fillDataGrid("");
            }
            
        }

        private void dateAuth_ValueChanged(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            string satePicked = dateAuth.Value.ToString().Split(' ')[0].Replace("/", "-").Trim();
           // Console.WriteLine(satePicked);
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView8.DataSource;
            bs.Filter = dataGridView8.Columns[4].HeaderText.ToString() + " LIKE '%" + satePicked + "%'";
            dataGridView8.DataSource = dataGridView8.DataSource = bs;

            dateAuth.Visible = false;
            comboBox1.Visible = true;
        }

        private void DocDestin_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            
            
        }

        private void comFileType_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtFileNo.Text = getFileNo(comFileType.SelectedIndex);
                txtIndivNo.Text = getFileNo(comFileType.SelectedIndex + 5);
            
             
        }

        private void txtFileNo_TextChanged(object sender, EventArgs e)
        {

        }

        private void sudan_affairs_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(DataSource);
            dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
            FormTimeLine formTimeLine = new FormTimeLine(attendedVC.SelectedIndex, DataSource, UserJobposition, FilespathIn, ArchFile + @"\", EmployeeName, GregorianDate);
            formTimeLine.ShowDialog();
        }

        private void comMandArea_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button30_Click_1(object sender, EventArgs e)
        {
            fillMandoubGrid();
            if (PanelMandounb.Visible == false)
            {
                PanelMandounb.Visible = true;
                panel4.Visible = false;
                SearchPanel.Visible = fileManagePanel2.Visible = panelAuthAknow.Visible = panelReceMess.Visible = ReportPanel.Visible = false;

            }
            else PanelMandounb.Visible = false;
        }

        private void reportpass_TextChanged(object sender, EventArgs e)
        {
            int hour = Convert.ToInt32(DateTime.Now.ToString("hh"));

            if (reportpass.Text == (hour*17).ToString())
            {
                reportpass.Text = "";
                reportpass.Visible = false;
                dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                DeepStatistics deepStatistics = new DeepStatistics(DataSource57, DataSource56, FilespathIn, FilespathOut);
                deepStatistics.ShowDialog();
            }
        }

        private void btnFileManage_Click(object sender, EventArgs e)
        {
            string item1Name = "",item2Name="";
            switch (comFileType.SelectedIndex)
            {
                case 0:
                    item1Name = "رقم_ملف_جدة";
                    item2Name = "عدد_الأفراد";
                    break;
                case 1:
                    item1Name = "رقم_ملف_مكة";
                    item2Name = "عدد_الأفراد_مكة";
                    break;

                case 2:
                    item1Name = "رقم_ملف_الوافدين";
                    item2Name = "عدد_الأفراد_الوافدين";
                    break;

                case 3:
                    item1Name = "رقم_ملف_اللجنة";
                    item2Name = "عدد_الأفراد_اللجنة";
                    break;
                case 4:
                    item1Name = "رقم_ملف_المقابل";
                    item2Name = "عدد_الأفراد_المقابل";
                    break;
            }

            upDateFilesInfo(item1Name, txtFileNo.Text, item2Name, txtIndivNo.Text);
            txtFileNo.Text = txtIndivNo.Text = "";
            fileManagePanel2.Visible = false;

        }

        private void upDateFilesInfo(string item1Name, string item1Value, string item2Name, string item2Value)
        {
            string at1 = "@" + item1Name;
            string at2 = "@" + item2Name;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("update TableSettings set "+ item1Name + "=@"+ item1Name +","+ item2Name + "=@" + item2Name + " where ID='1'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue(at1, item1Value);
            sqlCmd.Parameters.AddWithValue(at2, item2Value);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void Aprove_Click(object sender, EventArgs e)
        {
            dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
            SignUp signUp = new SignUp(EmployeeName, UserJobposition, DataSource);
            signUp.ShowDialog();
        }

        private void btnNewYear_Click(object sender, EventArgs e)
        {
            int form = 12;
            for (; form < 14; form++)
                NewYearEntry(form, DateTime.Now.Year.ToString().Replace("20", ""), GregorianDate);
                    if(form == 13)
                UpdateYear(DateTime.Now.Year.ToString());
        }

        private void UpdateYear(string text)
        {
            string qurey = "update TableSettings set CurrentYear=@CurrentYear where ID=@id";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", 1);
            sqlCmd.Parameters.AddWithValue("@CurrentYear", text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void flowLayoutPanel1_MouseMove(object sender, MouseEventArgs e)
        {
            i = 0;
        }



        private void pictureBox1_Click_2(object sender, EventArgs e)
        {

        }


        private void buttonClick(object sender, EventArgs e)
        {
            Button button = sender as Button;
        }

        private void labelarch_Click(object sender, EventArgs e)
        {

        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (button17.Text == "بحث")
            {
                fillDataGrid("");
                button17.Text = "الموِّقع/بحث";
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView8.DataSource;
                bs.Filter = dataGridView8.Columns["AppName"].HeaderText.ToString() + " LIKE '%" + txtHAAuthentic.Text + "%'";
                dataGridView8.DataSource = bs;
            }
            else
                button17.Text = "بحث";
        }


        private void applicant_TextChanged(object sender, EventArgs e)
        {
            if (nameNo) return;
            //MessageBox.Show("الاسم");
                FillDatafromGenArch(applicant.Text, "الاسم");
        }

        private void txtSearch_Click(object sender, EventArgs e)
        {
            nameNo = true;
        }

        private void applicant_Click(object sender, EventArgs e)
        {
            nameNo = false;
        }

        private void button41_Click(object sender, EventArgs e)
        {
            getInSettings(7);
        }

        private void button36_Click(object sender, EventArgs e)
        {
            getInSettings(8);
        }

        private void button40_Click(object sender, EventArgs e)
        {
            getInSettings(9);
        }

        private void button39_Click(object sender, EventArgs e)
        {
            getInSettings(10);
        }

        private void button38_Click(object sender, EventArgs e)
        {
            getInSettings(11);
        }

        

        private void timer5_Tick(object sender, EventArgs e)
        {
            //updataArchData1();
            //fillNonArchInfo();
        }


        private void timer3_Tick_1(object sender, EventArgs e)
        {
            
            //updataArchData1();
            //if (Server == "57") 
            //    updataArchData2();
            
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
        SqlCommand sqlCommand = new SqlCommand("insert into archives values (" + strList + ")", sqlConnection);
        //SqlCommand sqlCommand = new SqlCommand("insert into archives (docID, employName,archiveStat,databaseID,appType,appOldNew) " +
        //    "values (@docID, @employName,@archiveStat,@databaseID,@appType,@appOldNew)", sqlConnection);
        sqlCommand.CommandType = CommandType.Text;
        for (int i = 1; i < allList.Length; i++)
        {
            sqlCommand.Parameters.AddWithValue("@" + allList[i], text[i - 1]);
            //MessageBox.Show(text[i - 1]);
        }
        sqlCommand.ExecuteNonQuery();
    }

        

        private void persbtn10_Click(object sender, EventArgs e)
        {
            uploadDocx = false;
            dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
            //MessageBox.Show(HijriDate);

            FormAuth formAuth = new FormAuth(attendedVC.SelectedIndex, -1, "", DataSource, FilespathIn, FilespathOut, EmployeeName, UserJobposition, GregorianDate, HijriDate);
            formAuth.ShowDialog();
        }

        

        private void empUpdate_Click_1(object sender, EventArgs e)
        {
            File.Delete(primeryLink + "fileUpdate.txt");
            System.Diagnostics.Process.Start(getAppFolder() + @"\setup.exe");
            this.Close();
        }

        private void fileUpdate_Click_1(object sender, EventArgs e)
        {
            if (!backgroundWorker2.IsBusy) backgroundWorker2.RunWorkerAsync(); 
            if (!backgroundWorker1.IsBusy) backgroundWorker1.RunWorkerAsync();            
        }

        private void persbtn11_Click(object sender, EventArgs e)
        {
            if (mangerArch.CheckState == CheckState.Checked)
            {
                dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                string[] str = new string[persbtn3.Items.Count];
                for (int x = 0; x < persbtn3.Items.Count; x++) { str[x] = persbtn3.Items[x].ToString(); }
                string[] strSub = fileStrSub(DataSource, "ArabicGenIgrar", "TableListCombo");
                FormPics form2 = new FormPics(Server, EmployeeName, attendedVC.Text, UserJobposition, DataSource, 12, FormDataFile, FilespathOut, 10, str, strSub, true, MandoubM, GriDateM);
                form2.ShowDialog();

            }
            else
            {
                dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
                FormCollection formCollection = new FormCollection(attendedVC.SelectedIndex, IDNo, 0, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                formCollection.ShowDialog();
            }
        }

        private string[] fileStrSub(string source, string comlumnName, string tableName)
        {
            string[] strSub;
            int i = 0;
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName + " where " + comlumnName +" is not null";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;


                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                strSub = new string[table.Rows.Count];
                foreach (DataRow dataRow in table.Rows)
                {
                    if (!String.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                    { strSub[i] = dataRow[comlumnName].ToString();
                        i++; }
                }
                saConn.Close();
            }
            return strSub;  
        }

        private void docCollectCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] str = new string[docCollectCombo.Items.Count];
            for (int x = 0; x < docCollectCombo.Items.Count; x++)
            {
                str[x] = docCollectCombo.Items[x].ToString();
            }
            string[] strSub = new string[1] { "" };
            dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
            FormPics form2 = new FormPics(Server, EmployeeName, attendedVC.Text, UserJobposition, DataSource, docCollectCombo.SelectedIndex, FormDataFile, FilespathOut, 10, str, strSub, true, MandoubM, GriDateM);
            form2.ShowDialog();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            var settings = new Settings(Server, false, DataSource56, DataSource57, false, LocalModelFiles, ArchFile, ArchFile, LocalModelForms, "");
            settings.Show();
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();



                string[] serverfiles = Directory.GetFiles(ServerModelFiles);
            for (int i = 0; i < serverfiles.Length; i++)
            {
                //MessageBox.Show(serverfiles[i]);
                var serverfileinfo = new FileInfo(serverfiles[i]);
                string serverfilename = serverfileinfo.Name;
                string localFile = FilespathIn + serverfilename;

                if (!File.Exists(localFile))
                {
                    System.IO.File.Copy(serverfiles[i], localFile);
                    //MessageBox.Show(serverfiles[i]);
                }
                else if (File.Exists(localFile))
                {
                    DateTime serverLastWite = serverfileinfo.LastWriteTime.Date;
                    var localfileinfo = new FileInfo(localFile);
                    DateTime localLastWite = localfileinfo.LastWriteTime.Date;
                    if (serverLastWite != localLastWite)
                    {
                        try
                        {
                           File.Delete(localFile);
                        }
                        catch (Exception ex) { }
                        System.IO.File.Copy(serverfiles[i], localFile);
                    }
                }
            }


            foreach (string localFile in Directory.GetFiles(FilespathIn))
            {
                var localFileinfo = new FileInfo(localFile);
                string localFilename = localFileinfo.Name;
                string serverfile = ServerModelFiles + localFilename;
                if (File.Exists(localFile) && !File.Exists(serverfile))
                {
                    try
                    {
                        File.Delete(localFile);
                    }
                    catch (Exception ex) { }
                }
                
            }

            string[] formfiles = Directory.GetFiles(ServerModelForms);
            for (int i = 0; i < formfiles.Length; i++)
            {
                var serverforminfo = new FileInfo(formfiles[i]);
                string serverformName = serverforminfo.Name;
                string localForm = FormDataFile + serverformName;

                if (!File.Exists(localForm))
                    System.IO.File.Copy(formfiles[i], localForm);
                else if (File.Exists(localForm))
                {
                    DateTime serverLastWite = serverforminfo.LastWriteTime.Date;
                    var localforminfo = new FileInfo(localForm);
                    DateTime localLastWite = localforminfo.LastWriteTime.Date;
                    if (serverLastWite != localLastWite)
                    {
                        //MessageBox.Show(serverfiles[i]);
                        try
                        {
                            File.Delete(localForm);
                        }
                        catch (Exception ex) { }
                        System.IO.File.Copy(formfiles[i], localForm);
                    }
                }
            }
            foreach (string localForm in Directory.GetFiles(FormDataFile))
            {
                var localFileinfo = new FileInfo(localForm);
                string localFormname = localFileinfo.Name;
                string serverform = ServerModelForms + localFormname;
                if (File.Exists(localForm) && !File.Exists(serverform))
                {
                    try

                    {
                       File.Delete(localForm);
                    }
                    catch (Exception ex) { }
                }

            }
        }

        private string[] getColList(string table)
        {
            string[] allList = new string[1];
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return allList; }

            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('" + table + "')", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            allList = new string[dtbl.Rows.Count];
            int i = 0;
            foreach (DataRow row in dtbl.Rows)
            {
                allList[i] = row["name"].ToString();
                i++;
            }
            return allList;

        }

        private void updataArchData1()
        {
            i++;
            A = M = 0;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            int empCountN = 0;
            int empCountO = 0;
            int cuerrentRange = 0; 
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return ; }

                SqlDataAdapter sqlDa = new SqlDataAdapter("select * from  archives", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);
                sqlCon.Close();
            
            
            maxRange = getMaxRange(DataSource);
            foreach (DataRow dataRow in dtbl.Rows)
            {
                //if (dataRow["mandoubName"].ToString().Contains("مؤرشف نهائي"))
                //    deleteRowsData(dataRow["docID"].ToString());

                if (dataRow["appOldNew"].ToString() == "new" && dataRow["docDate"].ToString() != GregorianDate && dataRow["employName"].ToString().Contains(ConsulateEmployee.Text)) 
                    empCountN++;
                else if (dataRow["appOldNew"].ToString() == "old" && dataRow["docDate"].ToString() != GregorianDate && dataRow["employName"].ToString().Contains(ConsulateEmployee.Text)) 
                    empCountO++;
                //doubleCheckArch(dataRow["docID"].ToString());
                if (Career == "موظف ارشفة" || dataRow["employName"].ToString().Contains(ConsulateEmployee.Text))
                {
                    if (dataRow["mandoubName"].ToString() == "")
                    {
                        oldNewA[A] = dataRow["appOldNew"].ToString();
                        
                        DocA[A] = dataRow["docID"].ToString();
                        GriDateA[A] = dataRow["docDate"].ToString();
                        IDA[A] = Convert.ToInt32(dataRow["ID"].ToString());
                        if(oldNewA[A] == "old")
                            AppNameA[A] = dataRow["appName"].ToString() + " (نسخة معدلة)";
                        else AppNameA[A] = dataRow["appName"].ToString() ;
                        if (dataRow["appOldNew"].ToString() == "new" && dataRow["docDate"].ToString() != GregorianDate && dataRow["employName"].ToString().Contains(ConsulateEmployee.Text))
                            cuerrentRange++;
                        A++;
                    }
                    else if (dataRow["mandoubName"].ToString() != "" && dataRow["appName"].ToString() != "" )
                    {
                        oldNewM[M] = dataRow["appOldNew"].ToString();
                        DocIDM[M] = dataRow["docID"].ToString();
                        MandoubM[M] = dataRow["mandoubName"].ToString();
                        DocM[M] = dataRow["docID"].ToString();
                        GriDateM[M] = dataRow["docDate"].ToString();
                        IDM[M] = Convert.ToInt32(dataRow["ID"].ToString());
                        
                        if (oldNewM[M] == "old")
                            AppNameM[M] = dataRow["appName"].ToString() + " (نسخة معدلة)";
                        else AppNameM[M] = dataRow["appName"].ToString();
                        M++;
                    }
                }
            }
            }
            catch (Exception ex) { return; }
            if (A > 0)
            {
                labelarch.BackColor = Color.Red;
                labelarch.Text = "غير مؤرشف " + A.ToString();
            }
            else
                labelarch.BackColor = Color.Green;

            if (M > 0)
            {
                labelM.BackColor = Color.Red;
                labelM.Text = "غير مؤرشف " + M.ToString();
            }
            else
                labelM.BackColor = Color.Green;

            deleteEmptyRows = false;
            if (!Showed && Career != "موظف ارشفة")
            {
                Showed = true;
                string ReportName1 = "Report1" + DateTime.Now.ToString("mmss") + ".docx";
                string ReportName2 = "Report2" + DateTime.Now.ToString("mmss") + ".docx";
                if (empCountN != 0 && empCountO == 0)
                {
                    var selectedOption1 = MessageBox.Show( ConsulateEmployee.Text + " لديك عدد " + empCountN.ToString() + " معاملات غير مؤرشفة، يرجى أرشفتها...", "معاينة الأرشفة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (selectedOption1 == DialogResult.Yes)
                    {
                        if(A>0) CreateNotArchivedFiles(A, ReportName1, GriDateA, DocA, AppNameA, "رقم المعاملة المرجعي", "غير المؤرشفة");
                        
                        if (M>0) CreateMandounbFiles(M, ReportName2, GriDateM, DocIDM, AppNameM, MandoubM);
                    }
                }
                else if (empCountN != 0 && empCountO != 0)
                {
                    var selectedOption2 = MessageBox.Show(ConsulateEmployee.Text + " لديك عدد " + empCountN.ToString() + " معاملات غير مؤرشفة، يرجى أرشفتها، وعدد " + empCountO.ToString() + " نسخة معدلة من معاملات لم تتم إعادة إضافتها إلى ملف المعاملة...", "معاينة الأرشفة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (selectedOption2 == DialogResult.Yes)
                    {
                        if (A > 0) CreateNotArchivedFiles(A, ReportName1, GriDateA, DocA, AppNameA, "رقم المعاملة المرجعي", "غير المؤرشفة");
                        if (M > 0) CreateMandounbFiles(M, ReportName2, GriDateM, DocIDM, AppNameM, MandoubM);
                    }
                }
                else if (empCountN == 0 && empCountO != 0)
                {
                    var selectedOption3 = MessageBox.Show(ConsulateEmployee.Text + " لديك عدد " + empCountO.ToString() + " نسخة معدلة من معاملات لم تتم إعادة إضافتها إلى ملف المعاملة...", "معاينة الأرشفة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (selectedOption3 == DialogResult.Yes)
                    {
                        if (A > 0) CreateNotArchivedFiles(A, ReportName1, GriDateA, DocA, AppNameA, "رقم المعاملة المرجعي", "غير المؤرشفة");
                        if (M > 0) CreateMandounbFiles(M, ReportName2, GriDateM, DocIDM, AppNameM, MandoubM);

                    }
                }

            }
            if (!MessageShowed && (cuerrentRange > maxRange ))
            {
                MessageShowed = true;
                flowLayoutPanel3.BringToFront();
                Combtn2.BringToFront();
                MessageBox.Show("عدد المعاملات غير المؤرشفة تخطى الحد الأقصى لعدد المعاملات المسموح.. يرجى أرشفة المعاملات أولا للمتابعة");
            } else if ( (empCountN + empCountO) < maxRange)
            {
                MessageShowed = false;
                flowLayoutPanel3.SendToBack();
                
            }
        }


        private void updataArchData2()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string[] DocumentID = new string[40];
            string year = DateTime.Now.Year.ToString().Replace("20", "");

            try
            {
                 
            if (sqlCon.State == ConnectionState.Closed)
                   sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("select ID,المستند from TableGeneralArch where fileUpload='No'", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);
                dataGridView4.DataSource = dtbl;


                for (int x = 0; x < dtbl.Rows.Count; x++)
                {
                    string DocxFileName = dataGridView4.Rows[x].Cells[1].Value.ToString();
                    //string fileUpload = dataGridView4.Rows[x].Cells[11].Value.ToString();
                    if (uploadDocx)
                    {
                        if (File.Exists(DocxFileName) && !fileIsOpen(DocxFileName))
                        {
                            try
                            {
                                FinalDataArch(DataSource, DocxFileName, Convert.ToInt32(dataGridView4.Rows[x].Cells[0].Value.ToString()));
                                File.Delete(DocxFileName);
                            }
                            catch (Exception ex) { }
                        }
                    }
                }
            }
            catch (Exception ex) { return; }
        }
        //private void updataArchData()
        //{
        //    i++;
        //    V = A = M = 0;
        //    SqlConnection sqlCon = new SqlConnection(DataSource);
        //    string[] DocumentID = new string[40];
        //    string year = DateTime.Now.Year.ToString().Replace("20", "");

        //    if (sqlCon.State == ConnectionState.Closed)
        //    {
        //        sqlCon.Open();
        //        for (TableIndex = 10; TableIndex < 11; TableIndex++)
        //        {
        //            //ID,AppName,Viewed,ArchivedState,DocID,GriDate,DataInterType,FileName2,DataMandoubName,SpecType
        //            SqlDataAdapter sqlDa = new SqlDataAdapter(queryVA[TableIndex], sqlCon);
        //            if (!Pers_Peope && TableIndex < 7)
        //                sqlDa = new SqlDataAdapter("select ID,مقدم_الطلب,موقع_العاملة,حالة_الارشفة,رقم_المعاملة,التاريخ_الميلادي,طريقة_الطلب,المكاتبة_النهائية,اسم_المندوب,اسم_الموظف  from " + getFileTable(TableIndex - 1), sqlCon);
        //            else if (!Pers_Peope && TableIndex == 7) break;
        //            sqlDa.SelectCommand.CommandType = CommandType.Text;
        //            DataTable dtbl = new DataTable();
        //            sqlDa.Fill(dtbl);
        //            dataGridView4.DataSource = dtbl;


        //            for (int x = 0; x < dtbl.Rows.Count; x++)
        //            {
        //                bool spec = false;
        //                //if (TableIndex == 9)
        //                //{
        //                //    if (dataGridView4.Rows[x].Cells[9].Value.ToString() == "إقرار خروج نهائي بدون استحقاقات") spec = true;
        //                //}
        //                if (!spec && Pers_Peope && dataGridView4.Rows[x].Cells[6].Value.ToString() == "حضور مباشرة إلى القنصلية" && dataGridView4.Rows[x].Cells[2].Value.ToString() != "غير معالج" && dataGridView4.Rows[x].Cells[3].Value.ToString() == "غير مؤرشف")
        //                {
        //                    colIDs[0] = dataGridView4.Rows[x].Cells[4].Value.ToString();
        //                    colIDs[1] = dataGridView4.Rows[x].Cells[0].Value.ToString();
        //                    colIDs[2] = dataGridView4.Rows[x].Cells[5].Value.ToString();
        //                    colIDs[3] = dataGridView4.Rows[x].Cells[1].Value.ToString();
        //                    colIDs[4] = dataGridView4.Rows[x].Cells[9].Value.ToString();
        //                    colIDs[5] = "حضور مباشرة إلى القنصلية";
        //                    colIDs[6] = "";
        //                    colIDs[7] = "new";
        //                    addarchives(colIDs);
        //                    // DocA[A] = TableIndex;
        //                    GriDateA[A] = dataGridView4.Rows[x].Cells[5].Value.ToString();
        //                    IDA[A] = Convert.ToInt32(dataGridView4.Rows[x].Cells[0].Value.ToString());
        //                    AppNameA[A] = dataGridView4.Rows[x].Cells[1].Value.ToString();
        //                    DocumentID = dataGridView4.Rows[x].Cells[4].Value.ToString().Split('/');

        //                    if (DocumentID.Length == 4)
        //                    {
        //                        DocIDA[A] = year + DocumentID[2] + DocumentID[3];
        //                        string filePath1 = ArchFile + "text1.txt";
        //                        string filePath2, filePath3;
        //                        filePath2 = ArchFile + year + DocumentID[2] + DocumentID[3] + ".pdf";
        //                        filePath3 = ArchFile + year + DocumentID[2] + DocumentID[3] + "_0001.pdf";
        //                        if (TableIndex == 10) ArchType = false;
        //                        else
        //                        {
        //                            filePath2 = ArchFile + year + DocumentID[2] + DocumentID[3] + ".pdf";
        //                            filePath3 = ArchFile + year + DocumentID[2] + DocumentID[3] + "_0001.pdf";

        //                            ArchType = true;
        //                        }

        //                        if (File.Exists(filePath2))
        //                        {
        //                            AuthArch(ArchType, DataSource, IDA[A], qureyFunction(TableList[TableIndex], ArchType), filePath2, dataGridView4.Rows[x].Cells[1].Value.ToString());
        //                        }
        //                        if (File.Exists(filePath3))
        //                        {
        //                            AuthArch(ArchType, DataSource, IDA[A], qureyFunction(TableList[TableIndex], ArchType), filePath3, dataGridView4.Rows[x].Cells[1].Value.ToString());
        //                        }
        //                    }
        //                    A++;
        //                }
        //                else if (!Pers_Peope && dataGridView4.Rows[x].Cells[3].Value.ToString() == "غير مؤرشف")
        //                {

        //                    //DocA[A] = TableIndex;
        //                    GriDateA[A] = dataGridView4.Rows[x].Cells[5].Value.ToString();
        //                    IDA[A] = Convert.ToInt32(dataGridView4.Rows[x].Cells[0].Value.ToString());
        //                    AppNameA[A] = dataGridView4.Rows[x].Cells[1].Value.ToString();
        //                    DocumentID = dataGridView4.Rows[x].Cells[4].Value.ToString().Split('/');

        //                    A++;
        //                }
        //                else if (dataGridView4.Rows[x].Cells[3].Value.ToString().Contains("_") && dataGridView4.Rows[x].Cells[6].Value.ToString() != "حضور مباشرة إلى القنصلية" && !dataGridView4.Rows[x].Cells[6].Value.ToString().Contains("ملغي"))
        //                {
        //                    colIDs[0] = dataGridView4.Rows[x].Cells[4].Value.ToString();
        //                    colIDs[1] = dataGridView4.Rows[x].Cells[0].Value.ToString();
        //                    colIDs[2] = dataGridView4.Rows[x].Cells[5].Value.ToString();
        //                    colIDs[3] = dataGridView4.Rows[x].Cells[1].Value.ToString();
        //                    colIDs[4] = dataGridView4.Rows[x].Cells[9].Value.ToString();
        //                    colIDs[5] = "عن طريق أحد مندوبي القنصلية";
        //                    colIDs[6] = dataGridView4.Rows[x].Cells[3].Value.ToString();
        //                    colIDs[7] = "new";
        //                    addarchives(colIDs);
        //                    DocIDM[M] = dataGridView4.Rows[x].Cells[4].Value.ToString();
        //                    //DocM[M] = TableIndex;
        //                    GriDateM[M] = dataGridView4.Rows[x].Cells[5].Value.ToString();
        //                    IDM[M] = Convert.ToInt32(dataGridView4.Rows[x].Cells[0].Value.ToString());
        //                    AppNameM[M] = dataGridView4.Rows[x].Cells[1].Value.ToString();

        //                    if (dataGridView4.Rows[x].Cells[3].Value.ToString().Contains("-"))
        //                        MandoubM[M] = dataGridView4.Rows[x].Cells[3].Value.ToString().Split('-')[0].Split('_')[1] + "-" + dataGridView4.Rows[x].Cells[3].Value.ToString().Split('-')[1];
        //                    else
        //                        MandoubM[M] = dataGridView4.Rows[x].Cells[3].Value.ToString();

        //                    M++;
        //                }
        //                if (TableIndex == 10)
        //                {
        //                    string DocxFileName = dataGridView4.Rows[x].Cells[6].Value.ToString();
        //                    string fileUpload = dataGridView4.Rows[x].Cells[11].Value.ToString();
        //                    //queryVA[10] = "select ID,مقدم_الطلب,المعالجة,حالة_الارشفة,رقم_التوكيل,التاريخ_الميلادي,DocxData,Extension3,طريقة_الطلب,المكاتبة_النهائية,اسم_المندوب from TableAuth";
        //                    if (uploadDocx && fileUpload == "No")
        //                    {
        //                        //MessageBox.Show(file);
        //                        if (File.Exists(DocxFileName) && !fileIsOpen(DocxFileName))
        //                        {
        //                            //MessageBox.Show(file);
        //                            FinalDataArch(DataSource, DocxFileName);
        //                            File.Delete(DocxFileName);
        //                        }
        //                        //else Console.WriteLine("fileIsOpen " + DocxFileName);

        //                    }
        //                    //else Console.WriteLine("fileUpload " + fileUpload +" to " + DocxFileName);
        //                }

        //                if (dataGridView4.Rows[x].Cells[2].Value.ToString() == "غير معالج")
        //                {
        //                   // DocV[V] = TableIndex;
        //                    IDV[V] = Convert.ToInt32(dataGridView4.Rows[x].Cells[0].Value.ToString());
        //                    AppNameV[V] = dataGridView4.Rows[x].Cells[1].Value.ToString();
        //                    V++;
        //                }
        //            }
        //        }
        //    }
        //    if (A > 0)
        //    {
        //        labelarch.BackColor = Color.Red;
        //        labelarch.Text = "غير مؤرشف " + A.ToString();
        //    }
        //    else labelarch.BackColor = Color.Green;
        //    if (V > 1)
        //    {
        //        labelPro.BackColor = Color.Red;
        //        labelPro.Text = "غير معالج " + V.ToString();
        //        pictureBox2.Visible = labelPro.Visible = false;
        //    }
        //    else labelPro.BackColor = Color.Green;
        //    sqlCon.Close();

        //    deleteEmptyRows = false;
        //    MessageBox.Show("finish");
        //}

        public bool fileIsOpen(string path)
        {
            System.IO.FileStream a = null;

            try
            {
                a = System.IO.File.Open(path,
                System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None);
                return false;
            }
            catch (System.IO.IOException ex)
            {
                return true;
            }

            finally
            {
                if (a != null)
                {
                    a.Close();
                    a.Dispose();
                }
            }
        }
        private void UpdateState(int id, string text1, string table, string text2)
        {
            string qurey = "update " + table + " set "+text2+"=@"+ text2 + " where ID=@id";
            //MessageBox.Show(text2);
            string col = "@" + text2;
            try
            {
                SqlConnection sqlCon = new SqlConnection(DataSource);
                SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
                if (sqlCon.State == ConnectionState.Closed)
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return ; }
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@id", id);
                sqlCmd.Parameters.AddWithValue(col, text1);
                sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex) {
                //timer3.Enabled = false;
                //MessageBox.Show(qurey);
            }
            sqlCon.Close();
        }

        private void UpdateState( int id,string text, string table)
        {
            string qurey = "update "+table+" set GriDate=@GriDate where ID=@id";            
            SqlConnection sqlCon = new SqlConnection(DataSource);            
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);            
            sqlCmd.Parameters.AddWithValue("@GriDate", text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }


        private void UpdateMaririageColumn(string source, string type, string date)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string qurey = "insert into TableMarriageDocs (ProType,GriDate) values (@ProType,@GriDate)";

            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;

            sqlCmd.Parameters.AddWithValue("@ProType", type);
            sqlCmd.Parameters.AddWithValue("@GriDate", date);
            sqlCmd.ExecuteNonQuery();

            sqlCon.Close();
        }
        private void UpdateColumn(string source, string comlumnName, int id, string data, string table)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string column = "@" + comlumnName;
            string qurey = "UPDATE "+table+" SET " + comlumnName + " = " + column + " WHERE ID=@ID";

            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;

            sqlCmd.Parameters.AddWithValue("@ID", id);
                sqlCmd.Parameters.AddWithValue(column, data.Trim());
            try
            {
                sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex) { MessageBox.Show(column +"-"+ data); }
            
            sqlCon.Close();
        }
        
        private void InsertColumn(string source, string comlumnName, int id, string data, string table)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string column = "@" + comlumnName;
            string qurey = "SET IDENTITY_INSERT dbo.TableAuthRights ON;  insert into " + table + " (ID," + comlumnName+") values ('"+id.ToString()+"', N'" + data + "')";
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        
        private void insertRow(string source, string[] data)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string[] colList = new string[11];
            colList[0] = "رقم_المعاملة";
            colList[1] = "المعاملة";
            colList[2] = "المطلوب_رقم1";
            colList[3] = "المطلوب_رقم2";
            colList[4] = "المطلوب_رقم3";
            colList[5] = "المطلوب_رقم4";
            colList[6] = "المطلوب_رقم5";
            colList[7] = "المطلوب_رقم6";
            colList[8] = "المطلوب_رقم7";
            colList[9] = "المطلوب_رقم8";
            colList[10] = "المطلوب_رقم9";
            string item = "رقم_المعاملة";
            string value = "@رقم_المعاملة";
            for (int col = 1; col < 11; col++) {
                item = item + "," + colList[col];
                value = value + ",@" + colList[col];
            }

            string query = "INSERT INTO TableProcReq (" + item + ") values (" + value + ")";

            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            for (int col = 0; col < 11; col++)
            {
                //MessageBox.Show(colList[col] + ","+data[col]);

                sqlCmd.Parameters.AddWithValue(colList[col], data[col]);
            }
            try
            {
                sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex) 
            { 
            }
            sqlCon.Close();
        }
        

        private void AuthArchDocx(string source, int id, string filePath2, string name)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            
                    
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return; }
                SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET Data3=@Data3, المكاتبة_الاولية=@المكاتبة_الاولية  WHERE ID=@id", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@id", id);
                using (Stream stream = File.OpenRead(filePath2))
                {
                    byte[] buffer2 = new byte[stream.Length];
                    stream.Read(buffer2, 0, buffer2.Length);
                    var fileinfo2 = new FileInfo(filePath2);
                    string extn2 = fileinfo2.Extension;
                    string DocName2 = fileinfo2.Name;
                    sqlCmd.Parameters.Add("@Data3", SqlDbType.VarBinary).Value = buffer2;
                    sqlCmd.Parameters.Add("@المكاتبة_الاولية", SqlDbType.NVarChar).Value = DocName2;

                }
                sqlCmd.ExecuteNonQuery();
                sqlCon.Close();
            }
            catch (Exception ex) { }
        }

        private void AuthArch(bool state, string source, int id, string[] qureyData, string file2, string name)
        {

            SqlConnection sqlCon = new SqlConnection(source);
            
            var fileinfo2 = new FileInfo(file2);
            string extn1, extn2;
            string DocName1, DocName2;
            byte[] buffer1, buffer2;
            using (Stream stream = File.OpenRead(file2))
            {
                buffer2 = new byte[stream.Length];
                stream.Read(buffer2, 0, buffer2.Length);

                extn2 = fileinfo2.Extension;
                DocName2 = fileinfo2.Name;
            }

            SqlCommand sqlCmd = new SqlCommand(qureyData[3], sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.Parameters.AddWithValue("@Data2", buffer2);
            sqlCmd.ExecuteNonQuery();

            sqlCmd = new SqlCommand(qureyData[4], sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.Parameters.AddWithValue("@Extension2", extn2);
            sqlCmd.ExecuteNonQuery();

            sqlCmd = new SqlCommand(qureyData[5], sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", id);
            if (state)
                sqlCmd.Parameters.AddWithValue("@FileName2", DocName2);
            else sqlCmd.Parameters.AddWithValue("@المكاتبة_النهائية", DocName2);
            sqlCmd.ExecuteNonQuery();
            sqlCmd = new SqlCommand(qureyData[6], sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", id);
            if (state)
                sqlCmd.Parameters.AddWithValue("@ArchivedState", "مؤرشف");
            else sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", "مؤرشف");
            sqlCmd.ExecuteNonQuery();
            
            sqlCon.Close();
        }

        private void FinalDataArch(string dataSource, string filePath,int id)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return; }
                SqlCommand sqlCmd = new SqlCommand("UPDATE TableGeneralArch SET Data1=@Data1,fileUpload=@fileUpload WHERE ID=@ID", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.Add("@ID", SqlDbType.Int).Value = id;
                //MessageBox.Show(filePath);
                using (Stream stream = File.OpenRead(filePath))
                {
                    byte[] buffer1 = new byte[stream.Length];
                    stream.Read(buffer1, 0, buffer1.Length);
                    var fileinfo1 = new FileInfo(filePath);
                    string extn1 = fileinfo1.Extension;
                    string DocName1 = fileinfo1.Name;
                    sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                    sqlCmd.Parameters.Add("@fileUpload", SqlDbType.Char).Value = "Yes";
                    //Console.WriteLine("File uploaded " + filePath);
                }
                sqlCmd.ExecuteNonQuery();

                sqlCon.Close();
            }
            catch (Exception x) { }
        }


        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) 
        {
            
        }

        private void dateTimeTo_ValueChanged_1(object sender, EventArgs e)
        {

            LastDate = true;
            if (FirstDate)
            {
                int rows = DailyList(dateTimeFrom.Text, dateTimeTo.Text);
                if (rows > 0)
                {
                    PrintReport.Enabled = true;
                    PrintReport.Visible = true;
                    ReportPanel.Height = 205;
                }
                else
                {
                    PrintReport.Enabled = false;
                    PrintReport.Visible = false;
                    ReportPanel.Height = 42;
                    MessageBox.Show("لا يوجد قائمة بالتاريخ المحدد");
                }
            }
        }

        private void dateTimeFrom_ValueChanged_1(object sender, EventArgs e)
        {
            FirstDate = true;
            if (LastDate)
            {
                int rows = DailyList(dateTimeFrom.Text, dateTimeTo.Text);
                if (rows > 0)
                {
                    PrintReport.Enabled = true;
                    PrintReport.Visible = true;
                    ReportPanel.Height = 205;
                }
                else
                {
                    PrintReport.Enabled = false;
                    PrintReport.Visible = false;
                    ReportPanel.Height = 42;
                    MessageBox.Show("لا يوجد قائمة بالتاريخ المحدد");
                }
            }
            string Currentmonth = "", CurrentDay = "";
            int year, month, date, m = 0, d = 0;
            DateTime datetime = dateTimeFrom.Value;
            string[] YearMonthDayS = dateTimeFrom.Text.Split('-');
            year = Convert.ToInt16(YearMonthDayS[0]);
            m = Convert.ToInt16(YearMonthDayS[1]);
            d = Convert.ToInt16(YearMonthDayS[2]);


            if (m < 10) Currentmonth = "0" + m.ToString();
            else Currentmonth = m.ToString();
            if (d < 10) CurrentDay = "0" + d.ToString();
            else CurrentDay = d.ToString();
            string selecteddate =  Currentmonth.ToString() + "-" + CurrentDay.ToString() + "-" +year.ToString();
            DailyList(selecteddate);
            if (ReportType.SelectedIndex == 2 && (totalrowsAuth > 0 || totalrowsAffadivit > 0))
            {
                PrintReport.Enabled = true;
                PrintReport.Visible = true;
                ReportPanel.Height = 205;
            }
            else
            {
                PrintReport.Enabled = false;
                PrintReport.Visible = false;
                ReportPanel.Height = 42;
                MessageBox.Show("لا يوجد قائمة بالتاريخ المحدد");
            }
        }

        private void timer2_Tick_1(object sender, EventArgs e)
        { 
            //CultureInfo arSA = new CultureInfo("ar-SA");
            //arSA.DateTimeFormat.Calendar = new HijriCalendar();
            //Thread.CurrentThread.CurrentCulture = arSA;
            //int Ddiffer = HijriDateDifferment(DataSource, true);
            //int Mdiffer = HijriDateDifferment(DataSource, false);
            //string Stringdate, Stringmonth, StrHijriDate;
            //StrHijriDate = DateTime.Now.ToString("dd-MM-yyyy");
            //string[] YearMonthDay = StrHijriDate.Split('-');
            //int year, month, date;
            //year = Convert.ToInt16(YearMonthDay[2]);
            //month = Convert.ToInt16(YearMonthDay[1]) + Mdiffer;
            //date = Convert.ToInt16(YearMonthDay[0]) + Ddiffer;
            //if (month < 10) Stringmonth = "0" + month.ToString();
            //else Stringmonth = month.ToString();
            //if (date < 10) Stringdate = "0" + date.ToString();
            //else Stringdate = date.ToString();
            //HijriDate = Stringdate + "-" + Stringmonth + "-" + year.ToString();
            
            //CultureInfo arSA = new CultureInfo("ar-SA");
            //arSA.DateTimeFormat.Calendar = new HijriCalendar();
            //Thread.CurrentThread.CurrentCulture = arSA;
            //HijriDate = DateTime.Now.ToString("dd-MM-yyyy");
        }

        
        int countTimer = 0;
        private void timer1_Tick_1(object sender, EventArgs e)
        {
            
            //CultureInfo arSA = new CultureInfo("ar-SA");
            //arSA.DateTimeFormat.Calendar = new GregorianCalendar();
            //Thread.CurrentThread.CurrentCulture = arSA;
            //new System.Globalization.GregorianCalendar();
            //GregorianDate = DateTime.Now.ToString("dd-MM-yyyy");
            //if(dataGridView8.Visible && countTimer != 0) ColorFulGrid9();
        }

        private void txtModel_TextChanged(object sender, EventArgs e)
        {

        }


        public int FillDataGridView(String DataSource, int id)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return -1; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("AuthViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@مقدم_الطلب", "");
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView6.DataSource = dtbl;
            dataGridView6.Sort(dataGridView6.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            
            sqlCon.Close();
            int rows = 0;
            foreach (DataGridViewRow row in dataGridView6.Rows) {
                if (dataGridView6.Rows[rows].Cells[0].Value.ToString() == id.ToString())
                    return rows;
                else rows++;
            }
            return -1;
        }

        private void GoToForm(int indexNo, int locaIDNo)
        {
            dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Not Allowed");
            switch (indexNo)
            {
                case 0:
                    //Form1 form1 = new Form1(comboBox1.SelectedIndex,locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition);
                    //form1.ShowDialog();
                    break;
                case 1:

                    Form2 form2 = new Form2(attendedVC.SelectedIndex, locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form2.ShowDialog();
                    break;
                case 2:
                    Form3 form3 = new Form3(attendedVC.SelectedIndex, locaIDNo, 0, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form3.ShowDialog();
                    break;
                case 3:
                    Form4 form4 = new Form4(attendedVC.SelectedIndex, locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form4.ShowDialog();
                    break;
                case 4:
                    Form5 form5 = new Form5(attendedVC.SelectedIndex, locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form5.ShowDialog();
                    break;
                case 5:
                    Form6 form6 = new Form6(attendedVC.SelectedIndex, locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form6.ShowDialog();
                    break;
                case 6:
                    Form7 form7 = new Form7(attendedVC.SelectedIndex, locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form7.ShowDialog();
                    break;
                case 7:
                    Form8 form8 = new Form8(attendedVC.SelectedIndex, locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form8.ShowDialog();
                    break;
                case 8:
                    Form9 form9 = new Form9(attendedVC.SelectedIndex, locaIDNo, 0, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form9.ShowDialog();
                    break;
                case 9:
                    Form10 form10 = new Form10(attendedVC.SelectedIndex, locaIDNo, 0, EmployeeName, DataSource, FilespathIn, FilespathOut, UserJobposition, GregorianDate, HijriDate);
                    form10.ShowDialog();
                    break;
                case 10:
                    //Form11 form11  = new Form11(FillDataGridView(DataSource, locaIDNo), DataSource, FilespathIn, FilespathOut, EmployeeName, UserJobposition);
                    Form11 form11 = new Form11(attendedVC.SelectedIndex, locaIDNo, txtSearch.Text, DataSource, DataSource56,FilespathIn, FilespathOut, EmployeeName, UserJobposition, GregorianDate, HijriDate);
                    form11.ShowDialog();
                    break;
                default:
                    break;
            }
        }
        private void Open3File(int id, int fileNo)
        {
            string query;

            SqlConnection Con = new SqlConnection(DataSource);
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
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }
                else if (fileNo == 2)
                {
                    var name = reader["المكاتبة_النهائية"].ToString();
                    var Data = (byte[])reader["Data2"];
                    var ext = reader["Extension2"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }
                else
                {
                    var name = reader["DocxData"].ToString();
                    var Data = (byte[])reader["Data3"];
                    var ext = reader["Extension3"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }

            }
            Con.Close();


        }

    }

}

