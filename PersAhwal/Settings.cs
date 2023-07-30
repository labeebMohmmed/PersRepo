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
using System.IO;
using System.Data.SqlTypes;
using DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;
using static Azure.Core.HttpHeader;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using SixLabors.ImageSharp.Drawing;
using System.Globalization;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;
namespace PersAhwal
{
    public partial class Settings : Form
    {
        private string DataSource56, DataSource57, FilepathIn, FilepathOut, ArchFile, FormDataFile;
        private static bool NewSettings = false;
        string comboBoxOptions1 = "", comboBoxOptions2 = "";
        string[] txtComboOptions = new string[5] { "", "", "", "", "" };
        string[] DPTitle = new string[5];
        int pTextHieght = 42;
        int pComboHieght = 42;
        int pCheckHieght = 42;
        int pDateHieght = 42;
        int pbuttonHieght = 42;
        string ColumnName = "";
        bool NewColumn = false;
        string AuthBody1 = "لينوب عني ويقوم مقامي في ";
        int combo1index = 0, combo2index = 0, combo3index = 0, combo4index = 0, combo5index = 0;
        int id = 0;
        int idIndex = 1;
        bool review1 = false;
        string RightColumnName = "";
        int Nobox = 0;
        int listchecked = 0;
        DataTable checkboxdt;
        int LastID = 0, LastTabIndex = 0;
        static string[,] preffix = new string[10, 20];
        static string[] Text_statis = new string[5];
        static int[] statistic = new int[100];
        static int[] staticIndex = new int[100];
        static int[] times = new int[100];
        string[] allList = new string[100];
        static string[] Empty = new string[1] { "" };
        string Server = "M";
        string DataSource;
        int CombAuthTypeIndex = 1;
        string updateAll = "";
        string insertAll = "";
        string[] errors;
        string editRights;
        string errorList;
        string CurrentFile = "";
        bool readyToUpload = false;
        int ProcReqID = 0;
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;
        int rCnt;
        int cCnt;
        int rw = 0;
        int cl = 0;
        string revised = "";
        bool checkIndex = false;
        string formNo = "";
        string[] itemsicheck1 = new string[5];
        bool AuthType = true;
        bool reqGrid = false;
        string[] IDList = new string[100];
        string[] rightColNames;
        public Settings(string server, bool newSettings, string dataSource56, string dataSource57, bool setDataBase, string filepathIn, string filepathOut, string archFile, string formDataFile, string colName)
        {
            InitializeComponent();
            Server = server;
            DataSource56 = dataSource56;
            DataSource57 = dataSource57;

            if (Server == "57")
                DataSource = DataSource57;
            else if (Server == "56")
                DataSource = DataSource56;
            allList = getColList("TableAddContext");
            detectCharacter1(DataSource);
            detectCharacter2(DataSource);            
            NewSettings = newSettings;
            FilepathIn = filepathIn + @"\";
            FilepathOut = filepathOut;
            ArchFile = archFile;
            FormDataFile = formDataFile;
            if (!setDataBase)
            {
                if (!newSettings) loadSettings();
                else
                {
                    SaveSettings.Text = "أدخل بيانات قاعدة بيانات صحيحة";
                    MessageBox.Show("لا توجد قاعدة بيانات مسجلة");
                }
            }

            panelMainFiles.Location = Settingspanel.Location = missioInfopanel.Location = new System.Drawing.Point(4, 2);
            
            Suffex_preffixList();
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
        }

        private void testFiles()
        {            
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();



            string[] serverfiles = Directory.GetFiles(@"\\192.168.100.100\Users\Public\Documents\ModelFiles - Copy (3)");
            for (int i = 0; i < serverfiles.Length; i++)
            {
                MessageBox.Show(serverfiles[i]);
                CreateAuth(serverfiles[i], @"D:\ArchiveFiles\" + i.ToString() + ".docx");
                //var serverfileinfo = new FileInfo(serverfiles[i]);
                //string serverfilename = serverfileinfo.Name;
                //string serverLastWrite = serverfileinfo.LastWriteTime.ToShortTimeString();
                //string localFile = FilespathIn + serverfilename;

                //if (!File.Exists(localFile))
                //{
                //    System.IO.File.Copy(serverfiles[i], localFile);

                //}
                //else //if (File.Exists(localFile))
                //{
                //    //MessageBox.Show(serverfiles[i]);
                //    //MessageBox.Show(localFile);

                //    var localfileinfo = new FileInfo(localFile);
                //    string localLastWrite = localfileinfo.LastWriteTime.ToShortTimeString();

                //    //MessageBox.Show(serverLastWrite.Split(' ')[0] +"-" +localLastWrite.Split(' ')[0]);
                //    if (serverLastWrite.Split(' ')[0] != localLastWrite.Split(' ')[0])
                //    {

                //        try
                //        {
                //            File.Delete(localFile);
                //        }
                //        catch (Exception ex) { Console.WriteLine("الملف يحتاج إلى معالجة " + localFile); }
                //        System.IO.File.Copy(serverfiles[i], localFile);
                //    }
                //}
            }


        }
        private void CreateAuth(string DocxInFile, string DocxOutFile)
        {
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object objCurrentCopy = DocxInFile;
            Word.Document oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);
            oBMicroWord.Selection.Find.ClearFormatting();
            oBMicroWord.Selection.Find.Replacement.ClearFormatting();
            object ParaAuthIDNo = "MarkOtherTitle";
            try
            {
                Word.Range BookAuthIDNo = oBDoc.Bookmarks.get_Item(ref ParaAuthIDNo).Range;
            
            BookAuthIDNo.Text = "نائب قنصل";
            object rangeAuthIDNo = BookAuthIDNo;
            oBDoc.Bookmarks.Add("MarkOtherTitle", ref rangeAuthIDNo);
            }
            catch (Exception ex) {
                MessageBox.Show(DocxInFile);
                return;
            }
            oBDoc.SaveAs2(DocxOutFile);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(DocxOutFile);
        }
        private void detectCharacter1(string source)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource57);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return ; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('TableAuthRights')", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow row in dtbl.Rows)
            {
                string col = row["name"].ToString();
                //MessageBox.Show(col);
                if (col != "ID")
                {
                    sqlCon = new SqlConnection(DataSource57);
                    try
                    {
                        if (sqlCon.State == ConnectionState.Closed)
                            sqlCon.Open();
                    }
                    catch (Exception ex) { return; }
                    sqlDa = new SqlDataAdapter("SELECT * FROM TableAuthRights", sqlCon);
                    sqlDa.SelectCommand.CommandType = CommandType.Text;
                    DataTable dtblcol = new DataTable();
                    sqlDa.Fill(dtblcol);
                    sqlCon.Close();
                    foreach (DataRow rows in dtblcol.Rows)
                    {
                        //MessageBox.Show(rows[col].ToString());
                        string[] words = rows[col].ToString().Split('،')[0].Split(' ');                        
                        foreach (string str in words)
                        {
                            //MessageBox.Show(str);
                            if (str.Contains("_")) continue;
                            char[] strChar = str.ToCharArray();
                            foreach (char charItem in strChar)
                            {
                                if (!char.IsLetter(charItem) && charItem != '_' && charItem != '('&& charItem != ')' && !str.Contains("Col"))
                                {
                                    //if (!str.Split('،')[0].Contains("_"))
                                    //{
                                    if (!checkchar(str)) 
                                        insertChar(str);
                                        //MessageBox.Show(str);
                                        //break;
                                    //}
                                }
                            }
                        }
                    }
                }
            }
        }
        private void detectCharacter2(string source)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource57);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT TextModel FROM TableAddContext", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtblcol = new DataTable();
            sqlDa.Fill(dtblcol);
            sqlCon.Close();
            foreach (DataRow rows in dtblcol.Rows)
            {
                //MessageBox.Show(rows[col].ToString());
                string[] words = rows["TextModel"].ToString().Split('،')[0].Split(' ');
                foreach (string str in words)
                {
                    //MessageBox.Show(str);
                    if (str.Contains("_")) continue;
                    char[] strChar = str.ToCharArray();
                    foreach (char charItem in strChar)
                    {
                        if (!char.IsLetter(charItem) && charItem != '_' && charItem != '(' && charItem != ')' && !str.Contains("Col"))
                        {
                            //if (!str.Split('،')[0].Contains("_"))
                            //{
                            if (!checkchar(str))
                                insertChar(str);
                            //MessageBox.Show(str);
                            //break;
                            //}
                        }
                    }
                }
            }
        }

        private bool checkchar(string charStr)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource57);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT الرموز FROM Tablechar", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            foreach (DataRow row in dtbl.Rows)
            {

                if (row["الرموز"].ToString() == charStr)
                {
                    return true;
                }
            }
            //MessageBox.Show(table+" - "+ colName);
            return false;

        }
        private void insertChar(string charStr)
        {
            //MessageBox.Show(data[1]);
            SqlConnection sqlCon = new SqlConnection(DataSource);


            string query = "INSERT INTO Tablechar (الرموز) values (@الرموز)";

            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { }
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@الرموز", charStr);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        private void insertRow(string ColName)
        {
            //MessageBox.Show(data[1]);
            SqlConnection sqlCon = new SqlConnection(DataSource);


            string query = "INSERT INTO TableAddContext (ColName) values (@ColName)";

            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { }
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ColName", ColName);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }


        private bool PreReqFound(string proName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query = "SELECT رقم_المعاملة FROM TableProcReq where المعاملة=N'" + proName + "'";
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count > 0)
                return true;
            else return false;
        }

        private void sunInfo(string colName, int index)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query = "select [" + colName + "] from TableListCombo where [" + colName + "] is not null and [" + colName + "] <> ''";
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
                sqlCon.Close();
                foreach (DataRow row in dtbl.Rows)
                {
                    string formNo2 = row[colName].ToString().Trim() + "-" + index.ToString();

                    if (!TableAddContext(formNo2))
                    {
                        insertRow(formNo2);

                    }
                }
            }
            catch (Exception ex) { }
        }

        private bool TableAddContext(string proName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query = "SELECT ColName FROM TableAddContext where ColName=N'" + proName + "'";
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count > 0)
                return true;
            else return false;
        }

        private string[] getColList(string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource57);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return allList; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('" + table + "')", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            int i = 0;
            string insertItems = "";
            string insertValues = "";
            string updateValues = "";
            int colCount = 0;
            foreach (DataRow row in dtbl.Rows)
            {

                if (row["name"].ToString() != "ID" && row["name"].ToString() != "حالة_الارشفة" && row["name"].ToString() != "sms")
                {
                    colCount++;
                }
            }

            allList = new string[colCount];

            foreach (DataRow row in dtbl.Rows)
            {

                if (row["name"].ToString() != "ID" && row["name"].ToString() != "حالة_الارشفة" && row["name"].ToString() != "sms")
                {
                    allList[i] = row["name"].ToString();
                    //MessageBox.Show(row["name"].ToString());
                    if (i == 0)
                    {
                        insertItems = row["name"].ToString();
                        insertValues = "@" + row["name"].ToString();
                        updateValues = row["name"].ToString() + "=@" + row["name"].ToString();
                    }
                    else
                    {
                        insertItems = insertItems + "," + row["name"].ToString();
                        insertValues = insertValues + "," + "@" + row["name"].ToString();
                        updateValues = updateValues + "," + row["name"].ToString() + "=@" + row["name"].ToString();
                    }
                    i++;
                }
            }
            updateAll = "UPDATE " + table + " SET " + updateValues + " where ID = @id";
            insertAll = "INSERT INTO " + table + "(" + insertItems + ") values (" + insertValues + ")";

            return allList;

        }
        private bool checkColExist(string table, string colName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource57);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('" + table + "')", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            foreach (DataRow row in dtbl.Rows)
            {

                if (row["name"].ToString() == colName)
                {
                    return true;
                }
            }
            //MessageBox.Show(table+" - "+ colName);
            return false;

        }

        private void button17_Click(object sender, EventArgs e)
        {

        }

        private void txtModel_TextChanged(object sender, EventArgs e)
        {

        }

        private void button20_Click(object sender, EventArgs e)
        {

        }

        private void txtOutput_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void ArchiveFile_TextChanged(object sender, EventArgs e)
        {

        }

        private void button21_Click(object sender, EventArgs e)
        {

        }

        private void txtServerIP_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void txtDatabase_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void txtLogin_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void txtPass_TextChanged(object sender, EventArgs e)
        {

        }

        private void CreateColumn(string Columnname, string tableName)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("alter table " + tableName + " add " + Columnname + " nvarchar(150)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            try
            {
                sqlCmd.ExecuteNonQuery();
            }catch (Exception ex) { return; }   
            sqlCon.Close();
        }



        private string getAppFolder()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
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


        private void button3_Click(object sender, EventArgs e)
        {
            FolderApp.Text = getAppFolder();
            foreach(Control panel in this.Controls)
            {
                if(panel.Name.Contains("panel"))
                    panel.Visible = false;
            }
            Settingspanel.Visible = true;
        }



        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void button36_Click(object sender, EventArgs e)
        {
            
        }

        
        void deleteEmptyFields()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }

            SqlDataAdapter sqlDa = new SqlDataAdapter("delete from TableProcReq where المعاملة = '' or ((المطلوب_رقم1 = N'غير مدرج' or المطلوب_رقم1 = N'' ) and proForm1 is null)", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            
            try
            {
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);
            }
            catch (Exception ex) { }
            sqlCon.Close();
        }

        private string getColumnNames(string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return ""; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS " + table, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string cols = "";
            int index = 0;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(dataRow["COLUMN_NAME"].ToString()))
                {
                    if (!dataRow["COLUMN_NAME"].ToString().Contains("Data1"))
                    {
                        if (index == 0)
                            cols = dataRow["COLUMN_NAME"].ToString();
                        else cols = cols + "," + dataRow["COLUMN_NAME"].ToString();
                        index++;
                    }
                }
            }
            return cols;
        }

        
        private void Settings_Load(object sender, EventArgs e)
        {
            
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
                    if (!string.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                    {
                        for (int x = 0; x < Textboxtable.Rows.Count; x++)
                            if (dataRow[comlumnName].ToString().Equals(Textboxtable.Rows[x]))
                                newSrt = false;

                        if (newSrt) autoComplete.Add(dataRow[comlumnName].ToString());
                    }
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }

        private int getLastID(string source, string comlumnName, string tableName)
        {
            int x = 1;
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select max(ID) as maxID from " + tableName + " where " + comlumnName + " is not null";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;

                //MessageBox.Show(query);
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                foreach (DataRow dataRow in table.Rows)
                {
                    try
                    {
                        return Convert.ToInt32(dataRow["maxID"].ToString()) + 1;
                    }
                    catch (Exception) {
                        return 1;
                    }
                }
                saConn.Close();
            }
            return 1;
        }


        private int getCurrentID(string source, string comlumnName, string tableName, string text)
        {
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select ID from "+ tableName+"  where " + comlumnName+" = N'" + text+"'";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;

                //MessageBox.Show(query);
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                
                foreach (DataRow dataRow in table.Rows)
                {
                    try
                    {
                        return Convert.ToInt32(dataRow["ID"].ToString());
                    }
                    catch (Exception) {
                        return 0;
                    }
                }
                saConn.Close();
            }
            return 0;
        }

        private void fillSubComboBox(ComboBox combbox, string source, string comlumnName, string tableName, bool select)
        {
            //MessageBox.Show("source += "+source);
            combbox.Visible = true;
            //MessageBox.Show(source);
            //MessageBox.Show(Server);
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
                        if (!String.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                        {
                            //MessageBox.Show(dataRow[comlumnName].ToString());
                            combbox.Items.Add(dataRow[comlumnName].ToString());
                        }
                }
                }
                catch (Exception ex) { }
                
                saConn.Close();
            }
            if (select && combbox.Items.Count > 0) combbox.SelectedIndex = 0;
        }

        

        private bool checkRequInfo(string proID)
        {
            
            using (SqlConnection saConn = new SqlConnection(DataSource))
            {
                saConn.Open();

                string query = "select ID from TableProcReq where المعاملة =N'" + proID + "'";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();

                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                saConn.Close();
                if (table.Rows.Count > 0)
                {
                    Console.WriteLine("المعاملة موجودة " + proID);
                    return true;
                }
                else return false;
            }
            
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

                    if (dataRow["Lang"].ToString() == Language && dataRow["ColRight"].ToString() != "" && dataRow["ColName"].ToString().Contains("-"))
                    {

                        if (dataRow["ColName"].ToString().Split('-')[1].All(char.IsDigit))
                        {
                            try
                            {
                                if (id == dataRow["ColName"].ToString().Split('-')[1])
                                {
                                    //MessageBox.Show(dataRow["ColName"].ToString().Split('-')[0]);
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

        private void newFillComboBox2(ComboBox combbox, string source, string id, string Language)
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

                    if (dataRow["Lang"].ToString() == Language && dataRow["ColRight"].ToString() == "" && !String.IsNullOrEmpty(dataRow["ColName"].ToString()) && dataRow["ColName"].ToString().Contains("-"))
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

        private bool ComboTest(ComboBox combbox, string source, string comlumnName, string tableName)
        {
            combbox.Visible = true;
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

                    if (dataRow[comlumnName].ToString() == combbox.Text)
                    {
                        return true;
                    }
                }
                saConn.Close();
            }
            return false;
        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        


        private void newCombAuthType_TextChanged(object sender, EventArgs e)
        {

        }

        private void CombAuthType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ComboProcedure_SelectedIndexChanged(object sender, EventArgs e)
        {


        }

        private void button37_Click(object sender, EventArgs e)
        {

        }

        private void addMainAuth(string colText, string col)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableListCombo ("+ col+") values (@"+ col+")", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@"+col, colText);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private bool checkColumnName(string colNo)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return false; }
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

        private void getColumnName(string colNo)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableListCombo", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(dataRow["COLUMN_NAME"].ToString()))
                {
                    if (dataRow["COLUMN_NAME"].ToString().Contains(colNo))
                    {
                        ColumnName = dataRow["COLUMN_NAME"].ToString();
                        return;
                    }
                    else ColumnName = "";
                }
            }
        }

       

        private void addSubAuth(int id, string colText, string ColName)
        {
            string str = "@" + ColName;
            string query = "update TableListCombo set " + ColName + "=" + str + " where ID=@ID";
            //MessageBox.Show(query);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue(str, colText);
            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void FolderAppUpdate(string folderApp)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("update TableSettings set FolderApp=@FolderApp where ID='1'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@FolderApp", folderApp);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }


        private void checkSexType_CheckedChanged_2(object sender, EventArgs e)
        {

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


            preffix[0, 6] = "";//#5
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


            preffix[0, 9] = "نصيبي";//#6
            preffix[1, 9] = "نصيبي";
            preffix[2, 9] = "نصيبينا";
            preffix[3, 9] = "نصيبينا";
            preffix[4, 9] = "أنصبتنا";
            preffix[5, 9] = "أنصبتنا";


            preffix[0, 10] = "ت";//#*#
            preffix[1, 10] = "";



            preffix[0, 11] = "تنازلت تنازلاً نهائياً";//&&&
            preffix[1, 11] = "تنازلت تنازلاً نهائياً";
            preffix[2, 11] = "تنازلنا تنازلاً نهائياً";
            preffix[3, 11] = "تنازلنا تنازلاً نهائياً";
            preffix[4, 11] = "تنازلنا تنازلاً نهائياً";
            preffix[5, 11] = "تنازلنا تنازلاً نهائياً";


            preffix[0, 12] = "ي";//"%&%
            preffix[1, 12] = "ي";
            preffix[2, 12] = "نا";
            preffix[3, 12] = "نا";
            preffix[4, 12] = "نا";
            preffix[5, 12] = "نا";


            preffix[0, 13] = "نت";//#$#
            preffix[1, 13] = "نت";
            preffix[2, 13] = "نا";
            preffix[3, 13] = "نا";
            preffix[4, 13] = "نا";
            preffix[5, 13] = "نا";

            preffix[0, 14] = "أ";//&^&
            preffix[1, 14] = "إ";
            preffix[2, 14] = "ن";
            preffix[3, 14] = "ن";
            preffix[4, 14] = "ن";
            preffix[5, 14] = "ن";

            preffix[0, 15] = "ي";//&^^
            preffix[1, 15] = "ت";
            preffix[2, 15] = "ي";
            preffix[3, 15] = "ت";
            preffix[4, 15] = "ي";
            preffix[5, 15] = "ي";

            preffix[0, 16] = "";//*%*
            preffix[1, 16] = "";
            preffix[2, 16] = "ا";
            preffix[3, 16] = "ا";
            preffix[4, 16] = "ن";
            preffix[5, 16] = "وا";


        }



        private void UpdateColumn(string source, string comlumnName, int id, string data, bool datatype)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string column = "@" + comlumnName;
            string qurey;
            if (datatype) qurey = "INSERT INTO TableAuthRights (" + comlumnName + ") values(" + column + ")";
            else qurey = "UPDATE TableAuthRights SET " + comlumnName + " = " + column + " WHERE ID = @ID";

            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
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


        private void button39_Click(object sender, EventArgs e)
        {

        }



        private void button127_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            FolderApp.Text = dlg.FileName;
        }

        private void button121_Click(object sender, EventArgs e)
        {

        }

        private void button121_Click_1(object sender, EventArgs e)
        {
            FolderAppUpdate(FolderApp.Text);
            FolderApp.Text = "";
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            pTextHieght += 42;
            
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            pTextHieght -= 42;
            
        }

        private void IqrarPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void flowLayoutPanel4_MouseHover(object sender, EventArgs e)
        {
            //if(flowLayoutPanel4.Height == 42) 
            //    flowLayoutPanel4.Height = pTextHieght;
        }



        private string OpenFile(string documenNo, bool printOut, Button button)
        {
            string query = "SELECT ID, proForm1,Data1, Extension1 from TableProcReq where المعاملة=@المعاملة";
            
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@المعاملة", SqlDbType.NVarChar).Value = documenNo;
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            button.Enabled = false;
            if (!Directory.Exists(ArchFile + @"\formUpdated"))
            {
                System.IO.Directory.CreateDirectory(ArchFile + @"\formUpdated");
            }

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                string str = reader["proForm1"].ToString();
                Console.WriteLine(str);
                try
                {
                    var Data = (byte[])reader["Data1"];

                    CurrentFile = ArchFile + @"\formUpdated\" + str + ".docx";
                string filePath = ArchFile + @"\" + str +".docx";
                if (File.Exists(CurrentFile) && !fileIsOpen(CurrentFile)) {
                    File.Delete(CurrentFile);
                }
                
                if (!File.Exists(CurrentFile))
                {
                    try
                    {
                        //File.Delete(CurrentFile);

                        button.Enabled = true;

                        if (printOut)
                        {
                            File.WriteAllBytes(filePath, Data);

                            System.IO.File.Copy(filePath, CurrentFile);

                            FileInfo fileInfo = new FileInfo(CurrentFile);
                            if (fileInfo.IsReadOnly) fileInfo.IsReadOnly = false;
                                Console.WriteLine("CurrentFile " + CurrentFile);
                                System.Diagnostics.Process.Start(CurrentFile);
                        }
                        return CurrentFile;
                    }
                    catch (Exception ex)
                    {
                            Console.WriteLine("fail " +str);
                            button.Enabled = false;
                        return "";
                    }

                }
                else if (File.Exists(CurrentFile) && fileIsOpen(CurrentFile))
                {
                    button.Enabled = false;
                    MessageBox.Show("يرجى إغلاق الملف " + str + " أولا");
                }
                }
                catch (Exception ex)
                {
                    button.Enabled = false;
                    return "";
                }
            }
            else button.Enabled = false;
            Con.Close();
            return "";
        }


        private void uploadForms(string location)
        {
            if (location != "" && File.Exists(location) && !fileIsOpen(location))
            {
                using (Stream stream = File.OpenRead(location))
                {
                    byte[] buffer1 = new byte[stream.Length];
                    stream.Read(buffer1, 0, buffer1.Length);
                    var fileinfo1 = new FileInfo(location);
                    string formNo = fileinfo1.Name.Replace(".docx","");
                    string query = "UPDATE TableProcReq SET Data1=@Data1 WHERE المعاملة=N'" + formNo + "'";
                    //MessageBox.Show(query);
                    SqlConnection sqlCon = new SqlConnection(DataSource);
                    try
                    {
                        if (sqlCon.State == ConnectionState.Closed)
                            sqlCon.Open();
                    }
                    catch (Exception ex) { return; }
                    SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                    sqlCmd.ExecuteNonQuery();
                    sqlCon.Close();
                    try
                    {
                        //File.Delete(CurrentFile);
                    }
                    catch (Exception ex) { CurrentFile = ""; }
                    return;
                }
            }
        }
        
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


        private void checkBox2_CheckedChanged_1(object sender, EventArgs e)
        {

        }



        private void UpdateColumn(string source, string comlumnName, int id, string data)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string column = "@" + comlumnName;
            string qurey = "UPDATE TableAuthRights SET " + comlumnName + " = " + column + " WHERE ID=@ID";
            //MessageBox.Show(qurey);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;

            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.Parameters.AddWithValue(column, data.Trim());
            sqlCmd.ExecuteNonQuery();

            sqlCon.Close();
        }

        private void comboRights_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void newFillComboBox1(ComboBox combbox, string source, string colName)
        {
            combbox.Visible = true;
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();
                string query = "select " + colName + " from TableListCombo where " + colName + " is not null";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                foreach (DataRow dataRow in table.Rows)
                {
                    combbox.Items.Add(dataRow[colName].ToString());
                }
                saConn.Close();
            }
            //if (combbox.Items.Count > 0) combbox.SelectedIndex = 0;
        }
        private void Settings_FormClosed(object sender, FormClosedEventArgs e)
        {
            string primeryLink = @"D:\PrimariFiles\";
            if (!Directory.Exists(@"D:\"))
            {
                ///وله بموجب هذا التوكيل الحق في استخراج الشهادة، ومقابلة كافة الجهات المختصة في كل من مصلحة الأراضي والتسجيلات والمساحة، والوقوف والمقاضاة نيابة عني أمام كافة المحاكم والنيابات بمختلف أنواعها ودرجاتها، والقيام بكافة الإجراءات التي تتطلب حضوري، والتوقيع نيابة عني على كافة الأوراق والمستندات اللازمة لذلك، وله الحق في توكيل الغير في بعض أو كل مما أوكل فيه، وحظر قطعة الارض وأذنت لمن يشهد والله خير الشاهدين
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


        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName, bool clear)
        {

            if (clear) combbox.Items.Clear();
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


        private void insertRow(string source, string[] data)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string[] colList = new string[11];
            colList[0] = "المعاملة"; 
            colList[1] = "رقم_المعاملة";
            colList[2] = "المطلوب_رقم1";
            colList[3] = "المطلوب_رقم2";
            colList[4] = "المطلوب_رقم3";
            colList[5] = "المطلوب_رقم4";
            colList[6] = "المطلوب_رقم5";
            colList[7] = "المطلوب_رقم6";
            colList[8] = "المطلوب_رقم7";
            colList[9] = "المطلوب_رقم8";
            colList[10] = "المطلوب_رقم9";
            string item = "المعاملة";
            string value = "@المعاملة";
            for (int col = 1; col < 11; col++)
            {
                item = item + "," + colList[col];
                value = value + ",@" + colList[col];
            }

            string query = "INSERT INTO TableProcReq (" + item + ") values (" + value + ")";

            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            Console.WriteLine(query);
            //MessageBox.Show(query);
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
                MessageBox.Show(query);
            }
            sqlCon.Close();
        }


        private void txtSearch_MouseClick(object sender, MouseEventArgs e)
        {
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


        private void المعاملة_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(المعاملة.Text);
        }

       

        private void subTypeAuth_TextUpdate(object sender, EventArgs e)
        {

        }

        

        private void button34_Click(object sender, EventArgs e)
        {

        }



        private void button35_Click_1(object sender, EventArgs e)
        {
            FolderApp.Text = getAppFolder();
            try
            {
                string[] info = missionBasicInfo().Split('*');
                txtArabName.Text = info[0];
                txtEngName.Text = info[1];
                txtMissionAddress.Text = info[2];
                txtMissionCode.Text = info[3];
            }
            catch (Exception ex)
            {

            }

        }
        private string missionBasicInfo()
        {
            
            string infoDet = "";
            foreach (Control panel in this.Controls)
            {
                if (panel.Name.Contains("panel"))
                    panel.Visible = false;
            }
            missioInfopanel.Visible = true;
            string query = "select بيانات_البعثة from TableSettings";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return ""; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
            }
            catch (Exception ex)
            {
                return "";
            }



            sqlCon.Close();

            foreach (DataRow dataRow in dtbl.Rows)
            {
                try
                {
                    infoDet = dataRow["بيانات_البعثة"].ToString();                    
                }
                catch (Exception ex)
                {

                }
            }
            return infoDet;
        }

        private Excel.Range newFun()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            ColumnNamesLoad();

            for (cCnt = 2; cCnt <= cl; cCnt++)
            {
                //Console.WriteLine("rightColNames " + rightColNames.Length.ToString() + " cCnt " + cCnt.ToString());
                string cols = "";
                try
                {
                    string colname = (string)(range.Cells[1, cCnt] as Excel.Range).Value2;

                    if (string.IsNullOrEmpty(colname)) continue;

                    cols = colname.Replace(" ", "_").Replace("-", "_");
                    if (!checkColumnNames(cols, ""))
                    {
                        CreateColumns(cols);
                        if (checkID("1"))
                            UpdateColumn1(DataSource, cols, 1, cols, "TableAuthRights");
                        else InsertColumn(DataSource, cols, 1, cols, "TableAuthRights");
                    }
                    else
                    {
                        UpdateColumn1(DataSource, cols, 1, cols, "TableAuthRights");
                    }
                    for (rCnt = 2; rCnt < rw; rCnt++)
                    {
                        try
                        {
                            string strData = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                            if (String.IsNullOrEmpty(strData)) strData = "";

                            if (checkID(rCnt.ToString()))
                                UpdateColumn1(DataSource, cols, rCnt, strData, "TableAuthRights");
                            else InsertColumn(DataSource, cols, rCnt, strData, "TableAuthRights");
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
                catch (Exception ex)
                {
                }
            }

            return range;
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
        private void InsertColumn(string source, string comlumnName, int id, string data, string table)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string column = "@" + comlumnName;
            string qurey = "SET IDENTITY_INSERT dbo." + table + " ON;  insert into " + table + " (ID," + comlumnName + ") values ('" + id.ToString() + "', N'" + data + "')";
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

        private void UpdateColumn1(string source, string comlumnName, int id, string data, string table)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string column = "@" + comlumnName;
            string qurey = "UPDATE " + table + " SET " + comlumnName + " = " + column + " WHERE ID=@ID";

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
            catch (Exception ex) { MessageBox.Show(column + "-" + data); }

            sqlCon.Close();
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

        private void button6_Click(object sender, EventArgs e)
        {
            FolderApp.Text = getAppFolder();
            foreach (Control panel in this.Controls)
            {
                if (panel.Name.Contains("panel"))
                    panel.Visible = false;
            }
            panelMainFiles.Visible = true;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string missioninfo = txtArabName.Text +"*"+ txtEngName.Text +"*"+ txtMissionAddress.Text+"*"+ txtMissionCode.Text;
            string query = "update TableSettings set بيانات_البعثة " + "=N'" + missioninfo + "' where ID=@ID";            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;            
            sqlCmd.Parameters.AddWithValue("@ID", "1");
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("FilesAddorEdit", sqlCon);
            sqlCmd.CommandType = CommandType.StoredProcedure;
            if (SaveSettings.Text == "حفظ")
            {
                sqlCmd.Parameters.AddWithValue("@ID", 1);
                sqlCmd.Parameters.AddWithValue("@mode", "Add");
                sqlCmd.Parameters.AddWithValue("@FileArchive", FileArchive.Text);
                sqlCmd.Parameters.AddWithValue("@TempOutput", TempOutput.Text);
                sqlCmd.Parameters.AddWithValue("@Modelfilespath", Modelfilespath.Text);
                sqlCmd.Parameters.AddWithValue("@FolderApp", FolderApp.Text);
                sqlCmd.Parameters.AddWithValue("@archDisk", archDisk.Text);
                sqlCmd.Parameters.AddWithValue("@workingDisk", workingDisk.Text);

            }
            else
            {
                sqlCmd.Parameters.AddWithValue("@ID", 1);
                sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                sqlCmd.Parameters.AddWithValue("@FileArchive", FileArchive.Text);
                sqlCmd.Parameters.AddWithValue("@TempOutput", TempOutput.Text);
                sqlCmd.Parameters.AddWithValue("@Modelfilespath", Modelfilespath.Text);
                sqlCmd.Parameters.AddWithValue("@FolderApp", FolderApp.Text);
                sqlCmd.Parameters.AddWithValue("@archDisk", archDisk.Text);
                sqlCmd.Parameters.AddWithValue("@workingDisk", workingDisk.Text);

            }
            sqlCmd.ExecuteNonQuery();
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
            catch (Exception ex)
            {
                // MessageBox.Show("query " + query + "DataSource " + DataSource);
            }
            sqlCon.Close();
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
                Console.WriteLine("colName[" + index.ToString() + "] " + colName[index]);
                index++;
            }
            return colName;
        }

        private void fileComboBoxload(ComboBox combbox, string source, string comlumnName, string tableName, bool order)
        {
            //MessageBox.Show("source += "+source);
            combbox.Visible = true;
            //MessageBox.Show(source);
            //MessageBox.Show(Server);
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName;
                if (order) query = "select " + comlumnName + " from " + tableName + " order by " + comlumnName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                foreach (DataRow dataRow in table.Rows)
                {
                    if (!String.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                        combbox.Items.Add(dataRow[comlumnName].ToString().Replace("-", "_"));
                }
                saConn.Close();
            }
            //if (combbox.Items.Count > 0) combbox.SelectedIndex = 0;
        }

        
        private void button127_Click_1(object sender, EventArgs e)
        {

        }

        private void button121_Click_2(object sender, EventArgs e)
        {

        }

        private bool checkProReq(string proName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query = "SELECT * FROM TableProcReq where المعاملة=N'" + proName + "'";
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
            }
            catch (Exception ex) { }
            sqlCon.Close();
            if (dtbl.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }
        



        private void deleteRowsData(int v1, string v2, string source)
        {


            string query;

            SqlConnection Con = new SqlConnection(DataSource);
            query = "DELETE FROM " + v2 + " where ID = @ID";

            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", v1);
            sqlCmd.ExecuteNonQuery();
            Con.Close();
        }


        private void loadSettings()
        {
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select Modelfilespath,TempOutput,ServerName,Serverlogin,ServerPass,serverDatabase,FileArchive,archDisk,workingDisk  from TableSettings where ID=@id", Con);
            sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
            if (Con.State == ConnectionState.Closed)
                try
                {
                    Con.Open();

                    var reader = sqlCmd1.ExecuteReader();

                    if (reader.Read())
                    {
                        NewSettings = true;
                        Modelfilespath.Text = reader["Modelfilespath"].ToString();
                        TempOutput.Text = reader["TempOutput"].ToString();
                        ServerName.Text = reader["ServerName"].ToString();
                        Serverlogin.Text = reader["Serverlogin"].ToString();
                        ServerPass.Text = reader["ServerPass"].ToString();
                        serverDatabase.Text = reader["serverDatabase"].ToString();
                        FileArchive.Text = reader["FileArchive"].ToString();
                        archDisk.Text = reader["archDisk"].ToString();
                        workingDisk.Text = reader["workingDisk"].ToString();
                        if (NewSettings)
                        {
                            SaveSettings.Text = "تعديل";
                            NewSettings = false;
                        }
                    }
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    Con.Close();
                }

        }


        private void SaveSettings_Click_1(object sender, EventArgs e)
        {
            if (NewSettings)
            {
                DataSource = "Data Source=" + ServerName.Text + ";Network Library=DBMSSOCN;Initial Catalog=" + serverDatabase.Text + ";User ID=" + Serverlogin.Text + ";Password=" + ServerPass.Text;
                FilepathIn = Modelfilespath.Text;
                FilepathOut = TempOutput.Text;
            }
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    SqlCommand sqlCmd = new SqlCommand("NetConfAddorEdit", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    if (SaveSettings.Text == "حفظ")
                    {
                        sqlCmd.Parameters.AddWithValue("@ID", 1);
                        sqlCmd.Parameters.AddWithValue("@mode", "Add");
                        sqlCmd.Parameters.AddWithValue("@ServerName", ServerName.Text);
                        sqlCmd.Parameters.AddWithValue("@Serverlogin", Serverlogin.Text);
                        sqlCmd.Parameters.AddWithValue("@ServerPass", ServerPass.Text);
                        sqlCmd.Parameters.AddWithValue("@serverDatabase", serverDatabase.Text);
                        sqlCmd.ExecuteNonQuery();
                    }
                    else
                    {
                        sqlCmd.Parameters.AddWithValue("@ID", 1);
                        sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                        sqlCmd.Parameters.AddWithValue("@ServerName", ServerName.Text);
                        sqlCmd.Parameters.AddWithValue("@Serverlogin", Serverlogin.Text);
                        sqlCmd.Parameters.AddWithValue("@ServerPass", ServerPass.Text);
                        sqlCmd.Parameters.AddWithValue("@serverDatabase", serverDatabase.Text);
                        sqlCmd.ExecuteNonQuery();
                    }
                    sqlCon.Close();
                    this.Hide();
                    var formDataBase = new FormDataBase(Server, DataSource56, DataSource57, FilepathIn, FilepathOut, ArchFile, FormDataFile, "");
                    formDataBase.Closed += (s, args) => this.Close();
                    formDataBase.Show();
                }

                catch (Exception ex)
                {
                    MessageBox.Show("الوصول لقاعدة البيانات غير متاح");
                }
                finally
                {

                    clear_fields();

                }
        }

        private void clear_fields()
        {
            serverDatabase.Text = Serverlogin.Text = Modelfilespath.Text = TempOutput.Text = ServerPass.Text = ServerName.Text = FileArchive.Text = "";
        }
    }
}
