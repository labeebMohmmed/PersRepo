﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Globalization;
using System.Threading;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;

namespace PersAhwal
{
    public partial class Form6 : Form
    {
        public static string[] SponceDesc3 = new string[10];
        public static string[] SponceDoc3 = new string[10];
        public static string[] SponcerName3 = new string[10];
        public static string[] SponcePassIqama3 = new string[10];
        public static string[] ApplicantIdocNo3 = new string[10];
        public static string[] SponceIssueSource3 = new string[10];

        public static string[] DaughterMotheDocSource = new string[10];
        bool newData = false;
        bool SaveEdit = false;
        public static string route = "";
        string Viewed;
        string ConsulateEmpName;
        public static string ModelFileroute = "";
        String IfadaNumberPart;
        static string DataSource;
        string NewFileName;
        string PreAppId = "", PreRelatedID = "", NextRelId = "";
        static public string FamilySupport;
        private string[] FamelyMember = new string[10];
        string FilesPathIn, FilesPathOut;
        private int ApplicantID = 0;
        private bool fileloaded = false;
        string Jobposition;
        int ATVC = 0;
        string[] colIDs = new string[100];
        string GregorianDate = "";
        string HijriDate = "";
        string AuthTitle = "نائب قنصل";
        public Form6(int Atvc, int currentRow, string EmpName, string dataSource,  string filepathOut, string jobposition, string gregorianDate, string hijriDate)
        {
            InitializeComponent();
            //timer1.Enabled = true;
            //timer2.Enabled = true;
            التاريخ_الميلادي.Text = GregorianDate = gregorianDate;
            التاريخ_الهجري.Text = HijriDate = hijriDate;
            ATVC = Atvc;
            DataSource = dataSource;
            AttendViceConsul.SelectedIndex = 2;
            //FilesPathIn = filepathIn + @"\";
            FilesPathOut = filepathOut;
            colIDs [4] = ConsulateEmpName = EmpName;
            Jobposition = jobposition;
            Clear_Fields();
            FillDataGridView();
            if (currentRow == -1) Clear_Fields();
            else SetFieldswithData(currentRow);

            if (jobposition.Contains("قنصل"))
                btnEditID.Visible = button5.Visible = true;
            else btnEditID.Visible = button5.Visible = false;
            getTitle(DataSource, EmpName);
        }
        private void getTitle(string source, string empName)
        {
            string query = "select AuthenticType from TableUser where EmployeeName = N'" + empName + "'";
            SqlConnection sqlCon = new SqlConnection(source);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                AuthTitle = dataRow["AuthenticType"].ToString();
            }
        }

        private string loadRerNo(int id)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT DocID from TableForensicApp where ID=@ID", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ID", id);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string rowCnt = "";
            foreach (DataRow row in dtbl.Rows)
            {
                rowCnt = (Convert.ToInt32(row["DocID"].ToString().Split('/')[3]) + 1).ToString();
            }
            return rowCnt;

        }


        private int loadIDNo()
        {


            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from TableForensicApp order by ID desc", sqlCon);
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


        //private void OpenFileDoc(int id, int fileNo)
        //{
        //    string query;

        //    SqlConnection Con = new SqlConnection(DataSource);
        //    if (fileNo == 1)
        //    {
        //        query = "select Data1, Extension1,FileName1 from TableForensicApp  where ID=@id";
        //    }
        //    else if (fileNo == 2)
        //    {
        //        query = "select Data2, Extension2,FileName2 from TableForensicApp  where ID=@id";
        //    }
        //    else query = "select Data3, Extension3,FileName3 from TableForensicApp  where ID=@id";
        //    SqlCommand sqlCmd1 = new SqlCommand(query, Con);
        //    sqlCmd1.Parameters.Add("@Id", SqlDbType.Int).Value = id;
        //    if (Con.State == ConnectionState.Closed)
        //        Con.Open();

        //    var reader = sqlCmd1.ExecuteReader();
        //    if (reader.Read())
        //    {
        //        if (fileNo == 1)
        //        {
        //            var name = reader["FileName1"].ToString();
        //            var Data = (byte[])reader["Data1"];
        //            var ext = reader["Extension1"].ToString();
        //            var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
        //            File.WriteAllBytes(NewFileName, Data);
        //            System.Diagnostics.Process.Start(NewFileName);
        //        }
        //        else if (fileNo == 2)
        //        {
        //            var name = reader["FileName2"].ToString();
        //            var Data = (byte[])reader["Data2"];
        //            var ext = reader["Extension2"].ToString();
        //            var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
        //            File.WriteAllBytes(NewFileName, Data);
        //            System.Diagnostics.Process.Start(NewFileName);
        //        }
        //        else
        //        {
        //            var name = reader["FileName3"].ToString();
        //            var Data = (byte[])reader["Data3"];
        //            var ext = reader["Extension3"].ToString();
        //            var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
        //            File.WriteAllBytes(NewFileName, Data);
        //            System.Diagnostics.Process.Start(NewFileName);
        //        }

        //    }
        //    Con.Close();


        //}

        void FillDatafromGenArch(string doc, string id, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource.Replace("AhwalDataBase", "ArchFilesDB"));
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
        private void SetFieldswithData(int Rowindex)
        {
            Rowindex--;
            ApplicantID = Convert.ToInt32(dataGridView1.Rows[Rowindex].Cells[0].Value.ToString());
            PreAppId = dataGridView1.Rows[Rowindex].Cells[1].Value.ToString();
            مقدم_الطلب.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "ذكر") النوع.CheckState = CheckState.Unchecked;
            else if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "أنثى") النوع.CheckState = CheckState.Checked;
            رقم_الهوية.Text = dataGridView1.Rows[Rowindex].Cells[4].Value.ToString();
            مكان_الإصدار.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString();
            ApplicantIqamaNo.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString();
            IqamaIssuedSource.Text = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString();
            التاريخ_الميلادي.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString();
            التاريخ_الهجري.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[16].Value.ToString().ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                
            }
            else checkedViewed.CheckState = CheckState.Checked;

            AppType.Text = dataGridView1.Rows[Rowindex].Cells[12].Value.ToString();
            ConsulateEmployee.Text = dataGridView1.Rows[Rowindex].Cells[13].Value.ToString();
            if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked; else AppType.CheckState = CheckState.Unchecked;

            if (AppType.CheckState == CheckState.Unchecked)
            {
                mandoubVisibilty(); mandoubName.Text = dataGridView1.Rows[Rowindex].Cells[14].Value.ToString();
            }

            PreRelatedID = dataGridView1.Rows[Rowindex].Cells[15].Value.ToString();
            Comment.Text = dataGridView1.Rows[Rowindex].Cells[20].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[21].Value.ToString() != "غير مؤرشف")
            {
                ArchivedSt.CheckState = CheckState.Checked;
                ArchivedSt.Text = "مؤرشف";
                ArchivedSt.BackColor = Color.Green;
            }
            else
            {
                ArchivedSt.CheckState = CheckState.Unchecked;
                ArchivedSt.Text = "غير مؤرشف";
                ArchivedSt.BackColor = Color.Red;
            }
            ArchivedSt.Visible = true;
            
            btnSavePrint.Text = "حفظ";
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            MessageBox.Show("timer2");
            
        }

        private void timer1_Tick_1(object sender, EventArgs e)
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
            التاريخ_الهجري.Text = Stringdate + "-" + Stringmonth + "-" + year.ToString();
            timer1.Enabled = false;
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


        private void CreateWordFile()
        {
            string ReportName = DateTime.Now.ToString("mmss");
            if (النوع.CheckState == CheckState.Unchecked)
            {

                labelName.ForeColor = Color.Black;
                labelName.Text = "مقدم الطلب:";
                route = FilesPathIn + "ForesnecM.docx";
            }
            else if (النوع.CheckState == CheckState.Checked)
            {
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
                route = FilesPathIn + "ForesnecF.docx";
            }            
            string ActiveCopy;
            ActiveCopy = FilesPathOut + مقدم_الطلب.Text + ReportName + ".docx";
            if (!File.Exists(ActiveCopy))
            {
                System.IO.File.Copy(route, ActiveCopy);
                object oBMiss = System.Reflection.Missing.Value;
                Word.Application oBMicroWord = new Word.Application();
                object Routseparameter = ActiveCopy;
                Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);
                object ParaIfadaNo = "MarkIfadaNo";
                object ParaGreData = "MarkGreData";
                object ParaHijriData = "MarkHijriData";
                object Paraname = "MarkAppName";
                object ParaAplicantPass = "MarkAppPass";
                object ParaAppPassSource = "MarkAppPassSource";
                object ParaAplicantIqama = "MarkAppIqama";
                object ParaAppiIqamaSource = "MarkAppIqamaSource";
                object ParavConsul = "MarkViseConsul";

                Word.Range BookIfadaNo = oBDoc.Bookmarks.get_Item(ref ParaIfadaNo).Range;
                Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
                Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
                Word.Range Bookname = oBDoc.Bookmarks.get_Item(ref Paraname).Range;
                Word.Range BookAppPass = oBDoc.Bookmarks.get_Item(ref ParaAplicantPass).Range;
                Word.Range BookAppPassSource = oBDoc.Bookmarks.get_Item(ref ParaAppPassSource).Range;
                Word.Range BookAppIqama = oBDoc.Bookmarks.get_Item(ref ParaAplicantIqama).Range;
                Word.Range BookAppIqamaSource = oBDoc.Bookmarks.get_Item(ref ParaAppiIqamaSource).Range;
                Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;

                BookIfadaNo.Text = colIDs[0] = Ifadaid.Text;
                BookGreData.Text = التاريخ_الميلادي_off.Text;
                colIDs[2] = التاريخ_الميلادي.Text;
                BookHijriData.Text = التاريخ_الهجري.Text;
                Bookname.Text = colIDs[3] = مقدم_الطلب.Text;
                BookAppPass.Text = رقم_الهوية.Text;
                BookAppPassSource.Text = مكان_الإصدار.Text;
                BookAppIqama.Text = ApplicantIqamaNo.Text;
                BookAppIqamaSource.Text = IqamaIssuedSource.Text;
                BookvConsul.Text = AttendViceConsul.Text + Environment.NewLine + AuthTitle;
                colIDs[5] = AppType.Text;
                colIDs[6] = mandoubName.Text;
                object rangeIfadaNo = BookIfadaNo;
                object rangeGreData = BookGreData;
                object rangeHijriData = BookHijriData;
                object rangeName = Bookname;
                object rangeAppPass = BookAppPass;
                object rangeAppPassSource = BookAppPassSource;
                object rangeAppIqama = BookAppIqama;
                object rangeAppIqamaSource = BookAppIqamaSource;
                object rangevConsul = BookvConsul;

                oBDoc.Bookmarks.Add("MarkIfadaNo", ref rangeIfadaNo);
                oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
                oBDoc.Bookmarks.Add("MarkHijiData", ref rangeHijriData);
                oBDoc.Bookmarks.Add("MarkAppName", ref rangeName);
                oBDoc.Bookmarks.Add("MarkAppPass", ref rangeAppPass);
                oBDoc.Bookmarks.Add("MarkAppPassSource", ref rangeAppPassSource);
                oBDoc.Bookmarks.Add("MarkAppIqama", ref rangeAppIqama);
                oBDoc.Bookmarks.Add("MarkAppIqamaSource", ref rangeAppIqamaSource);
                oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);



                string docxouput = FilesPathOut + مقدم_الطلب.Text + DateTime.Now.ToString("ssmm") + ".docx";
                string pdfouput = FilesPathOut + مقدم_الطلب.Text + DateTime.Now.ToString("ssmm") + ".pdf";
                oBDoc.SaveAs2(docxouput);
                oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
                oBDoc.Close(false, oBMiss);
                oBMicroWord.Quit(false, false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
                System.Diagnostics.Process.Start(pdfouput);
                object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;

            }
            else
            {
                MessageBox.Show("يرجى حذف الملف الموجودأولاً");
                btnSavePrint.Enabled = true;

            }
            addarchives(colIDs);

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
                if(allList[i] == "appType") 
                    sqlCommand.Parameters.AddWithValue("@" + allList[i], "حضور مباشرة إلى القنصلية");
                else sqlCommand.Parameters.AddWithValue("@" + allList[i], text[i - 1]);
                //MessageBox.Show(text[i - 1]);
            }
            sqlCommand.ExecuteNonQuery();
        }
        private string[] getColList(string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('" + table + "')", sqlCon);
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


        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

        }

        private void mandoubVisibilty()
        {
            if (AppType.CheckState == CheckState.Checked)
            {
                AppType.Text = "حضور مباشرة إلى القنصلية";
                mandoubName.Visible = false;
                mandoubLabel.Visible = false;
            }
            else
            {
                AppType.Text = "عن طريق أحد مندوبي القنصلية";
                mandoubName.Visible = true;
                mandoubLabel.Visible = true;
            }
        }
        private void Form6_Load(object sender, EventArgs e)
        {
            
            
            
            autoCompleteTextBox(IqamaIssuedSource, DataSource, "SDNIssueSource", "TableListCombo");
            autoCompleteTextBox(مكان_الإصدار, DataSource, "SDNIssueSource", "TableListCombo");
            fileComboBox(AttendViceConsul, DataSource, "ArabicAttendVC", "TableListCombo");
            AttendViceConsul.SelectedIndex = ATVC;
            autoCompleteTextBox1(مقدم_الطلب, DataSource, "الاسم", "TableGenNames");

            fileComboBoxMandoub(mandoubName, DataSource, "TableMandoudList");
        }
        private void fileComboBoxMandoub(ComboBox combbox, string source, string tableName)
        {
            combbox.Visible = true;
            combbox.Items.Clear();
            combbox.Items.Add("حضور مباشرة إلى القنصلية");
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();
                string query = "select MandoubNames,MandoubAreas,وضع_المندوب from " + tableName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                foreach (DataRow dataRow in table.Rows)
                {
                    if (dataRow["MandoubNames"].ToString() != "" && dataRow["وضع_المندوب"].ToString() == "الحساب مفعل")
                        combbox.Items.Add(dataRow["MandoubNames"].ToString() + " - " + dataRow["MandoubAreas"].ToString());
                }
                saConn.Close();
            }
            if (combbox.Items.Count > 0)
                combbox.SelectedIndex = 0;
        }
        private void autoCompleteTextBox1(TextBox textbox, string source, string comlumnName, string tableName)
        {
            textbox.Multiline = false;
            //MessageBox.Show(textbox.Name);
            using (SqlConnection saConn = new SqlConnection(source))
            {
                if (saConn.State == ConnectionState.Closed)
                    try
                    {
                        saConn.Open();
                    }
                    catch (Exception ex) { }

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
                    string text = dataRow[comlumnName].ToString().Trim();
                    Console.WriteLine("autoCompleteTextBox " + text);
                    autoComplete.Add(text);
                }
                textbox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }

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
        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName)
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

        private void Review_Click(object sender, EventArgs e)
        {

        }


        private void Save2DataBase()
        {
            
            SqlConnection sqlCon = new SqlConnection(DataSource);

            string AppGender;
            if (النوع.CheckState == CheckState.Unchecked) AppGender = "ذكر"; else AppGender = "أنثى";
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                else Viewed = "غير معالج";
                SqlCommand sqlCmd = new SqlCommand("ForensicAddorEdit", sqlCon);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                
                if ( newData)
                {
                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");
                    sqlCmd.Parameters.AddWithValue("@DocID", Ifadaid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", مقدم_الطلب.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@PassNo", رقم_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@PassSource", مكان_الإصدار.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@IqamaNo", ApplicantIqamaNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@IqamaSource", IqamaIssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", PreAppId.Trim());
                    sqlCmd.Parameters.AddWithValue("@purpose", txtPurpose.Text.Trim());
                    
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");
                    sqlCmd.ExecuteNonQuery();
                }
                else
                {
                    sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                    sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                    sqlCmd.Parameters.AddWithValue("@DocID", Ifadaid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", مقدم_الطلب.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@PassNo", رقم_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@PassSource", مكان_الإصدار.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@IqamaNo", ApplicantIqamaNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@IqamaSource", IqamaIssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", PreAppId.Trim());
                    sqlCmd.Parameters.AddWithValue("@purpose", txtPurpose.Text.Trim());
                    
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
                    if (fileloaded)
                        sqlCmd.Parameters.AddWithValue("@ArchivedState", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    else sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف"); sqlCmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
            }
            finally
            {
                sqlCon.Close();
            }

        }

        private void FillDataGridView()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("ForensicViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ApplicantName", ListSearch.Text.Trim());
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            IfadaNumberPart = loadRerNo(loadIDNo());
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 200;
            sqlCon.Close();
            NewFileName = IfadaNumberPart + "_06";
        }

        private void AppType_CheckedChanged_1(object sender, EventArgs e)
        {
            mandoubVisibilty();
        }

        private void btnSavePrint_Click_1(object sender, EventArgs e)
        {
            getTitle(DataSource, AttendViceConsul.Text); 
            التاريخ_الميلادي.Text = GregorianDate;
            التاريخ_الهجري.Text = HijriDate;
            if (!checkGender(PanelMain, "مقدم_الطلب", "النوع"))
            {
                return;
            }
            else addNewAppNameInfo(مقدم_الطلب); 
            
            if (txtPurpose.Text == "" || txtPurpose.Text == "إختر أو أكتب الغرض") { MessageBox.Show("يرجى توضيح الغرض من الإجراء ");return; }
            Save2DataBase();
            CreateWordFile();
            
            btnSavePrint.Enabled = false;
            btnSavePrint.Text = "جاري المعالجة";
            
            this.Close();
            //Clear_Fields();
        }
        private void addNewAppNameInfo(TextBox textName)
        {

            string query = "insert into TableGenNames ([الاسم], رقم_الهوية,النوع,نوع_الهوية,مكان_الإصدار) values (@col1,@col2,@col5,@col6,@col7) ;SELECT @@IDENTITY as lastid";
            string id = checkExist(textName.Text);
            if (id != "0")
            {
                query = "update TableGenNames set [الاسم] =  @col1,[رقم_الهوية] = @col2,النوع = @col5,نوع_الهوية = @col6,مكان_الإصدار = @col7 where ID = " + id;
                //MessageBox.Show(query);
            }
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();

            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@col1", مقدم_الطلب.Text);
            sqlCommand.Parameters.AddWithValue("@col2", رقم_الهوية.Text);
            sqlCommand.Parameters.AddWithValue("@col5", النوع.Text);
            sqlCommand.Parameters.AddWithValue("@col6", "جواز سفر");
            sqlCommand.Parameters.AddWithValue("@col7", مكان_الإصدار.Text);
            var reader = sqlCommand.ExecuteReader();
            if (reader.Read())
            {
                //MessageBox.Show(reader["lastid"].ToString());
            }
            try
            {


            }
            catch (Exception ex)
            {
                MessageBox.Show("addNewAppNameInfo");
            }
        }
        private bool checkGender(Panel panel, string controlType, string control2type)
        {
            int index = 0;
            foreach (Control control in panel.Controls)
            {
                if (control.Name == controlType + index + ".")
                {
                    string gender = getGender(control.Text.Split(' ')[0]);
                    foreach (Control control2 in panel.Controls)
                    {
                        if (control2.Name == control2type + index + ".")
                        {
                            if (gender != control2.Text)
                            {
                                var selectedOption = MessageBox.Show("هل تود تغيير إعدادات البرنامج الداخلية والمتابعة للصفحة التالية؟", "يرجى مراحعة جنس   " + control.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                if (selectedOption == DialogResult.No)
                                {
                                    return false;
                                }
                                else if (selectedOption == DialogResult.Yes)
                                {
                                    updateGender(control2.Text, getSexIndex);
                                    return true;
                                }
                            }
                        }
                    }
                    index++;
                }
            }
            return true;
        }
        string getSexIndex = "0";
        public string getGender(string name)
        {
            string sex = "ذكر";
            string query = "SELECT ID,النوع FROM TableGenGender where الاسم = N'" + name + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            foreach (DataRow row in dtbl.Rows)
            {
                getSexIndex = row["ID"].ToString();
                sex = row["النوع"].ToString();
            }
            return sex;
        }

        private void updateGender(string newGender, string id)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    SqlCommand sqlCmd = new SqlCommand("UPDATE TableGenGender SET النوع=N'" + newGender + "' WHERE ID=" + id, sqlCon);
                    MessageBox.Show("UPDATE TableGenGender SET النوع=N'" + newGender + "' WHERE ID=" + id);
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.ExecuteNonQuery();
                    sqlCon.Close();

                }

                catch (Exception ex)
                {
                    return;
                }
                finally
                {
                }
        }
        public string checkExist(string name)
        {
            string id = "0";
            string query = "SELECT ID FROM TableGenNames where الاسم like N'" + name + "%'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            foreach (DataRow row in dtbl.Rows)
            {
                id = row["ID"].ToString();
            }
            return id;
        }
        private void Clear_Fields()
        {
            IqamaIssuedSource.Text = ApplicantIqamaNo.Text = مكان_الإصدار.Text = مقدم_الطلب.Text = "";
            النوع.CheckState = CheckState.Checked;
            رقم_الهوية.Text = "P0";
            النوع.CheckState = CheckState.Checked;
            AttendViceConsul.SelectedIndex = 2;
            mandoubName.Text = ListSearch.Text = "";
            AppType.CheckState = CheckState.Unchecked;
            mandoubVisibilty();
            newData = false;
            btnSavePrint.Text = "طباعة وحفظ";
            btnSavePrint.Enabled = true;
            Comment.Text = "لا تعليق";
            FillDataGridView();
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            
            ArchivedSt.BackColor = Color.Red;
            fileloaded = false;
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();

            التاريخ_الميلادي.Text = DateTime.Now.ToString("dd-MM-yyyy");
            ConsulateEmployee.Text = ConsulateEmpName;
        }

        private void btnprintOnly_Click(object sender, EventArgs e)
        {
            
            CreateWordFile();
            this.Close();
            //Clear_Fields();
        }

        

        private void button2_Click(object sender, EventArgs e)
        {
            var selectRows = dataGridView1.SelectedRows;
            foreach (var row in selectRows)
            {
                int id = (int)((DataGridViewRow)row).Cells[0].Value;
                OpenFile(id, 1);
            }
            if (ApplicantID != 0) OpenFile(ApplicantID, 1);
            ApplicantID = 0;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var selectRows = dataGridView1.SelectedRows;
            foreach (var row in selectRows)
            {
                int id = (int)((DataGridViewRow)row).Cells[0].Value;
                OpenFile(id, 2);
            }
            if (ApplicantID != 0) OpenFile(ApplicantID, 2);
            ApplicantID = 0;
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
            FillDataGridView();
        }


        private void timer2_Tick_1(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            التاريخ_الميلادي.Text = DateTime.Now.ToString("MM-dd-yyyy");
            timer2.Enabled = false;
        }

        private void ApplicantSex_CheckedChanged_1(object sender, EventArgs e)
        {
            if (النوع.CheckState == CheckState.Unchecked)
            {
                النوع.Text = "ذكر";
                labelName.Text = "مقدم الطلب:";
            }
            else if (النوع.CheckState == CheckState.Checked)
            {
                النوع.Text = "إنثى";
                labelName.Text = "مقدمة الطلب:";
            }
        }

        private void deleteRow_Click_1(object sender, EventArgs e)
        {
            
        }

        private void ListSearch_TextChanged(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '" + ListSearch.Text + "%'";
            dataGridView1.DataSource = bs;
            ColorFulGrid9();
        }
        private void ColorFulGrid9()
        {
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            dataGridView1.BackgroundColor = Color.White;

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;
                if (dataGridView1.Rows[i].Cells[21].Value.ToString() == "مؤرشف نهائي") dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;

                //else dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.;

            }
            //
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void deleteRow_Click(object sender, EventArgs e)
        {
            deleteRowsData(ApplicantID, "TableForensicApp", DataSource);
            deleteRow.Visible = false;
        }

        private void btnFile3_Click(object sender, EventArgs e)
        {
            //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 3);
            //FillDatafromGenArch("data2", colIDs[1], "TableFamilySponApp");
        }

        private void btnFile2_Click(object sender, EventArgs e)
        {
            //OpenFileDoc(/Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 2);
            FillDatafromGenArch("data2", colIDs[1], "TableForensicApp");
        }

        private void btnFile1_Click(object sender, EventArgs e)
        {
            //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 1);
            FillDatafromGenArch("data1", colIDs[1], "TableForensicApp");
        }

        private void SearchDoc_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            Search.Text = dlg.FileName;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                dataGridView1.Visible = false;
                PanelFiles.Visible = true;
                PanelMain.Visible = true;
                colIDs[1] = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                colIDs[0] = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                try
                {
                    txtEditID2.Text = colIDs[0].Split('/')[4];
                    txtEditID1.Text = colIDs[0].Replace(txtEditID2.Text, "");
                }
                catch (Exception ex)
                {

                }

                if (dataGridView1.CurrentRow.Cells[2].Value.ToString() == "")
                {
                    ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                    newData = false;
                    SaveEdit = true;
                    colIDs[7] = "new";
                    Ifadaid.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    txtPurpose.Text = dataGridView1.CurrentRow.Cells[22].Value.ToString();
                    //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 1);
                    FillDatafromGenArch("data1", colIDs[1], "TableForensicApp");
                    if (Jobposition.Contains("قنصل")) deleteRow.Visible = true;
                    return;
                }
                colIDs[7] = "old";
                SaveEdit = newData = false;
                ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                if (Jobposition.Contains("قنصل")) deleteRow.Visible = true;
                Ifadaid.Text = PreAppId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                مقدم_الطلب.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "ذكر") النوع.CheckState = CheckState.Unchecked;
                else if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "أنثى") النوع.CheckState = CheckState.Checked;
                رقم_الهوية.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                مكان_الإصدار.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                ApplicantIqamaNo.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                IqamaIssuedSource.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                التاريخ_الميلادي.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                التاريخ_الهجري.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[16].Value.ToString().ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked; 

                }
                else checkedViewed.CheckState = CheckState.Checked;

                AppType.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                //MessageBox.Show(AppType.Text); 
                mandoubName.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                //MessageBox.Show(mandoubName.Text);
                ConsulateEmployee.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                if (mandoubName.Text == "") 
                    AppType.CheckState = CheckState.Checked; 
                else AppType.CheckState = CheckState.Unchecked;

                if (AppType.CheckState == CheckState.Unchecked)
                {
                    mandoubVisibilty(); 
                }

                PreRelatedID = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                Comment.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();
                
                txtPurpose.Text= dataGridView1.CurrentRow.Cells[22].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[21].Value.ToString() != "غير مؤرشف")
                {
                    ArchivedSt.CheckState = CheckState.Checked;
                    ArchivedSt.Text = "مؤرشف";
                    ArchivedSt.BackColor = Color.Green;
                }
                else
                {
                    ArchivedSt.CheckState = CheckState.Unchecked;
                    ArchivedSt.Text = "غير مؤرشف";
                    ArchivedSt.BackColor = Color.Red;
                }
                ArchivedSt.Visible = true;

                
                btnSavePrint.Text = "حفظ";
            }
        }

        private void btnListView_Click(object sender, EventArgs e)
        {
            Clear_Fields();
            FillDataGridView();
            dataGridView1.Visible = true;
            PanelFiles.Visible = true;
            PanelMain.Visible = false;
        }

        private void Form6_FormClosed(object sender, FormClosedEventArgs e)
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

        private void btnEditID_Click(object sender, EventArgs e)
        {
            if (btnEditID.Text == "إجراء")
            {
                btnEditID.Text = "تعديل";
                SqlConnection sqlCon = new SqlConnection(DataSource);
                SqlCommand sqlCmd = new SqlCommand("update TableForensicApp SET DocID = @DocID WHERE ID = @ID", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                sqlCmd.Parameters.AddWithValue("@DocID", txtEditID1.Text + txtEditID2.Text);
                sqlCmd.ExecuteNonQuery();
                sqlCon.Close();
                txtEditID1.Visible = txtEditID2.Visible = false;
            }
            else
            {
                btnEditID.Text = "إجراء";
                txtEditID1.Visible = txtEditID2.Visible = true;
            }
        }

        private void txtEditID2_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtEditID1_TextChanged(object sender, EventArgs e)
        {

        }

        private void التاريخ_الميلادي_TextChanged(object sender, EventArgs e)
        {
            التاريخ_الميلادي_off.Text = التاريخ_الميلادي.Text.Split('-')[1] + " - " + التاريخ_الميلادي.Text.Split('-')[0] + " - " + التاريخ_الميلادي.Text.Split('-')[2];
        }

        private void مقدم_الطلب_TextChanged(object sender, EventArgs e)
        {
            getID(رقم_الهوية, مكان_الإصدار, النوع, مقدم_الطلب.Text);
        }
        bool gridFill = false;
        public void getID(TextBox رقم_الهوية_1 , TextBox مكان_الإصدار_1, CheckBox النوع_1, string name)
        {
            if (gridFill) return;
            string query = "SELECT * FROM TableGenNames where الاسم like N'" + name + "%'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);

            رقم_الهوية_1.Text = "P0";
            مكان_الإصدار_1.Text = "";
            النوع_1.Text = "ذكر";
            foreach (DataRow row in dtbl.Rows)
            {
                رقم_الهوية_1.Text = row["رقم_الهوية"].ToString();
                مكان_الإصدار_1.Text = row["مكان_الإصدار"].ToString();
                النوع_1.Text = row["النوع"].ToString();
            }
        }

        private void ListSearch_TextChanged_1(object sender, EventArgs e)
        {
            if (ListSearch.Text.Length != 0)
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '" + ListSearch.Text + "%'";
                dataGridView1.DataSource = bs;
            }else FillDataGridView();
        }

        private void OpenFile(int id, int fileNo)
        {
            string query;

            SqlConnection Con = new SqlConnection(DataSource);
            if (fileNo == 1)
            {
                query = "select Data1, Extension1,FileName1 from TableForensicApp where ID=@id";
            }
            else
            {
                query = "select Data2, Extension2,FileName2 from TableForensicApp where ID=@id";
            }
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@Id", SqlDbType.Int).Value = id;
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

        private void ResetAll_Click_1(object sender, EventArgs e)
        {
            
        }

        private void SaveOnly_Click_1(object sender, EventArgs e)
        {
            
            Save2DataBase();
            this.Close();
            //Clear_Fields();
        }
    }
}
