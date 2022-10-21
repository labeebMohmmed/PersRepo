
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
using System.Globalization;
using System.Threading;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using DocumentFormat.OpenXml.Office2010.Excel;
using Color = System.Drawing.Color;

namespace PersAhwal
{
    public partial class Form10 : Form
    {
        static string[,] preffix = new string[10, 20];
        static public bool ApplicantSexStatus = false;
        public static string route = "";
        static bool[] All_okey = new bool[10];
        static string[] DocxArabic= new string[10];
        static string[] DocxEnglish = new string[10];
        string textSave1 = "";
        string[] dataGrid = new string[50];
        string strlistdata = "", strlistdata1 = "";
        string textSave2 = "";
        string WitNessesInfo = "";
        public static string[] ChildName = new string[10];
        static string[,] Muaamla = new string[3, 5];
        static bool[] Son_Daughter = new bool[10];
        private string title = "حامل ";
        string Viewed;
        string textItems = "";
        string StrSpecPur = "";
        string[] DPTitle = new string[5];
        string activeCopy = "";
        string ConsulateEmpName;
        public static string ModelFileroute = "";
       String IqrarNumberPart;
        static string DataSource;
        string pronoun = "his";
        int ApplicantID = 0;
        private bool fileloaded = false;
        string NewFileName = "";
        string PreAppId = "", PreRelatedID = "", NextRelId = "";
        static public string FamilySupport;
        private string[] FamelyMember = new string[10];
        string FormTypeDocx, Auth;
        string FilesPathIn, FilesPathOut;
        string Jobposition;
        bool newData = true;
        bool SaveEdit = true;
        string[] colIDs = new string[100];
        int ATVC = 0;
        public Form10(int Atvc, int currentRow, int DocumentType, string EmpName, string dataSource, string filepathIn, string filepathOut, string jobposition, string gregorianDate, string hijriDate)
        {
            InitializeComponent();
            //timer1.Enabled = true;
            //timer2.Enabled = true;
            التاريخ_الميلادي.Text = gregorianDate;
            التاريخ_الهجري.Text = hijriDate;
            ATVC = Atvc;
            Muaamla[0, 0] = "أنا المواطن/";
            Muaamla[0, 1] = "، أقر وبكامل قـــواي العقليـــة وحالــــتي المعتبــرة شــرعاً وقانوناً وبطوعي واختياري بأنه";
            Muaamla[0, 2] = "، وهذا إقرار مني بذلك .";
            DataSource = dataSource;
            FilesPathIn = filepathIn + @"\";
            FilesPathOut = filepathOut;
            Jobposition = jobposition;
            colIDs[4] = ConsulateEmpName = EmpName;
            FillDataGridView();
            if (currentRow == -1) Clear_Fields();
            else SetFieldswithData(currentRow);
            //FormType.SelectedIndex = DocumentType;
            if (jobposition.Contains("قنصل"))
                btnEditID.Visible = deleteRow.Visible = true;
            else btnEditID.Visible = deleteRow.Visible = false;
        }
        private string loadRerNo(int id)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT DocID from TableFreeForm where ID=@ID", sqlCon);
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
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from TableFreeForm order by ID desc", sqlCon);
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
        //        query = "select Data1, Extension1,FileName1 from TableFreeForm  where ID=@id";
        //    }
        //    else if (fileNo == 2)
        //    {
        //        query = "select Data2, Extension2,FileName2 from TableFreeForm  where ID=@id";
        //    }
        //    else 
        //        query = "select Data3, Extension3,FileName3 from TableFreeForm  where ID=@id";

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
        private void SetFieldswithData(int Rowindex)
        {
            Rowindex--;
            ApplicantID = Convert.ToInt32(dataGridView1.Rows[Rowindex].Cells[0].Value.ToString());
            PreAppId = dataGridView1.Rows[Rowindex].Cells[1].Value.ToString();
            ApplicantName.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
            else if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
            DocType.Text = dataGridView1.Rows[Rowindex].Cells[4].Value.ToString();
            DocNo.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString();
            DocSource.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString();
            text.Text = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString();
            subFormType.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString();
            التاريخ_الميلادي.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString();
            التاريخ_الهجري.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[11].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[12].Value.ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                
            }
            else checkedViewed.CheckState = CheckState.Checked;
            AppType.Text = dataGridView1.Rows[Rowindex].Cells[13].Value.ToString();
            if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked;
            else AppType.CheckState = CheckState.Unchecked;
            ConsulateEmployee.Text = dataGridView1.Rows[Rowindex].Cells[14].Value.ToString();
            if (AppType.CheckState == CheckState.Unchecked)
            {
                mandoubVisibilty(); mandoubName.Text = dataGridView1.Rows[Rowindex].Cells[15].Value.ToString();
            }
            PreRelatedID = dataGridView1.Rows[Rowindex].Cells[20].Value.ToString();
            Comment.Text = dataGridView1.Rows[Rowindex].Cells[21].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[22].Value.ToString() != "غير مؤرشف")
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
            btnSavePrint.Visible = false;
        }

        private void FillDataGridView()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("FreeFormViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ApplicantName", ListSearch.Text.Trim());
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            //dataGridView1.Rows[1].DefaultCellStyle.BackColor = Color.Green;
            IqrarNumberPart = loadRerNo(loadIDNo()); ;
            sqlCon.Close();
            NewFileName = IqrarNumberPart + "_10";
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 200;

            int x = 0;
            foreach (DataGridViewRow dataRow in dataGridView1.Rows)
            {
                dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.LightGreen;
                x++;
                //if (dataGridView1.Rows[x].Cells[22].Value.ToString() == "مؤرشف")
                //{
                //    dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.LightGreen;
                //}
                //else
                //{
                //    dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.Red;
                //}
                //x++;
                if (dataGridView1.Rows.Count - 1 == x) return;
            }
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

        private void timer2_Tick_1(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            التاريخ_الميلادي.Text = DateTime.Now.ToString("MM-dd-yyyy");
            timer2.Enabled = false;
        }

        private void Review_Click(object sender, EventArgs e)
        {
        }

        private void text_Click(object sender, EventArgs e)
        {

            //string Auth = ApplicantName.Text + "،المقيم بالمملكة العربية الســـعودية " + title + PassIqama.Text + " رقم " + AppDocNo.Text+ " إصدار " + IssuedSource.Text;
            //Authtext.Text = Muaamla[0, 0] + Auth+ Muaamla[0, 1];
        }

        private void ApplicantSex_CheckedChanged_1(object sender, EventArgs e)
        {
            if (ApplicantSex.CheckState == CheckState.Unchecked) ApplicantSex.Text = "ذكر";
            else ApplicantSex.Text = "أنثى";
            //if (ApplicantSex.CheckState == CheckState.Checked) title = "حاملة "; else title = "حامل";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Authtext.Text = Authtext.Text + Muaamla[0, 2];
        }

        private void Form10_Load(object sender, EventArgs e)
        {
            fileComboBox2(FormType, DataSource, "ArabicGenIgrar", "TableListCombo");
            fileComboBox(mandoubName, DataSource, "MandoubNames", "TableListCombo");
            fileComboBox(comboHistory, DataSource, "SpecText", "TableFreeForm");
            fileComboBoxAttend(DocType, DataSource, "DocType", "TableListCombo");
            autoCompleteTextBox(DocSource, DataSource, "SDNIssueSource", "TableListCombo");
            fileComboBox(AttendViceConsul, DataSource, "ArabicAttendVC", "TableListCombo");
            
            AttendViceConsul.SelectedIndex = ATVC;
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
                textbox.AutoCompleteCustomSource.Clear();
                
                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    if (!string.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                    {
                        for (int x = 0; x < Textboxtable.Rows.Count; x++)
                            if (dataRow[comlumnName].ToString().Equals(Textboxtable.Rows[x]))
                                newSrt = false;

                        if (newSrt) 
                            autoComplete.Add(dataRow[comlumnName].ToString());
                    }
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }


        private void fileComboBox2(ComboBox combbox, string source, string comlumnName, string tableName)
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
                    if (!String.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                        combbox.Items.Add(dataRow[comlumnName].ToString());
                }
                saConn.Close();
            }
            if (combbox.Items.Count > 0) combbox.SelectedIndex = 0;
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
                        if(dataRow[comlumnName].ToString() != "")
                        combbox.Items.Add(dataRow[comlumnName].ToString());
                    }
                }
                saConn.Close();
            }
        }

        private void fileComboBoxAttend(ComboBox combbox, string source, string comlumnName, string tableName)
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
                    if(dataRow[comlumnName].ToString() != "")
                    combbox.Items.Add(dataRow[comlumnName].ToString());

                }
                saConn.Close();
            }
        }

        
        private void Save2DataBase(string DocName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            collectItemsData();
            string AppGender;
            if (language.CheckState == CheckState.Unchecked)
            {
                if (ApplicantSex.CheckState == CheckState.Unchecked)
                    AppGender = "ذكر";
                else AppGender = "أنثى";
            }
            else AppGender = titleEng.Text;
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                if (btnSavePrint.Text == "طباعة وحفظ" && newData)
                {
                    if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                    else Viewed = "غير معالج";
                    SqlCommand sqlCmd = new SqlCommand("FreeFormAddorEdit", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");
                    sqlCmd.Parameters.AddWithValue("@DocID", Iqrarid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", DocType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", DocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocIssueSource", DocSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SpecText", text.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SpecType", subFormType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", PreAppId.Trim());
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
                    string filePath1 = FilesPathIn + "text1.txt";
                    string filePath2 = FilesPathIn + "text2.txt";
                    sqlCmd.Parameters.Add("@FileName3", SqlDbType.NVarChar).Value = DocName;
                    using (Stream stream = File.OpenRead(filePath2))
                    {
                        byte[] buffer2 = new byte[stream.Length];
                        stream.Read(buffer2, 0, buffer2.Length);
                        var fileinfo2 = new FileInfo(filePath2);
                        string extn2 = fileinfo2.Extension;
                        string DocName2 = fileinfo2.Name;
                        sqlCmd.Parameters.Add("@Data2", SqlDbType.VarBinary).Value = buffer2;
                        sqlCmd.Parameters.Add("@Extension2", SqlDbType.Char).Value = extn2;
                        sqlCmd.Parameters.Add("@FileName2", SqlDbType.NVarChar).Value = DocName2;


                    }

                    sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");
                    sqlCmd.Parameters.AddWithValue("@Lang", language.Text);
                    sqlCmd.Parameters.AddWithValue("@textToSave", textSave2);
                    sqlCmd.Parameters.AddWithValue("@Witnesses", WitNessesInfo);
                    sqlCmd.Parameters.AddWithValue("@mainForm", FormType.SelectedIndex.ToString());
                    sqlCmd.ExecuteNonQuery();
                }
                else
                {
                    if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                    else Viewed = "غير معالج";
                    SqlCommand sqlCmd = new SqlCommand("FreeFormAddorEdit", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                    sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                    sqlCmd.Parameters.AddWithValue("@DocID", Iqrarid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", DocType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", DocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocIssueSource", DocSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SpecText", text.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SpecType", subFormType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", PreAppId.Trim());
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
                    string filePath1 = FilesPathIn + "text1.txt";
                    string filePath2 = FilesPathIn + "text2.txt";
                    sqlCmd.Parameters.Add("@FileName3", SqlDbType.NVarChar).Value = DocName;
                    if (SearchFile.Text != "") { filePath2 = SearchFile.Text; fileloaded = true; }
                    using (Stream stream = File.OpenRead(filePath2))
                    {
                        byte[] buffer2 = new byte[stream.Length];
                        stream.Read(buffer2, 0, buffer2.Length);
                        var fileinfo2 = new FileInfo(filePath2);
                        string extn2 = fileinfo2.Extension;
                        string DocName2 = fileinfo2.Name;
                        sqlCmd.Parameters.Add("@Data2", SqlDbType.VarBinary).Value = buffer2;
                        sqlCmd.Parameters.Add("@Extension2", SqlDbType.Char).Value = extn2;
                        sqlCmd.Parameters.Add("@FileName2", SqlDbType.NVarChar).Value = DocName2;
                        if (fileloaded)
                        {
                            ArchivedSt.CheckState = CheckState.Checked;
                            Clear_Fields();
                        }
                    }

                    if (fileloaded)
                        sqlCmd.Parameters.AddWithValue("@ArchivedState", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    else sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");

                    sqlCmd.Parameters.AddWithValue("@Lang", language.Text);
                    sqlCmd.Parameters.AddWithValue("@textToSave",textSave2);
                    sqlCmd.Parameters.AddWithValue("@Witnesses", WitNessesInfo);
                    sqlCmd.Parameters.AddWithValue("@mainForm", FormType.SelectedIndex.ToString());
                    sqlCmd.ExecuteNonQuery();
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
            FillDataGridView();
        }



        private void AppType_CheckedChanged(object sender, EventArgs e)
        {
        }
        private void button2_Click(object sender, EventArgs e)
        {
            
            Save2DataBase(activeCopy);
            CreateWordFile(activeCopy);


        }

        
        private void EngCreateWordFile(string ActiveCopy)
        {
            strAuth();
            string ReportName = DateTime.Now.ToString("mmss");
            route = FilesPathIn + FormType.Text;

            if (ApplicantSex.CheckState == CheckState.Checked)
            {
                ApplicantSexStatus = false;
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
            }



            if (!File.Exists(ActiveCopy))
            {
                System.IO.File.Copy(route, ActiveCopy);
                object oBMiss = System.Reflection.Missing.Value;
                Word.Application oBMicroWord = new Word.Application();

                object Routseparameter = ActiveCopy;

                Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

                object ParaIqrarNo = "MarkIqrarNo";
                object ParaGreData = "MarkGreData";
                object ParaHijriData = "MarkHijriData";
                object ParaName = "MarkAppName";
                object ParaDocType = "MarkDocType";
                object Parapronoun = "Markpronoun";
                object ParaDocSource = "MarkDocSource";
                object ParaDocNo = "MarkDocNo";
                object ParaPurpose = "MarkPurpose";
                object ParaPurposeText = "MarkPurposeText";
                object ParavConsul = "MarkViseConsul";
                object ParaAuthAppName = "MarkAuthAppName";
                object ParaAuthorization = "MarkAuthorization";
                object ParaAppTitle = "MarkMarkAppTitle";
                object ParaAppTitle1 = "MarkMarkAppTitle1";
                object ParaAppTitle2 = "MarkAppTitle";


                Word.Range BookIqrarNo = oBDoc.Bookmarks.get_Item(ref ParaIqrarNo).Range;
                Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
                Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
                Word.Range BookName;
                Word.Range BookDocType;
                Word.Range BookDocSource;
                Word.Range BookDocNo;
                Word.Range BookAppTitle;
                Word.Range BookAppTitle1;
                Word.Range BookApptitle2;
                Word.Range BookPurpose;
                Word.Range Bookpronoun;
                BookName = oBDoc.Bookmarks.get_Item(ref ParaName).Range;
                BookDocType = oBDoc.Bookmarks.get_Item(ref ParaDocType).Range;
                BookDocSource = oBDoc.Bookmarks.get_Item(ref ParaDocSource).Range;
                BookDocNo = oBDoc.Bookmarks.get_Item(ref ParaDocNo).Range;


                BookName.Text = ApplicantName.Text;
                BookDocType.Text = DocType.Text;
                BookDocSource.Text = DocSource.Text;
                BookDocNo.Text = DocNo.Text;


                object rangeName = BookName;
                object rangeDocSource = BookDocSource;
                object rangeDocType = BookDocType;
                object rangeDocNo = BookDocNo;


                oBDoc.Bookmarks.Add("MarkAppName", ref rangeName);
                oBDoc.Bookmarks.Add("MarkDocType", ref rangeDocSource);
                oBDoc.Bookmarks.Add("MarkDocSource", ref rangeDocType);
                oBDoc.Bookmarks.Add("MarkDocNo", ref rangeDocNo);


                if (FormType.SelectedIndex >= 0 && FormType.SelectedIndex <= 1 && language.CheckState == CheckState.Unchecked)
                {

                    Word.Range BookAuthAppName = oBDoc.Bookmarks.get_Item(ref ParaAuthAppName).Range;
                    BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;
                    Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;

                    BookAuthAppName.Text = ApplicantName.Text;
                    BookAuthorization.Text = Auth;
                    BookPurpose.Text = subFormType.Text;

                    object RangeAuthappname = BookAuthAppName;
                    object rangeAuthorization = BookAuthorization;
                    object rangePurpose = BookPurpose;

                    oBDoc.Bookmarks.Add("MarkAuthAppName", ref RangeAuthappname);
                    oBDoc.Bookmarks.Add("MarkPurpose", ref rangePurpose);
                    oBDoc.Bookmarks.Add("MarkAuthorization", ref rangeAuthorization);
                }




                else if (language.CheckState == CheckState.Checked && (FormType.SelectedIndex == 0 || FormType.SelectedIndex == 1))
                {
                    BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;

                    BookPurpose.Text = subFormType.Text;

                    object rangePurpose = BookPurpose;
                    oBDoc.Bookmarks.Add("MarkPurpose", ref rangePurpose);

                }
                else if (language.CheckState == CheckState.Checked && FormType.SelectedIndex == 2)
                {

                    BookAppTitle = oBDoc.Bookmarks.get_Item(ref ParaAppTitle).Range;
                    BookAppTitle1 = oBDoc.Bookmarks.get_Item(ref ParaAppTitle1).Range;
                    BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;

                    BookAppTitle.Text = ApplicantName.Text;
                    BookAppTitle1.Text = DocType.Text;
                    BookPurpose.Text = FormType.Text;

                    object rangeAppTitle = BookAppTitle;
                    object rangeAppTitle1 = BookAppTitle1;
                    object rangePurpose = BookPurpose;
                    oBDoc.Bookmarks.Add("MarkAppTitle", ref rangeAppTitle);
                    oBDoc.Bookmarks.Add("MarkAppTitle1", ref rangeAppTitle1);
                    oBDoc.Bookmarks.Add("MarkPurpose", ref rangePurpose);
                }
                else if (language.CheckState == CheckState.Checked && FormType.SelectedIndex == 3)
                {
                    //MessageBox.Show(titleEng.Text);
                    if (titleEng.SelectedIndex != 0) pronoun = "her";
                    Bookpronoun = oBDoc.Bookmarks.get_Item(ref Parapronoun).Range;
                    Bookpronoun.Text = pronoun;
                    object rangepronoun = Bookpronoun;
                    oBDoc.Bookmarks.Add("Markpronoun", ref rangepronoun);

                    BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;
                    BookPurpose.Text = FormType.Text;
                    object rangePurpose = BookPurpose;
                    oBDoc.Bookmarks.Add("MarkPurpose", ref rangePurpose);

                    BookApptitle2 = oBDoc.Bookmarks.get_Item(ref ParaAppTitle2).Range;
                    BookApptitle2.Text = titleEng.Text;
                    object rangetitle2 = BookApptitle2;
                    oBDoc.Bookmarks.Add("MarkAppTitle", ref rangetitle2);

                }

                Word.Range BookPurposeText = oBDoc.Bookmarks.get_Item(ref ParaPurposeText).Range;
                Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;



                BookIqrarNo.Text = Iqrarid.Text;
                BookGreData.Text = التاريخ_الميلادي.Text;
                BookHijriData.Text = التاريخ_الهجري.Text;



                BookPurposeText.Text = text.Text;

                BookvConsul.Text = AttendViceConsul.Text;


                object rangeIqrarNo = BookIqrarNo;
                object rangeGreData = BookGreData;
                object rangeHijriData = BookHijriData;


                object rangePurposeText = BookPurposeText;
                object rangevConsul = BookvConsul;



                oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);
                oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
                oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);


                oBDoc.Bookmarks.Add("MarkPurposeText", ref rangePurposeText);
                oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);
                string docxouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".docx";
                string pdfouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
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
        }

        private void CreateWordFile(string ActiveCopy)
        {
            strAuth();
            string ReportName = DateTime.Now.ToString("mmss");
            route = FilesPathIn + FormType.Text + ".docx";
            if (txtWitName1.Text != "" && language.Text == "العربية") {
                route = FilesPathIn + "إقرار مشهود.docx";
            } 
            else if (txtWitName1.Text != "" && language.Text == "الانجليزية")
            {
                route = FilesPathIn + "witnessed Affidavit.docx";
            }
            System.IO.File.Copy(route, ActiveCopy);
            //MessageBox.Show(route);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

            if (txtWitName1.Text != "")
            {
                object ParaWitName1 = "MarkWitName1";
                object ParaWitName2 = "MarkWitName2";
                object ParaWitPass1 = "MarkWitPass1";
                object ParaWitPass2 = "MarkWitPass2";

                Word.Range BookWitName1 = oBDoc.Bookmarks.get_Item(ref ParaWitName1).Range;
                Word.Range BookWitName2 = oBDoc.Bookmarks.get_Item(ref ParaWitName2).Range;
                Word.Range BookWitPass1 = oBDoc.Bookmarks.get_Item(ref ParaWitPass1).Range;
                Word.Range BookWitPass2 = oBDoc.Bookmarks.get_Item(ref ParaWitPass2).Range;

                BookWitName1.Text = combTitle13.Text + ". " + txtWitName1.Text;
                BookWitName2.Text = combTitle13.Text + ". " + txtWitName2.Text;
                BookWitPass1.Text = txtWitPass1.Text;
                BookWitPass2.Text = txtWitPass2.Text;

                object rangeWitName1 = BookWitName1;
                object rangeWitName2 = BookWitName2;
                object rangeWitPass1 = BookWitPass1;
                object rangeWitPass2 = BookWitPass2;

                oBDoc.Bookmarks.Add("MarkWitName1", ref rangeWitName1);
                oBDoc.Bookmarks.Add("MarkWitName2", ref rangeWitName2);
                oBDoc.Bookmarks.Add("MarkWitPass1", ref rangeWitPass1);
                oBDoc.Bookmarks.Add("MarkWitPass2", ref rangeWitPass2);
            }

            object ParaIqrarNo = "MarkIqrarNo";
            object ParaGreData = "MarkGreData";
            object ParaHijriData = "MarkHijriData";
            object ParaPurpose = "MarkPurpose";//message لعناية ... note verbal فارغ
            object ParaPurposeText = "MarkPurposeText";
            object ParavConsul = "MarkViseConsul";  //note verbal لتعرب       
            object ParaAuthorization = "MarkAuthorization";//note verbal الى الوجهة
            
            Word.Range BookIqrarNo = oBDoc.Bookmarks.get_Item(ref ParaIqrarNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;            
            Word.Range BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;
            Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;
            Word.Range BookPurposeText = oBDoc.Bookmarks.get_Item(ref ParaPurposeText).Range;
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;
            Word.Range BookApplicantName2;

            if (FormType.SelectedIndex == 0)
            {
                object ParaApplicantName2 = "MarkApplicantName2";
                BookApplicantName2 = oBDoc.Bookmarks.get_Item(ref ParaApplicantName2).Range;

                if(language.CheckState == CheckState.Unchecked)
                    BookApplicantName2.Text = ApplicantName.Text;
                else BookApplicantName2.Text = titleEng.Text + ". "+ApplicantName.Text;

                object rangeApplicantName2 = BookApplicantName2;
                oBDoc.Bookmarks.Add("MarkApplicantName2", ref rangeApplicantName2);
            }


            string str = "";
            if (!language.Checked && ApplicantSex.Text != "ذكر") str = "ة";
            else if(language.Checked && titleEng.SelectedIndex == 0) str = "his";
            else if (language.Checked && titleEng.SelectedIndex != 0) str = "her";
            if (FormType.Text == "إفادة لمن يهمه الأمر")
            {
                Auth = "قد حررت هذه الإفادة بناءً على طلب المذكور" + str+ " أعلاه لاستخدامها على الوجه المشروع";
            }
            else if (FormType.Text == "شهادة عدم ممانعة" || FormType.Text == "شهادة لمن يهمه الأمر")
                Auth = "قد حررت هذه الشهادة بناءً على طلب المذكور" + str + "أعلاه لاستخدامها على الوجه المشروع";
            if (FormType.Text == "TO WHOM IT MAY CONCERN")
            {
                Auth = "This certificate has been issued upon " + str + " request "; 
            }
            BookAuthorization.Text = Auth;
            BookPurpose.Text = comboTitle.Text;
            BookIqrarNo.Text = colIDs[0] = Iqrarid.Text;
            BookGreData.Text = colIDs[2] = التاريخ_الميلادي.Text;
            colIDs[3] = ApplicantName.Text;
            colIDs[5] =AppType.Text;
            colIDs[6] = mandoubName.Text;
            BookHijriData.Text = التاريخ_الهجري.Text;
            BookPurposeText.Text = text.Text;
            BookvConsul.Text = AttendViceConsul.Text;
                

            object rangeAuthorization = BookAuthorization;
            object rangePurpose = BookPurpose;
            object rangeIqrarNo = BookIqrarNo;
            object rangeGreData = BookGreData;
            object rangeHijriData = BookHijriData;
            object rangePurposeText = BookPurposeText;
            object rangevConsul = BookvConsul;


            oBDoc.Bookmarks.Add("MarkPurpose", ref rangePurpose);
            oBDoc.Bookmarks.Add("MarkAuthorization", ref rangeAuthorization);
            oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);
            oBDoc.Bookmarks.Add("MarkPurposeText", ref rangePurposeText);
            oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);

            string docxouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(pdfouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;

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
                sqlCommand.Parameters.AddWithValue("@" + allList[i], text[i - 1]);
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
        void FillDatafromGenArch(string doc, string id)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableGeneralArch where  رقم_المرجع='" + id + "' and نوع_المستند='" + doc + "'", sqlCon);
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

        //private void OpenFile(int id, int fileNo)
        //{
        //    string query;

        //    SqlConnection Con = new SqlConnection(DataSource);
        //    if (fileNo == 1)
        //    {
        //        query = "select Data1, Extension1,FileName1 from TableFreeForm where ID=@id";
        //    }
        //    else if (fileNo == 2)
        //    {
        //        query = "select Data2, Extension2,FileName2 from TableFreeForm where ID=@id";
        //    }
        //    else
        //        query = "select Data3, Extension3,FileName3 from TableFreeForm where ID=@id";

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
        //            var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
        //            File.WriteAllBytes(NewFileName, Data);
        //            System.Diagnostics.Process.Start(NewFileName);
        //        }
        //        else if (fileNo == 2)
        //        {
        //            var name = reader["FileName2"].ToString();
        //            var Data = (byte[])reader["Data2"];
        //            var ext = reader["Extension2"].ToString();
        //            var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
        //            File.WriteAllBytes(NewFileName, Data);
        //            System.Diagnostics.Process.Start(NewFileName);
        //        }
        //        else 
        //        {
        //            var name = reader["FileName3"].ToString();
        //            var Data = (byte[])reader["Data3"];
        //            var ext = reader["Extension3"].ToString();
        //            var NewFileName = name;
        //            updateExtension3(id);
        //            File.WriteAllBytes(NewFileName, Data);
        //            System.Diagnostics.Process.Start(NewFileName);
        //        }
        //    }
        //    Con.Close();

        //}

        private void updateExtension3(int id)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("update TableFreeForm set Extension3=@Extension3 where ID=@ID", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@Extension3", ".txt");
            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            var selectRows = dataGridView1.SelectedRows;
            foreach (var row in selectRows)
            {
                int id = (int)((DataGridViewRow)row).Cells[0].Value;
                //OpenFile(id, 1);
                FillDatafromGenArch("data1", id.ToString());
            }
            if (ApplicantID != 0) FillDatafromGenArch("data1", ApplicantID.ToString());
            //ApplicantID = 0;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var selectRows = dataGridView1.SelectedRows;
            foreach (var row in selectRows)
            {
                int id = (int)((DataGridViewRow)row).Cells[0].Value;
                FillDatafromGenArch("data2", ApplicantID.ToString());
            }
            if (ApplicantID != 0) FillDatafromGenArch("data2", ApplicantID.ToString());
           // ApplicantID = 0;
        }

        
        private void btnSavePrint_Click_1(object sender, EventArgs e)
        {
            WitNessesInfo = txtWitName1.Text + "_" + combTitle13.Text + "_" + txtWitPass1.Text + "_" + txtWitName2.Text + "_" + combTitle14.Text + "_" + txtWitPass2.Text;
            btnSavePrint.Text = "جاري المعالجة";
            btnSavePrint.Enabled = false;
            string activeCopy = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("mmss") + ".docx";
            Save2DataBase(activeCopy);
            CreateWordFile(activeCopy);
            this.Close();
            //Clear_Fields();
        }


        private void ResetAll_Click_1(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void PassIqama_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DocType.Text == "اقامة ") { DocType.Text = ""; } else DocType.Text = "P0";
        }

        private void Clear_Fields()
        {
            ApplicantName.Text = DocSource.Text = DocSource.Text = "";

            ApplicantSex.CheckState = CheckState.Unchecked;
            labeldoctype.Text = "رقم جواز السفر: ";
            DocNo.Text = "P0";
            
            text.Text = "";
            mandoubName.Text = ListSearch.Text = "";
            mandoubVisibilty();            
            btnSavePrint.Text = "طباعة وحفظ";
            btnSavePrint.Enabled = true;
            btnSavePrint.Visible = true;
            Comment.Text = "لا تعليق";
            FillDataGridView();
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            
            ArchivedSt.BackColor = Color.Red;
            SearchFile.Visible = false;
            fileloaded = false;
            if(combo1.Items.Count > 25)
                combo1.SelectedIndex = 26;
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            التاريخ_الميلادي.Text = DateTime.Now.ToString("dd-MM-yyyy");
            Console.WriteLine(التاريخ_الميلادي.Text);
            //Iqrarid.Text = "ق س ج/80/" + GregorianDate.Text.Split('-')[2].Replace("20", "") + "/10/" + loadRerNo(loadIDNo()); ;
            ConsulateEmployee.Text = ConsulateEmpName;            
            foreach (Control control in PanelItemsboxes.Controls)
            {
                if (control is TextBox)
                {
                    ((TextBox)control).Text = "";
                }

                if (control is CheckBox)
                {
                    ((CheckBox)control).CheckState = CheckState.Unchecked;
                }
                if (control is ComboBox)
                {
                    ((ComboBox)control).Items.Clear();
                }
            }            
            fileComboBox(FormType, DataSource, "ArabicGenIgrar", "TableListCombo");
            fileComboBox(AttendViceConsul, DataSource, "ArabicAttendVC", "TableListCombo");
            fileComboBoxAttend(DocType, DataSource, "DocType", "TableListCombo");
            DocType.SelectedIndex = 0;
            AttendViceConsul.SelectedIndex = 2;
            language.CheckState = CheckState.Unchecked;
            language.Text = "العربية";
            //FormType.SelectedIndex = 0;
            titleEng.Visible = false;
            ApplicantSex.Visible = true;
            newData = true;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

      

        private void deleteRow_Click_1(object sender, EventArgs e)
        {
            
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

        private void comboHistory_SelectedIndexChanged(object sender, EventArgs e)
        {
            text.Text = comboHistory.Text;
        }

        private void language_CheckedChanged_1(object sender, EventArgs e)
        {
            if (language.CheckState == CheckState.Checked)
            {
                language.Text = "الانجليزية";
                txtWitPass1.Width = 120;
                txtWitPass2.Width = 120;
                labeltitle12.Visible = true;
                combTitle13.Visible = true;
                labeltitle13.Visible = true;
                combTitle14.Visible = true;
                titleEng.Visible = true;
                ApplicantSex.Visible = false;
                FormType.RightToLeft = RightToLeft.No;
                //Iqrarid.Text = "CGSJ/80/" + GregorianDate.Text.Split('-')[2].Replace("20", "") + "/02/" + loadRerNo(loadIDNo());
                fileComboBox(FormType, DataSource, "EnglishGenIgrar", "TableListCombo"); 
                fileComboBoxAttend(AttendViceConsul, DataSource, "EnglishAttendVC", "TableListCombo");
                fileComboBoxAttend(DocType, DataSource, "EngDocType", "TableListCombo");
                autoCompleteTextBox(DocSource, DataSource, "KSAIssureSource", "TableListCombo");
                fileComboBox(comboHistory, DataSource, "SpecText", "TableFreeForm");
                DocType.SelectedIndex = 0;
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("en-US");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);

                PanelItemsboxes.RightToLeft = text.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
                labeltitle12.Visible = labeltitle13.Visible = true;
                
            }
            else if (language.CheckState == CheckState.Unchecked)
            {
                language.Text = "العربية";
                fileComboBox(AttendViceConsul, DataSource, "ArabicAttendVC", "TableListCombo");
                fileComboBoxAttend(FormType, DataSource, "ArabicGenIgrar", "TableListCombo");
                fileComboBoxAttend(DocType, DataSource, "DocType", "TableListCombo");
                autoCompleteTextBox(DocSource, DataSource, "SDNIssueSource", "TableListCombo");
                fileComboBox(comboHistory, DataSource, "SpecText", "TableFreeForm");

                labeltitle12.Visible = labeltitle13.Visible = false;

                txtWitPass1.Width = 214;
                txtWitPass2.Width = 214;
                labeltitle12.Visible = false;
                combTitle13.Visible = false;
                labeltitle13.Visible = false;
                combTitle14.Visible = false;

                FormType.RightToLeft = RightToLeft.Yes;
                //Iqrarid.Text = "ق س ج/80/" + GregorianDate.Text.Split('-')[2].Replace("20", "") + "/10/" + loadRerNo(loadIDNo());
                titleEng.Visible = false;
                ApplicantSex.Visible = true;
                DocType.SelectedIndex = 0;
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
                PanelItemsboxes.RightToLeft = text.RightToLeft = System.Windows.Forms.RightToLeft.No;
            }
            AttendViceConsul.SelectedIndex = 2;
            FormType.SelectedIndex = 0;
        }

        private bool checkColumnName(string colNo)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
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



        private void FormType_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (FormType.SelectedIndex == 0) 
                PanelWit.Visible = true;
            else 
                PanelWit.Visible = false;
            if (checkColumnName(FormType.Text.Replace(" ", "_")))
            {
                subFormType.Items.Clear();
                newFillComboBox2(subFormType, DataSource, FormType.SelectedIndex.ToString(), language.Text);
                if (subFormType.Items.Count > 0) subFormType.SelectedIndex = 0;
                comboTitle.Items.Clear();
                comboTitle.Items.Add(subFormType.Text);
                comboTitle.Items.Add(FormType.Text);
                comboTitle.SelectedIndex = 1;
                return;
            }
        }

        private void FormType_TextChanged(object sender, EventArgs e)
        {
            string[] strList = new string[2];
            strList[0] = "Son";
            strList[1] = "Daughter";
            strlistdata = strList[0];
            combo1.Items.Clear();
            for (int x = 1; x < 2; x++)
            {
                combo1.Items.Add(strList[x]);
                strlistdata = strlistdata + "_" + strList[x];
            }
            combo1.SelectedIndex = 0;
            switch (FormType.Text) {
                //case "Affidavit of Financial support":
                //    optionPanel("اسم المكفول:", 270, "رقم الجواز:", 174, "اسم الجامعة:", 270, "", 174, "", 174, "", "", strList, "بحث الدولة المراد زيارتها:");
                //    textSave1 = "اسم المكفول:" +"-"+ "270" +"-"+ "رقم الجواز:" +"-"+ "174" +"-"+ "اسم الجامعة:" +"-"+ "270" +"-"+ "" +"-"+ "174" +"-"+ "" +"-"+ "174" +"-"+ "" +"-"+ "" +"-"+ strList +"-"+ "بحث الدولة المراد زيارتها:";
                //    break;
                //case "Note verbal":
                //    optionPanel("", 270, "", 174, "", 270, "", 174, "", 174, "", "", strList, "بحث الدولة:");
                //    textSave1 = "" + "-" + "270" + "-" + "" + "-" + "174" + "-" + "" + "-" + "270" + "-" + "" + "-" + "174" + "-" + "" + "-" + "174" + "-" + "" + "-" + "" + "-" + strList + "-" + "بحث الدولة:";
                //    break;
                //case "مذكرة لسفارة عربية":
                //    optionPanel("", 270, "", 174, "", 270, "", 174, "", 174, "", "", strList, "بحث الدولة:");
                //    textSave1 = "" + "-" + "270" + "-" + "" + "-" + "174" + "-" + "" + "-" + "270" + "-" + "" + "-" + "174" + "-" + "" + "-" + "174" + "-" + "" + "-" + "" + "-" + strList + "-" + "بحث الدولة:";
                //    break;
                //case "برقية":
                //    optionPanel("", 270, "", 174, "", 270, "", 174, "", 174, "", "", strList, "بحث الدولة:");
                //    textSave1 = "" + "-" + "270" + "-" + "" + "-" + "174" + "-" + "" + "-" + "270" + "-" + "" + "-" + "174" + "-" + "" + "-" + "174" + "-" + "" + "-" + "" + "-" + strList + "-" + "بحث الدولة:";
                //    break;
            }
        }

        private void FormType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void AppType_CheckedChanged_1(object sender, EventArgs e)
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
            }
            else
            {
                AppType.Text = "عن طريق أحد مندوبي القنصلية";
                mandoubName.Visible = true;
                mandoubLabel.Visible = true;
            }
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
                if (dataGridView1.Rows[i].Cells[22].Value.ToString() == "مؤرشف نهائي") dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;

                //else dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.;

            }
            //
        }

        private void btnListView_Click(object sender, EventArgs e)
        {
            Clear_Fields(); 
            FillDataGridView();
            dataGridView1.Visible = true;
            PanelFiles.Visible = false;
            PanelMain.Visible = false;
        }

        private void SearchDoc_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            SearchFile.Text = dlg.FileName;
        }

        private void btnFile1_Click(object sender, EventArgs e)
        {
            FillDatafromGenArch("data1", ApplicantID.ToString());
        }

        private void btnFile2_Click(object sender, EventArgs e)
        {
            FillDatafromGenArch("data2", ApplicantID.ToString());
        }

        private void btnFile3_Click(object sender, EventArgs e)
        {
         //   OpenFile(ApplicantID, 3);
        }

        private void deleteRow_Click(object sender, EventArgs e)
        {
            var selectedOption = MessageBox.Show("", "تأكيد عملية الحذف", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                deleteRowsData(ApplicantID, "TableFreeForm", DataSource);
                deleteRow.Visible = false;
                FillDataGridView();
                dataGridView1.Visible = true;
                PanelFiles.Visible = false;
                PanelMain.Visible = false;
            }
        }


        

        private string SuffPrefReplacements(string text)
        {
            Suffex_preffixList();
            string str = "";
            if (ApplicantSex.Text != "ذكر") str = "ة";
            
            if (text.Contains("tN"))
                return text.Replace("tN", ApplicantName.Text);
            if (text.Contains("tP"))
                return text.Replace("tP", DocNo.Text);
            if (text.Contains("tS"))
                return text.Replace("tS", DocSource.Text);
            if (text.Contains("tX"))
                return text.Replace("tX", str);            
            if (text.Contains("tT")) 
                return text.Replace("tT", titleEng.Text);
            if (text.Contains("tD"))
                return text.Replace("tD", DocType.Text);
            if (text.Contains("t1"))
                return text.Replace("t1", txt1.Text);
            if (text.Contains("t2"))
                return text.Replace("t2", txt2.Text);
            if (text.Contains("t3"))
                return text.Replace("t3", txt3.Text);
            if (text.Contains("t4"))
                return text.Replace("t4", txt4.Text);
            if (text.Contains("t5"))
                return text.Replace("t5", txt5.Text);
            if (text.Contains("c1"))
                return text.Replace("c1", check1.Text);

            if (text.Contains("m1"))
                return text.Replace("m1", combo1.Text);
            if (text.Contains("m2"))
                return text.Replace("m2", combo2.Text);

            if (text.Contains("a1"))
                return text.Replace("a1", addName1.Text);

            if (text.Contains("n1"))
                return text.Replace("n1", " " + txtD1.Text + "/" + txtM1.Text + "/" + txtY1.Text + " ");
            if (text.Contains("#*#"))
                return text.Replace("#*#", preffix[0, 10]);

            if (text.Contains("#1"))
                return text.Replace("#1", preffix[0, 11]);
            if (text.Contains("#2"))
                return text.Replace("#2", preffix[0, 12]);

            if (text.Contains("@*@"))
                return text.Replace("@*@", "لدى  برقم الايبان ()");
            if (text.Contains("$$$"))
                return text.Replace("$$$", preffix[0, 0]);
            if (text.Contains("&&&"))
                return text.Replace("&&&", preffix[0, 1]);
            if (text.Contains("^^^"))
                return text.Replace("^^^", preffix[0, 2]);
            if (text.Contains("###"))
                return text.Replace("###", preffix[0, 4]);
            if (text.Contains("***"))
                return text.Replace("***", preffix[0, 3]);
            else return text;
        }




        private void Suffex_preffixList()
        {

            preffix[0, 0] = "ي"; //$$$ "ي/نا";
            preffix[1, 0] = "ي";
            preffix[2, 0] = "نا";
            preffix[3, 0] = "نا";
            preffix[4, 0] = "نا";
            preffix[5, 0] = "نا";

            preffix[0, 1] = "ت";//&&& "ت/نا";
            preffix[1, 1] = "ت";
            preffix[2, 1] = "نا";
            preffix[3, 1] = "نا";
            preffix[4, 1] = "نا";
            preffix[5, 1] = "نا";

            preffix[0, 2] = "ني";//^^^ "ني/نا";
            preffix[1, 2] = "ني";
            preffix[2, 2] = "نا";
            preffix[3, 2] = "نا";
            preffix[4, 2] = "نا";
            preffix[5, 2] = "نا";

            preffix[0, 3] = "";//*** "/ت/ا/تا/ن/وا                               
            preffix[1, 3] = "ت";
            preffix[2, 3] = "ا";
            preffix[3, 3] = "تا";
            preffix[4, 3] = "ن";
            preffix[5, 3] = "وا";

            preffix[0, 4] = "ه";//### "ه/ها/هما/هما/من/هم"
            preffix[1, 4] = "ها";
            preffix[2, 4] = "هما";
            preffix[3, 4] = "هما";
            preffix[4, 4] = "هن";
            preffix[5, 4] = "هم";

            preffix[0, 5] = ""; //
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
        }

        public void DetermineData(string v1, string v2, string v3, string v4, string v5, string v6, string v7, string v8, string v9, string v10)
        {
            txt1.Text = v1;
            txt2.Text = v2;
            txt3.Text = v3;
            txt4.Text = v4;
            txt5.Text = v5;
            check1.Text = v6;
            if (v7.Contains("-"))
            {
                txtD1.Text = v7.Split('-')[0];
                txtM1.Text = v7.Split('-')[1];
                txtY1.Text = v7.Split('-')[2];
            }
            combo1.Text = v8;
            combo1.Text = v10;
            addName1.Text = v9;
            //MessageBox.Show(v1);
        }
        public void DetermineData(string v1, string v2, string v3, string v4, string v5, string v6, string v7)
        {
            txt1.Text = v1;
            txt2.Text = v2;
            txt3.Text = v3;
            txt4.Text = v4;
            txt5.Text = v5;
            check1.Text = v6;
            if (v7.Contains("-"))
            {
                txtD1.Text = v7.Split('-')[0];
                txtM1.Text = v7.Split('-')[1];
                txtY1.Text = v7.Split('-')[2];
            }
            combo1.Text = v7;
        }


        private void collectItemsData()
        {
            if (language.CheckState == CheckState.Checked)
                strlistdata1 = combo1.Text;
            else strlistdata1 = combo2.Text;
            textSave2 = txt1.Text + "_" + txt2.Text + "_" + txt3.Text + "_" + txt4.Text + "_" + txt5.Text + "_" + check1.Text + "_" + txtD1.Text + "/" + txtM1.Text + "/" + txtY1.Text + "_" + combo1.Text + "_" + addName1.Text + "_" + combo2.Text;
        }


        private void text_MouseHover(object sender, EventArgs e)
        {
            if (!checkupDate.Checked) return;
            text.Text = StrSpecPur;
            for (int x = 0; x < 30; x++)
                text.Text = SuffPrefReplacements(text.Text);
        }

        void ContextView(string text)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("ContextViewSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ColName", text);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView2.DataSource = dtbl;
            dataGridView2.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            dataGridView2.Columns["ID"].Visible = false;
            sqlCon.Close();
        }

        private void subFormType_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            
            if (ShowRowNo())
            {                
                if (!String.IsNullOrEmpty(textItems))
                {
                    string[] SI = textItems.Split('_');
                    if (SI[0].Contains("-"))
                        SI[0]= SI[0].Split('-')[1];
                        if (SI.Length == 10)
                        DetermineData(SI[0], SI[1], SI[2], SI[3], SI[4], SI[5], SI[6], SI[7], SI[8], SI[9]);
                    else
                        DetermineData(SI[0], SI[1], SI[2], SI[3], SI[4], SI[5], SI[6]);
                }
            }

            comboTitle.Items.Clear();
            comboTitle.Items.Add(subFormType.Text);
            comboTitle.Items.Add(FormType.Text);
            comboTitle.SelectedIndex = 1;
        }

        private bool ShowRowNo()
        {
            //label1,lenght1,label2,lenght2,label3,lenght3,label4,lenght4,label5,lenght5,
            //labelcheck,optionscheck,12
            //labelcomb1,optionscombo1,lenghtscombo1,labelcomb2,optionscombo2,lenghtscombo2,13
            //labelbtn,lenghtsbtn,19
            //dateYN,dateType,TextModel,ColRight,ColName 21

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("ContextViewSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ColName", "");
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView2.DataSource = dtbl;
            //dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            sqlCon.Close();
            
            for (int id = 0; id < dataGridView2.Rows.Count; id++)
            {

                if (dataGridView2.Rows[id].Cells[25].Value.ToString() == subFormType.Text + "-" + FormType.SelectedIndex.ToString())
                {
                    for (int col = 0; col < 26; col++)
                        dataGrid[col] = dataGridView2.Rows[id].Cells[col].Value.ToString();
                    if (dataGrid[21] == "لا") dataGrid[21] = "";                    
                    StrSpecPur = dataGridView2.Rows[id].Cells[23].Value.ToString();
                    
                    DetermineCheckBox(dataGrid[1], Convert.ToInt32(dataGrid[2]), dataGrid[3], Convert.ToInt32(dataGrid[4]), dataGrid[5], Convert.ToInt32(dataGrid[6]), dataGrid[7], Convert.ToInt32(dataGrid[8]), dataGrid[9], Convert.ToInt32(dataGrid[10]), dataGrid[11], dataGrid[21], dataGrid[13], dataGrid[14].Split('_'), dataGrid[19], dataGrid[16], dataGrid[17].Split('_'));
                    return true;
                }
            }
            return false;
        }

        private void DetermineCheckBox(string v1, int s1, string v2, int s2, string v3, int s3, string v4, int s4, string v5, int s5, string v6, string v7, string v81, string[] v8, string button, string v91, string[] v9)
        {
            restShowingItems();
            if (v1 != "")
            {
                //MessageBox.Show(s1.ToString());
                txt1.RightToLeft = RightToLeft.Yes;
                labeltxt1.Text = v1;
                labeltxt1.Visible = true;
                if (s1 < 700)
                    txt1.Size = new System.Drawing.Size(s1, 35);
                else
                {
                    txt1.Multiline = true;
                    txt1.Size = new System.Drawing.Size(s1, 135);
                }
                txt1.Visible = true;
            }
            if (v2 != "")
            {
                labeltxt2.Text = v2;
                labeltxt2.Visible = true;
                if (s2 < 700)
                    txt2.Size = new System.Drawing.Size(s2, 35);
                else
                {
                    txt2.Multiline = true;
                    txt2.Size = new System.Drawing.Size(s2, 135);
                }
                txt2.Visible = true;
            }
            if (v3 != "")
            {
                labeltxt3.Text = v3;
                labeltxt3.Visible = true;
                if (s3 < 700)
                    txt3.Size = new System.Drawing.Size(s3, 35);
                else
                {
                    txt3.Multiline = true;
                    txt3.Size = new System.Drawing.Size(s3, 135);
                }
                txt3.Visible = true;
            }
            if (v4 != "")
            {
                labeltxt4.Text = v4;
                labeltxt4.Visible = true;
                if (s4 < 700)
                    txt4.Size = new System.Drawing.Size(s4, 35);
                else
                {
                    txt4.Multiline = true;
                    txt4.Size = new System.Drawing.Size(s4, 146);
                }
                txt4.Visible = true;
            }
            if (v5 != "")
            {
                labeltxt5.Text = v5;
                labeltxt5.Visible = true;
                if (s5 < 700)
                    txt5.Size = new System.Drawing.Size(s5, 35);
                else
                {
                    txt5.Multiline = true;
                    txt5.Size = new System.Drawing.Size(s5, 135);
                }
                txt5.Visible = true;
            }


            if (v6 != "")
            {
                DPTitle[0] = dataGrid[12];
                labelcheck1.Text = v6;
                labelcheck1.Visible = true;
                if (DPTitle[0].Contains("_")) check1.Text = DPTitle[0].Split('_')[0];
                else check1.Text = DPTitle[0];
                check1.Visible = true;
            }

            if (v7 != "")
            {
                labeldate1.Visible = true;
                lblD1.Visible = txtD1.Visible = lalM1.Visible = txtM1.Visible = lalY1.Visible = txtY1.Visible = true;
                labeldate1.Text = v7;
            }


            if (v8[0] != "")
            {
                labelcomb1.Visible = true;
                combo1.Visible = true;
                labelcomb1.Text = v81;

                combo1.Items.Clear();
                for (int x = 0; x < v8.Length; x++)
                    combo1.Items.Add(v8[x]);
                combo1.SelectedIndex = 0;
            }
            if (button != "")
            {
                addName1.Text = button;
                addName1.Visible = true;
            }
            if (v9[0] != "")
            {
                labelcomb2.Visible = true;
                combo2.Visible = true;
                labelcomb2.Text = v91;

                combo2.Items.Clear();
                for (int x = 0; x < v9.Length; x++)
                    combo2.Items.Add(v9[x]);
                combo2.SelectedIndex = 0;
            }
        }

        private void subFormType_TextChanged(object sender, EventArgs e)
        {

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
                    newData = false;
                    SaveEdit = true;
                    colIDs[7] = "new";
                    activeCopy = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("mmss") + ".docx";
                    Iqrarid.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                    //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 1);
                    FillDatafromGenArch("data1", ApplicantID.ToString());
                    if (Jobposition.Contains("قنصل")) deleteRow.Visible = true;
                    return;
                }
                colIDs[7] = "old";
                SaveEdit = false;
                ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());

                Iqrarid.Text = PreAppId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                if (Jobposition.Contains("قنصل")) deleteRow.Visible = true;
                ApplicantName.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                DocType.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                DocNo.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                DocSource.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();

                text.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();


                التاريخ_الميلادي.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                التاريخ_الهجري.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[12].Value.ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;

                }
                else checkedViewed.CheckState = CheckState.Checked;
                AppType.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked;
                else AppType.CheckState = CheckState.Unchecked;
                ConsulateEmployee.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                if (AppType.CheckState == CheckState.Unchecked)
                {
                    mandoubVisibilty(); mandoubName.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                }
                PreRelatedID = dataGridView1.CurrentRow.Cells[20].Value.ToString();
                Comment.Text = dataGridView1.CurrentRow.Cells[21].Value.ToString();
                language.Text = dataGridView1.CurrentRow.Cells[23].Value.ToString();
                if (language.Text == "العربية")
                    language.CheckState = CheckState.Unchecked;
                else
                    language.CheckState = CheckState.Checked;
                if (!language.Checked)
                {
                    if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
                    else if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
                }
                else {
                    titleEng.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                }
                if (dataGridView1.CurrentRow.Cells[22].Value.ToString() != "غير مؤرشف")
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
                textItems = dataGridView1.CurrentRow.Cells[24].Value.ToString();
                
                if (!String.IsNullOrEmpty(textItems))
                {
                    string[] SI = textItems.Split('_');
                    if (SI[0].Contains("-"))
                        SI[0] = SI[0].Split('-')[1];
                    if (SI.Length == 10)
                    DetermineData(SI[0], SI[1], SI[2], SI[3], SI[4], SI[5], SI[6], SI[7], SI[8], SI[9]);
                    else if(SI.Length == 7)
                        DetermineData(SI[0], SI[1], SI[2], SI[3], SI[4], SI[5], SI[6]);
                }

                activeCopy = dataGridView1.CurrentRow.Cells[25].Value.ToString();
                WitNessesInfo = dataGridView1.CurrentRow.Cells[26].Value.ToString();
                if (WitNessesInfo != "") 
                {
                    string[] witinfo = WitNessesInfo.Split('_');
                    txtWitName1.Text = witinfo[0];
                    combTitle13.Text = witinfo[1]; 
                    txtWitPass1.Text = witinfo[2];
                    txtWitName2.Text = witinfo[3];
                    combTitle14.Text = witinfo[4]; 
                    txtWitPass2.Text = witinfo[5];

                }
                int mainIndex = 0;
                if (dataGridView1.CurrentRow.Cells[27].Value.ToString() != "")
                    mainIndex = Convert.ToInt32(dataGridView1.CurrentRow.Cells[27].Value.ToString());
                FormType.SelectedIndex= mainIndex;

                subFormType.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                //if (FormType.Text == "Affidavit of Financial support")
                //{
                //    string[] strText = textSave.Split('-');
                //    txt1.Text = strText[14];
                //    txt2.Text = strText[15];
                //    txt3.Text = strText[16];
                //    txt4.Text = strText[17];
                //    txt5.Text = strText[18];
                //    checkSexType.Text = strText[19];
                //    string[] dateStr = strText[20].Split('/');
                //    dd1.Text = dateStr[0];
                //    mm1.Text = dateStr[1];
                //    yy1.Text = dateStr[2];

                //    strlistdata = strText[21];
                //    strlistdata1 = strText[22];

                //    if (language.CheckState == CheckState.Checked) countryNonArab.Text = strlistdata1;
                //    countryArab.Text = strlistdata1;
                //    combo1.Items.Clear();
                //    for (int x = 0; x < strlistdata.Split('_').Length; x++)
                //    {
                //        combo1.Items.Add(strlistdata.Split('_')[x]);
                //    }
                //    combo1.SelectedIndex = 0;
                //    optionPanel(strText[0], Convert.ToInt32(strText[1]), strText[2], Convert.ToInt32(strText[3]), strText[4], Convert.ToInt32(strText[5]), strText[6], Convert.ToInt32(strText[7]), strText[8], Convert.ToInt32(strText[9]), strText[10], strText[11], strText[12].Split('_'), strText[13]);
                //}
                strAuth();
                ArchivedSt.Visible = true;
            }
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

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void mandoubName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void mandoubLabel_Click(object sender, EventArgs e)
        {

        }

        private void SearchFile_TextChanged(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '%" + ListSearch.Text + "%'";
            dataGridView1.DataSource = bs;
        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void ListSearch_TextChanged_1(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '%" + ListSearch.Text + "%'";
            dataGridView1.DataSource = bs;
        }

        private void btnEditID_Click(object sender, EventArgs e)
        {
            if (btnEditID.Text == "إجراء")
            {
                btnEditID.Text = "تعديل";
                SqlConnection sqlCon = new SqlConnection(DataSource);
                SqlCommand sqlCmd = new SqlCommand("update TableFreeForm SET DocID = @DocID WHERE ID = @ID", sqlCon);
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

        void strAuth()
        {
            if (language.CheckState == CheckState.Unchecked)
            {
                if (FormType.Text == "إقرار" || FormType.Text == "إقرار مشفوع باليمين")
                {
                    FormTypeDocx = "GenIqrarM.docx";
                    if (AppType.CheckState == CheckState.Checked)
                    {
                        if (ApplicantSex.CheckState == CheckState.Unchecked)
                            Auth = "أشهد أنا/" + AttendViceConsul.Text + " نائب قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد حضر للقنصلية ووقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                        else
                        {
                            Auth = "أشهد أنا/ " + AttendViceConsul.Text + " نائب قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد حضرت للقنصلية ووقعت بتوقيعها على هذا الإقرار بعد تلاوته عليها وبعد أن فهمت مضمونه ومحتواه. ";
                            FormTypeDocx = "GenIqrarF.docx";
                        }
                    }
                    else
                    {
                        string[] strmandoub = new string[2];
                        strmandoub = mandoubName.Text.Split('-');
                        if (ApplicantSex.CheckState == CheckState.Unchecked)

                            if (strmandoub[1].Trim() != "القنصلية العامة لجمهورية السودان بجدة")
                            {
                                Auth = "المواطن المذكور أعلاه حضر ووقع بتوقيعه على هذا الإقرار أمام مندوب جالية منطقة" + strmandoub[1] + " السيد/ " + strmandoub[0] + "، وذلك بموجب التفويض الممنوح له من القنصلية العامة، ";
                            }
                            else { Auth = "المواطن المذكور أعلاه حضر ووقع بتوقيعه على هذا الإقرار أمام مندوب " + strmandoub[1] + " السيد/ " + strmandoub[0] + "، وذلك بموجب التفويض الممنوح له من القنصلية العامة، "; }
                        if (ApplicantSex.CheckState == CheckState.Checked)
                        {
                            FormTypeDocx = "GenIqrarF.docx";
                            if (strmandoub[1].Trim() != "القنصلية العامة لجمهورية السودان بجدة")
                            {
                                Auth = "المواطنة المذكورة أعلاه حضرت ووقعت بتوقيعها على هذا الإقرار أمام مندوب جالية منطقة" + strmandoub[1] + " السيد/ " + strmandoub[0] + "، وذلك بموجب التفويض الممنوح له من القنصلية العامة، ";
                            }
                            else { Auth = "المواطن المذكور أعلاه حضر ووقع بتوقيعه على هذا الإقرار أمام مندوب " + strmandoub[1] + " السيد/ " + strmandoub[0] + "، وذلك بموجب التفويض الممنوح له من القنصلية العامة، "; }
                        }



                    }
                }
                else if (FormType.SelectedIndex == 2)
                {
                    FormTypeDocx = "GenIfada.docx";
                    Auth = "";
                }
                else if (FormType.SelectedIndex == 3)
                {
                    FormTypeDocx = "GenCertiNoObj.docx";
                    Auth = "";
                }
                

            }
            else
            {
                if (FormType.SelectedIndex == 0)
                    FormTypeDocx = "GenIqrarEng.docx";
                if (FormType.SelectedIndex == 1)
                    FormTypeDocx = "GenIqrarEng.docx";
                else if (FormType.SelectedIndex == 3)
                    FormTypeDocx = "EngGenIfada.docx";
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
                if (control is ComboBox)
                {
                    ((ComboBox)control).Visible = false;
                }
            }
            check1.Visible = false;
            combo1.Visible = false;            
            addName1.Visible = false;
        }

        private void ExtendedFillBox(string v1, int s1, string v2, int s2, string v3, int s3, string v4, int s4, string v5, int s5, string vE1, int sE1, string vE2, int sE2, string vE3, int sE3, string vE4, int sE4, string vE5, int sE5, string vE61, string vE62, string vE63, string vE64, string vE65, string vE71, string vE72, string vE73, string vE74, string vE75, string v81, string[] vE81, string v82, string[] vE82, string v83, string[] vE83, string v84, string[] vE84, string v85, string[] vE85, string button1, string button2, string button3, string button4, string button5)
        {
            restShowingItems();
            if (v1 != "")
            {
                labeltxt1.Text = v1;
                labeltxt1.Visible = true;
                txt1.Width = s1;
                txt1.Visible = true;
            }
            if (v2 != "")
            {
                labeltxt2.Text = v2;
                labeltxt2.Visible = true;
                txt2.Width = s2;
                txt2.Visible = true;
            }
            if (v3 != "")
            {
                labeltxt3.Text = v3;
                labeltxt3.Visible = true;
                txt3.Width = s3;
                txt3.Visible = true;
            }
            if (v4 != "")
            {
                labeltxt4.Text = v4;
                labeltxt4.Visible = true;
                txt4.Width = s4;
                txt4.Visible = true;
            }
            if (v5 != "")
            {
                labeltxt5.Text = v5;
                labeltxt5.Visible = true;
                txt5.Width = s5;
                txt5.Visible = true;
            }



            if (vE1 != "")
            {
                labeltxt6.Text = vE1;
                labeltxt6.Visible = true;
                txt6.Width = sE1;
                txt6.Visible = true;
            }
            if (vE2 != "")
            {
                labeltxt7.Text = vE2;
                labeltxt7.Visible = true;
                txt7.Width = sE2;
                txt7.Visible = true;
            }
            if (vE3 != "")
            {
                labeltxt8.Text = vE3;
                labeltxt8.Visible = true;
                txt8.Width = sE3;
                txt8.Visible = true;
            }
            if (vE4 != "")
            {
                labeltxt9.Text = vE4;
                labeltxt9.Visible = true;
                txt9.Width = sE4;
                txt9.Visible = true;
            }
            if (vE5 != "")
            {
                labeltxt10.Text = vE5;
                labeltxt10.Visible = true;
                txt10.Width = sE5;
                txt10.Visible = true;
            }


            if (vE61 != "")
            {
                labelcheck1.Text = vE61;
                labelcheck1.Visible = true;
                if (DPTitle[0].Contains("_")) check1.Text = DPTitle[0].Split('_')[0];
                else check1.Text = DPTitle[0];
                check1.Visible = true;
            }
            if (vE62 != "")
            {
                labelcheck2.Text = vE62;
                labelcheck2.Visible = true;
                if (DPTitle[1].Contains("_")) check2.Text = DPTitle[1].Split('_')[0];
                else check2.Text = DPTitle[1];
                check3.Visible = true;
            }
            if (vE63 != "")
            {
                labelcheck3.Text = vE63;
                labelcheck3.Visible = true;
                if (DPTitle[2].Contains("_")) check3.Text = DPTitle[2].Split('_')[0];
                else check3.Text = DPTitle[2];
                check3.Visible = true;
            }
            if (vE64 != "")
            {
                labelcheck4.Text = vE64;
                labelcheck4.Visible = true;
                if (DPTitle[3].Contains("_")) check4.Text = DPTitle[3].Split('_')[0];
                else check4.Text = DPTitle[3];
                check4.Visible = true;
            }
            if (vE65 != "")
            {
                labelcheck5.Text = vE65;
                labelcheck5.Visible = true;
                if (DPTitle[4].Contains("_")) check5.Text = DPTitle[4].Split('_')[0];
                else check5.Text = DPTitle[4];
                check5.Visible = true;
            }

            if (vE71 != "")
            {
                labeldate1.Text = vE71;
                labeldate1.Visible = true;
                lblD1.Visible = true;
                txtD1.Visible = true;
                lalM1.Visible = true;
                txtM1.Visible = true;
                lalY1.Visible = true;
                txtY1.Visible = true;
            }

            if (vE72 != "")
            {
                labeldate2.Text = vE71;
                labeldate2.Visible = true;
                lblD2.Visible = true;
                txtD2.Visible = true;
                lalM2.Visible = true;
                txtM2.Visible = true;
                lalY2.Visible = true;
                txtY2.Visible = true;
            }
            if (vE73 != "")
            {
                labeldate3.Text = vE71;
                labeldate3.Visible = true;
                lblD3.Visible = true;
                txtD3.Visible = true;
                lalM3.Visible = true;
                txtM3.Visible = true;
                lalY3.Visible = true;
                txtY3.Visible = true;
            }
            if (vE74 != "")
            {
                labeldate4.Text = vE71;
                labeldate4.Visible = true;
                lblD4.Visible = true;
                txtD4.Visible = true;
                lalM4.Visible = true;
                txtM4.Visible = true;
                lalY4.Visible = true;
                txtY4.Visible = true;
            }
            if (vE75 != "")
            {
                labeldate5.Text = vE71;
                labeldate5.Visible = true;
                lblD5.Visible = true;
                txtD5.Visible = true;
                lalM5.Visible = true;
                txtM5.Visible = true;
                lalY5.Visible = true;
                txtY5.Visible = true;
            }

            if (v81 != "")
            {
                labelcomb1.Visible = true;
                combo1.Visible = true;
                labelcomb1.Text = v81;

                combo1.Items.Clear();
                for (int x = 0; x < vE81.Length; x++)
                    combo1.Items.Add(vE81[x]);
                combo1.SelectedIndex = 0;
            }

            if (v82 != "")
            {
                labelcomb2.Visible = true;
                combo1.Visible = true;
                labelcomb2.Text = v81;

                combo1.Items.Clear();
                for (int x = 0; x < vE82.Length; x++)
                    combo1.Items.Add(vE82[x]);
                combo1.SelectedIndex = 0;
            }

            if (v83 != "")
            {
                labelcomb3.Visible = true;
                combo1.Visible = true;
                labelcomb3.Text = v81;

                combo1.Items.Clear();
                for (int x = 0; x < vE83.Length; x++)
                    combo1.Items.Add(vE83[x]);
                combo1.SelectedIndex = 0;
            }

            if (v84 != "")
            {
                labelcomb4.Visible = true;
                combo1.Visible = true;
                labelcomb4.Text = v81;

                combo1.Items.Clear();
                for (int x = 0; x < vE84.Length; x++)
                    combo1.Items.Add(vE84[x]);
                combo1.SelectedIndex = 0;
            }

            if (v85 != "")
            {
                labelcomb5.Visible = true;
                combo1.Visible = true;
                labelcomb5.Text = v81;

                combo1.Items.Clear();
                for (int x = 0; x < vE85.Length; x++)
                    combo1.Items.Add(vE85[x]);
                combo1.SelectedIndex = 0;
            }
            if (button1 != "")
            {
                addName1.Text = button1;
                addName1.Visible = true;
            }
            if (button2 != "")
            {
                addName2.Text = button2;
                addName2.Visible = true;
            }
            if (button3 != "")
            {
                addName3.Text = button3;
                addName3.Visible = true;
            }
            if (button4 != "")
            {
                addName4.Text = button4;
                addName4.Visible = true;
            }
            if (button5 != "")
            {
                addName5.Text = button5;
                addName5.Visible = true;
            }
        }
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

        }

        private void SearchDoc_Click(object sender, EventArgs e)
        {

        }
    }
}
