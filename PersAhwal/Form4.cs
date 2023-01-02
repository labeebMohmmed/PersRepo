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
namespace PersAhwal
{
    public partial class Form4 : Form
    {
        static int i = 0;
        public static string[] DaughterMother = new string[10];
        public static string VisaLine = "";
        bool gridFill = false;
        int VisaIndex = 0;
        public static string[] titleEngFamily = new string[10];
        AutoCompleteStringCollection autoComplete;
        public static string[] Pass = new string[10];
        public static string[] IssueDate = new string[10];
        public static string[] Source = new string[10];
        string Viewed;
        bool colored = false;
        int ShowIndex = 0;
        DataTable dtbl;
        bool EditSave = false;
        DataTable dtbl2;
        string ConsulateEmpName;
        public static string ModelFileroute = "";

        String IqrarNumberPart;
        static string DataSource;
        int ApplicantID = 0;
        bool fileloaded = false;
        string NewFileName;
        string PreVisaAppId = "";
        static public string titleFam, FamilySupport;
        private string[] FamelyMember = new string[10];
        static bool Firstline = false, archived = false;
        static public string title = "";
        string PreRelatedID = "", NextRelId = "";
        int rowIndexTodelete = 0;
        string FilesPathIn, FilesPathOut;
        string Jobposition;
        int ATVC = 0;
        string[] colIDs = new string[100];
        public Form4(int Atvc, int currentRow, string EmpName, string dataSource, string filepathIn, string filepathOut, string jobposition, string gregorianDate, string hijriDate)
        {
            InitializeComponent();
            //timer1.Enabled = true;
            //timer2.Enabled = true;
            التاريخ_الميلادي.Text = gregorianDate;
            التاريخ_الهجري.Text = hijriDate;
            ATVC = Atvc;
            DataSource = dataSource;
            FilesPathIn = filepathIn + @"\";
            FilesPathOut = filepathOut;
            colIDs[4] = ConsulateEmpName = EmpName;
            Jobposition = jobposition;
            FillDataGridView();
            if (currentRow == -1) Clear_Fields();
            else SetFieldswithData(currentRow);
            if (jobposition.Contains("قنصل"))
                btnEditID.Visible = deleteRow.Visible = true;
            else btnEditID.Visible = deleteRow.Visible = false;
        }

        private string loadRerNo(int id)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT DocID from TableVisaApp where ID=@ID", sqlCon);
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
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from TableVisaApp order by ID desc", sqlCon);
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


        private void OpenFileDoc(int id, int fileNo)
        {
            string query;

            SqlConnection Con = new SqlConnection(DataSource);
            if (fileNo == 1)
            {
                query = "select Data1, Extension1,FileName1 from TableVisaApp  where ID=@id";
            }
            else if (fileNo == 2)
            {
                query = "select Data2, Extension2,FileName2 from TableVisaApp  where ID=@id";
            }
            else query = "select Data3, Extension3,FileName3 from TableVisaApp  where ID=@id";
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
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }
                else if (fileNo == 2)
                {
                    var name = reader["FileName2"].ToString();
                    var Data = (byte[])reader["Data2"];
                    var ext = reader["Extension2"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }
                else
                {
                    var name = reader["FileName3"].ToString();
                    var Data = (byte[])reader["Data3"];
                    var ext = reader["Extension3"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }

            }
            Con.Close();


        }
        private void SetFieldswithData(int Rowindex)
        {
            Rowindex--;
            ApplicantID = Convert.ToInt32(dataGridView1.Rows[Rowindex].Cells[0].Value.ToString());
            NextRelId = dataGridView1.Rows[Rowindex].Cells[1].Value.ToString();
            مقدم_الطلب_1.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString();
            language.Text = dataGridView1.Rows[Rowindex].Cells[3].Value.ToString();
            رقم_الهوية_1.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString();
            مكان_الإصدار_1.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString();
            string IssueDate = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString();
            string[] YearMonthDay = IssueDate.Split('/');
            //yy1.Text = YearMonthDay[2];
            dd1.Text = YearMonthDay[1];
            mm1.Text = YearMonthDay[0];
            if (language.Text == "الانجليزية")
            {
                language.CheckState = CheckState.Unchecked;
                txtCountry.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString();
            }
            else
            {
                txtCountry.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString();
                language.CheckState = CheckState.Checked;
            }
            التاريخ_الميلادي.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString();
            التاريخ_الهجري.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString();

            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[12].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[13].Value.ToString().ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;

            }
            else checkedViewed.CheckState = CheckState.Checked;

            AppType.Text = dataGridView1.Rows[Rowindex].Cells[14].Value.ToString();
            if (AppType.Text == "حضور مباشرة إلى القنصلية")
                AppType.CheckState = CheckState.Checked;
            else AppType.CheckState = CheckState.Unchecked;
            ConsulateEmployee.Text = dataGridView1.Rows[Rowindex].Cells[15].Value.ToString();
            if (AppType.CheckState == CheckState.Unchecked)
            {
                mandoubVisibilty(); mandoubName.Text = dataGridView1.Rows[Rowindex].Cells[16].Value.ToString();
            }
            TravPurpose.Text = dataGridView1.Rows[Rowindex].Cells[17].Value.ToString();
            PreRelatedID = dataGridView1.Rows[Rowindex].Cells[19].Value.ToString();
            Comment.Text = dataGridView1.Rows[Rowindex].Cells[24].Value.ToString();
            if (dataGridView1.CurrentRow.Cells[25].Value.ToString() != "غير مؤرشف")
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
            labelArch.Visible = true;
            btnprintOnly.Visible = true;
            SaveOnly.Visible = true;
            btnSavePrint.Text = "حفظ";
            btnSavePrint.Visible = false;
        }

        private void FillDataGridView()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("VisaViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ApplicantName", Search.Text.Trim());
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            IqrarNumberPart = (dtbl.Rows.Count + 1).ToString();
            dataGridView1.Columns[27].Visible = false;
            sqlCon.Close();
            NewFileName = IqrarNumberPart + "_04";
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 450;
        }

        private void Save2DataBase(bool newData)
        {
            addNewAppNameInfo(مقدم_الطلب_1, رقم_الهوية_1, مكان_الإصدار_1);
            if (مقدم_الطلب_2.Text != "") addNewAppNameInfo(مقدم_الطلب_2, رقم_الهوية_2, مكان_الإصدار_2);
            if (مقدم_الطلب_3.Text != "") addNewAppNameInfo(مقدم_الطلب_3, رقم_الهوية_3, مكان_الإصدار_3);
            if (مقدم_الطلب_4.Text != "") addNewAppNameInfo(مقدم_الطلب_4, رقم_الهوية_4, مكان_الإصدار_4);
            if (مقدم_الطلب_5.Text != "") addNewAppNameInfo(مقدم_الطلب_5, رقم_الهوية_5, مكان_الإصدار_5);

            SqlConnection sqlCon = new SqlConnection(DataSource);
            string[] strLines = VisaLine.Split('*');
            string[] str = strLines[0].Split('-');
            string AppGender;
            try
            {
                if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                else Viewed = "غير معالج";

                if (newData)
                {
                    SqlCommand sqlCmd = new SqlCommand("VisaAddorEdit", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    if (sqlCon.State == ConnectionState.Closed)
                        sqlCon.Open();

                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");
                    sqlCmd.Parameters.AddWithValue("@DocID", VisaAppId.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", str[2].Trim());
                    sqlCmd.Parameters.AddWithValue("@lang", str[0].Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", str[1].Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", str[3].Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", str[4].Trim());
                    sqlCmd.Parameters.AddWithValue("@DocIssueDate", str[5]);
                    sqlCmd.Parameters.AddWithValue("@CountryDest", txtCountry.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AllFamilyMembers", "");
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@TravelPurpose", str[7]);
                    sqlCmd.Parameters.AddWithValue("@RelatedVisaApp", PreRelatedID);
                    string filePath1 = FilesPathIn + "text1.txt";
                    string filePath2 = FilesPathIn + "text2.txt";
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");
                    sqlCmd.Parameters.AddWithValue("@Apptitle", str[1]);
                    sqlCmd.Parameters.AddWithValue("@FullTextData", VisaLine);

                    sqlCmd.ExecuteNonQuery();


                }
                else
                {
                    SqlCommand sqlCmd = new SqlCommand("UPDATE TableVisaApp SET DocID=@DocID,AppName=@AppName,lang=@lang,Gender=@Gender,DocNo=@DocNo,DocSource=@DocSource,DocIssueDate=@DocIssueDate,CountryDest=@CountryDest,GriDate=@GriDate,Hijri=@Hijri,AllFamilyMembers=@AllFamilyMembers,AtteVicCo=@AtteVicCo,Viewed =@Viewed,DataInterType =@DataInterType,DataInterName =@DataInterName, DataMandoubName =@DataMandoubName, TravelPurpose =@TravelPurpose,RelatedVisaApp = @RelatedVisaApp,Comment=@Comment,ArchivedState=@ArchivedState,Apptitle=@Apptitle,FullTextData=@FullTextData WHERE ID = @ID", sqlCon);
                    sqlCmd.CommandType = CommandType.Text;
                    if (sqlCon.State == ConnectionState.Closed)
                        sqlCon.Open();
                    sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                    sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                    sqlCmd.Parameters.AddWithValue("@DocID", VisaAppId.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@lang", str[0].Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", str[2].Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", str[1].Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", str[3].Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", str[4].Trim());
                    sqlCmd.Parameters.AddWithValue("@DocIssueDate", str[5]);
                    sqlCmd.Parameters.AddWithValue("@CountryDest", txtCountry.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AllFamilyMembers", "");
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@TravelPurpose", str[7].Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedVisaApp", PreVisaAppId);
                    string filePath1 = FilesPathIn + "text1.txt";
                    string filePath2 = FilesPathIn + "text2.txt";

                    if (Search.Text != "") { filePath2 = Search.Text; fileloaded = true; }
                    
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
                    if (fileloaded)
                        sqlCmd.Parameters.AddWithValue("@ArchivedState", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    else sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");
                    sqlCmd.Parameters.AddWithValue("@Apptitle", str[1]);
                    sqlCmd.Parameters.AddWithValue("@FullTextData", VisaLine);
                    sqlCmd.ExecuteNonQuery();
                    btnprintOnly.Visible = false;
                    SaveOnly.Visible = false;
                    btnSavePrint.Visible = true;
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

        void ReverseDGVRows(DataGridView dgv)
        {
            List<DataGridViewRow> rows = new List<DataGridViewRow>();
            rows.AddRange(dgv.Rows.Cast<DataGridViewRow>());
            rows.Reverse();
            dgv.Rows.Clear();
            dgv.Rows.AddRange(rows.ToArray());
        }

        private void language_CheckedChanged_1(object sender, EventArgs e)
        {
            Check_lang();
        }

        private void Check_lang()
        {
            if (language.CheckState == CheckState.Unchecked)
            {
                language.Text = "الانجليزية";
                title1.Visible = true;
                title2.Visible = true;
                title3.Visible = true;
                title4.Visible = true;
                title5.Visible = true;
                
                autoCompleteTextBox(textBox1, DataSource, "ForiegnCountries", "TableListCombo");
                autoCompleteTextBox(مكان_الإصدار_1, DataSource, "KSAIssureSource", "TableListCombo");
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("en-US");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
                dataGridView3.Columns[0].Visible = false;
                dataGridView3.Columns[1].Visible = true;




            }
            else
            {
                language.Text = "العربية";
                title1.Visible = false;
                title2.Visible = false;
                title3.Visible = false;
                title4.Visible = false;
                title5.Visible = false;
                autoCompleteTextBox(textBox1, DataSource, "ArabCountries", "TableListCombo");
                autoCompleteTextBox(مكان_الإصدار_1, DataSource, "SDNIssueSource", "TableListCombo");
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
                dataGridView3.Columns[0].Visible = true;
                dataGridView3.Columns[1].Visible = false;
            }
        }

        private void Review_Click(object sender, EventArgs e)
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

        private void timer2_Tick(object sender, EventArgs e)
        {
            //CultureInfo arSA = new CultureInfo("ar-SA");
            //arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            //Thread.CurrentThread.CurrentCulture = arSA;
            //new System.Globalization.GregorianCalendar();
            //التاريخ_الميلادي.Text = DateTime.Now.ToString("MM-dd-yyyy");
            //timer2.Enabled = false;
        }



        private void ApplicantSex_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void ApplicantSex_CheckedChanged_2(object sender, EventArgs e)
        {

            int y1 = 0, y2 = 0, y3 = 0, y4 = 0;
            foreach (Control control in panel1.Controls)
            {
                if (control is CheckBox)
                {
                    if (language.CheckState == CheckState.Checked)
                    {
                        if (((CheckBox)control).CheckState  == CheckState.Unchecked)
                        {
                            ((CheckBox)control).Text = "ذكر";
                        }
                        else if (((CheckBox)control).CheckState == CheckState.Checked)
                        {
                            ((CheckBox)control).Text = "إنثى";
                        }
                    }
                }
            }
            
        }

        private void titleEng_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void motherDaughter_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form4_Load(object sender, EventArgs e)
        {
            FillDataGridAdd();
            autoCompleteTextBox1(مقدم_الطلب_1, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox1(مقدم_الطلب_2, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox1(مقدم_الطلب_3, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox1(مقدم_الطلب_4, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox1(مقدم_الطلب_5, DataSource, "الاسم", "TableGenNames");

            fileComboBox(mandoubName, DataSource, "MandoubNames", "TableListCombo");
            //autoCompleteTextBox(countryNonArab);
            //autoCompleteTextBox(countryArab);

            //SqlConnection sqlCon = new SqlConnection(DataSource);
            //if (sqlCon.State == ConnectionState.Closed)
            //    sqlCon.Open();
            //SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT ArabCountries from TableListCombo", sqlCon);
            //sqlDa.SelectCommand.CommandType = CommandType.Text;           
            //dtbl = new DataTable();
            //sqlDa.Fill(dtbl);

            //sqlDa = new SqlDataAdapter("SELECT ForiegnCountries from TableListCombo", sqlCon);
            //sqlDa.SelectCommand.CommandType = CommandType.Text;
            //dtbl = new DataTable();
            //sqlDa.Fill(dtbl);
            //dataGridView2.DataSource = dtbl;
            //sqlCon.Close();
            fileComboBox(AttendViceConsul, DataSource, "ArabicAttendVC", "TableListCombo");

            autoCompleteTextBox(مكان_الإصدار_1, DataSource, "SDNIssueSource", "TableListCombo");
            //autoCompleteTextBox(textBox1, DataSource, "ArabCountries", "TableListCombo");
            AttendViceConsul.SelectedIndex = ATVC;
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
                        if (newSrt) autoComplete.Add(dataRow[comlumnName].ToString());
                    }
                }
                textbox.AutoCompleteCustomSource.Clear();
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

        private void search_TextChanged(object sender, EventArgs e)
        {
            FillDataGridView();
        }

        private void AppType_CheckedChanged_1(object sender, EventArgs e)
        {
            mandoubVisibilty();
        }

        private void ResetAll_Click_1(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void Clear_Fields()
        {
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();

            التاريخ_الميلادي.Text = DateTime.Now.ToString("dd-MM-yyyy"); language.Enabled = true;
            VisaIndex = 0;
            VisaLine = "";
            مقدم_الطلب_1.Text = مكان_الإصدار_1.Text = "";
            AttendViceConsul.SelectedIndex = 2;
            رقم_الهوية_1.Text = "P0";yy1.Text = "";

            language.Text = "العربية";

            TravPurpose.SelectedIndex = 0;
            EditSave = false;
            رقم_الهوية_1.Text = "P0";
            txtCountry.Text = "";
            language.CheckState = CheckState.Checked;

            //VisaAppId.Text = "ق س ج/80/" + GregorianDate.Text.Split('-')[2].Replace("20", "") + "/04/" + loadRerNo(loadIDNo());
            mandoubName.Text = Search.Text = "";
            AppType.CheckState = CheckState.Checked;
            mandoubVisibilty();
            btnprintOnly.Visible = false;
            btnSavePrint.Text = "حفظ وطباعة";
            btnSavePrint.Visible = true;
            SaveOnly.Visible = false;
            Comment.Text = "لا تعليق";
            FillDataGridView();
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            labelArch.Visible = false;
            ArchivedSt.BackColor = Color.Red;
            ConsulateEmployee.Text = ConsulateEmpName;
            i = 0;
            dataGridView1.Visible = true;
            PanelMain.Visible = false;
            VisaAppId.Text = "";
        }

        private void printOnly_Click(object sender, EventArgs e)
        {

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
                var Data = (byte[])reader["Data1"];
                var ext = reader["Extension1"].ToString();
                var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                File.WriteAllBytes(NewFileName, Data);
                System.Diagnostics.Process.Start(NewFileName);
            }


            sqlCon.Close();
        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            var selectRows = dataGridView1.SelectedRows;
            foreach (var row in selectRows)
            {
                int id = (int)((DataGridViewRow)row).Cells[0].Value;
                //OpenFile(id, 1);
                FillDatafromGenArch("data1", colIDs[1], "TableVisaApp");
            }
            if (ApplicantID != 0) FillDatafromGenArch("data1", colIDs[1], "TableVisaApp");
            //ApplicantID = 0;
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            var selectRows = dataGridView1.SelectedRows;
            foreach (var row in selectRows)
            {
                int id = (int)((DataGridViewRow)row).Cells[0].Value;
                FillDatafromGenArch("data2", colIDs[1], "TableVisaApp");
            }
            if (ApplicantID != 0) FillDatafromGenArch("data2", colIDs[1], "TableVisaApp");
            //ApplicantID = 0;
        }

        //private void OpenFile(int id, int fileNo)
        //{
        //    string query;

        //    SqlConnection Con = new SqlConnection(DataSource);
        //    if (fileNo == 1)
        //    {
        //        query = "select Data1, Extension1,FileName1 from TableVisaApp where ID=@id";
        //    }
        //    else
        //    {
        //        query = "select Data2, Extension2,FileName2 from TableVisaApp where ID=@id";
        //    }
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
        //        else
        //        {
        //            var name = reader["FileName2"].ToString();
        //            var Data = (byte[])reader["Data2"];
        //            var ext = reader["Extension2"].ToString();
        //            var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
        //            File.WriteAllBytes(NewFileName, Data);
        //            System.Diagnostics.Process.Start(NewFileName);
        //        }

        //    }
        //    Con.Close();

        //}

        private void btnSavePrint_Click_1(object sender, EventArgs e)
        {
            if (VisaAppId.Text == "")
            {
                MessageBox.Show("رقم خطاب خاطئ");
                return;
            }
            AddData(true);
            btnSavePrint.Enabled = false;
            btnSavePrint.Text = "جاري المعالجة";
            if (!EditSave)
                CreateWordFile(false);
            else CreateWordFile(true);
            btnSavePrint.Text = "حفظ وطباعة";
            btnSavePrint.Enabled = true;
            this.Close();
        }

        private void btnprintOnly_Click(object sender, EventArgs e)
        {
            
            addNewAppNameInfo(مقدم_الطلب_1, رقم_الهوية_1, مكان_الإصدار_1);
            if(مقدم_الطلب_2.Text != "") addNewAppNameInfo(مقدم_الطلب_2, رقم_الهوية_2, مكان_الإصدار_2);
            if(مقدم_الطلب_3.Text != "") addNewAppNameInfo(مقدم_الطلب_3, رقم_الهوية_3, مكان_الإصدار_3);
            if(مقدم_الطلب_4.Text != "") addNewAppNameInfo(مقدم_الطلب_4, رقم_الهوية_4, مكان_الإصدار_4);
            if(مقدم_الطلب_5.Text != "") addNewAppNameInfo(مقدم_الطلب_5, رقم_الهوية_5, مكان_الإصدار_5);
            
            if (VisaAppId.Text == "")
            {
                MessageBox.Show("رقم خطاب خاطئ");
                return;
            }
            AddData(true);
            btnprintOnly.Enabled = false;
            btnprintOnly.Text = "طباعة";
            //if (ApplicantName.Text != "") AddData(false);
            CreateWordFile(true);
            btnSavePrint.Text = "حفظ وطباعة";
            btnprintOnly.Enabled = true;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void PassIssueDate_ValueChanged(object sender, EventArgs e)
        {

        }


        private void deleteRow_Click_1(object sender, EventArgs e)
        {
            deleteRowsData(ApplicantID, "TableVisaApp", DataSource);
            deleteRow.Visible = false;
            Clear_Fields();
        }

        private void autoCompleteTextBox(ComboBox combbox)
        {
            autoComplete = new AutoCompleteStringCollection();
            for (int item = 0; item < combbox.Items.Count; item++)
            {
                autoComplete.Add(combbox.Items[item].ToString());
            }
            combbox.AutoCompleteMode = AutoCompleteMode.Suggest;
            combbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
            combbox.AutoCompleteCustomSource = autoComplete;
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

        private void button1_Click_2(object sender, EventArgs e)
        {
            if (مقدم_الطلب_1.Text == "") return;
            AddData(true);
        }

        private void AddData(bool reset)
        {
            VisaIndex = 0;
            VisaLine = "";
            string title = "";

            if (language.Text == "العربية")
                title = sex1.Text;
            else title = title1.Text;
            VisaLine = language.Text + "-" + title + "-" + مقدم_الطلب_1.Text + "-" + رقم_الهوية_1.Text + "-" + مكان_الإصدار_1.Text + "-" + dd1.Text + "/" + mm1.Text + "/" + yy1.Text + "-" + txtCountry.Text + "-" + TravPurpose.Text;
            VisaIndex = 1;
            if (language.Text == "العربية")
                title = sex2.Text;
            else title = title2.Text;
            if (مقدم_الطلب_2.Text != "")
            {
                VisaLine = VisaLine + "*" + language.Text + "-" + title + "-" + مقدم_الطلب_2.Text + "-" + رقم_الهوية_2.Text + "-" + مكان_الإصدار_2.Text + "-" + dd2.Text + "/" + mm2.Text + "/" + yy2.Text + "-" + txtCountry.Text + "-" + TravPurpose.Text;
                VisaIndex = 2;
            }

            if (language.Text == "العربية")
                title = sex3.Text;
            else title = title3.Text;
            if (مقدم_الطلب_2.Text != "" && مقدم_الطلب_3.Text != "")
            {
                VisaLine = VisaLine + "*" + language.Text + "-" + title + "-" + مقدم_الطلب_3.Text + "-" + رقم_الهوية_3.Text + "-" + مكان_الإصدار_3.Text + "-" + dd3.Text + "/" + mm3.Text + "/" + yy3.Text + "-" + txtCountry.Text + "-" + TravPurpose.Text;
                VisaIndex = 3;
            }
            if (language.Text == "العربية")
                title = sex4.Text;
            else title = title4.Text;
            if (مقدم_الطلب_2.Text != "" && مقدم_الطلب_3.Text != ""&& مقدم_الطلب_4.Text != "")
            {
                VisaLine = VisaLine + "*" + language.Text + "-" + title + "-" + مقدم_الطلب_4.Text + "-" + رقم_الهوية_4.Text + "-" + مكان_الإصدار_4.Text + "-" + dd4.Text + "/" + mm4.Text + "/" + yy4.Text + "-" + txtCountry.Text + "-" + TravPurpose.Text;
                VisaIndex = 4;
            }
            if (language.Text == "العربية")
                title = sex5.Text;
            else title = title5.Text;
            if (مقدم_الطلب_2.Text != "" && مقدم_الطلب_3.Text != "" && مقدم_الطلب_4.Text != "" && مقدم_الطلب_5.Text != "")
            {
                VisaLine = VisaLine + "*" + language.Text + "-" + title + "-" + مقدم_الطلب_5.Text + "-" + رقم_الهوية_5.Text + "-" + مكان_الإصدار_5.Text + "-" + dd5.Text + "/" + mm5.Text + "/" + yy5.Text + "-" + txtCountry.Text + "-" + TravPurpose.Text;

                VisaIndex = 5;
            }


            //VisaIndex++;
            //addbtn(VisaIndex.ToString(), ApplicantName1.Text);
            //if (reset)
            //{
            //    language.Enabled = TravPurpose.Enabled = false;
            //    ApplicantName1.Text = IssuedSource1.Text = dd1.Text = mm1.Text = yy1.Text = "";
            //    ApplicantIdoc1.Text = "P0";
            //}
            //btnAddInfo.Text = "اضافة (" + VisaIndex.ToString() + "/" + ShowIndex.ToString() + ")";
        }


        private void SetData(int visaIndex, string str)
        {

            string[] strLines = str.Split('*');
            string[] strLine = new string[8];
            if (visaIndex < strLines.Length && visaIndex > -1)
                strLine = strLines[visaIndex].Split('-');
            else return;
            string Str = strLines[visaIndex];
            if (strLine[0] == "العربية")
            {
                language.CheckState = CheckState.Checked;
                language.Text = "العربية";
                
                txtCountry.Text = strLine[6];
            }
            else
            {
                language.CheckState = CheckState.Unchecked;                
                txtCountry.Text = strLine[6];
            }

            مقدم_الطلب_1.Text = strLine[2];
            رقم_الهوية_1.Text = strLine[3];
            مكان_الإصدار_1.Text = strLine[4];
            string[] YearMonthDay = strLine[5].Split('/');
            yy1.Text = YearMonthDay[2];
            dd1.Text = YearMonthDay[0];
            mm1.Text = YearMonthDay[1];
            TravPurpose.Text = strLine[7];
        }

        private void countryNonArab_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        void FillDataGridAdd()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();

            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT ArabCountries,ForiegnCountries from TableListCombo", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            dtbl2 = new DataTable();
            sqlDa.Fill(dtbl2);
            dataGridView3.DataSource = dtbl2;
            sqlCon.Close();
            dataGridView3.Columns[0].Width = 150;
            dataGridView3.Columns[1].Width = 150;
            if (language.CheckState == CheckState.Unchecked)
            {
                dataGridView3.Columns[0].Visible = false;
                dataGridView3.Columns[1].Visible = true;
            }
            else
            {
                dataGridView3.Columns[0].Visible = true;
                dataGridView3.Columns[1].Visible = false;
            }

        }

        private void countryArab_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void countryNonArab_TextChanged(object sender, EventArgs e)
        {
            // FillDataGridAdd(countryNonArab.Text, true);
        }

        private void countryArab_TextChanged(object sender, EventArgs e)
        {
            // FillDataGridAdd(countryArab.Text, false);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            foreach (DataRow dataRow in dtbl2.Rows)
            {
                if (language.CheckState == CheckState.Checked)
                {
                    if (dataRow["ArabCountries"].ToString().Contains(textBox1.Text))
                    {
                        txtCountry.Text = dataRow["ArabCountries"].ToString();
                    }
                }
                else if (language.CheckState == CheckState.Unchecked)
                {
                    if (dataRow["ForiegnCountries"].ToString().Contains(textBox1.Text))
                    {
                        txtCountry.Text = dataRow["ForiegnCountries"].ToString();
                    }
                }
            }
        }

        
        
        

        private void SaveOnly_Click(object sender, EventArgs e)
        {
            //if (!checkGender(PanelMain, "مقدم_الطلب", "النوع"))
            //{
            //    return;
            //}
            //else addNewAppNameInfo(مقدم_الطلب_1); 
            
            if (VisaAppId.Text == "")
            {
                MessageBox.Show("رقم خطاب خاطئ");
                return;
            }
            AddData(true);
            if (VisaIndex > 1)
            {
                string[] strLines = VisaLine.Split('*');
                Save2DataBase(false);
            }
            else Save2DataBase(false);
            
        }

        private void addNewAppNameInfo(TextBox مقدم_الطلب, TextBox رقم_الهوية, TextBox مكان_الإصدار)
        {

            string query = "insert into TableGenNames ([الاسم], رقم_الهوية,نوع_الهوية,مكان_الإصدار) values (@col1,@col2,@col6,@col7) ;SELECT @@IDENTITY as lastid";
            string id = checkExist(مقدم_الطلب.Text);
            if (id != "0")
            {
                query = "update TableGenNames set [الاسم] =  @col1,[رقم_الهوية] = @col2,نوع_الهوية = @col6,مكان_الإصدار = @col7 where ID = " + id;
                //MessageBox.Show(query);
            }
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();

            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@col1", مقدم_الطلب.Text);
            sqlCommand.Parameters.AddWithValue("@col2", رقم_الهوية.Text);            
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
        private void dataGridView1_Click(object sender, EventArgs e)
        {
            
        }

       

        private void Search_TextChanged_1(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '" + Search.Text + "%'";
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

                if (dataGridView1.Rows[i].Cells[25].Value.ToString() == "مؤرشف نهائي")
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                if (dataGridView1.Rows[i].Cells[27].Value.ToString().Contains("*") && dataGridView1.Rows[i].Cells[25].Value.ToString() == "مؤرشف نهائي")
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Green;

                //else dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.;
                
            }
            //
        }

        private void SearchDoc_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            SearchFile.Text = dlg.FileName;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PanelMain.Visible = false;
            dataGridView1.Visible = true;
        }



        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                gridFill = true;
                PanelMain.Visible = true;
                dataGridView1.Visible = false;
                colIDs[1] = dataGridView1.CurrentRow.Cells[0].Value.ToString();

                colIDs[0] = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                try
                {
                    txtEditID2.Text = colIDs[0].Split('/')[4];
                    txtEditID1.Text = colIDs[0].Replace(txtEditID2.Text, "");
                }
                catch (Exception ex) {
                }

                if (dataGridView1.CurrentRow.Cells[2].Value.ToString() == "")
                {
                    EditSave = true;
                    colIDs[7] = "new";
                    NextRelId = VisaAppId.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    rowIndexTodelete = ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                    //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 1);
                    FillDatafromGenArch("data1", colIDs[1], "TableVisaApp");
                    if (Jobposition.Contains("قنصل")) deleteRow.Visible = true;
                    gridFill = false;
                    return;
                }
                gridFill = false;
                colIDs[7] = "old";
                rowIndexTodelete = ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                if (Jobposition.Contains("قنصل")) deleteRow.Visible = true;
                NextRelId = VisaAppId.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                مقدم_الطلب_1.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                language.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                رقم_الهوية_1.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                مكان_الإصدار_1.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                string IssueDate = dataGridView1.CurrentRow.Cells[7].Value.ToString();

                string[] YearMonthDay = IssueDate.Split('/');

                yy1.Text = YearMonthDay[2];
                dd1.Text = YearMonthDay[1];
                mm1.Text = YearMonthDay[0];
                if (language.Text == "الانجليزية")
                {
                    language.CheckState = CheckState.Unchecked;
                    txtCountry.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                }
                else
                {
                    txtCountry.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                    language.CheckState = CheckState.Checked;
                }
                التاريخ_الميلادي.Text = "0-0-0";
                التاريخ_الميلادي.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                //MessageBox.Show();
                التاريخ_الهجري.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[13].Value.ToString().ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;

                }
                else checkedViewed.CheckState = CheckState.Checked;

                AppType.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                if (AppType.Text == "حضور مباشرة إلى القنصلية")
                    AppType.CheckState = CheckState.Checked;
                else AppType.CheckState = CheckState.Unchecked;
                ConsulateEmployee.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                if (AppType.CheckState == CheckState.Unchecked)
                {
                    mandoubVisibilty(); mandoubName.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();
                }
                TravPurpose.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
                PreRelatedID = dataGridView1.CurrentRow.Cells[19].Value.ToString();
                Comment.Text = dataGridView1.CurrentRow.Cells[24].Value.ToString();


                if (dataGridView1.CurrentRow.Cells[25].Value.ToString() != "غير مؤرشف")
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
                VisaLine = dataGridView1.CurrentRow.Cells[27].Value.ToString();
                string[] strList = VisaLine.Split('*');
                if (strList[0].Split('-')[0] == "العربية")
                {
                    language.Text = "العربية";
                    language.CheckState = CheckState.Checked;
                }
                else
                {
                    language.Text = "الانجليزية";
                    language.CheckState = CheckState.Unchecked;
                }
                TravPurpose.Text = strList[0].Split('-')[7];
                textBox3.Text = strList[0].Split('-')[5];
                مقدم_الطلب_1.Text = strList[0].Split('-')[2];
                رقم_الهوية_1.Text = strList[0].Split('-')[3];
                مكان_الإصدار_1.Text = strList[0].Split('-')[4];

                yy1.Text = textBox3.Text.Split('/')[2];
                //MessageBox.Show(strList[0].Split('-')[5].Split('/')[2]);
                dd1.Text = textBox3.Text.Split('/')[1];
                mm1.Text = textBox3.Text.Split('/')[0];
                if (VisaLine.Contains("*"))
                {
                    

                    //VisaLine = language.Text + "-" + title + "-" + ApplicantName1.Text + "-" + ApplicantIdoc1.Text + "-" + IssuedSource1.Text + "-" + dd1.Text + "/" + mm1.Text + "/" + yy1.Text + "-" + txtCountry.Text + "-" + TravPurpose.Text;
                    //الانجليزية - Miss - HUYAM SAIFELDIN IBRAHIM MOHAMEDALI - P08295472 - JEDDAH - 28 / 08 / 2021 - the Republic of Turkey - زيارة عائلية
                    if (VisaLine.Split('*').Length == 2)
                        {
                            مقدم_الطلب_1.Text = strList[0].Split('-')[2];
                            رقم_الهوية_1.Text = strList[0].Split('-')[3];
                            مكان_الإصدار_1.Text = strList[0].Split('-')[4];
                            yy1.Text = strList[0].Split('-')[5].Split('/')[2];
                            dd1.Text = strList[0].Split('-')[5].Split('/')[1];
                            mm1.Text = strList[0].Split('-')[5].Split('/')[0];

                            مقدم_الطلب_2.Text = strList[1].Split('-')[2];
                            رقم_الهوية_2.Text = strList[1].Split('-')[3];
                            مكان_الإصدار_2.Text = strList[1].Split('-')[4];
                            yy2.Text = strList[1].Split('-')[5].Split('/')[2];
                            dd2.Text = strList[1].Split('-')[5].Split('/')[1];
                            mm2.Text = strList[1].Split('-')[5].Split('/')[0];
                        }

                        if (VisaLine.Split('*').Length == 3)
                        {
                            مقدم_الطلب_1.Text = strList[0].Split('-')[2];
                            رقم_الهوية_1.Text = strList[0].Split('-')[3];
                            مكان_الإصدار_1.Text = strList[0].Split('-')[4];
                            yy1.Text = strList[0].Split('-')[5].Split('/')[2];
                            dd1.Text = strList[0].Split('-')[5].Split('/')[1];
                            mm1.Text = strList[0].Split('-')[5].Split('/')[0];

                            مقدم_الطلب_2.Text = strList[1].Split('-')[2];
                            رقم_الهوية_2.Text = strList[1].Split('-')[3];
                            مكان_الإصدار_2.Text = strList[1].Split('-')[4];
                            yy2.Text = strList[1].Split('-')[5].Split('/')[2];
                            dd2.Text = strList[1].Split('-')[5].Split('/')[1];
                            mm2.Text = strList[1].Split('-')[5].Split('/')[0];

                            مقدم_الطلب_3.Text = strList[2].Split('-')[2];
                            رقم_الهوية_3.Text = strList[2].Split('-')[3];
                            مكان_الإصدار_3.Text = strList[2].Split('-')[4];
                            yy3.Text = strList[2].Split('-')[5].Split('/')[2];
                            dd3.Text = strList[2].Split('-')[5].Split('/')[1];
                            mm3.Text = strList[2].Split('-')[5].Split('/')[0];
                        }

                        if (VisaLine.Split('*').Length == 4)
                        {
                            مقدم_الطلب_1.Text = strList[0].Split('-')[2];
                            رقم_الهوية_1.Text = strList[0].Split('-')[3];
                            مكان_الإصدار_1.Text = strList[0].Split('-')[4];
                            yy1.Text = strList[0].Split('-')[5].Split('/')[2];
                            dd1.Text = strList[0].Split('-')[5].Split('/')[1];
                            mm1.Text = strList[0].Split('-')[5].Split('/')[0];

                            مقدم_الطلب_2.Text = strList[1].Split('-')[2];
                            رقم_الهوية_2.Text = strList[1].Split('-')[3];
                            مكان_الإصدار_2.Text = strList[1].Split('-')[4];
                            yy2.Text = strList[1].Split('-')[5].Split('/')[2];
                            dd2.Text = strList[1].Split('-')[5].Split('/')[1];
                            mm2.Text = strList[1].Split('-')[5].Split('/')[0];

                            مقدم_الطلب_3.Text = strList[2].Split('-')[2];
                            رقم_الهوية_3.Text = strList[2].Split('-')[3];
                            مكان_الإصدار_3.Text = strList[2].Split('-')[4];
                            yy3.Text = strList[2].Split('-')[5].Split('/')[2];
                            dd3.Text = strList[2].Split('-')[5].Split('/')[1];
                            mm3.Text = strList[2].Split('-')[5].Split('/')[0];

                            مقدم_الطلب_4.Text = strList[3].Split('-')[2];
                            رقم_الهوية_4.Text = strList[3].Split('-')[3];
                            مكان_الإصدار_4.Text = strList[3].Split('-')[4];
                            yy4.Text = strList[3].Split('-')[5].Split('/')[2];
                            dd4.Text = strList[3].Split('-')[5].Split('/')[1];
                            mm4.Text = strList[3].Split('-')[5].Split('/')[0];
                        }

                        if (VisaLine.Split('*').Length == 5)
                        {
                            مقدم_الطلب_1.Text = strList[0].Split('-')[2];
                            رقم_الهوية_1.Text = strList[0].Split('-')[3];
                            مكان_الإصدار_1.Text = strList[0].Split('-')[4];
                            yy1.Text = strList[0].Split('-')[5].Split('/')[2];
                            dd1.Text = strList[0].Split('-')[5].Split('/')[1];
                            mm1.Text = strList[0].Split('-')[5].Split('/')[0];

                            مقدم_الطلب_2.Text = strList[1].Split('-')[2];
                            رقم_الهوية_2.Text = strList[1].Split('-')[3];
                            مكان_الإصدار_2.Text = strList[1].Split('-')[4];
                            yy2.Text = strList[1].Split('-')[5].Split('/')[2];
                            dd2.Text = strList[1].Split('-')[5].Split('/')[1];
                            mm2.Text = strList[1].Split('-')[5].Split('/')[0];

                            مقدم_الطلب_3.Text = strList[2].Split('-')[2];
                            رقم_الهوية_3.Text = strList[2].Split('-')[3];
                            مكان_الإصدار_3.Text = strList[2].Split('-')[4];
                            yy3.Text = strList[2].Split('-')[5].Split('/')[2];
                            dd3.Text = strList[2].Split('-')[5].Split('/')[1];
                            mm3.Text = strList[2].Split('-')[5].Split('/')[0];

                            مقدم_الطلب_4.Text = strList[3].Split('-')[2];
                            رقم_الهوية_4.Text = strList[3].Split('-')[3];
                            مكان_الإصدار_4.Text = strList[3].Split('-')[4];
                            yy4.Text = strList[3].Split('-')[5].Split('/')[2];
                            dd4.Text = strList[3].Split('-')[5].Split('/')[1];
                            mm4.Text = strList[3].Split('-')[5].Split('/')[0];

                            مقدم_الطلب_5.Text = strList[4].Split('-')[2];
                            رقم_الهوية_5.Text = strList[4].Split('-')[3];
                            مكان_الإصدار_5.Text = strList[4].Split('-')[4];
                            yy5.Text = strList[4].Split('-')[5].Split('/')[2];
                            dd5.Text = strList[4].Split('-')[5].Split('/')[1];
                            mm5.Text = strList[4].Split('-')[5].Split('/')[0];
                        }
                    
                }
                //else {
                //    string strList = VisaLine;
                //    if (strList.Split('-')[0] == "العربية")
                //    {
                //        language.Text = "العربية";
                //        language.CheckState = CheckState.Checked;
                //    }
                //    else
                //    {
                //        language.Text = "الانجليزية";
                //        language.CheckState = CheckState.Unchecked;
                //    }
                //    TravPurpose.Text = strList.Split('-')[7];
                //    //VisaLine = language.Text + "-" + title + "-" + ApplicantName1.Text + "-" + ApplicantIdoc1.Text + "-" + IssuedSource1.Text + "-" + dd1.Text + "/" + mm1.Text + "/" +
                //    + "-" + txtCountry.Text + "-" + TravPurpose.Text;
                //    ApplicantName1.Text = strList.Split('-')[2];

                //    ApplicantIdoc1.Text = strList.Split('-')[3];

                //    IssuedSource1.Text = strList.Split('-')[4];
                //    //MessageBox.Show(strList.Split('-')[5].Split('/')[2]);



                //    yy2.Text = strList.Split('-')[5].Split('/')[2];
                //    dd2.Text = strList.Split('-')[5].Split('/')[1];
                //    mm2.Text = strList.Split('-')[5].Split('/')[0];

                //}
                if (VisaLine != "")
                {
                    VisaIndex = ShowIndex = VisaLine.Split('*').Length;
                }
                if (VisaIndex == 1)
                {
                    checkBox1.Text = "فردي";
                }
                else
                {
                    checkBox1.Text = "متعدد";
                }
                //language.Enabled = TravPurpose.Enabled = false;
                ArchivedSt.Visible = true;
                labelArch.Visible = true;
                btnprintOnly.Visible = true;
                SaveOnly.Visible = true;
                btnSavePrint.Text = "حفظ";
                btnSavePrint.Visible = false;
                gridFill = true;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.CheckState == CheckState.Unchecked)
            {
                checkBox1.Text = "فردي";
            }
            else 
            {
                checkBox1.Text = "متعدد";
            }
        }

        private void textBox1_TextChanged_2(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 0)
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView3.DataSource;
                if (language.CheckState == CheckState.Unchecked)
                {
                    bs.Filter = dataGridView3.Columns[1].HeaderText.ToString() + " LIKE '%" + textBox1.Text + "%'";
                }else
                    bs.Filter = dataGridView3.Columns[0].HeaderText.ToString() + " LIKE '%" + textBox1.Text + "%'";
                dataGridView3.DataSource = bs;

                if (dataGridView3.Rows.Count == 2) {
                    if (language.CheckState == CheckState.Unchecked)
                    {
                        txtCountry.Text = dataGridView3.Rows[0].Cells[1].Value.ToString();
                    }
                    else
                        txtCountry.Text = dataGridView3.Rows[0].Cells[0].Value.ToString();
                }
            }
            else {
                FillDataGridAdd();
            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView3.CurrentRow.Index != -1)
            {
                if (language.CheckState == CheckState.Unchecked)
                {
                    txtCountry.Text = dataGridView3.CurrentRow.Cells[1].Value.ToString();
                }else 
                    txtCountry.Text = dataGridView3.CurrentRow.Cells[0].Value.ToString();
            }               
        }

        private void timer5_Tick(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 0) 
                dataGridView3.Visible = true;
            else 
                dataGridView3.Visible = false;

            if (colored) return;
            ColorFulGrid9();
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                //dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;

                if (dataGridView1.Rows[i].DefaultCellStyle.BackColor != Color.White)
                {
                    colored = true;
                    return;
                }


            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //Button btn = sender as Button;
            ////MessageBox.Show(btn.Name);
            ////language.Text + "-" + title + "-" + ApplicantName.Text + "-" + ApplicantIdoc.Text + "-" + IssuedSource.Text + "-" + dd1.Text + "/" + mm1.Text + "/" + yy1.Text + "-" + country + "-" + TravPurpose.Text;
            //string[] strList = VisaLine.Split('*');
            //for (int x = 0; x < VisaLine.Split('*').Length; x++)
            //{
            //    if (x == Convert.ToInt32(btn.Name))
            //    {
            //        language.Text = strList[x].Split('-')[0];
            //        title = strList[x].Split('-')[1];
            //        ApplicantName1.Text = strList[x].Split('-')[2];
            //        ApplicantIdoc1.Text = strList[x].Split('-')[3];
            //        IssuedSource1.Text = strList[x].Split('-')[4];
            //        txtCountry.Text = strList[x].Split('-')[6];
            //        string[] YearMonthDay = strList[x].Split('-')[5].Split('/');
            //        yy1.Text = YearMonthDay[2];
            //        dd1.Text = YearMonthDay[1];
            //        mm1.Text = YearMonthDay[0];
            //        TravPurpose.Text = strList[x].Split('-')[7];
            //    }
            //}
            
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        //private void dd1_TextChanged(object sender, EventArgs e)
        //{
        //    dd1.Text = "";
        //}

        //private void dd2_TextChanged(object sender, EventArgs e)
        //{
        //    dd2.Text = "";
        //}

        //private void dd3_TextChanged(object sender, EventArgs e)
        //{
        //    dd3.Text = "";
        //}

        //private void dd4_TextChanged(object sender, EventArgs e)
        //{
        //    dd4.Text = "";
        //}

        //private void dd5_TextChanged(object sender, EventArgs e)
        //{
        //    dd5.Text = "";
        //}

        //private void mm1_TextChanged(object sender, EventArgs e)
        //{
        //    mm1.Text = "";
        //}

        //private void mm2_TextChanged(object sender, EventArgs e)
        //{
        //    mm2.Text = "";
        //}

        //private void mm3_TextChanged(object sender, EventArgs e)
        //{
        //    mm3.Text = "";
        //}

        //private void mm4_TextChanged(object sender, EventArgs e)
        //{
        //    mm4.Text = "";
        //}

        //private void mm5_TextChanged(object sender, EventArgs e)
        //{
        //    mm5.Text = "";
        //}

        //private void yy1_TextChanged(object sender, EventArgs e)
        //{
        //    //yy1.Text = "";
        //}

        //private void yy2_TextChanged(object sender, EventArgs e)
        //{
        //    yy2.Text = "";
        //}

        //private void yy3_TextChanged(object sender, EventArgs e)
        //{
        //    yy3.Text = "";
        //}

        //private void yy4_TextChanged(object sender, EventArgs e)
        //{
        //    yy4.Text = "";
        //}

        //private void yy5_TextChanged(object sender, EventArgs e)
        //{
        //    yy5.Text = "";
        //}

        private void pictureBox11_Click(object sender, EventArgs e)
        {            
            
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
           
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
           
        }

        private void pictureBox13_Click(object sender, EventArgs e)
        {
            مقدم_الطلب_1.Text = "";
            رقم_الهوية_1.Text = "";
            title1.Text = "";
            sex1.Text = "";
            مكان_الإصدار_1.Text = "";
            dd1.Text = "يوم";
            mm1.Text = "شهر";
            yy1.Text = "عام";
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            مقدم_الطلب_2.Text = "";
            رقم_الهوية_2.Text = "";
            title2.Text = "";
            sex2.Text = "";
            مكان_الإصدار_2.Text = "";
            dd2.Text = "يوم";
            mm2.Text = "شهر";
            yy2.Text = "عام";
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            مقدم_الطلب_3.Text = "";
            رقم_الهوية_3.Text = "";
            title3.Text = "";
            sex3.Text = "";
            مكان_الإصدار_3.Text = "";
            dd3.Text = "يوم";
            mm3.Text = "شهر";
            yy3.Text = "عام";
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            مقدم_الطلب_4.Text = "";
            رقم_الهوية_4.Text = "";
            title4.Text = "";
            sex4.Text = "";
            مكان_الإصدار_4.Text = "";
            dd4.Text = "يوم";
            mm4.Text = "شهر";
            yy4.Text = "عام";
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            مقدم_الطلب_5.Text = "";
            رقم_الهوية_5.Text = "";
            title5.Text = "";
            sex5.Text = "";
            مكان_الإصدار_5.Text = "";
            dd5.Text = "يوم";
            mm5.Text = "شهر";
            yy5.Text = "عام";
        }

        private void Form4_FormClosed(object sender, FormClosedEventArgs e)
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
        private void SearchDoc_Click(object sender, EventArgs e)
        {
            
        }

        private void CreateWordFile(bool printOnly)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();

            string ReportName = DateTime.Now.ToString("mmss");
            string route = "";
            if (language.CheckState == CheckState.Unchecked)
            {
                if (VisaIndex > 1)
                    route = FilesPathIn + "MultiForiegnVisaM_F.docx";
                else route = FilesPathIn + "ForiegnVisaM_F.docx";
            }
            else
            {
                if (VisaIndex > 1)
                    route = FilesPathIn + "MultiArabVisaM_F.docx";
                else route = FilesPathIn + "ArabVisaM_F.docx";
            }


            string ActiveCopy;
            ActiveCopy = FilesPathOut + مقدم_الطلب_1.Text + ReportName + ".docx";
            

            System.IO.File.Copy(route, ActiveCopy);

            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();

            object Routseparameter = ActiveCopy;

            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);
            object ParaGreData = "MarkGreData";
            object ParaIqrarNo = "MarkIqrarNo";
            object ParaHijriData = "MarkHijriData";
            object ParaName = "MarkApplicantName";
            object ParaPass = "MarkPass";
            object ParaPassIssueDate = "MarkPassIssueDate";
            object ParaDestinationCountry1 = "MarkDestinationCountry1";
            object ParaDestinationCountry2 = "MarkDestinationCountry2";
            object ParaDestinationCountry3 = "MarkDestinationCountry3";
            object ParaDestinationCountry4 = "MarkDestinationCountry4";
            object ParaTitle = "MarkTitle";
            object ParaPassISource = "MarkPassISource";
            object ParaAppTitle = "MarkAppTitle";

            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookIqrarNo = oBDoc.Bookmarks.get_Item(ref ParaIqrarNo).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            if (VisaIndex <= 1)
            {
                Word.Range BookName = oBDoc.Bookmarks.get_Item(ref ParaName).Range;
                Word.Range BookPass = oBDoc.Bookmarks.get_Item(ref ParaPass).Range;
                Word.Range BookPassIssueDate = oBDoc.Bookmarks.get_Item(ref ParaPassIssueDate).Range;
                Word.Range BookPassISource = oBDoc.Bookmarks.get_Item(ref ParaPassISource).Range;
                Word.Range BookTitle = oBDoc.Bookmarks.get_Item(ref ParaTitle).Range;
                string[] strLines = VisaLine.Split('*');
                string[] str = strLines[0].Split('-');

                //VisaLine = language.Text + "-" + title + "-" + ApplicantName.Text + "-" + ApplicantIdoc.Text + "-" + IssuedSource.Text + "-" + dd1.Text + "/" + mm1.Text + "/" + yy1.Text + "-" + country + "-" + TravPurpose.Text;
                BookName.Text = colIDs[3] = str[2];
                BookPass.Text = str[3];
                BookPassISource.Text = str[4];
                BookPassIssueDate.Text = str[5];
                if (language.CheckState == CheckState.Unchecked)
                {
                    BookTitle.Text = str[1];
                }
                else
                {
                    if (sex1.CheckState == CheckState.Checked)
                        BookTitle.Text = "المواطنة السودانية";
                    else BookTitle.Text = "المواطن السوداني";
                    Word.Range BookAppTitle = oBDoc.Bookmarks.get_Item(ref ParaAppTitle).Range;
                    BookAppTitle.Text = "";


                    if (language.CheckState == CheckState.Checked)
                    {
                        if (sex1.Text == "ذكر") BookAppTitle.Text = "";
                        else BookAppTitle.Text = "ة";
                    }

                    object rangeAppTitle = BookAppTitle;
                    oBDoc.Bookmarks.Add("MarkAppTitle", ref rangeAppTitle);

                }
                //if (personalNonPersonal.CheckState == CheckState.Unchecked)
                //{
                //    BookPass.Text = Pass[x];
                //    BookPassIssueDate.Text = IssueDate[x];
                //    BookPassISource.Text = Source[x];
                //    BookName.Text = DaughterMother[x];
                //}

                object rangeName = BookName;
                object rangePass = BookPass;
                object rangePassIssueDate = BookPassIssueDate;
                object rangePassISource = BookPassISource;
                object rangeTitle = BookTitle;

                oBDoc.Bookmarks.Add("MarkTitle", ref rangeTitle);
                oBDoc.Bookmarks.Add("MarkName", ref rangeName);
                oBDoc.Bookmarks.Add("MarkPass", ref rangePass);
                oBDoc.Bookmarks.Add("MarkPassIssueDate", ref rangePassIssueDate);
                oBDoc.Bookmarks.Add("MarkPassISource", ref rangePassISource);
                Save2DataBase(!printOnly);
            }
            else
            {

                Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
                string[] strLines = VisaLine.Split('*');
                Save2DataBase(!printOnly);
                colIDs[3] = strLines[0].Split('-')[2];
                if (strLines[0].Split('-')[0] == "العربية")
                {
                    for (int x = 0; x < strLines.Length; x++)
                    {
                        string[] str = strLines[x].Split('-');
                        table.Rows.Add();
                        table.Rows[x + 2].Cells[1].Range.Text = (x + 1).ToString();
                        table.Rows[x + 2].Cells[2].Range.Text = str[2];
                        table.Rows[x + 2].Cells[3].Range.Text = str[3];
                        table.Rows[x + 2].Cells[4].Range.Text = str[5];
                        table.Rows[x + 2].Cells[5].Range.Text = str[4];
                    }
                }
                else {
                    //dd1.Text + "/" + mm1.Text + "/" + yy1.Text 
                    for (int x = 0; x < strLines.Length; x++)
                    {
                        string[] str = strLines[x].Split('-');
                        table.Rows.Add();
                        table.Rows[x + 2].Cells[5].Range.Text = (x + 1).ToString();
                        table.Rows[x + 2].Cells[4].Range.Text = str[2];
                        table.Rows[x + 2].Cells[3].Range.Text = str[3];
                        if(str[5].Split('/')[2].Length == 4) 
                            table.Rows[x + 2].Cells[2].Range.Text = str[5].Split('/')[2] + "/" + str[5].Split('/')[1] + "/" + str[5].Split('/')[0];
                        else 
                            table.Rows[x + 2].Cells[2].Range.Text = str[5].Split('/')[0] + "/" + str[5].Split('/')[1] + "/" +str[5].Split('/')[2];
                        table.Rows[x + 2].Cells[1].Range.Text = str[4];
                    }
                }
                
            }

            Word.Range BookDestinationCountry1 = oBDoc.Bookmarks.get_Item(ref ParaDestinationCountry1).Range;
            Word.Range BookDestinationCountry2 = oBDoc.Bookmarks.get_Item(ref ParaDestinationCountry2).Range;
            Word.Range BookDestinationCountry3 = oBDoc.Bookmarks.get_Item(ref ParaDestinationCountry3).Range;
            Word.Range BookDestinationCountry4;
            if (language.CheckState == CheckState.Checked)
                BookDestinationCountry4 = oBDoc.Bookmarks.get_Item(ref ParaDestinationCountry4).Range;
            else BookDestinationCountry4 = oBDoc.Bookmarks.get_Item(ref ParaDestinationCountry3).Range;


            BookIqrarNo.Text = colIDs[0] = VisaAppId.Text;
            BookGreData.Text = التاريخ_الميلادي_off.Text;
            colIDs[2] = التاريخ_الميلادي.Text;
            BookHijriData.Text = التاريخ_الهجري.Text;
            colIDs[5] = AppType.Text;
            colIDs[6] = mandoubName.Text;
            string country;
            string empty = "";
            country = txtCountry.Text; 
            if (country.StartsWith("ا"))
            {
                BookDestinationCountry1.Text = country.Remove(0, empty.Length);
                string somestring = country;
                StringBuilder sb = new StringBuilder(somestring);
                sb[0] = 'ل'; // index starts at 0!
                if (language.CheckState == CheckState.Checked)
                    BookDestinationCountry4.Text = sb.ToString();
                BookDestinationCountry1.Text = BookDestinationCountry3.Text = sb.ToString();
            }
            else
            {
                if (language.CheckState == CheckState.Checked)
                    BookDestinationCountry3.Text = BookDestinationCountry1.Text = BookDestinationCountry4.Text = "ل" + country;
                else
                    BookDestinationCountry1.Text = BookDestinationCountry3.Text = country;
            }
            BookDestinationCountry2.Text = country;

            if (language.CheckState == CheckState.Unchecked)
            {

                BookDestinationCountry3.Text = country;
            }
            else
            {

            }



            object rangeIqrarNo = BookIqrarNo;
            object rangeGreData = BookGreData;
            object rangeHijriData = BookHijriData;

            object rangeDestinationCountry1 = BookDestinationCountry1;
            object rangeDestinationCountry2 = BookDestinationCountry2;
            object rangeDestinationCountry3 = BookDestinationCountry3;
            object rangeDestinationCountry4 = BookDestinationCountry4;
            oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);

            oBDoc.Bookmarks.Add("MarkDestinationCountry1", ref rangeDestinationCountry1);
            oBDoc.Bookmarks.Add("MarkDestinationCountry2", ref rangeDestinationCountry2);
            oBDoc.Bookmarks.Add("MarkDestinationCountry3", ref rangeDestinationCountry3);
            if (language.CheckState == CheckState.Checked)
                oBDoc.Bookmarks.Add("MarkDestinationCountry4", ref rangeDestinationCountry4);

            //oBDoc.Bookmarks.Add("MarkAuthorization", ref rangeAuthorization);

            string docxouput = FilesPathOut + مقدم_الطلب_1.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilesPathOut + مقدم_الطلب_1.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(pdfouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;


            i = 0;
            Firstline = false;
            Clear_Fields();

            addarchives(colIDs);

        }

        private void btnEditID_Click(object sender, EventArgs e)
        {
            if (btnEditID.Text == "إجراء")
            {
                btnEditID.Text = "تعديل";
                SqlConnection sqlCon = new SqlConnection(DataSource);
                SqlCommand sqlCmd = new SqlCommand("update TableVisaApp SET DocID = @DocID WHERE ID = @ID", sqlCon);
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
            //MessageBox.Show(التاريخ_الميلادي_off.Text);
        }

        private void التاريخ_الهجري_TextChanged(object sender, EventArgs e)
        {
            //التاريخ_الميلادي_off.Text = التاريخ_الميلادي.Text.Split('-')[1] + " - " + التاريخ_الميلادي.Text.Split('-')[0] + " - " + التاريخ_الميلادي.Text.Split('-')[2];
        }

        private void مقدم_الطلب_1_TextChanged(object sender, EventArgs e)
        {
            getID(رقم_الهوية_1, مكان_الإصدار_1, مقدم_الطلب_1.Text);
        }

        public void getID(TextBox رقم_الهوية_1 , TextBox مكان_الإصدار_1, string name)
        {
            if (gridFill) {
                //MessageBox.Show("show"); 
                return; }
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
            
            foreach (DataRow row in dtbl.Rows)
            {
                رقم_الهوية_1.Text = row["رقم_الهوية"].ToString();
                مكان_الإصدار_1.Text = row["مكان_الإصدار"].ToString();            
            }
        }

        private void مقدم_الطلب_2_TextChanged(object sender, EventArgs e)
        {
            getID(رقم_الهوية_2, مكان_الإصدار_2, مقدم_الطلب_2.Text);
        }

        private void مقدم_الطلب_3_TextChanged(object sender, EventArgs e)
        {
            getID(رقم_الهوية_3, مكان_الإصدار_3, مقدم_الطلب_3.Text);
        }

        private void مقدم_الطلب_4_TextChanged(object sender, EventArgs e)
        {
            getID(رقم_الهوية_4, مكان_الإصدار_4, مقدم_الطلب_4.Text);
        }

        private void مقدم_الطلب_5_TextChanged(object sender, EventArgs e)
        {
            getID(رقم_الهوية_5, مكان_الإصدار_5, مقدم_الطلب_5.Text);
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

        private void Review_Click_1(object sender, EventArgs e)
        {

        }

        private void dataGridView1_DoubleClick_1(object sender, EventArgs e)
        {

        }


    }
}