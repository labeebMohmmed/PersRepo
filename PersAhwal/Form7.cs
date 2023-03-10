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
    public partial class Form7 : Form
    {
        public static string route = "";
        string Viewed;
        string ConsulateEmpName;
        public static string ModelFileroute = "";
        String IqrarNumberPart;
        static string DataSource;
        int ApplicantID = 0;
        string FilesPathIn, FilesPathOut;
        string NewFileName;
        private bool fileloaded = false;
        string PreAppId = "", PreRelatedID = "", NextRelId = "";
        static public string FamilySupport;
        private string[] FamelyMember = new string[10];
        string Jobposition;
        bool newData = true;
        bool SaveEdit = true;
        int ATVC = 0;
        string[] colIDs = new string[100];
        string GregorianDate = "";
        string HijriDate = "";
        string AuthTitle = "نائب قنصل";
        public Form7(int Atvc, int currentRow, string EmpName, string dataSource, string filepathIn, string filepathOut, string jobposition, string gregorianDate, string hijriDate)
        {
            InitializeComponent();
            //timer1.Enabled = true;
            //timer2.Enabled = true;
            التاريخ_الميلادي.Text = GregorianDate = gregorianDate;
            التاريخ_الهجري.Text = HijriDate = hijriDate;
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
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT DocID from TableTRName where ID=@ID", sqlCon);
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
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from TableTRName order by ID desc", sqlCon);
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
                query = "select Data1, Extension1,FileName1 from TableTRName  where ID=@id";
            }
            else if (fileNo == 2)
            {
                query = "select Data2, Extension2,FileName2 from TableTRName  where ID=@id";
            }
            else query = "select Data3, Extension3,FileName3 from TableTRName  where ID=@id";
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
            PreAppId = dataGridView1.Rows[Rowindex].Cells[1].Value.ToString();
            مقدم_الطلب.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "ذكر") النوع.CheckState = CheckState.Unchecked;
            else if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "أنثى") النوع.CheckState = CheckState.Checked;
            نوع_الهوية.Text = dataGridView1.Rows[Rowindex].Cells[4].Value.ToString();
            رقم_الهوية.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString();
            مكان_الإصدار.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString();
            AppTrueName.Text = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString();
            AppWrongName.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString();
            التاريخ_الميلادي.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString();
            التاريخ_الهجري.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[11].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[16].Value.ToString().ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                
            }
            else checkedViewed.CheckState = CheckState.Checked;

            AppType.Text = dataGridView1.Rows[Rowindex].Cells[13].Value.ToString();
            ConsulateEmployee.Text = dataGridView1.Rows[Rowindex].Cells[14].Value.ToString();
            if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked;
            else AppType.CheckState = CheckState.Unchecked;

            if (AppType.CheckState == CheckState.Unchecked)
            {
                mandoubVisibilty(); mandoubName.Text = dataGridView1.Rows[Rowindex].Cells[15].Value.ToString();
            }
            PreRelatedID = dataGridView1.Rows[Rowindex].Cells[16].Value.ToString();
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
            labelArch.Visible = true;
            btnSavePrint.Text = "حفظ";
            btnSavePrint.Visible = false;
        }

        //private void OpenFile(int id, int fileNo)
        //{
        //    string query;

        //    SqlConnection Con = new SqlConnection(DataSource);
        //    if (fileNo == 1)
        //    {
        //        query = "select Data1, Extension1,FileName1 from TableTRName where ID=@id";
        //    }
        //    else
        //    {
        //        query = "select Data2, Extension2,FileName2 from TableTRName where ID=@id";
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


        private void Save2DataBase()
        {
            if (!checkGender(PanelMain, "مقدم_الطلب", "النوع"))
            {
                return;
            }
            else addNewAppNameInfo(مقدم_الطلب);

            SqlConnection sqlCon = new SqlConnection(DataSource);

            string AppGender;
            if (النوع.CheckState == CheckState.Unchecked) AppGender = "ذكر"; else AppGender = "أنثى";
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                else Viewed = "غير معالج";
                SqlCommand sqlCmd = new SqlCommand("TRNameAddorEdit", sqlCon);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                if (btnSavePrint.Text == "طباعة وحفظ" && newData)
                {
                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");
                    sqlCmd.Parameters.AddWithValue("@DocID", Iqrarid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", مقدم_الطلب.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", نوع_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", رقم_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocIssueSource", مكان_الإصدار.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppTrueName", AppTrueName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppWrongName", AppWrongName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", PreAppId.Trim());
                    string filePath1 = FilesPathIn + "text1.txt";
                    string filePath2 = FilesPathIn + "text2.txt";
                    
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
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");
                    sqlCmd.Parameters.AddWithValue("@document", document.Text);
                    sqlCmd.Parameters.AddWithValue("@IqrarType", PersonDesc.Text);
                    sqlCmd.ExecuteNonQuery();
                }
                else
                {
                    sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                    sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                    sqlCmd.Parameters.AddWithValue("@DocID", Iqrarid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", مقدم_الطلب.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", نوع_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", رقم_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocIssueSource", مكان_الإصدار.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppTrueName", AppTrueName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppWrongName", AppWrongName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", PreAppId.Trim());
                    string filePath1 = FilesPathIn + "text1.txt";
                    string filePath2 = FilesPathIn + "text2.txt";
                    
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
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
                    if (fileloaded)
                        sqlCmd.Parameters.AddWithValue("@ArchivedState", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    else sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف"); 
                    sqlCmd.Parameters.AddWithValue("@document", document.Text);
                    sqlCmd.Parameters.AddWithValue("@IqrarType", PersonDesc.Text);
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

        private void FillDataGridView()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("TRNameViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ApplicantName", ListSearch.Text.Trim());
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            IqrarNumberPart = loadRerNo(loadIDNo());
            dataGridView1.Columns[0].Visible = false;
            
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 200;
            sqlCon.Close();
            NewFileName = مقدم_الطلب.Text + IqrarNumberPart + "_07";
        }
        private void timer2_Tick_1(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            التاريخ_الميلادي.Text = DateTime.Now.ToString("MM-dd-yyyy");
            timer1.Enabled = false;
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
            timer2.Enabled = false;
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

        private void ApplicantSex_CheckedChanged_1(object sender, EventArgs e)
        {
            if (النوع.CheckState == CheckState.Unchecked)
            {

                النوع.Text = "ذكر";
                labelName.Text = "الاسم الصحيح لمقدم الطلب:";
            }
            else if (النوع.CheckState == CheckState.Checked)
            {
                النوع.Text = "إنثى";
                labelName.Text = "الاسم الصحيح لمقدمة الطلب:";
            }
        }

        private void Review_Click(object sender, EventArgs e)
        {

        }

        private void Form7_Load(object sender, EventArgs e)
        {
            fileComboBox(نوع_الهوية, DataSource, "DocType", "TableListCombo");            
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

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void CreateWordFile()
        {
            string ReportName = DateTime.Now.ToString("mmss");
            if (النوع.CheckState == CheckState.Unchecked)
            {

                labelName.ForeColor = Color.Black;
                labelName.Text = "مقدم الطلب:";
                if (PersonDesc.SelectedIndex == 1)
                    route = FilesPathIn + "IqrarTwoRightdocM.docx";
                else route = FilesPathIn + "IqrarTrueNameM.docx";
            }
            else if (النوع.CheckState == CheckState.Checked)
            {
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
                if (PersonDesc.SelectedIndex == 1)
                    route = FilesPathIn + "IqrarTwoRightdocF.docx";
                else route = FilesPathIn + "IqrarTrueNameF.docx";
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

                object ParaIqrarNo = "MarkIqrarNo";
                object ParaGreData = "MarkGreData";
                object ParaHijriData = "MarkHijriData";
                object ParaAppDocName = "MarkAppDocName";
                object ParaAppDocName2 = "MarkAppName2";
                object ParaAppTrueName1 = "MarkAppTrueName1";
                object ParaAppTrueName2 = "MarkAppTrueName2";
                object ParaAppWrongName = "MarkAppWorngName";
                object ParaAppPass = "MarkAppDoc"; 
                object ParaAppDocType = "MarkAppDocType";
                object ParaAppPassSource = "MarkAppPassSource";
                object ParavConsul = "MarkViseConsul";
                object ParaAuthorization = "MarkAuthorization";


                Word.Range BookIqrarNo = oBDoc.Bookmarks.get_Item(ref ParaIqrarNo).Range;
                Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
                Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
                Word.Range BookDocName = oBDoc.Bookmarks.get_Item(ref ParaAppDocName).Range;
                Word.Range BookDocName2 = oBDoc.Bookmarks.get_Item(ref ParaAppDocName2).Range;
                Word.Range BookTrueName1 = oBDoc.Bookmarks.get_Item(ref ParaAppTrueName1).Range;
                Word.Range BookTrueName2 = oBDoc.Bookmarks.get_Item(ref ParaAppTrueName2).Range;
                if (PersonDesc.SelectedIndex != 1)
                {
                    Word.Range BookWrongName = oBDoc.Bookmarks.get_Item(ref ParaAppWrongName).Range;
                    BookWrongName.Text = AppWrongName.Text;
                    object rangeWrongPass = BookWrongName;
                    oBDoc.Bookmarks.Add("MarkWrongName", ref rangeWrongPass);
                }
                Word.Range BookAppPass = oBDoc.Bookmarks.get_Item(ref ParaAppPass).Range;
                Word.Range BookAppDocType = oBDoc.Bookmarks.get_Item(ref ParaAppDocType).Range;
                Word.Range BookAppPassSource = oBDoc.Bookmarks.get_Item(ref ParaAppPassSource).Range;
                Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;
                Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;

                BookIqrarNo.Text = colIDs[0] = Iqrarid.Text;
                BookGreData.Text = التاريخ_الميلادي_off.Text;
                colIDs[2] = التاريخ_الميلادي.Text;
                BookHijriData.Text = التاريخ_الهجري.Text;
                BookDocName.Text = BookDocName2.Text = colIDs[3] = مقدم_الطلب.Text;
                colIDs[5] = AppType.Text;
                colIDs[6] = mandoubName.Text;
                string[] docs = new string[2];
                docs = document.Text.Split('*');
                if (PersonDesc.SelectedIndex == 1)
                {
                    BookTrueName1.Text = docs[0];
                    BookTrueName2.Text = docs[1];
                }
                else
                    BookTrueName1.Text = BookTrueName2.Text = AppTrueName.Text;
                BookAppDocType.Text = نوع_الهوية.Text;
                BookAppPass.Text = رقم_الهوية.Text;
                BookAppPassSource.Text = مكان_الإصدار.Text;
                if (AppType.CheckState == CheckState.Checked)
                {
                    if (النوع.CheckState == CheckState.Unchecked) BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text+ " "+AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد حضر للقنصلية ووقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                    if (النوع.CheckState == CheckState.Checked) BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text+" "+ AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد حضرت للقنصلية ووقعت بتوقيعها على هذا الإقرار بعد تلاوته عليها وبعد أن فهمت مضمونه ومحتواه. ";
                }
                else
                {
                    if (النوع.CheckState == CheckState.Unchecked)
                        BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد وقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                    if (النوع.CheckState == CheckState.Checked)
                        BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد وقعت بتوقيعها على هذا الإقرار بعد تلاوته عليها وبعد أن فهمت مضمونه ومحتواه. ";

                }
                BookvConsul.Text = AttendViceConsul.Text + Environment.NewLine + AuthTitle;

                object rangeIqrarNo = BookIqrarNo;
                object rangeGreData = BookGreData;
                object rangeHijriData = BookHijriData;
                object rangeDocName = BookDocName;
                object rangeDocName2 = BookDocName2;
                object rangeTrueName1 = BookTrueName1;
                object rangeTrueName2 = BookTrueName2;
                object rangeAppDocType = BookAppDocType;
                object rangeAppPass = BookAppPass;
                object rangeAppPassSource = BookAppPassSource;
                object rangevConsul = BookvConsul;
                object rangeAuthorization = BookAuthorization;

                oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);
                oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
                oBDoc.Bookmarks.Add("MarkHijiData", ref rangeHijriData);
                oBDoc.Bookmarks.Add("MarkAppName", ref rangeDocName);
                oBDoc.Bookmarks.Add("MarkAppName2", ref rangeDocName2);
                oBDoc.Bookmarks.Add("MarkTrueName1", ref rangeTrueName1);
                oBDoc.Bookmarks.Add("MarkTrueName2", ref rangeTrueName2);
                oBDoc.Bookmarks.Add("MarkAppDocType", ref rangeAppDocType);
                oBDoc.Bookmarks.Add("MarkAppPass", ref rangeAppPass);
                oBDoc.Bookmarks.Add("MarkAppPassSource", ref rangeAppPassSource);
                oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);
                oBDoc.Bookmarks.Add("MarkAuthorization", ref rangeAuthorization);

                string docxouput = FilesPathOut + AppTrueName.Text + DateTime.Now.ToString("ssmm") + ".docx";
                string pdfouput = FilesPathOut + AppTrueName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
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
            }
            else
            {
                AppType.Text = "عن طريق أحد مندوبي القنصلية";
                mandoubName.Visible = true;
                mandoubLabel.Visible = true;
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

        }

        private void ApplicantIdocName_TextChanged(object sender, EventArgs e)
        {
            if (PersonDesc.SelectedIndex != 1) AppTrueName.Text = مقدم_الطلب.Text;
        }

        private void DocType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (نوع_الهوية.SelectedIndex == 0)
            {
                labeldoctype.Text = "رقم جواز السفر: ";
                رقم_الهوية.Text = "P";
            }
            else
            {
                رقم_الهوية.Text = ""; labeldoctype.Text = "رقم الإقامة:";
            }
        }

        private void btnSavePrint_Click_1(object sender, EventArgs e)
        {
            getTitle(DataSource, AttendViceConsul.Text); 
            التاريخ_الميلادي.Text = GregorianDate;
            التاريخ_الهجري.Text = HijriDate; 
            
            Save2DataBase();
            btnSavePrint.Text = "جاري المعالجة";
            CreateWordFile();
            btnSavePrint.Text = "حفظ وطباعة";
            btnSavePrint.Enabled = true;
            Clear_Fields();
        }

        

        private void SaveOnly_Click_1(object sender, EventArgs e)
        {
            Save2DataBase();
            Clear_Fields();
        }

        private void ResetAll_Click_1(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void Clear_Fields()
        {
            AppWrongName.Text = AppTrueName.Text = مقدم_الطلب.Text = مكان_الإصدار.Text = PersonDesc.Text = مقدم_الطلب.Text = "";
            النوع.CheckState = CheckState.Checked;
            labeldoctype.Text = "رقم جواز السفر: ";
            رقم_الهوية.Text = "P";
            AttendViceConsul.SelectedIndex = 2;
            نوع_الهوية.SelectedIndex = 0;
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();

            التاريخ_الميلادي.Text = DateTime.Now.ToString("dd-MM-yyyy");
            Iqrarid.Text = "ق س ج/80/" + التاريخ_الميلادي.Text.Split('-')[2].Replace("20", "") + "/07/" + loadRerNo(loadIDNo());
            mandoubName.Text = ListSearch.Text = "";
            النوع.CheckState = CheckState.Checked;
            mandoubVisibilty();
            btnSavePrint.Text = "طباعة وحفظ";
            btnSavePrint.Visible = true;
            Comment.Text = "لا تعليق";
            FillDataGridView();
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            labelArch.Visible = false;
            ArchivedSt.BackColor = Color.Red;
            PersonDesc.SelectedIndex = 0;
            SearchFile.Visible = false;
            fileloaded = false;
            
            Iqrarid.Text = "ق س ج/80/" + التاريخ_الميلادي.Text.Split('-')[2].Replace("20", "") + "/07/" + IqrarNumberPart;
            ConsulateEmployee.Text = ConsulateEmpName;
            IqrarType.SelectedIndex  = 0;
            newData = true;
        }

        private void deleteRow_Click_1(object sender, EventArgs e)
        {
            deleteRowsData(ApplicantID, "TableTRName", DataSource);
            deleteRow.Visible = false;
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


        private void PersonDesc_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (PersonDesc.SelectedIndex != 1)
            {
                labelName.Text = "الاسم الصحيح :";
                labelWrongName.Text = "الاسم الوارد بصورة غير صحيحة:";
                AppWrongName.Visible = labelWrongName.Visible = true;
                btnAddDoc.Visible = false;
                document.Visible = false;
                AppTrueName.Size = new System.Drawing.Size(266, 35);
            }
            else
            {
                AppWrongName.Visible = labelWrongName.Visible = false;
                labelWrongName.Visible = false;
                btnAddDoc.Visible = true;
                document.Visible = true;
                labelName.Text = "بيانات الوثيقة :";
                AppTrueName.Multiline = true;//size  //location 
                AppTrueName.Size = new System.Drawing.Size(266, 123);
            }

        }

        private void btnAddDoc_Click_1(object sender, EventArgs e)
        {
            document.Text = document.Text + AppTrueName.Text + " * ";
            AppTrueName.Text = "";

        }

        
        private void btnprintOnly_Click(object sender, EventArgs e)
        {
            if (!checkGender(PanelMain, "مقدم_الطلب", "النوع"))
            {
                return;
            }
            else addNewAppNameInfo(مقدم_الطلب); 
            CreateWordFile();
            Clear_Fields();
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
            sqlCommand.Parameters.AddWithValue("@col6", نوع_الهوية.Text);
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
        private void SearchFile_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_Click(object sender, EventArgs e)
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
                if (dataGridView1.Rows[i].Cells[22].Value.ToString() == "مؤرشف نهائي") dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;

                //else dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.;

            }
            //
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                gridFill = true;
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
                    SaveEdit = true;
                    newData = false;
                    colIDs[7] = "new";
                    Iqrarid.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                    //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 1);
                    FillDatafromGenArch("data1", colIDs[1], "TableTRName");
                    if (Jobposition.Contains("قنصل")) deleteRow.Visible = true;
                    gridFill = false;
                    return;
                }
                gridFill = false;
                colIDs[7] = "old";
                SaveEdit = false;
                ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                if (Jobposition.Contains("قنصل")) deleteRow.Visible = true;
                Iqrarid.Text = PreAppId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                مقدم_الطلب.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "ذكر") النوع.CheckState = CheckState.Unchecked;
                else if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "أنثى") النوع.CheckState = CheckState.Checked;
                نوع_الهوية.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                رقم_الهوية.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                مكان_الإصدار.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                AppTrueName.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                AppWrongName.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                التاريخ_الميلادي.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                التاريخ_الهجري.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[16].Value.ToString().ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;

                }
                else checkedViewed.CheckState = CheckState.Checked;

                AppType.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                ConsulateEmployee.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked;
                else AppType.CheckState = CheckState.Unchecked;

                if (AppType.CheckState == CheckState.Unchecked)
                {
                    mandoubVisibilty(); mandoubName.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                }
                PreRelatedID = dataGridView1.CurrentRow.Cells[16].Value.ToString();
                Comment.Text = dataGridView1.CurrentRow.Cells[21].Value.ToString();
                document.Text = dataGridView1.CurrentRow.Cells[23].Value.ToString();
                PersonDesc.Text = dataGridView1.CurrentRow.Cells[24].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[24].Value.ToString() != "وثائق رسمية")
                {
                    labelName.Text = "الاسم الصحيح :";
                    labelWrongName.Text = "الاسم الوارد بصورة غير صحيحة:";
                    AppWrongName.Visible = labelWrongName.Visible = true;
                    btnAddDoc.Visible = false;
                    document.Visible = false;
                    AppTrueName.Size = new System.Drawing.Size(266, 35);
                    //PersonDesc.SelectedIndex = 0;
                }
                else
                {
                    PersonDesc.SelectedIndex = 1;
                    AppWrongName.Visible = labelWrongName.Visible = false;
                    labelWrongName.Visible = false;
                    btnAddDoc.Visible = true;
                    document.Visible = true;
                    labelName.Text = "بيانات الوثيقة :";
                    AppTrueName.Multiline = true;//size  //location 
                    AppTrueName.Size = new System.Drawing.Size(266, 123);
                }

                //if (document.Text == "")
                //{
                //    labelName.Text = "الاسم الصحيح :";
                //    labelWrongName.Text = "الاسم الوارد بصورة غير صحيحة:";
                //    AppWrongName.Visible = labelWrongName.Visible = true;
                //    btnAddDoc.Visible = false;
                //    document.Visible = false;
                //}
                //else
                //{
                //    PersonDesc.SelectedIndex = 1;
                //    AppTrueName.Text = "";
                //    AppWrongName.Visible = labelWrongName.Visible = false;
                //    labelWrongName.Visible = false;
                //    btnAddDoc.Visible = true;
                //    document.Visible = true;
                //    labelName.Text = "بيانات الوثيقة :";
                //    AppTrueName.Multiline = true;//size  //location 
                //    AppTrueName.Size = new System.Drawing.Size(266, 123);
                //}

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

                ArchivedSt.Visible = true;
                labelArch.Visible = true;
                btnSavePrint.Text = "حفظ";
                btnSavePrint.Visible = false;
            }
        }

        private void Form7_FormClosed(object sender, FormClosedEventArgs e)
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
                SqlCommand sqlCmd = new SqlCommand("update TableTRName SET DocID = @DocID WHERE ID = @ID", sqlCon);
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

        private void SearchDoc_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            SearchFile.Visible = true;
            SearchFile.Text = dlg.FileName;
            if (SearchFile.Text != "") fileloaded = true;
        }

        private void التاريخ_الميلادي_TextChanged(object sender, EventArgs e)
        {
            التاريخ_الميلادي_off.Text = التاريخ_الميلادي.Text.Split('-')[1] + " - " + التاريخ_الميلادي.Text.Split('-')[0] + " - " + التاريخ_الميلادي.Text.Split('-')[2];
        }

        private void مقدم_الطلب_TextChanged(object sender, EventArgs e)
        {
            getID(رقم_الهوية, نوع_الهوية, مكان_الإصدار, النوع, مقدم_الطلب.Text);
        }
        bool gridFill = false;
        public void getID(TextBox رقم_الهوية_1, ComboBox نوع_الهوية_1, TextBox مكان_الإصدار_1, CheckBox النوع_1  , string name)
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
            نوع_الهوية_1.Text = "جواز سفر";
            مكان_الإصدار_1.Text = "";
            النوع_1.Text = "ذكر";
            foreach (DataRow row in dtbl.Rows)
            {
                رقم_الهوية_1.Text = row["رقم_الهوية"].ToString();
                نوع_الهوية_1.Text = row["نوع_الهوية"].ToString();
                مكان_الإصدار_1.Text = row["مكان_الإصدار"].ToString();
                النوع_1.Text = row["النوع"].ToString();
            }
        }

        private void AttendViceConsul_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            var selectRows = dataGridView1.SelectedRows;
            foreach (var row in selectRows)
            {
                int id = (int)((DataGridViewRow)row).Cells[0].Value;
                //OpenFile(id, 1);
                FillDatafromGenArch("data1", colIDs[1], "TableTRName");
            }
            if (ApplicantID != 0) FillDatafromGenArch("data1", colIDs[1], "TableTRName");  //OpenFile(ApplicantID, 1);
            //ApplicantID = 0;
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            var selectRows = dataGridView1.SelectedRows;
            foreach (var row in selectRows)
            {
                int id = (int)((DataGridViewRow)row).Cells[0].Value;
                //OpenFile(id, 2);
                FillDatafromGenArch("data2", colIDs[1], "TableTRName");
            }
            if (ApplicantID != 0) FillDatafromGenArch("data2", colIDs[1], "TableTRName");  //OpenFile(ApplicantID, 2);
            //ApplicantID = 0;
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
    }
}
