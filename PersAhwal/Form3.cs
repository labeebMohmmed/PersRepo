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
    public partial class Form3 : Form
    {
        public static int thirdPartyIndex = 0, x = 0;
        public static string AllFamilyList = "";
        public static string[] DaughterMother = new string[10];
        public static string[] DaughterMotherdocNo = new string[10];
        public static string[] DaughterMotherdocIssue = new string[10];

        public static string[] AllFamilyMemberList = new string[10];
        public static string[] FamMebersName = new string[10];
        public static string[] MothDaughter = new string[10];
        public static string[] DocumentType = new string[10];
        public static string[] DocumentNo = new string[10];
        public static string[] DocumentIssue = new string[10];
        public static string[] TitleFam = new string[10];

        public static string[] DaughterMotheDocType = new string[10];
        public static string[] DaughterMotheDoc = new string[10];
        public static string[] DaughterMotheDocSource = new string[10];
        static bool Firstline = false;
        int rowIndexTodelete = 0;
        int idShow = 0;
        static public string titleFam, FamilySupport;
        public static string route = "";
        string Viewed;
        string ConsulateEmpName;
        public static string ModelFileroute = "";
        
        String IqrarNumberPart;
        static string DataSource;
        bool fileloaded = false;
        int ApplicantID = 0;
        string NewFileName, CurrentIqrarId = "";
        string FilesPathIn, FilesPathOut, PreAppId = "", PreRelatedID = "", NextRelId = "";
        string Jobposition;
        bool newData = true;
        int ATVC = 0;
        static string[] colIDs = new string[100];
        bool gridFill = true;
        string GregorianDate = "";
        string HijriDate = "";
        string AuthTitle = "نائب قنصل";
        public Form3(int Atvc ,int currentRow, int IqrarType, string EmpName, string dataSource, string filepathIn, string filepathOut, string jobposition, string gregorianDate, string hijriDate)
        {
            InitializeComponent();
            colIDs[4] = ConsulateEmpName = EmpName;
            IqrarPurpose.SelectedIndex = IqrarType + 1;
            ATVC = Atvc ;
            if (IqrarType == 7)
            {
                panel1.Visible = true;
                thirdPartyIndex = 0;
            }
            التاريخ_الميلادي.Text = GregorianDate = gregorianDate;
            التاريخ_الهجري.Text = HijriDate= hijriDate;
            DataSource = dataSource;
            //timer1.Enabled = true;
            //timer2.Enabled = true;
            FilesPathIn = filepathIn + @"\";
            FilesPathOut = filepathOut;
            Jobposition = jobposition;
            FillDataGridView();
            if (currentRow == -1) Clear_Fields();
            else SetFieldswithData(currentRow);
            IqrarPurpose.SelectedIndex = IqrarType + 1;
            purposeType();

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
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT DocID from TableMultiIqrar where ID=@ID", sqlCon);
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
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from TableMultiIqrar order by ID desc", sqlCon);
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

        private void FillDataGridView()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("MultiViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ApplicantName", Search.Text.Trim());
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[2].Width = 200;
            IqrarNumberPart = loadRerNo(loadIDNo());
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 200;
            sqlCon.Close();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            fileComboBox(نوع_الهوية, DataSource, "DocType", "TableListCombo");
            fileComboBox(نوع_الهوية_2, DataSource, "DocType", "TableListCombo");
            autoCompleteTextBox(مكان_الإصدار, DataSource, "SDNIssueSource", "TableListCombo");
            autoCompleteTextBox(مكان_الإصدار_2, DataSource, "SDNIssueSource", "TableListCombo");
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            purposeType();
        }

        private void purposeType()
        {
            if (النوع.CheckState == CheckState.Unchecked && IqrarPurpose.SelectedIndex == 2)
            {
                personalNonPersonal.CheckState = CheckState.Unchecked;
                label2.Visible = panel1.Visible = true;

            }

            if (IqrarPurpose.SelectedIndex == 2 || IqrarPurpose.SelectedIndex == 3)
                label2.Visible = personalNonPersonal.Visible = true;
            else
            {
                label2.Visible = personalNonPersonal.Visible = false;
                panel1.Visible = false;
            }

            if (IqrarPurpose.Text == "إعالة أسرية")
            {
                panel1.Visible = true;
                thirdPartyIndex = 0;
                deatHusbend.Visible = true;

            }
            else deatHusbend.Visible = false;
        }

        private void SetFieldswithData(int Rowindex)
        {
            Rowindex--;
            ApplicantID = Convert.ToInt32(dataGridView1.Rows[Rowindex].Cells[0].Value.ToString());
            IqrarNo.Text = CurrentIqrarId = dataGridView1.Rows[Rowindex].Cells[1].Value.ToString();
            مقدم_الطلب.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString().ToString();
            if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "ذكر") النوع.CheckState = CheckState.Unchecked;
            else if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "أنثى") النوع.CheckState = CheckState.Checked;
            نوع_الهوية.Text = dataGridView1.Rows[Rowindex].Cells[4].Value.ToString().ToString();
            رقم_الهوية.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString().ToString();
            مكان_الإصدار.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString().ToString();
            IqrarPurpose.Text = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString().ToString();
            if (IqrarPurpose.Text == "إعالة أسرية") panel1.Visible = true; else panel1.Visible = false;
            AllFamilyMembers.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString().ToString();
            التاريخ_الميلادي.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString().ToString();
            التاريخ_الهجري.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString().ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[11].Value.ToString().ToString();
            if (dataGridView1.Rows[Rowindex].Cells[12].Value.ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
               
            }
            else checkedViewed.CheckState = CheckState.Checked;
            AppType.Text = dataGridView1.Rows[Rowindex].Cells[13].Value.ToString();
            if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked; else AppType.CheckState = CheckState.Unchecked;

            if (AppType.CheckState == CheckState.Unchecked)
            {
                mandoubVisibilty();
                mandoubName.Text = dataGridView1.Rows[Rowindex].Cells[15].Value.ToString();
            }
            ConsulateEmployee.Text = dataGridView1.Rows[Rowindex].Cells[14].Value.ToString();
            PreRelatedID = dataGridView1.Rows[Rowindex].Cells[16].Value.ToString();
            Comment.Text = dataGridView1.Rows[Rowindex].Cells[21].Value.ToString();
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
        }


        private void personalNonPersonal_CheckedChanged_1(object sender, EventArgs e)
        {
            if (personalNonPersonal.CheckState == CheckState.Unchecked)
            {
                panel1.Visible = true;
                personalNonPersonal.Text = "غير شخصي";
            }
            else
            {
                panel1.Visible = false;
                personalNonPersonal.Text = "شخصي";
            }
        }

        private void motherDaughter_CheckedChanged(object sender, EventArgs e)
        {


        }




        private void timer1_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new HijriCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            int differ = HijriDateDifferment(DataSource, true);
            string Stringdate, StrHijriDate;
            StrHijriDate = DateTime.Now.ToString("dd-MM-yyyy");
            string[] YearMonthDay = StrHijriDate.Split('-');
            int year, month, date;
            year = Convert.ToInt16(YearMonthDay[2]);
            month = Convert.ToInt16(YearMonthDay[1]);
            date = Convert.ToInt16(YearMonthDay[0]) + differ;

            if (date < 10) Stringdate = "0" + date.ToString();
            else Stringdate = date.ToString();
            التاريخ_الهجري.Text = Stringdate + "-" + month.ToString() + "-" + year.ToString();
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

        private void ApplicantSex_CheckedChanged(object sender, EventArgs e)
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

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void AddChildren_Click(object sender, EventArgs e)
        {
            string newLine = Environment.NewLine;
            titleFam = " حامل ";
            if (motherDaughter.Text == "ابني" || motherDaughter.Text == "والدي" || motherDaughter.Text == "شقيقي") titleFam = " حامل "; 
            else if (motherDaughter.Text == "ابنتي" || motherDaughter.Text == "والدتي" || motherDaughter.Text == "شقيقتي") titleFam = " حاملة ";
            AllFamilyMemberList[thirdPartyIndex] = FamilyMebersName.Text + "/" + motherDaughter.Text + "/" + نوع_الهوية_2.Text + "/" + رقم_الهوية_2.Text + "/" + مكان_الإصدار_2.Text + "/" + titleFam;
            FamMebersName[thirdPartyIndex] = FamilyMebersName.Text;
            MothDaughter[thirdPartyIndex] = motherDaughter.Text;
            DocumentType[thirdPartyIndex] = نوع_الهوية_2.Text;
            DocumentNo[thirdPartyIndex] = رقم_الهوية_2.Text;
            DocumentIssue[thirdPartyIndex] = مكان_الإصدار_2.Text;
            TitleFam[thirdPartyIndex] = titleFam; 
            //datasumFamily(AllFamilyMemberList[thirdPartyIndex], thirdPartyIndex, thirdPartyIndex + 1);
            //MessageBox.Show(thirdPartyIndex.ToString());

            if (thirdPartyIndex == 0)
            {
                AllFamilyList = FamMebersName[0] + "/" + MothDaughter[thirdPartyIndex] + "/" + DocumentType[0] + "/" + DocumentNo[0] + "/" + DocumentIssue[0] + "/" + TitleFam[0];
                if (رقم_الهوية_2.Text == "") {
                    DaughterMother[thirdPartyIndex] = MothDaughter[0] + " " + FamMebersName[0];
                    AllFamilyMembers.Text = "1- " + MothDaughter[0] + " " + FamMebersName[0];
                }
                else
                {
                    DaughterMother[thirdPartyIndex] = MothDaughter[0] + " " + FamMebersName[0] + TitleFam[0] + DocumentType[0] + " رقم " + DocumentNo[0] + " إصدار " + DocumentIssue[0];
                    AllFamilyMembers.Text = "1- " + MothDaughter[0] + " " + FamMebersName[0] + TitleFam[0] + DocumentType[0] + " رقم " + DocumentNo[0] + " إصدار " + DocumentIssue[0];
                }
                
            }

            //if (thirdPartyIndex == 0) {
            //    AllFamilyList = AllFamilyMemberList[thirdPartyIndex];
            //    if (documentNo.Text == "") AllFamilyMembers.Text = (thirdPartyIndex + 1).ToString() + "- " + DaughterMother[thirdPartyIndex];
            //    else AllFamilyMembers.Text = (thirdPartyIndex + 1).ToString() + "- " + DaughterMother[thirdPartyIndex] + titleFam + documentType.Text + " رقم " + DaughterMotherdocNo[thirdPartyIndex] + " إصدار " + DaughterMotherdocIssue[thirdPartyIndex];

            //}
            else
            {
                if (رقم_الهوية_2.Text == "")
                {
                    DaughterMother[thirdPartyIndex] = MothDaughter[thirdPartyIndex];
                AllFamilyMembers.Text = AllFamilyMembers.Text + Environment.NewLine + (thirdPartyIndex + 1).ToString() + "- " + MothDaughter[thirdPartyIndex];
                }
                else
                {
                    DaughterMother[thirdPartyIndex] = MothDaughter[thirdPartyIndex] + TitleFam[thirdPartyIndex] + DocumentType[thirdPartyIndex] + " رقم " + DocumentNo[thirdPartyIndex] + " إصدار " + DocumentIssue[thirdPartyIndex];
                AllFamilyMembers.Text = AllFamilyMembers.Text + Environment.NewLine + (thirdPartyIndex + 1).ToString() + "- " + MothDaughter[thirdPartyIndex] + TitleFam[thirdPartyIndex] + DocumentType[thirdPartyIndex] + " رقم " + DocumentNo[thirdPartyIndex] + " إصدار " + DocumentIssue[thirdPartyIndex];
                }

                AllFamilyList = AllFamilyList + "*" + AllFamilyMemberList[thirdPartyIndex];
            }

            Firstline = true;
            idShow = thirdPartyIndex;
            thirdPartyIndex++;
            FamilyMebersName.Clear();
            رقم_الهوية_2.Text = مكان_الإصدار_2.Text = "";
            نوع_الهوية_2.SelectedIndex = 0;
        }

        private void datasumFamily(string FamilyMemberList, int idlist, int totalNo)
        {
            FamilyMebersName.Text = FamMebersName[idlist];
            motherDaughter.Text = MothDaughter[idlist];
            نوع_الهوية_2.Text = DocumentType[idlist];
            رقم_الهوية_2.Text = DocumentNo[idlist];
            مكان_الإصدار_2.Text = DocumentIssue[idlist];

            //string[] memberdata = FamilyMemberList.Split('/');
            //DaughterMother[idlist] = memberdata[1] + " " + memberdata[0];
            //FamilySupport = DaughterMother[0] + memberdata[5] + memberdata[2] + " رقم " + memberdata[3] + " إصدار " + memberdata[4] + "،";

            //FamilyMebersName.Text = memberdata[0];
            //motherDaughter.Text = memberdata[1];
            //documentType.Text = memberdata[2];
            //documentNo.Text = memberdata[3];
            //documentIssue.Text = memberdata[4];
            //titleFam = memberdata[5];

            //AllFamilyMembers.Text = "";
            //for (int id = 0; id < totalNo; id++)
            //{ 
            //    if (!Firstline)
            //    {
            //        if (memberdata[3] == "") AllFamilyMembers.Text = (id +1).ToString() + "- " + DaughterMother[id];
            //        else AllFamilyMembers.Text = (id+1 ).ToString() + "- " + DaughterMother[id] + memberdata[5] + memberdata[2] + " رقم " + memberdata[3] + " إصدار " + memberdata[4];
            //        Firstline = true;
            //    }
            //    else
            //    {
            //        if (memberdata[3] == "")
            //            AllFamilyMembers.Text = AllFamilyMembers.Text + Environment.NewLine + (id +1).ToString() + "- " + DaughterMother[id];
            //        else 
            //            AllFamilyMembers.Text = AllFamilyMembers.Text + Environment.NewLine + (id +1).ToString() + "- " + DaughterMother[id] + memberdata[5] + memberdata[2] + " رقم " + memberdata[3] + " إصدار " + memberdata[4];
            //    }
            //}

        }

        private void ApplicantName_TextChanged(object sender, EventArgs e)
        {

        }

        private void CreateWordFile()
        {
            string newLine = Environment.NewLine;
            string ReportName = DateTime.Now.ToString("mmss");
            Firstline = false;
            if (النوع.CheckState == CheckState.Unchecked)
            {

                labelName.ForeColor = Color.Black;
                labelName.Text = "مقدم الطلب:";
                route = FilesPathIn + "Igrar_SocialStatusM.docx";
            }
            else if (النوع.CheckState == CheckState.Checked)
            {
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
                route = FilesPathIn + "Igrar_SocialStatusF.docx";
            }
            //
            for (x = 0; x <= thirdPartyIndex || personalNonPersonal.CheckState == CheckState.Checked; x++)
            {

                string ActiveCopy;
                if (x == 0) ActiveCopy = FilesPathOut + مقدم_الطلب.Text + ReportName + ".docx";
                else ActiveCopy = FilesPathOut + مقدم_الطلب.Text + NewFileName + x.ToString() + ".docx";

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
                    object Paraname1 = "MarkApplicantName1";
                    object Paraname2 = "MarkApplicantName2";
                    object ParaPassIqama = "MarkPassIqama";
                    object Paraigama = "MarkAppliigamaNo";
                    object ParaFamilyName = "MarkFamilyName";
                    object ParaPurpose = "MarkPurpose";
                    object ParaAppiIssSource = "MarkAppIssSource";
                    object ParaAuthorization = "MarkAuthorization";
                    object ParavConsul = "MarkViseConsul";

                    Word.Range BookIqrarNo = oBDoc.Bookmarks.get_Item(ref ParaIqrarNo).Range;
                    Word.Range Bookname1 = oBDoc.Bookmarks.get_Item(ref Paraname1).Range;
                    Word.Range Bookname2 = oBDoc.Bookmarks.get_Item(ref Paraname2).Range;
                    Word.Range Bookigama = oBDoc.Bookmarks.get_Item(ref Paraigama).Range;
                    Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;
                    Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;
                    Word.Range BookFamilyName = oBDoc.Bookmarks.get_Item(ref ParaFamilyName).Range;
                    Word.Range BookAppiIssSource = oBDoc.Bookmarks.get_Item(ref ParaAppiIssSource).Range;
                    Word.Range BookPassIqama = oBDoc.Bookmarks.get_Item(ref ParaPassIqama).Range;
                    Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
                    Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
                    Word.Range BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;

                    BookIqrarNo.Text = colIDs[0] = IqrarNo.Text;
                    Bookname1.Text = Bookname2.Text = colIDs[3] = مقدم_الطلب.Text;
                    Bookigama.Text = رقم_الهوية.Text;
                    BookvConsul.Text = AttendViceConsul.Text + Environment.NewLine + AuthTitle;
                    BookAppiIssSource.Text = مكان_الإصدار.Text;
                    BookPassIqama.Text = نوع_الهوية.Text;
                    BookGreData.Text = التاريخ_الميلادي_off.Text;
                    colIDs[2] = التاريخ_الميلادي.Text;
                    BookHijriData.Text = التاريخ_الهجري.Text;
                    if (IqrarPurpose.Text == "إعالة أسرية")
                    {
                        if (thirdPartyIndex == 1) BookPurpose.Text = "العائل الوحيد بعد الله عز وجل ل" + FamilySupport + "،";
                        else BookPurpose.Text = "العائل الوحيد بعد الله عز وجل لأسرتي المكونة من الآتي:" + Environment.NewLine + AllFamilyMembers.Text + "." + Environment.NewLine + " ولا مانع لدي من نقل كفالتهم إلى كفالتي، ";
                        if (deatHusbend.CheckState == CheckState.Checked)
                            BookPurpose.Text = BookPurpose.Text + "وذلك بعد وفاة / " + exGarder.Text + " " + husbandtxt.Text;
                        if (personalNonPersonal.CheckState == CheckState.Checked) BookFamilyName.Text = "";
                        else BookFamilyName.Text = " ";

                    }
                    else if (IqrarPurpose.Text == "إعفاء خروج جزئي")
                    {
                        BookPurpose.Text = "برغبتي في الاستفادة من حقي في إعفاء خروج جزئي من الخروج النهائي وبأن لا أطالب به مستقبلاً عند عودتي النهائية إلى السودان،";
                    }
                    else if (IqrarPurpose.Text == "خطة إسكانية")
                    {
                        BookPurpose.Text = "أقر بأنني لم أمنح قطعة أرض سكنية فى خطة إسكانية أو سكن شعبي بأي ولاية من ولايات السودان، والله على ما أقول شهيد";
                    }
                    else if (IqrarPurpose.Text == "إثبات حياة")
                    {
                        BookPurpose.Text = "أقر بأنني على قيد الحياة، والله على ما أقول شهيد";
                    }
                    else if (IqrarPurpose.Text == "بلوغ سن الرشد")
                    {
                        BookPurpose.Text = "أقر بأنني قد بلغت سن الرشد، والله على ما أقول شهيد";
                    }
                    else if (IqrarPurpose.SelectedIndex == 2)
                    {
                        if (personalNonPersonal.CheckState == CheckState.Checked)
                        {
                            if (النوع.CheckState == CheckState.Unchecked) BookPurpose.Text = "أقر بأنني متزوج، والله على ما أقول شهيد";
                            else BookPurpose.Text = "أقر بأنني متزوجة، والله على ما أقول شهيد";
                        }
                        else
                        {
                            if (titleFam == "حامل") BookPurpose.Text = " أقر بأن " + DaughterMother[x] + " متزوج، والله على ما أقول شهيد ";
                            else BookPurpose.Text = " أقر بأن " + DaughterMother[x] + " متزوجة، والله على ما أقول شهيد ";

                        }
                    }
                    else if (IqrarPurpose.SelectedIndex == 3)
                    {
                        if (personalNonPersonal.CheckState == CheckState.Checked)
                        {
                            if (النوع.CheckState == CheckState.Unchecked) BookPurpose.Text = "أقر بأنني غير متزوج، والله على ما أقول شهيد";
                            else BookPurpose.Text = "أقر بأنني غير متزوجة، والله على ما أقول شهيد";
                        }
                        else
                        {
                            if (titleFam == "حامل") BookPurpose.Text = " أقر بأن " + DaughterMother[x] + " غير متزوج، والله على ما أقول شهيد ";
                            else BookPurpose.Text = " أقر بأن " + DaughterMother[x] + " غير متزوجة، والله على ما أقول شهيد ";

                        }
                    }
                    else if (IqrarPurpose.SelectedIndex == 4)
                    {
                        if (النوع.CheckState == CheckState.Checked) BookPurpose.Text = "أقر بأنني أرملة وبأني لم اتزوج بعد وفاة زوجي، والله على ما أقول شهيد";
                        else MessageBox.Show("اختيار خاظئ لجنس مقدم الطلب");

                    }
                    else
                    {
                        BookPurpose.Text = IqrarPurpose.Text + "، والله على ما أقول شهيد";
                        if (personalNonPersonal.CheckState == CheckState.Checked) BookFamilyName.Text = " أقر بأني" + DaughterMother[x];
                        else BookFamilyName.Text = " أقر بأن " + DaughterMother[x];
                    }
                    colIDs[5] = AppType.Text;
                    colIDs[6] = mandoubName.Text;
                    if (AppType.CheckState == CheckState.Checked)
                    {
                        if (النوع.CheckState == CheckState.Unchecked) BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text +" "+ AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد حضر للقنصلية ووقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                        if (النوع.CheckState == CheckState.Checked) BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle+ "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد حضرت للقنصلية ووقعت بتوقيعها على هذا الإقرار بعد تلاوته عليها وبعد أن فهمت مضمونه ومحتواه. ";
                    }
                    else
                    {
                        if (النوع.CheckState == CheckState.Unchecked)
                            BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد وقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                        if (النوع.CheckState == CheckState.Checked)
                            BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد وقعت بتوقيعها على هذا الإقرار بعد تلاوته عليها وبعد أن فهمت مضمونه ومحتواه. ";
                    }

                    object rangeIqrarNo = BookIqrarNo;
                    object rangeName1 = Bookname1;
                    object rangeName2 = Bookname2;
                    object rangeigama = Bookigama;
                    object rangevConsul = BookvConsul;
                    object rangeAuthorization = BookAuthorization;
                    object rangeAppiIssSource = BookAppiIssSource;
                    object rangePassIqama = BookPassIqama;
                    object rangeGreData = BookGreData;
                    object rangeHijriData = BookHijriData;
                    object rangePurpose = BookPurpose;
                    object rangeFamilyName = BookFamilyName;

                    oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);
                    oBDoc.Bookmarks.Add("MarkApplicantName1", ref rangeName1);
                    oBDoc.Bookmarks.Add("MarkApplicantName2", ref rangeName2);
                    oBDoc.Bookmarks.Add("MarkAppliigamaNo", ref rangeigama);
                    oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);
                    oBDoc.Bookmarks.Add("MarkAuthorization", ref rangeAuthorization);
                    oBDoc.Bookmarks.Add("MarkFamilyName", ref rangeFamilyName);
                    oBDoc.Bookmarks.Add("MarkAppIssSource", ref rangeAppiIssSource);
                    oBDoc.Bookmarks.Add("MarkPassIqama", ref rangePassIqama);
                    oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
                    oBDoc.Bookmarks.Add("MarkHijiData", ref rangeHijriData);
                    oBDoc.Bookmarks.Add("MarkPurpose", ref rangePurpose);
                    if (AppType.Checked)
                    {
                        Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
                        table.Delete();
                    }
                    else
                    {
                        object Paraالشاهد_الأول = "الشاهد_الأول";
                        object Paraالشاهد_الثاني = "الشاهد_الثاني";
                        object Paraهوية_الأول = "هوية_الأول";
                        object Paraهوية_الثاني = "هوية_الثاني";
                        Word.Range Bookالشاهد_الأول = oBDoc.Bookmarks.get_Item(ref Paraالشاهد_الأول).Range;
                        Word.Range Bookالشاهد_الثاني = oBDoc.Bookmarks.get_Item(ref Paraالشاهد_الثاني).Range;
                        Word.Range Bookهوية_الأول = oBDoc.Bookmarks.get_Item(ref Paraهوية_الأول).Range;
                        Word.Range Bookهوية_الثاني = oBDoc.Bookmarks.get_Item(ref Paraهوية_الثاني).Range;
                        Bookالشاهد_الأول.Text = الشاهد_الأول.Text;
                        Bookالشاهد_الثاني.Text = الشاهد_الثاني.Text;
                        Bookهوية_الأول.Text = هوية_الأول.Text;
                        Bookهوية_الثاني.Text = هوية_الثاني.Text;
                        object rangeالشاهد_الأول = Bookالشاهد_الأول;
                        object rangeالشاهد_الثاني = Bookالشاهد_الثاني;
                        object rangeهوية_الأول = Bookهوية_الأول;
                        object rangeهوية_الثاني = Bookهوية_الثاني;
                        oBDoc.Bookmarks.Add("الشاهد_الأول", ref rangeالشاهد_الأول);
                        oBDoc.Bookmarks.Add("الشاهد_الثاني", ref rangeالشاهد_الثاني);
                        oBDoc.Bookmarks.Add("هوية_الأول", ref rangeهوية_الأول);
                        oBDoc.Bookmarks.Add("هوية_الثاني", ref rangeهوية_الثاني);

                    }
                    string docxouput = FilesPathOut + مقدم_الطلب.Text + DateTime.Now.ToString("ssmm") + ".docx";
                    string pdfouput = FilesPathOut + مقدم_الطلب.Text + DateTime.Now.ToString("ssmm") + ".pdf";
                    oBDoc.SaveAs2(docxouput);
                    oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
                    oBDoc.Close(false, oBMiss);
                    oBMicroWord.Quit(false, false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
                    System.Diagnostics.Process.Start(pdfouput);
                    object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;

                    if (personalNonPersonal.CheckState == CheckState.Checked) break;
                }
                else
                {
                    MessageBox.Show("يرجى حذف الملف الموجودأولاً");
                    btnSavePrint.Enabled = true;
                    thirdPartyIndex = 0;
                    break;
                }
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

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

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
                if (btnSavePrint.Text == "حفظ وطباعة" && newData)
                {


                    if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                    else Viewed = "غير معالج";
                    SqlCommand sqlCmd = new SqlCommand("MultiAddorEdit", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");
                    sqlCmd.Parameters.AddWithValue("@DocID", "ق س ج/80/" + التاريخ_الميلادي.Text.Split('-')[2].Replace("20", "") + "/03/" + loadRerNo(loadIDNo()));
                    sqlCmd.Parameters.AddWithValue("@AppName", مقدم_الطلب.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", نوع_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", رقم_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", مكان_الإصدار.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@IqrarPurpose", IqrarPurpose.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@FamilyName", AllFamilyList);
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@Husband", husbandtxt.Text);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text);
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", "");
                    string filePath1 = FilesPathIn + "text1.txt";
                    string filePath2 = FilesPathIn + "text2.txt";
                    
                    if (Search.Text != "") filePath2 = Search.Text;
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
                        Search.Clear();
                    }
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");
                    sqlCmd.ExecuteNonQuery();



                }
                else
                {
                    if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                    else Viewed = "غير معالج";

                    SqlCommand sqlCmd = new SqlCommand("MultiAddorEdit", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                    sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                    sqlCmd.Parameters.AddWithValue("@DocID", IqrarNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", مقدم_الطلب.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", نوع_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", رقم_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", مكان_الإصدار.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@IqrarPurpose", IqrarPurpose.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@FamilyName", AllFamilyList);
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@Husband", husbandtxt.Text);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text);
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", PreAppId.Trim());
                    string filePath1 = FilesPathIn + "text1.txt";
                    string filePath2 = FilesPathIn + "text2.txt";
                    
                    if (Search.Text != "") { filePath2 = Search.Text; fileloaded = true; }
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
        private void addNewAppNameInfo(TextBox textName)
        {

            string query = "insert into TableGenNames ([الاسم], رقم_الهوية,تاريخ_الميلاد,المهنة,النوع,نوع_الهوية,مكان_الإصدار) values (@col1,@col2,@col3,@col4,@col5,@col6,@col7) ;SELECT @@IDENTITY as lastid";
            string id = checkExist(textName.Text);
            if (id != "0")
            {
                query = "update TableGenNames set [الاسم] =  @col1,[رقم_الهوية] = @col2,[تاريخ_الميلاد] = @col3,[المهنة] = @col4,النوع = @col5,نوع_الهوية = @col6,مكان_الإصدار = @col7 where ID = " + id;
                //MessageBox.Show(query);
            }
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();

            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@col1", مقدم_الطلب.Text);
            sqlCommand.Parameters.AddWithValue("@col2", رقم_الهوية.Text);
            sqlCommand.Parameters.AddWithValue("@col3", تاريخ_الميلاد.Text);
            sqlCommand.Parameters.AddWithValue("@col4", المهنة.Text);
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
                mandoubLabel.Visible = panel3.Visible = false;
            }
            else
            {
                AppType.Text = "عن طريق أحد مندوبي القنصلية";
                mandoubName.Visible = panel3.Visible = true;
                mandoubLabel.Visible = true;
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

        }

        private void PassIqama_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (نوع_الهوية.Text == "اقامة ") { رقم_الهوية.Text = ""; } else رقم_الهوية.Text = "P";
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            FillDataGridView();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void printOnly_Click(object sender, EventArgs e)
        {

        }

        private void SearchDoc_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
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

        private void button5_Click_1(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void printOnly_Click_1(object sender, EventArgs e)
        {
            if (!checkGender(PanelMain, "مقدم_الطلب", "النوع"))
            {
                return;
            }
            else addNewAppNameInfo(مقدم_الطلب);
            CreateWordFile();
            this.Close();
            Clear_Fields();
        }

        private void btnSavePrint_Click_1(object sender, EventArgs e)
        {
            getTitle(DataSource, AttendViceConsul.Text); 
            التاريخ_الميلادي.Text = GregorianDate;
            التاريخ_الهجري.Text = HijriDate; 
            Save2DataBase();
            if (btnSavePrint.Text != "حفظ وطباعة") return;
            btnSavePrint.Text = "جاري المعالجة";
            btnSavePrint.Enabled = false;
            CreateWordFile();
            this.Close();
            Clear_Fields();
        }

        private void ApplicantIdoc_TextChanged(object sender, EventArgs e)
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void deatHusbend_CheckedChanged_1(object sender, EventArgs e)
        {
            if (deatHusbend.CheckState == CheckState.Checked)
            {
                husbendlabel.Visible = true;
                husbandtxt.Visible = true;
                exGarder.Visible = true;
                label15.Visible = true;
            }
            else
            {
                husbendlabel.Visible = false;
                husbandtxt.Visible = false;
                exGarder.Visible = false;
                label15.Visible = false;
            }
        }

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {
            //minus
            if (idShow > 0)
            {
                idShow--;
                datasumFamily(AllFamilyMemberList[idShow], idShow, thirdPartyIndex);
            }
        }

        private void pictureBox2_Click_1(object sender, EventArgs e)
        {
            //plus

            if (idShow < thirdPartyIndex - 1)
            {
                idShow++;
                datasumFamily(AllFamilyMemberList[idShow], idShow, thirdPartyIndex);

            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            string newLine = Environment.NewLine;

            if (motherDaughter.Text == "ابني" || motherDaughter.Text == "والدي" || motherDaughter.Text == "شقيقي") titleFam = " حامل "; else titleFam = " حاملة ";
            AllFamilyMemberList[idShow] = FamilyMebersName.Text + "/" + motherDaughter.Text + "/" + نوع_الهوية_2.Text + "/" + رقم_الهوية_2.Text + "/" + مكان_الإصدار_2.Text + "/" + titleFam;

            datasumFamily(AllFamilyMemberList[idShow], idShow, thirdPartyIndex);
            AllFamilyList = "";
            for (int i = 0; i < thirdPartyIndex; i++)
            {
                if (i == 0) AllFamilyList = AllFamilyMemberList[i];
                else AllFamilyList = AllFamilyList + "*" + AllFamilyMemberList[i];
            }
            FamilyMebersName.Clear();
            رقم_الهوية_2.Text = مكان_الإصدار_2.Text = "";
            نوع_الهوية_2.SelectedIndex = 0;
        }

        private void dataGridView1_CellClick_1(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                gridFill = true;
                dataGridView1.Visible = false;
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
                AppType.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked; else AppType.CheckState = CheckState.Unchecked;

                if (AppType.CheckState == CheckState.Unchecked)
                {
                    mandoubVisibilty(); mandoubName.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                }

                //mandoubName.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
                //MessageBox.Show(AppType.Text); 
                if (AppType.Text == "حضور مباشرة إلى القنصلية")
                {
                    AppType.CheckState = CheckState.Checked;
                    mandoubName.Text = "";
                }
                else AppType.CheckState = CheckState.Unchecked;
                if (dataGridView1.CurrentRow.Cells[2].Value.ToString() == "")
                {
                    newData = false ;
                    colIDs[7] = "new";
                    IqrarNo.Text = CurrentIqrarId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    rowIndexTodelete = ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                    //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 1);
                    FillDatafromGenArch("data1", colIDs[1], "TableMultiIqrar");
                    IqrarPurpose.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString().ToString();
                    if (Jobposition.Contains("قنصل")) deleteRow.Visible = true;
                    if (IqrarPurpose.Text == "إعالة أسرية")
                    {
                        panel1.Visible = true;
                        thirdPartyIndex = 0;
                    }
                    else panel1.Visible = false;
                    gridFill = false;
                    return;
                }
                colIDs[7] = "old";
                rowIndexTodelete = ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                if (Jobposition.Contains("قنصل")) deleteRow.Visible = true;
                IqrarNo.Text = CurrentIqrarId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                مقدم_الطلب.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString().ToString();
                gridFill = false;
                if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "ذكر") النوع.CheckState = CheckState.Unchecked;
                else if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "أنثى") النوع.CheckState = CheckState.Checked;
                نوع_الهوية.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().ToString();
                رقم_الهوية.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString().ToString();
                مكان_الإصدار.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString().ToString();
                IqrarPurpose.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString().ToString();
                if (IqrarPurpose.Text == "إعالة أسرية")
                {
                    panel1.Visible = true;
                    thirdPartyIndex = 0;
                }
                else panel1.Visible = false;
                string[] firstData = new string[10];
                AllFamilyList = dataGridView1.CurrentRow.Cells[8].Value.ToString().ToString();
                //MessageBox.Show(AllFamilyList.Split('*').Length.ToString());
                if (AllFamilyList != "")
                {
                    thirdPartyIndex = 0;
                   
                    for (; thirdPartyIndex < AllFamilyList.Split('*').Length; thirdPartyIndex++)
                    {
                        //MessageBox.Show(thirdPartyIndex.ToString()); ;
                        FamMebersName[thirdPartyIndex] = AllFamilyList.Split('*')[thirdPartyIndex].Split('/')[0];
                        MothDaughter[thirdPartyIndex] = AllFamilyList.Split('*')[thirdPartyIndex].Split('/')[1];
                        DocumentType[thirdPartyIndex] = AllFamilyList.Split('*')[thirdPartyIndex].Split('/')[2];
                        DocumentNo[thirdPartyIndex] = AllFamilyList.Split('*')[thirdPartyIndex].Split('/')[3];
                        DocumentIssue[thirdPartyIndex] = AllFamilyList.Split('*')[thirdPartyIndex].Split('/')[4];
                        TitleFam[thirdPartyIndex] = AllFamilyList.Split('*')[thirdPartyIndex].Split('/')[5];
                        //datasumFamily(AllFamilyMemberList[thirdPartyIndex], thirdPartyIndex, thirdPartyIndex + 1);
                        

                        if (thirdPartyIndex == 0)
                        {
                            FamilyMebersName.Text = FamMebersName[0];
                            motherDaughter.Text = MothDaughter[0];
                            نوع_الهوية_2.Text = DocumentType[0];
                            رقم_الهوية_2.Text = DocumentNo[0];
                            مكان_الإصدار_2.Text = DocumentIssue[0];

                            
                            AllFamilyMemberList[thirdPartyIndex] = FamMebersName[0] + "/" + MothDaughter[0] + "/" + DocumentType[0] + "/" + DocumentNo[0] + "/" + DocumentIssue[0] + "/" + TitleFam[0];
                            if (DocumentNo[thirdPartyIndex] == "")
                            {
                                AllFamilyMembers.Text = "1- " + MothDaughter[0] + " " + FamMebersName[0];
                                DaughterMother[thirdPartyIndex] = MothDaughter[0] + " " + FamMebersName[0];
                            }
                            else
                            {
                                DaughterMother[thirdPartyIndex] = MothDaughter[0] + " " + FamMebersName[0] + TitleFam[0] + DocumentType[0] + " رقم " + DocumentNo[0] + " إصدار " + DocumentIssue[0];
                            AllFamilyMembers.Text = "1- " + MothDaughter[0] + " " + FamMebersName[0] + TitleFam[0] + DocumentType[0] + " رقم " + DocumentNo[0] + " إصدار " + DocumentIssue[0];
                            }
                            
                        }

                        //if (thirdPartyIndex == 0) {
                        //    AllFamilyList = AllFamilyMemberList[thirdPartyIndex];
                        //    if (documentNo.Text == "") AllFamilyMembers.Text = (thirdPartyIndex + 1).ToString() + "- " + DaughterMother[thirdPartyIndex];
                        //    else AllFamilyMembers.Text = (thirdPartyIndex + 1).ToString() + "- " + DaughterMother[thirdPartyIndex] + titleFam + documentType.Text + " رقم " + DaughterMotherdocNo[thirdPartyIndex] + " إصدار " + DaughterMotherdocIssue[thirdPartyIndex];

                        //}
                        else
                        {
                            if (DocumentNo[thirdPartyIndex] == "")
                            {
                                DaughterMother[thirdPartyIndex] = MothDaughter[thirdPartyIndex];
                            AllFamilyMembers.Text = AllFamilyMembers.Text + Environment.NewLine + (thirdPartyIndex + 1).ToString() + "- " + MothDaughter[thirdPartyIndex];
                            }
                            else
                            {
                                DaughterMother[thirdPartyIndex] = MothDaughter[thirdPartyIndex] + TitleFam[thirdPartyIndex] + DocumentType[thirdPartyIndex] + " رقم " + DocumentNo[thirdPartyIndex] + " إصدار " + DocumentIssue[thirdPartyIndex];
                            AllFamilyMembers.Text = AllFamilyMembers.Text + Environment.NewLine + (thirdPartyIndex + 1).ToString() + "- " + MothDaughter[thirdPartyIndex] + TitleFam[thirdPartyIndex] + DocumentType[thirdPartyIndex] + " رقم " + DocumentNo[thirdPartyIndex] + " إصدار " + DocumentIssue[thirdPartyIndex];
                            }
                            
                        }

                    }
                    //firstData = AllFamilyList.Split('*');
                    ////MessageBox.Show(AllFamilyList);
                    //thirdPartyIndex = firstData.Length;
                    //for (int y = 0; y < thirdPartyIndex; y++) 
                    //    AllFamilyMemberList[y] = firstData[y];


                    //AllFamilyMembers.Text = AllFamilyList.Replace("*", Environment.NewLine + thirdPartyIndex.ToString() + " - ");
                    //AllFamilyMembers.Text = AllFamilyMembers.Text.Replace("/", " ");
                    //Firstline = false;
                    //idShow = thirdPartyIndex - 1;

                    //for (int x = 0; x < thirdPartyIndex; x++)
                    //    datasumFamily(AllFamilyMemberList[x], x, thirdPartyIndex);
                }

                if (AllFamilyList != "" && (IqrarPurpose.Text == "إثبات حالة إجتماعية (غير متزوج)" || IqrarPurpose.Text == "إثبات حالة إجتماعية (متزوج)"))
                {
                    firstData = AllFamilyList.Split('*');
                    thirdPartyIndex = firstData.Length;
                    for (int y = 0; y < thirdPartyIndex; y++) AllFamilyMemberList[y] = firstData[y];
                    AllFamilyMembers.Text = AllFamilyList.Replace("*", Environment.NewLine + thirdPartyIndex.ToString() + " - ");
                    AllFamilyMembers.Text = AllFamilyMembers.Text.Replace("/", " ");
                    Firstline = false;
                    idShow = thirdPartyIndex - 1;
                    personalNonPersonal.CheckState = CheckState.Unchecked;
                    panel1.Visible = true;
                    personalNonPersonal.Visible = label2.Visible = true;
                    for (int x = 0; x < thirdPartyIndex; x++)
                        datasumFamily(AllFamilyMemberList[x], x, thirdPartyIndex);
                }
                التاريخ_الميلادي.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString().ToString();
                التاريخ_الهجري.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString().ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString().ToString();
                if (dataGridView1.CurrentRow.Cells[12].Value.ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;
                    
                }
                else checkedViewed.CheckState = CheckState.Checked;
                
                ConsulateEmployee.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                PreRelatedID = dataGridView1.CurrentRow.Cells[16].Value.ToString();
                Comment.Text = dataGridView1.CurrentRow.Cells[21].Value.ToString();
                husbandtxt.Text = dataGridView1.CurrentRow.Cells[22].Value.ToString();
                if (husbandtxt.Text != "")
                {
                    husbandtxt.Visible = true;
                    husbendlabel.Visible = true;
                }
                else
                {
                    husbandtxt.Visible = false;
                    husbendlabel.Visible = false;
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
                idShow = 0;
                ArchivedSt.Visible = true;
                AllFamilyMembers.Enabled = true;
                gridFill = false;
            }
        }

     
       
        private void SaveOnly_Click_1(object sender, EventArgs e)
        {
            Save2DataBase();
            Clear_Fields();
        }

        private void Search_TextChanged(object sender, EventArgs e)
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
                if (dataGridView1.Rows[i].Cells[22].Value.ToString() == "مؤرشف نهائي") dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;

                //else dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.;

            }
            //
        }

        private void deleteRow_Click(object sender, EventArgs e)
        {
            deleteRowsData(rowIndexTodelete, "TableMultiIqrar", DataSource);
            deleteRow.Visible = false;
        }

        private void btnListView_Click(object sender, EventArgs e)
        {
            Clear_Fields();
            FillDataGridView();
            dataGridView1.Visible = true;
            PanelMain.Visible = false;

        }

        private void btnFile3_Click(object sender, EventArgs e)
        {
           // OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 3);
        }

        private void btnFile2_Click(object sender, EventArgs e)
        {
            //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 2);
            FillDatafromGenArch("data1", colIDs[1], "TableMultiIqrar");
        }

        private void btnFile1_Click(object sender, EventArgs e)
        {
            //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 1);
            FillDatafromGenArch("data2", colIDs[1], "TableMultiIqrar");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void SearchDoc_Click_1(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click_1(object sender, EventArgs e)
        {
            
            
        }

        private void Form3_FormClosed(object sender, FormClosedEventArgs e)
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
                SqlCommand sqlCmd = new SqlCommand("update TableMultiIqrar SET DocID = @DocID WHERE ID = @ID", sqlCon);
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

        private void التاريخ_الميلادي_TextChanged(object sender, EventArgs e)
        {
            التاريخ_الميلادي_off.Text = التاريخ_الميلادي.Text.Split('-')[1] + " - " + التاريخ_الميلادي.Text.Split('-')[0] + " - " + التاريخ_الميلادي.Text.Split('-')[2];
        }

        private void ApplicantName_TextChanged_1(object sender, EventArgs e)
        {
            //writeIDChanged(PanelMain, ApplicantName.Text, "1");
            
            getID(رقم_الهوية, نوع_الهوية, مكان_الإصدار, النوع, تاريخ_الميلاد, المهنة, مقدم_الطلب.Text);
        }

        public void getID(TextBox رقم_الهوية_1, ComboBox نوع_الهوية_1, TextBox مكان_الإصدار_1, CheckBox النوع_1, TextBox تاريخ_الميلاد_1, TextBox المهنة_1, string name)
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
            المهنة_1.Text = "";
            تاريخ_الميلاد_1.Text = "";
            النوع_1.Text = "ذكر";
            foreach (DataRow row in dtbl.Rows)
            {
                رقم_الهوية_1.Text = row["رقم_الهوية"].ToString();
                نوع_الهوية_1.Text = row["نوع_الهوية"].ToString();
                مكان_الإصدار_1.Text = row["مكان_الإصدار"].ToString();
                المهنة_1.Text = row["المهنة"].ToString();
                تاريخ_الميلاد_1.Text = row["تاريخ_الميلاد"].ToString();
                النوع_1.Text = row["النوع"].ToString();                
            }            
        }

        private void writeIDChanged(Panel panel, string name, string index)
        {
            if (gridFill) return;
            //MessageBox.Show(name);
            foreach (Control control2 in panel.Controls)
            {
                if (control2.Name == "رقم_الهوية_" + index + ".")
                    getID((TextBox)control2, name.Trim(), "رقم_الهوية", "P0");
                if (control2.Name == "تاريخ_الميلاد_" + index + ".")
                    getID((TextBox)control2, name.Trim(), "تاريخ_الميلاد", "");
                if (control2.Name == "المهنة_" + index + ".")
                    getID((TextBox)control2, name.Trim(), "المهنة", "");
                if (control2.Name == "نوع_الهوية_" + index + ".")
                    getID((ComboBox)control2, name.Trim(), "نوع_الهوية", "جواز سفر");
                if (control2.Name == "مكان_الإصدار_" + index + ".")
                    getID((TextBox)control2, name.Trim(), "مكان_الإصدار", "");
                if (control2.Name == "النوع_" + index + ".")
                    getID((CheckBox)control2, name.Trim(), "النوع", "ذكر");
            }
        }
        public void getID(ComboBox textTo, string name, string controlType, string def)
        {
            
            string query = "SELECT " + controlType + " FROM TableGenNames where الاسم like N'" + name + "%'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            int index = 0;
            textTo.Text = "";
            foreach (DataRow row in dtbl.Rows)
            {
                if (index == 0)
                    textTo.Text = row[controlType].ToString();
                else if (!textTo.Text.Contains(row[controlType].ToString()))
                    textTo.Text = textTo.Text + "_" + row[controlType].ToString();
                index++;
            }
            int AllIndex = textTo.Text.Split('_').Length;
            textTo.Text = textTo.Text.Split('_')[AllIndex - 1];
            if (index == 0)
                textTo.Text = def;
        }public void getID(CheckBox textTo, string name, string controlType, string def)
        {
            if (gridFill) return;
            string query = "SELECT " + controlType + " FROM TableGenNames where الاسم like N'" + name + "%'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            int index = 0;
            textTo.Text = "";
            foreach (DataRow row in dtbl.Rows)
            {
                if (index == 0)
                    textTo.Text = row[controlType].ToString();
                else if (!textTo.Text.Contains(row[controlType].ToString()))
                    textTo.Text = textTo.Text + "_" + row[controlType].ToString();
                index++;
            }
            int AllIndex = textTo.Text.Split('_').Length;
            textTo.Text = textTo.Text.Split('_')[AllIndex - 1];
            if (index == 0)
                textTo.Text = def;
        }public void getID(TextBox textTo, string name, string controlType, string def)
        {
            if (gridFill) return;
            string query = "SELECT " + controlType + " FROM TableGenNames where الاسم like N'" + name + "%'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            int index = 0;
            textTo.Text = "";
            foreach (DataRow row in dtbl.Rows)
            {
                if (index == 0)
                    textTo.Text = row[controlType].ToString();
                else if (!textTo.Text.Contains(row[controlType].ToString()))
                    textTo.Text = textTo.Text + "_" + row[controlType].ToString();
                index++;
            }
            int AllIndex = textTo.Text.Split('_').Length;
            textTo.Text = textTo.Text.Split('_')[AllIndex - 1];
            if (index == 0)
                textTo.Text = def;
        }

        private void mandoubName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void mandoubName_TextChanged(object sender, EventArgs e)
        {
            الشاهد_الأول.Text = mandoubName.Text.Split('-')[0].Trim();
            هوية_الأول.Text = getMandoubPass(DataSource, mandoubName.Text.Split('-')[0].Trim());
        }
        private string getMandoubPass(string source, string empName)
        {
            string pass = "";
            string query = "select رقم_الجواز from TableMandoudList where MandoubNames = N'" + empName + "'";
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
                pass = dataRow["رقم_الجواز"].ToString();
            }
            return pass;
        }

        private void button4_Click_1(object sender, EventArgs e)
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

        private void OpenFile(int id, int fileNo)
        {
            string query;

            SqlConnection Con = new SqlConnection(DataSource);
            if (fileNo == 1)
            {
                query = "select Data1, Extension1,FileName1 from TableMultiIqrar where ID=@id";
            }
            else
            {
                query = "select Data2, Extension2,FileName2 from TableMultiIqrar where ID=@id";
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


        private void ResetAll_Click(object sender, EventArgs e)
        {

        }

        private void Clear_Fields()
        {
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();

            التاريخ_الميلادي.Text = DateTime.Now.ToString("dd-MM-yyyy");
            مقدم_الطلب.Text = مكان_الإصدار.Text = رقم_الهوية.Text = نوع_الهوية.Text = "";
            AttendViceConsul.SelectedIndex = 2;
            نوع_الهوية.SelectedIndex = 0;
            رقم_الهوية.Text = "P";
            نوع_الهوية.Text = "جواز سفر";
            mandoubName.Text = Search.Text = "";
            AppType.CheckState = CheckState.Checked;
            mandoubVisibilty();
            btnSavePrint.Text = "حفظ وطباعة";
            Comment.Clear();
            panel1.Visible = false;
            FillDataGridView();
            AttendViceConsul.SelectedIndex = 2;
            thirdPartyIndex = 0;
            NewFileName = IqrarNumberPart + "_03";
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            ArchivedSt.BackColor = Color.Red;
            IqrarPurpose.SelectedIndex = 0;
            motherDaughter.SelectedIndex = 0;
            newData = true;
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (نوع_الهوية_2.Text == "إقامة") labelIqama.Text = "رقم الاقامة:";
            else labelIqama.Text = "رقم جواز السفر:";
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            التاريخ_الميلادي.Text = DateTime.Now.ToString("MM-dd-yyyy");
            timer2.Enabled = false;
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
        //private void OpenFileDoc(int id, int fileNo)
        //{
        //    string query;

        //    SqlConnection Con = new SqlConnection(DataSource);
        //    if (fileNo == 1)
        //    {
        //        query = "select Data1, Extension1,FileName1 from TableMultiIqrar  where ID=@id";
        //    }
        //    else if (fileNo == 2)
        //    {
        //        query = "select Data2, Extension2,FileName2 from TableMultiIqrar  where ID=@id";
        //    }
        //    else query = "select Data3, Extension3,FileName3 from TableMultiIqrar  where ID=@id";
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

    }
}
