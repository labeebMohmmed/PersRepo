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
        public static string[] titleEngFamily = new string[10];

        public static string[] Pass = new string[10];
        public static string[] IssueDate = new string[10];
        public static string[] Source = new string[10];
        string Viewed;
        string ConsulateEmpName;
        public static string ModelFileroute = "";
        String IqrarStaticPartAr = "ق س ج/160/04/";
        String IqrarStaticPartEn = "CGSJ/160/04/";
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
        string  PreRelatedID = "", NextRelId = "";
        string FilesPathIn, FilesPathOut;
        public Form4(int currentRow, string EmpName, string dataSource, string filepathIn, string filepathOut)
        {
            InitializeComponent();
            timer1.Enabled = true;
            timer2.Enabled = true;
            DataSource = dataSource;
            FilesPathIn = filepathIn;
            FilesPathOut = filepathOut;
            ConsulateEmpName = EmpName;
            FillDataGridView();
            if (currentRow == -1) Clear_Fields();
            else SetFieldswithData(currentRow);

        }

        private void SetFieldswithData(int Rowindex)
        {
            Rowindex--;
            ApplicantID = Convert.ToInt32(dataGridView1.Rows[Rowindex].Cells[0].Value.ToString());
            NextRelId = dataGridView1.Rows[Rowindex].Cells[1].Value.ToString();
            ApplicantName.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString();
            language.Text = dataGridView1.Rows[Rowindex].Cells[3].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[4].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
            else if (dataGridView1.Rows[Rowindex].Cells[4].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
            ApplicantIdoc.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString();
            IssuedSource.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString();
            string IssueDate = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString();
            string[] YearMonthDay = IssueDate.Split('/');
            int year, month, date;
            year = Convert.ToInt16(YearMonthDay[2]);
            month = Convert.ToInt16(YearMonthDay[0]);
            date = Convert.ToInt16(YearMonthDay[1]);

            PassIssueDate.Value = new DateTime(year, month, date);
            if (language.Text == "الانجليزية")
            {
                language.CheckState = CheckState.Unchecked;
                countryNonArab.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString();
            }
            else
            {
                countryArab.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString();
                language.CheckState = CheckState.Checked;
            }
            GregorianDate.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString();
            HijriDate.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString();
            AllFamilyMembers.Text = dataGridView1.Rows[Rowindex].Cells[11].Value.ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[12].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[13].Value.ToString().ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                VisaAppId.Text = NextRelId;
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
            personalNonPersonal.Text = dataGridView1.Rows[Rowindex].Cells[18].Value.ToString();
            if (personalNonPersonal.Text == "شخصي") personalNonPersonal.CheckState = CheckState.Checked;
            else personalNonPersonal.CheckState = CheckState.Unchecked;
            personalNonPersonalChecking();
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
            IqrarNumberPart = (dtbl.Rows.Count + 1).ToString();
            sqlCon.Close();
            NewFileName = IqrarNumberPart + "_04";
        }

        private void Save2DataBase(string FamilyMember)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);

            string AppGender;
            if (ApplicantSex.CheckState == CheckState.Unchecked) AppGender = "ذكر"; else AppGender = "أنثى";
            try
            {
                SqlCommand sqlCmd = new SqlCommand("VisaAddorEdit", sqlCon);
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                else Viewed = "غير معالج";
                sqlCmd.CommandType = CommandType.StoredProcedure;
                if (btnSavePrint.Visible == true)
                {
                    
                    
                    
                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");                
                    sqlCmd.Parameters.AddWithValue("@DocID", VisaAppId.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@lang", language.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());                    
                    sqlCmd.Parameters.AddWithValue("@DocNo", ApplicantIdoc.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", IssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocIssueDate", PassIssueDate.Text.Trim());
                    if(language.CheckState == CheckState.Checked)
                    sqlCmd.Parameters.AddWithValue("@CountryDest", countryArab.Text.Trim());
                    else sqlCmd.Parameters.AddWithValue("@CountryDest", countryNonArab.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AllFamilyMembers", FamilyMember.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@TravelPurpose", TravPurpose.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@PersonalNonPersonal", personalNonPersonal.Text);
                    sqlCmd.Parameters.AddWithValue("@RelatedVisaApp", PreRelatedID);
                    string filePath1 = FilesPathIn + "text1.txt";
                    string filePath2 = FilesPathIn + "text2.txt";
                    using (Stream stream = File.OpenRead(filePath1))
                    {
                        byte[] buffer1 = new byte[stream.Length];
                        stream.Read(buffer1, 0, buffer1.Length);
                        var fileinfo1 = new FileInfo(filePath1);
                        string extn1 = fileinfo1.Extension;
                        string DocName1 = fileinfo1.Name;
                        sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                        sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
                        sqlCmd.Parameters.Add("@FileName1", SqlDbType.NVarChar).Value = DocName1;
                    }                    
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
                    sqlCmd.ExecuteNonQuery();


                }
                else 
                {                    
                    if (checkedViewed.CheckState == CheckState.Checked)
                        sqlCmd.Parameters.AddWithValue("@ID", 0);
                    else
                        sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                    sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                    sqlCmd.Parameters.AddWithValue("@DocID", VisaAppId.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@lang", language.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", ApplicantIdoc.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", IssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocIssueDate", PassIssueDate.Text.Trim());
                    if (language.CheckState == CheckState.Checked)
                        sqlCmd.Parameters.AddWithValue("@CountryDest", countryArab.Text.Trim());
                    else sqlCmd.Parameters.AddWithValue("@CountryDest", countryNonArab.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AllFamilyMembers", FamilyMember.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@TravelPurpose", TravPurpose.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@PersonalNonPersonal", personalNonPersonal.Text);
                    sqlCmd.Parameters.AddWithValue("@RelatedVisaApp", PreVisaAppId);
                    string filePath1 = FilesPathIn + "text1.txt";
                    string filePath2 = FilesPathIn + "text2.txt";
                    using (Stream stream = File.OpenRead(filePath1))
                    {
                        byte[] buffer1 = new byte[stream.Length];
                        stream.Read(buffer1, 0, buffer1.Length);
                        var fileinfo1 = new FileInfo(filePath1);
                        string extn1 = fileinfo1.Extension;
                        string DocName1 = fileinfo1.Name;
                        sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                        sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
                        sqlCmd.Parameters.Add("@FileName1", SqlDbType.NVarChar).Value = DocName1;
                    }
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
                        ArchivedSt.CheckState = CheckState.Checked;
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


        private void language_CheckedChanged(object sender, EventArgs e)
        {
            Check_lang();
        }

        private void Check_lang()
        {
            if (language.CheckState == CheckState.Unchecked)
            {
                VisaAppId.Text = IqrarStaticPartEn + IqrarNumberPart;
                language.Text = "الانجليزية";
                countryNonArab.Visible = true;
                countryArab.Visible = false;
                titleEng.Visible = true;
                ApplicantSex.Visible = false;
                titleFamily.Visible = true;
                labeltitleFamily.Visible = true;
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("en-US");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            }
            else
            {
                VisaAppId.Text = IqrarStaticPartAr + IqrarNumberPart;
                language.Text = "العربية";
                countryNonArab.Visible = false;
                countryArab.Visible = true;
                titleEng.Visible = false;
                ApplicantSex.Visible = true;
                titleFamily.Visible = false;
                labeltitleFamily.Visible = false;
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            }
        }

        private void Review_Click(object sender, EventArgs e)
        {
            

        }

        private void personalNonPersonal_CheckedChanged(object sender, EventArgs e)
        {
            personalNonPersonalChecking();
        }

        private void personalNonPersonalChecking()
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

        private void AddChildren_Click(object sender, EventArgs e)
        {
            string newLine = Environment.NewLine;
            if (language.CheckState == CheckState.Unchecked)
            {
                titleEngFamily[i] = titleFamily .Text;
            }
            else {
                titleEngFamily[i] = motherDaughter.Text;
            }
            DaughterMother[i] = FamilyMebersName.Text;

            Pass[i] = FamilyPass.Text;
            IssueDate[i] = dateTimePicker1.Text;
            Source[i] = FamiyIssue.Text;
            if (motherDaughter.Text == "ابني") titleFam = " حامل "; 
            else titleFam = " حاملة ";

            FamelyMember[i] = DaughterMother[i] +" "+titleFam + " جواز سفر رقم " + Pass[i] + " إصدار " + Source[i] + " بتاريخ " + IssueDate[i];
            if (!Firstline) AllFamilyMembers.Text = FamelyMember[i];
            else AllFamilyMembers.Text = AllFamilyMembers.Text + newLine+ FamelyMember[i];
            Firstline = true;
            i++;
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
            HijriDate.Text = Stringdate + "-" + Stringmonth + "-" + year.ToString();
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
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            GregorianDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
        }

        

        private void ApplicantSex_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void ApplicantSex_CheckedChanged_1(object sender, EventArgs e)
        {
            if (ApplicantSex.CheckState == CheckState.Unchecked || titleEng.Text == "Mr")
            {

                ApplicantSex.Text = "ذكر";
                labelName.Text = "مقدم الطلب:";
            }
            else if (ApplicantSex.CheckState == CheckState.Checked)
            {
                ApplicantSex.Text = "إنثى";
                labelName.Text = "مقدمة الطلب:";
            }
        }

        private void titleEng_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (titleEng.Text == "Mr")
            {

                
                labelName.Text = "اسم طالب التاشيرة:";
            }
            else 
            {
                
                labelName.Text = "اسم طالبة التاشيرة::";
            }
        }

        private void motherDaughter_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form4_Load(object sender, EventArgs e)
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

        private void search_TextChanged(object sender, EventArgs e)
        {
            FillDataGridView();
        }

        private void AppType_CheckedChanged(object sender, EventArgs e)
        {
            mandoubVisibilty();
        }

        private void ResetAll_Click(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void Clear_Fields()
        {
            ApplicantName.Text = AllFamilyMembers.Text = FamiyIssue.Text = FamilyMebersName.Text = IssuedSource.Text = "";
            AttendViceConsul.SelectedIndex = 2;
            ApplicantSex.CheckState = CheckState.Unchecked;
            ApplicantIdoc.Text = FamilyPass.Text = "P";
            dateTimePicker1.Text = PassIssueDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            titleFamily.Visible = false;
            labeltitleFamily.Visible = false;
            language.Text = "العربية";
            titleEng.SelectedIndex = 0;
            titleFamily.SelectedIndex = 0;
            TravPurpose.SelectedIndex = 0;
            panel1.Visible = false;
            ApplicantIdoc.Text = "P";
            personalNonPersonal.CheckState = CheckState.Checked;
            countryArab.SelectedIndex = 0;
            language.CheckState = CheckState.Checked;
            countryNonArab.SelectedIndex = 0;
            VisaAppId.Text = IqrarStaticPartAr + IqrarNumberPart;
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
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            ConsulateEmployee.Text = ConsulateEmpName;
            i = 0;
        }

        private void printOnly_Click(object sender, EventArgs e)
        {
            
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

        private void OpenFile(int id, int fileNo)
        {
            string query;

            SqlConnection Con = new SqlConnection(DataSource);
            if (fileNo == 1)
            {
                query = "select Data1, Extension1,FileName1 from TableVisaApp where ID=@id";
            }
            else
            {
                query = "select Data2, Extension2,FileName2 from TableVisaApp where ID=@id";
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

        private void btnSavePrint_Click(object sender, EventArgs e)
        {
            btnSavePrint.Enabled = false;
            btnSavePrint.Text = "جاري المعالجة";
            CreateWordFile(false);
            btnSavePrint.Text = "حفظ وطباعة";
            btnSavePrint.Enabled = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            btnprintOnly.Enabled = false;
            btnprintOnly.Text = "طباعة";
            CreateWordFile(true);
            btnSavePrint.Text = "حفظ وطباعة";
            btnSavePrint.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Save2DataBase(FamilyMebersName.Text);
        }

        private void PassIssueDate_ValueChanged(object sender, EventArgs e)
        {

        }

        private void SearchDoc_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            SearchFile.Text = dlg.FileName;
        }

        private void CreateWordFile(bool printOnly)
        {
            string route = "";
            if (language.CheckState == CheckState.Unchecked)
            {

                route = FilesPathIn + "ForiegnVisaM_F.docx";
            }
            else
            {

                route = FilesPathIn + "ArabVisaM_F.docx";
            }

            for (int x = 0; x < i || personalNonPersonal.CheckState == CheckState.Checked; x++)
            {

                
                string ActiveCopy;
                ActiveCopy = FilesPathOut + ApplicantName.Text + NewFileName + ".docx";
                if (!printOnly)
                {
                    if (personalNonPersonal.CheckState == CheckState.Unchecked) Save2DataBase(FamelyMember[x]);
                    else Save2DataBase("");
                }
                if (!File.Exists(ActiveCopy))
                {
                
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
                
                Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
                Word.Range BookIqrarNo = oBDoc.Bookmarks.get_Item(ref ParaIqrarNo).Range;
                Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
                Word.Range BookName = oBDoc.Bookmarks.get_Item(ref ParaName).Range;
                Word.Range BookPass = oBDoc.Bookmarks.get_Item(ref ParaPass).Range;
                Word.Range BookPassIssueDate = oBDoc.Bookmarks.get_Item(ref ParaPassIssueDate).Range;
                Word.Range BookDestinationCountry1 = oBDoc.Bookmarks.get_Item(ref ParaDestinationCountry1).Range;
                Word.Range BookDestinationCountry2 = oBDoc.Bookmarks.get_Item(ref ParaDestinationCountry2).Range;
                Word.Range BookDestinationCountry3 = oBDoc.Bookmarks.get_Item(ref ParaDestinationCountry3).Range;
                Word.Range BookDestinationCountry4;
                    if (language.CheckState == CheckState.Checked)
                    BookDestinationCountry4 = oBDoc.Bookmarks.get_Item(ref ParaDestinationCountry4).Range;
                    else BookDestinationCountry4 = oBDoc.Bookmarks.get_Item(ref ParaDestinationCountry3).Range;
                Word.Range BookTitle = oBDoc.Bookmarks.get_Item(ref ParaTitle).Range;
                Word.Range BookPassISource = oBDoc.Bookmarks.get_Item(ref ParaPassISource).Range;
                
                BookIqrarNo.Text = VisaAppId.Text;
                BookGreData.Text = GregorianDate.Text;
                BookHijriData.Text = HijriDate.Text;
                BookName.Text = ApplicantName.Text;
                BookPass.Text = ApplicantIdoc.Text;
                BookPassISource.Text = IssuedSource.Text;
                BookPassIssueDate.Text = PassIssueDate.Text;
                string country;
                if (language.CheckState == CheckState.Checked) country = countryArab.Text; else country = countryNonArab.Text;
                BookDestinationCountry1.Text = country;
                BookDestinationCountry2.Text = country;
                BookDestinationCountry3.Text = country;
                if (language.CheckState == CheckState.Checked)
                    BookDestinationCountry4.Text = country;
                if (language.CheckState == CheckState.Unchecked)
                {
                    BookTitle.Text = titleEng.Text;
                }
                else
                {
                    if (ApplicantSex.CheckState == CheckState.Checked)
                        BookTitle.Text = "المواطنة السودانية";
                    else BookTitle.Text = "المواطن السوداني";
                }

                if (personalNonPersonal.CheckState == CheckState.Unchecked)
                {
                    BookPass.Text = Pass[x];
                    BookPassIssueDate.Text = IssueDate[x];
                    BookPassISource.Text = Source[x];

                    BookName.Text = DaughterMother[x];
                    if (language.CheckState == CheckState.Unchecked)
                    {
                        BookTitle.Text = titleEngFamily[x];
                    }
                    else
                    {
                        if (motherDaughter.Text == "ابنتي")
                            BookTitle.Text = "المواطنة السودانية";
                        else BookTitle.Text = "المواطن السوداني";
                    }
                }
                
                object rangeIqrarNo = BookIqrarNo;
                object rangeGreData = BookGreData;
                object rangeHijriData = BookHijriData;
                object rangeName = BookName;
                object rangePass = BookPass;
                object rangePassIssueDate = BookPassIssueDate;
                object rangeDestinationCountry1 = BookDestinationCountry1;
                object rangeDestinationCountry2 = BookDestinationCountry2;
                object rangeDestinationCountry3 = BookDestinationCountry3;
                object rangeDestinationCountry4 = BookDestinationCountry4;
                object rangeTitle = BookTitle;
                object rangePassISource = BookPassISource;
               
                oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo); 
                oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
                oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);
                oBDoc.Bookmarks.Add("MarkName", ref rangeGreData);
                oBDoc.Bookmarks.Add("MarkPass", ref rangePass);
                oBDoc.Bookmarks.Add("MarkPassIssueDate", ref rangePassIssueDate);
                oBDoc.Bookmarks.Add("MarkDestinationCountry1", ref rangeDestinationCountry1);
                oBDoc.Bookmarks.Add("MarkDestinationCountry2", ref rangeDestinationCountry2);
                oBDoc.Bookmarks.Add("MarkDestinationCountry3", ref rangeDestinationCountry3);
                if (language.CheckState == CheckState.Checked)
                    oBDoc.Bookmarks.Add("MarkDestinationCountry4", ref rangeDestinationCountry4);
                oBDoc.Bookmarks.Add("MarkTitle", ref rangeTitle);
                oBDoc.Bookmarks.Add("MarkPassISource", ref rangePassISource);
                //oBDoc.Bookmarks.Add("MarkAuthorization", ref rangeAuthorization);
                oBDoc.Activate();

                oBDoc.Save();
                oBMicroWord.Visible = true;
                if (personalNonPersonal.CheckState == CheckState.Checked) break;
                    }
                    else
                    {
                        MessageBox.Show("يرجى حذف الملف الموجودأولاً");
                    btnprintOnly.Enabled = true;
                        btnSavePrint.Enabled = true;
                        i = 0;
                    }
                }
            i = 0;
            Firstline = false;
            
        
    }

        private void Review_Click_1(object sender, EventArgs e)
        {
            
        }

        private void dataGridView1_DoubleClick_1(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                NextRelId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                ApplicantName.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                language.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[4].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
                else if (dataGridView1.CurrentRow.Cells[4].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
                ApplicantIdoc.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                IssuedSource.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                string IssueDate = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                string[] YearMonthDay = IssueDate.Split('/');
                int year, month, date;
                year = Convert.ToInt16(YearMonthDay[2]);
                month = Convert.ToInt16(YearMonthDay[0]);
                date = Convert.ToInt16(YearMonthDay[1]);

                PassIssueDate.Value = new DateTime(year, month, date);
                if (language.Text == "الانجليزية")
                {
                    language.CheckState = CheckState.Unchecked;
                    countryNonArab.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                }
                else
                {
                    countryArab.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                    language.CheckState = CheckState.Checked;
                }
                GregorianDate.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                HijriDate.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                AllFamilyMembers.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[13].Value.ToString().ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;
                    VisaAppId.Text = NextRelId;
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
                personalNonPersonal.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
                if (personalNonPersonal.Text == "شخصي") personalNonPersonal.CheckState = CheckState.Checked;
                else personalNonPersonal.CheckState = CheckState.Unchecked;
                personalNonPersonalChecking();
                PreRelatedID = dataGridView1.CurrentRow.Cells[19].Value.ToString();
                Comment.Text = dataGridView1.CurrentRow.Cells[24].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[25].Value.ToString() != "غير مؤرشف")
                {
                    ArchivedSt.CheckState = CheckState.Checked;
                    ArchivedSt.Text = "مؤرشف";
                    ArchivedSt.BackColor = Color.Green;
                }
                else {
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
        }

        
    }
}