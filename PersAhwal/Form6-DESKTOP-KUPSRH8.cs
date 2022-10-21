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
    public partial class Form6 : Form
    {
        public static string[] SponceDesc3 = new string[10];
        public static string[] SponceDoc3 = new string[10];
        public static string[] SponcerName3 = new string[10];
        public static string[] SponcePassIqama3 = new string[10];
        public static string[] ApplicantIdocNo3 = new string[10];
        public static string[] SponceIssueSource3 = new string[10];

        public static string[] DaughterMotheDocSource = new string[10];

        public static string route = "";
        string Viewed;
        string ConsulateEmpName;
        public static string ModelFileroute = "";
        String IfadaStaticPart = "ق س ج/160/06/";
        String IfadaNumberPart;
        static string DataSource;
        string NewFileName;
        string PreAppId = "", PreRelatedID="", NextRelId="";
        static public string FamilySupport;
        private string[] FamelyMember = new string[10];
        string FilesPathIn, FilesPathOut;
        private int ApplicantID = 0;
        private bool fileloaded = false;
        public Form6(int currentRow, string EmpName, string dataSource, string filepathIn, string filepathOut)
        {
            InitializeComponent();
            timer1.Enabled = true;
            timer2.Enabled = true;
            DataSource = dataSource;
            AttendViceConsul.SelectedIndex = 2;
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
            PreAppId = dataGridView1.Rows[Rowindex].Cells[1].Value.ToString();
            ApplicantIdocName.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
            else if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
            ApplicantPassNo.Text = dataGridView1.Rows[Rowindex].Cells[4].Value.ToString();
            PassIssuedSource.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString();
            ApplicantIqamaNo.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString();
            IqamaIssuedSource.Text = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString();
            GregorianDate.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString();
            HijriDate.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[16].Value.ToString().ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                Ifadaid.Text = NextRelId;
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
            labelArch.Visible = true;
            btnprintOnly.Visible = true;
            SaveOnly.Visible = true;
            btnSavePrint.Text = "حفظ";
            btnSavePrint.Visible = false;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            GregorianDate.Text = DateTime.Now.ToString("dd-MM-yyyy");
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

        private void ApplicantSex_CheckedChanged(object sender, EventArgs e)
        {
            if (ApplicantSex.CheckState == CheckState.Unchecked)
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

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void CreateWordFile()
        {
            if (ApplicantSex.CheckState == CheckState.Unchecked)
            {

                labelName.ForeColor = Color.Black;
                labelName.Text = "مقدم الطلب:";
                route = FilesPathIn + "SaudiForesnecM.docx";
            }
            else if (ApplicantSex.CheckState == CheckState.Checked)
            {
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
                route = FilesPathIn + "SaudiForesnecF.docx";
            }

            Save2DataBase();
            string ActiveCopy;
            ActiveCopy = FilesPathOut + ApplicantIdocName.Text + NewFileName + ".docx";
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

            BookIfadaNo.Text = Ifadaid.Text;
            BookGreData.Text = GregorianDate.Text;
            BookHijriData.Text = HijriDate.Text;
            Bookname.Text = ApplicantIdocName.Text;
            BookAppPass.Text = ApplicantPassNo.Text;
            BookAppPassSource.Text = PassIssuedSource.Text;
            BookAppIqama.Text = ApplicantIqamaNo.Text;
            BookAppIqamaSource.Text = IqamaIssuedSource.Text;
            BookvConsul.Text = AttendViceConsul.Text;

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



            oBDoc.Activate();

            oBDoc.Save();
            oBMicroWord.Visible = true;
                }
                else
                {
                    MessageBox.Show("يرجى حذف الملف الموجودأولاً");
                    btnprintOnly.Enabled = true;
                    btnSavePrint.Enabled = true;
                    
                }
            }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                PreAppId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                ApplicantIdocName.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
                else if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
                ApplicantPassNo.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                PassIssuedSource.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                ApplicantIqamaNo.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                IqamaIssuedSource.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                GregorianDate.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                HijriDate.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[16].Value.ToString().ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;
                    Ifadaid.Text = NextRelId;
                }
                else checkedViewed.CheckState = CheckState.Checked;

                AppType.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                ConsulateEmployee.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked; else AppType.CheckState = CheckState.Unchecked;

                if (AppType.CheckState == CheckState.Unchecked)
                {
                    mandoubVisibilty(); mandoubName.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                }

                PreRelatedID = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                Comment.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();
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
                labelArch.Visible = true;
                btnprintOnly.Visible = true;
                btnprintOnly.Text = "طباعة";
                SaveOnly.Visible = true;
                btnSavePrint.Text = "حفظ";
                btnSavePrint.Visible = false;
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
        private void Form6_Load(object sender, EventArgs e)
        {

        }

        private void Review_Click(object sender, EventArgs e)
        {

        }


        private void Save2DataBase()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);

            string AppGender;
            if (ApplicantSex.CheckState == CheckState.Unchecked) AppGender = "ذكر"; else AppGender = "أنثى";
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                else Viewed = "غير معالج";
                SqlCommand sqlCmd = new SqlCommand("ForensicAddorEdit", sqlCon);
                sqlCmd.CommandType = CommandType.StoredProcedure;

                if (btnSavePrint.Text == "طباعة وحفظ")
                {
                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");
                    sqlCmd.Parameters.AddWithValue("@DocID", Ifadaid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantIdocName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@PassNo", ApplicantPassNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@PassSource", PassIssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@IqamaNo", ApplicantIqamaNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@IqamaSource", IqamaIssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", PreAppId.Trim());
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
                    sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                    sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                    sqlCmd.Parameters.AddWithValue("@DocID", Ifadaid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantIdocName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@PassNo", ApplicantPassNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@PassSource", PassIssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@IqamaNo", ApplicantIqamaNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@IqamaSource", IqamaIssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", PreAppId.Trim());
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
            IfadaNumberPart = (dtbl.Rows.Count + 1).ToString();
            dataGridView1.Columns[0].Visible = false;
            sqlCon.Close();
            NewFileName = IfadaNumberPart + "_06";
        }

        private void AppType_CheckedChanged(object sender, EventArgs e)
        {
            mandoubVisibilty();
        }

        private void btnSavePrint_Click(object sender, EventArgs e)
        {
            Save2DataBase();
            btnSavePrint.Enabled = false;
            btnSavePrint.Text = "جاري المعالجة";
            CreateWordFile();

            Clear_Fields();
        }
        private void Clear_Fields()
        {
            IqamaIssuedSource.Text = ApplicantIqamaNo.Text = PassIssuedSource.Text = ApplicantIdocName.Text = "";
            ApplicantSex.CheckState = CheckState.Checked;
            ApplicantPassNo.Text = "P";
            ApplicantSex.CheckState = CheckState.Checked;
            AttendViceConsul.SelectedIndex = 2;
            Ifadaid.Text = IfadaStaticPart + IfadaNumberPart;
            mandoubName.Text = ListSearch.Text = "";
            AppType.CheckState = CheckState.Unchecked;
            mandoubVisibilty();
            btnprintOnly.Visible = false;
            btnSavePrint.Text = "طباعة وحفظ";
            btnprintOnly.Enabled = true;
            btnprintOnly.Text = "طباعة";
            btnSavePrint.Visible = true;
            SaveOnly.Visible = false;
            Comment.Text = "لا تعليق";
            FillDataGridView();
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            labelArch.Visible = false;
            ArchivedSt.BackColor = Color.Red;
            fileloaded = false;
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            Ifadaid.Text = IfadaStaticPart + IfadaNumberPart;
            ConsulateEmployee.Text = ConsulateEmpName;
        }

        private void printOnly_Click(object sender, EventArgs e)
        {
            btnprintOnly.Enabled = false;
            btnprintOnly.Text = "طباعة";
            CreateWordFile();
            Clear_Fields();
        }

        private void SearchDoc_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            SearchFile.Text = dlg.FileName;
            if (SearchFile.Text != "") fileloaded = true;
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

        private void ResetAll_Click(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void SaveOnly_Click(object sender, EventArgs e)
        {
            Save2DataBase();
            Clear_Fields();
        }
    }
}
