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
        String IqrarStaticPart = "ق س ج/160/07/";
        String IqrarNumberPart;
        static string DataSource;
        int ApplicantID = 0;
        string FilesPathIn, FilesPathOut;
        string NewFileName;
        private bool fileloaded = false;
        string PreAppId = "", PreRelatedID="", NextRelId="";
        static public string FamilySupport;
        private string[] FamelyMember = new string[10];

        public Form7(int currentRow, string EmpName, string dataSource, string filepathIn, string filepathOut)
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
            PreAppId = dataGridView1.Rows[Rowindex].Cells[1].Value.ToString();
            ApplicantIdocName.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
            else if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
            DocType.Text = dataGridView1.Rows[Rowindex].Cells[4].Value.ToString();
            DocNo.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString();
            IssuedSource.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString();
            AppTrueName.Text = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString();
            AppWrongName.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString();
            GregorianDate.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString();
            HijriDate.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[11].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[16].Value.ToString().ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                Iqrarid.Text = NextRelId;
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
            btnprintOnly.Visible = true;
            SaveOnly.Visible = true;
            btnSavePrint.Text = "حفظ";
            btnSavePrint.Visible = false;
        }

        private void OpenFile(int id, int fileNo)
        {
            string query;

            SqlConnection Con = new SqlConnection(DataSource);
            if (fileNo == 1)
            {
                query = "select Data1, Extension1,FileName1 from TableTRName where ID=@id";
            }
            else
            {
                query = "select Data2, Extension2,FileName2 from TableTRName where ID=@id";
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
                SqlCommand sqlCmd = new SqlCommand("TRNameAddorEdit", sqlCon);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                if (btnSavePrint.Text == "طباعة وحفظ")
                {
                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");
                    sqlCmd.Parameters.AddWithValue("@DocID", Iqrarid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantIdocName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", DocType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", DocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocIssueSource", IssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppTrueName", AppTrueName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppWrongName", AppWrongName.Text.Trim());
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
                else {
                    sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                    sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                    sqlCmd.Parameters.AddWithValue("@DocID", Iqrarid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantIdocName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", DocType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", DocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocIssueSource", IssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppTrueName", AppTrueName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppWrongName", AppWrongName.Text.Trim());
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
            IqrarNumberPart = (dtbl.Rows.Count + 1).ToString();
            dataGridView1.Columns[0].Visible = false;
            sqlCon.Close();
            NewFileName = ApplicantIdocName.Text + IqrarNumberPart + "_07";
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
                labelName.Text = "الاسم الصحيح لمقدم الطلب:";
            }
            else if (ApplicantSex.CheckState == CheckState.Checked)
            {
                ApplicantSex.Text = "إنثى";
                labelName.Text = "الاسم الصحيح لمقدمة الطلب:";
            }
        }

        private void Review_Click(object sender, EventArgs e)
        {
           
        }

        private void Form7_Load(object sender, EventArgs e)
        {

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
                route = FilesPathIn + "IqrarTrueNameM.docx";
            }
            else if (ApplicantSex.CheckState == CheckState.Checked)
            {
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
                route = FilesPathIn + "IqrarTrueNameF.docx";
            }
            string ActiveCopy;
            ActiveCopy = FilesPathOut + ApplicantIdocName.Text + NewFileName + ".docx";
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
            object ParaAppTrueName1 = "MarkAppTrueName1";
            object ParaAppTrueName2 = "MarkAppTrueName2";
            object ParaAppWrongName = "MarkAppWorngName";
            object ParaAppPass = "MarkAppDoc";
            object ParaAppPassSource = "MarkAppPassSource";
            object ParavConsul = "MarkViseConsul";
            object ParaAuthorization = "MarkAuthorization";

            Word.Range BookIqrarNo = oBDoc.Bookmarks.get_Item(ref ParaIqrarNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            Word.Range BookDocName = oBDoc.Bookmarks.get_Item(ref ParaAppDocName).Range;
            Word.Range BookTrueName1 = oBDoc.Bookmarks.get_Item(ref ParaAppTrueName1).Range;
            Word.Range BookTrueName2 = oBDoc.Bookmarks.get_Item(ref ParaAppTrueName2).Range;
            Word.Range BookWrongName = oBDoc.Bookmarks.get_Item(ref ParaAppWrongName).Range;
            Word.Range BookAppPass = oBDoc.Bookmarks.get_Item(ref ParaAppPass).Range;
            Word.Range BookAppPassSource = oBDoc.Bookmarks.get_Item(ref ParaAppPassSource).Range;
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;
            Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;

            BookIqrarNo.Text = Iqrarid.Text; 
            BookGreData.Text = GregorianDate.Text;
            BookHijriData.Text = HijriDate.Text;
            BookDocName.Text = ApplicantIdocName.Text;
            BookTrueName1.Text = BookTrueName2.Text = AppTrueName.Text;
            BookWrongName.Text = AppWrongName.Text;
            BookAppPass.Text = DocNo.Text;
            BookAppPassSource.Text = IssuedSource.Text;
            if (AppType.CheckState == CheckState.Checked)
            {
                if (ApplicantSex.CheckState == CheckState.Unchecked) BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " نائب قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد حضر للقنصلية ووقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                if (ApplicantSex.CheckState == CheckState.Checked) BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " نائب قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد حضرت للقنصلية ووقعت بتوقيعها على هذا الإقرار بعد تلاوتها عليها وبعد أن فهمت مضمونه ومحتواه. ";
            }
            else
            {
                if (ApplicantSex.CheckState == CheckState.Unchecked) BookAuthorization.Text = "المواطن المذكور أعلاه حضر ووقع بتوقيعه على هذا الإقرار أمام مندوب الجالية لدى القنصلية السيد/ " + mandoubName.Text + "، وذلك بموجب التفويض الممنوح له، ";
                if (ApplicantSex.CheckState == CheckState.Checked) BookAuthorization.Text = "المواطنة المذكورة أعلاه حضرت ووقعت بتوقيعها على هذا الإقرار أمام مندوب الجالية لدى القنصلية السيد/ " + mandoubName.Text + "، وذلك بموجب التفويض الممنوح له، ";
            }
            BookvConsul.Text = AttendViceConsul.Text;

            object rangeIqrarNo = BookIqrarNo;
            object rangeGreData = BookGreData;
            object rangeHijriData = BookHijriData;
            object rangeDocName = BookDocName;
            object rangeTrueName1 = BookTrueName1;
            object rangeTrueName2 = BookTrueName2;
            object rangeWrongPass = BookWrongName;
            object rangeAppPass = BookAppPass;
            object rangeAppPassSource = BookAppPassSource;
            object rangevConsul = BookvConsul;
            object rangeAuthorization = BookAuthorization;

            oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkHijiData", ref rangeHijriData);
            oBDoc.Bookmarks.Add("MarkAppName", ref rangeDocName);
            oBDoc.Bookmarks.Add("MarkTrueName1", ref rangeTrueName1);
            oBDoc.Bookmarks.Add("MarkTrueName2", ref rangeTrueName2);
            oBDoc.Bookmarks.Add("MarkWrongName", ref rangeWrongPass);
            oBDoc.Bookmarks.Add("MarkAppPass", ref rangeAppPass);
            oBDoc.Bookmarks.Add("MarkAppPassSource", ref rangeAppPassSource);
            oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);
            oBDoc.Bookmarks.Add("MarkAuthorization", ref rangeAuthorization);



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
            if (dataGridView1.CurrentRow.Index != -1)
            {
                ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                PreAppId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                ApplicantIdocName.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
                else if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
                DocType.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                DocNo.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                IssuedSource.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                AppTrueName.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                AppWrongName.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                GregorianDate.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                HijriDate.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[16].Value.ToString().ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;
                    Iqrarid.Text = NextRelId;
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
                btnprintOnly.Visible = true;
                SaveOnly.Visible = true;
                btnSavePrint.Text = "حفظ";
                btnSavePrint.Visible = false;
            }
        }

        private void ApplicantIdocName_TextChanged(object sender, EventArgs e)
        {
            AppTrueName.Text = ApplicantIdocName.Text;
        }

        private void DocType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DocType.SelectedIndex == 0)
            {
                labeldoctype.Text = "رقم جواز السفر: ";
                DocNo.Text = "P";
            }
            else
            {
                DocNo.Text = ""; labeldoctype.Text = "رقم الإقامة:";
            }
        }

        private void btnSavePrint_Click(object sender, EventArgs e)
        {
            Save2DataBase();
            btnSavePrint.Text = "جاري المعالجة";
            CreateWordFile();
            btnSavePrint.Text = "حفظ وطباعة";
            btnSavePrint.Enabled = true;
            Clear_Fields();
        }

        private void printOnly_Click(object sender, EventArgs e)
        {
            btnprintOnly.Text = "طباعة";
            CreateWordFile();
            btnprintOnly.Text = "حفظ وطباعة";
            btnprintOnly.Enabled = true;
            Clear_Fields();
        }

        private void SaveOnly_Click(object sender, EventArgs e)
        {
            Save2DataBase();
            Clear_Fields();
        }

        private void ResetAll_Click(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void Clear_Fields()
        {
            AppWrongName.Text = AppTrueName.Text = ApplicantIdocName.Text = IssuedSource.Text = PersonDesc.Text = ApplicantIdocName.Text = "";
            ApplicantSex.CheckState = CheckState.Checked;
            labeldoctype.Text = "رقم جواز السفر: ";
            DocNo.Text = "P";
            AttendViceConsul.SelectedIndex = 2;
            DocType.SelectedIndex = 0;
            Iqrarid.Text = IqrarStaticPart + IqrarNumberPart;
            mandoubName.Text = ListSearch.Text = "";
            ApplicantSex.CheckState = CheckState.Checked;
            mandoubVisibilty();
            btnprintOnly.Visible = false;
            btnSavePrint.Text = "طباعة وحفظ";
            btnSavePrint.Visible = true;
            SaveOnly.Visible = false;
            Comment.Text = "لا تعليق";
            FillDataGridView();
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            labelArch.Visible = false;
            ArchivedSt.BackColor = Color.Red;
            PersonDesc.SelectedIndex = 0;
            SearchFile.Visible = false;
            fileloaded = false;
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);

            Iqrarid.Text = IqrarStaticPart + IqrarNumberPart;
            ConsulateEmployee.Text = ConsulateEmpName;
        }
        private void SearchDoc_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            SearchFile.Visible = true;
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
    }
}
