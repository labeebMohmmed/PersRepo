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
    public partial class Form10 : Form
    {
        static public bool ApplicantSexStatus = false;
        public static string route = "";
        static bool[] All_okey = new bool[10];

        public static string[] ChildName = new string[10];
        static string [,] Muaamla = new string[3,5];
        static bool[] Son_Daughter = new bool[10];
        private string title = "حامل ";
        string Viewed;
        string ConsulateEmpName;
        public static string ModelFileroute = "";
        String IqrarStaticPart = "ق س ج/160/10/";
        String IqrarNumberPart;
        static string DataSource;
        int ApplicantID = 0;
        private bool fileloaded = false;
        string NewFileName="";
        string PreAppId = "", PreRelatedID="",NextRelId="";
        static public string FamilySupport;
        private string[] FamelyMember = new string[10];
        string  FormTypeDocx, Auth;
        string FilesPathIn, FilesPathOut;
        public Form10(int currentRow, int DocumentType, string EmpName, string dataSource, string filepathIn, string filepathOut)
        {
            InitializeComponent();
            timer1.Enabled = true;
            timer2.Enabled = true;
            Muaamla[0,0] = "أنا المواطن/";
            Muaamla[0,1] = "، أقر وبكامل قـــواي العقليـــة وحالــــتي المعتبــرة شــرعاً وقانوناً وبطوعي واختياري بأنه";
            Muaamla[0,2] = "، وهذا إقرار مني بذلك .";
            DataSource = dataSource;
            FilesPathIn = filepathIn;
            FilesPathOut = filepathOut;
            FormType.SelectedIndex = DocumentType;
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
            ApplicantName.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
            else if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
            DocType.Text = dataGridView1.Rows[Rowindex].Cells[4].Value.ToString();
            AppDocNo.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString();
            IssuedSource.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString();
            text.Text = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString();
            FormType.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString();
            GregorianDate.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString();
            HijriDate.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[11].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[12].Value.ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                Iqrarid.Text = NextRelId;
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
            SqlDataAdapter sqlDa = new SqlDataAdapter("FreeFormViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ApplicantName", ListSearch.Text.Trim());
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            IqrarNumberPart = (dtbl.Rows.Count + 1).ToString();
            sqlCon.Close();
            NewFileName = IqrarNumberPart + "_10";
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
            GregorianDate.Text = DateTime.Now.ToString("dd-MM-yyyy");
        }

        private void Review_Click(object sender, EventArgs e)
        {
                    }

        private void text_Click(object sender, EventArgs e)
        {
            
            //string Auth = ApplicantName.Text + "،المقيم بالمملكة العربية الســـعودية " + title + PassIqama.Text + " رقم " + AppDocNo.Text+ " إصدار " + IssuedSource.Text;
            //Authtext.Text = Muaamla[0, 0] + Auth+ Muaamla[0, 1];
        }

        private void ApplicantSex_CheckedChanged(object sender, EventArgs e)
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

        }


        private void Save2DataBase()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);

            string AppGender;
            if (ApplicantSex.CheckState == CheckState.Unchecked)
                AppGender = "ذكر"; else AppGender = "أنثى";
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                if (btnSavePrint.Text == "طباعة وحفظ")
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
                    sqlCmd.Parameters.AddWithValue("@DocNo", AppDocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocIssueSource", IssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SpecText", text.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SpecType", FormType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", PreAppId.Trim());
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
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
                    
                    sqlCmd.Parameters.AddWithValue("@ArchivedState","غير مؤرشف");
                    sqlCmd.ExecuteNonQuery();
                }
                else {
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
                    sqlCmd.Parameters.AddWithValue("@DocNo", AppDocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocIssueSource", IssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SpecText", text.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SpecType", FormType.Text.Trim());                   
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", PreAppId.Trim());
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
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

        private void button2_Click(object sender, EventArgs e)
        {
            Save2DataBase();
            CreateWordFile();
            

        }

        private void CreateWordFile()
        {
            route = FilesPathIn + FormTypeDocx;

            if (ApplicantSex.CheckState == CheckState.Checked)
            {
                ApplicantSexStatus = false;
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
            }
            
            string ActiveCopy = FilesPathOut + ApplicantName.Text + NewFileName + ".docx";
            
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
            object ParaDocSource = "MarkDocSource";
            object ParaDocNo = "MarkDocNo";
            object ParaPurpose = "MarkPurpose";
            object ParaPurposeText = "MarkPurposeText";
            object ParavConsul = "MarkViseConsul";
            object ParaAuthorization = "MarkAuthorization";

               

                Word.Range BookIqrarNo = oBDoc.Bookmarks.get_Item(ref ParaIqrarNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            Word.Range BookName;
            Word.Range BookDocType;
            Word.Range BookDocSource;
            Word.Range BookDocNo;
            if (FormType.SelectedIndex == 0 || FormType.SelectedIndex == 1)
            {
                BookName = oBDoc.Bookmarks.get_Item(ref ParaName).Range;
                BookDocType = oBDoc.Bookmarks.get_Item(ref ParaDocType).Range;
                BookDocSource = oBDoc.Bookmarks.get_Item(ref ParaDocSource).Range;
                BookDocNo = oBDoc.Bookmarks.get_Item(ref ParaDocNo).Range;
                Word.Range BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;
                BookName.Text = ApplicantName.Text;
                BookDocType.Text = DocType.Text;
                BookDocSource.Text = IssuedSource.Text;
                BookDocNo.Text = AppDocNo.Text;
                BookPurpose.Text = FormType.Text;
                object rangeName = BookName;
                object rangeDocSource = BookDocSource;
                object rangeDocType = BookDocType;
                object rangeDocNo = BookDocNo;
                object rangePurpose = BookPurpose;
                oBDoc.Bookmarks.Add("MarkAppName", ref rangeName);
                oBDoc.Bookmarks.Add("MarkDocType", ref rangeDocSource);
                oBDoc.Bookmarks.Add("MarkDocSource", ref rangeDocType);
                oBDoc.Bookmarks.Add("MarkDocNo", ref rangeDocNo);
                oBDoc.Bookmarks.Add("MarkPurpose", ref rangePurpose);
            }
            
            Word.Range BookPurposeText = oBDoc.Bookmarks.get_Item(ref ParaPurposeText).Range;
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;
            Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;
 
            BookIqrarNo.Text = Iqrarid.Text;
            BookGreData.Text = GregorianDate.Text;
            BookHijriData.Text = HijriDate.Text;

            
            
            BookPurposeText.Text = text.Text;
            BookAuthorization.Text = Auth;
            BookvConsul.Text  = AttendViceConsul.Text;

            object rangeIqrarNo = BookIqrarNo;
            object rangeGreData = BookGreData;
            object rangeHijriData = BookHijriData;
            
            
            object rangePurposeText = BookPurposeText;
            object rangevConsul = BookvConsul;
            object rangeAuthorization = BookAuthorization;

            oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);
            
            
            oBDoc.Bookmarks.Add("MarkPurposeText", ref rangePurposeText);
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

     



        private void OpenFile(int id, int fileNo)
        {
            string query;
            
            SqlConnection Con = new SqlConnection(DataSource);
            if (fileNo == 1)
            {
                query = "select Data1, Extension1,FileName1 from TableFreeForm where ID=@id";
            }
            else {
                query = "select Data2, Extension2,FileName2 from TableFreeForm where ID=@id"; 
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

        private void btnSavePrint_Click(object sender, EventArgs e)
        {
            Save2DataBase();
            btnSavePrint.Text = "جاري المعالجة";
            btnSavePrint.Enabled = false;
            CreateWordFile();
            Clear_Fields();
        }

        private void printOnly_Click(object sender, EventArgs e)
        {
            btnprintOnly.Text = "طباعة";
            btnprintOnly.Enabled = false;
            CreateWordFile();
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

        private void PassIqama_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DocType.Text == "اقامة ") { DocType.Text = ""; } else DocType.Text = "P";
        }

        private void Clear_Fields()
        {
            ApplicantName.Text = IssuedSource.Text = IssuedSource.Text = "";

            ApplicantSex.CheckState = CheckState.Unchecked;
            labeldoctype.Text = "رقم جواز السفر: ";
            AppDocNo.Text = "P";
            FormType.SelectedIndex = 0;
            text.Text = "";
            AttendViceConsul.SelectedIndex = 2;
            DocType.SelectedIndex = 0;
            Iqrarid.Text = IqrarStaticPart + IqrarNumberPart;
            mandoubName.Text = ListSearch.Text = "";
            ApplicantSex.CheckState = CheckState.Checked;
            mandoubVisibilty();
            btnprintOnly.Visible = false;
            btnSavePrint.Text = "طباعة وحفظ";
            btnSavePrint.Enabled = true;
            btnSavePrint.Visible = true;
            SaveOnly.Visible = false;
            Comment.Text = "لا تعليق";
            FillDataGridView();
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            labelArch.Visible = false;
            ArchivedSt.BackColor = Color.Red;
            SearchFile.Visible = false;
            fileloaded = false;
            AttendViceConsul.SelectedIndex = 2;
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            Iqrarid.Text = IqrarStaticPart + IqrarNumberPart;
            ConsulateEmployee.Text = ConsulateEmpName;
            countryNonArab.SelectedIndex = 0;
            countryArab.SelectedIndex = 0;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void FormType_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void strAuth()
        {            
            if (FormType.Text == "إقرار" || FormType.Text == "إقرار مشفوع باليمين")
            {
                if (ApplicantSex.CheckState == CheckState.Unchecked)
                {
                    FormTypeDocx = "GenIqrarM.docx";
                    if (AppType.CheckState == CheckState.Checked)
                        Auth = "أشهد أنا/" + AttendViceConsul.Text + " نائب قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد حضر للقنصلية ووقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                    else Auth = "المواطن المذكور أعلاه حضر ووقع بتوقيعه على هذا الإقرار أمام مندوب الجالية لدى القنصلية السيد/ " + mandoubName.Text + "، وذلك بموجب التفويض الممنوح له، ";
                }
                else
                {
                    FormTypeDocx = "GenIqrarF.docx";
                    if (AppType.CheckState == CheckState.Checked)
                        Auth = "أشهد أنا/ " + AttendViceConsul.Text + " نائب قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد حضرت للقنصلية ووقعت بتوقيعها على هذا الإقرار بعد تلاوتها عليها وبعد أن فهمت مضمونه ومحتواه. ";
                    else Auth = "المواطنة المذكورة أعلاه حضرت ووقعت بتوقيعها على هذا الإقرار أمام مندوب الجالية لدى القنصلية السيد/ " + mandoubName.Text + "، وذلك بموجب التفويض الممنوح له، ";
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
            else if (FormType.SelectedIndex == 4)
            {
                countryArab.Visible = true;
                countryNonArab.Visible = false;
                labelCountry.Visible = true;
                PassIssueDate.Visible = true;
                labelIssue.Visible = true;
            }
            else if (FormType.SelectedIndex == 5)
            {
                countryNonArab.Visible = true;
                countryArab.Visible = false;
                labelCountry.Visible = true;
                PassIssueDate.Visible = true;
                labelIssue.Visible = true;
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                PreAppId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                ApplicantName.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
                else if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
                DocType.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                AppDocNo.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                IssuedSource.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                text.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                FormType.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                GregorianDate.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                HijriDate.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[12].Value.ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;
                    Iqrarid.Text = NextRelId;
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
                strAuth();
                ArchivedSt.Visible = true;
                labelArch.Visible = true;
                btnprintOnly.Visible = true;
                SaveOnly.Visible = true;
                btnSavePrint.Text = "حفظ";
                btnSavePrint.Visible = false;

            }
        }

        private void SearchDoc_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            SearchFile.Visible = true;
            SearchFile.Text = dlg.FileName;
            if (SearchFile.Text != "") fileloaded = true;
        }
    }
}
