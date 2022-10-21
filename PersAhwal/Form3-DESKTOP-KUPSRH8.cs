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
    public partial class Form3 : Form
    {
        public static int i = 0,x=0;
        public static string[] DaughterMother = new string[10];
        public static string[] DaughterMotheDocType = new string[10];
        public static string[] DaughterMotheDoc = new string[10];
        public static string[] DaughterMotheDocSource = new string[10];
        static bool Firstline = false;
        
        static public string titleFam, FamilySupport;
        public static string route = "";
        string Viewed;
        string ConsulateEmpName;
        public static string ModelFileroute = "";
        String IqrarStaticPart = "ق س ج/160/03/";
        String IqrarNumberPart;
        static string DataSource;
        bool fileloaded = false;
        int ApplicantID = 0;
        string NewFileName,CurrentIqrarId="";
        string FilesPathIn, FilesPathOut, PreAppId="", PreRelatedID="",NextRelId="";
        public Form3(int currentRow, int IqrarType, string EmpName, string dataSource, string filepathIn, string filepathOut)
        {
            InitializeComponent();
            ConsulateEmpName = EmpName;
            IqrarPurpose.SelectedIndex = IqrarType+1;
            if (IqrarType  == 7) {
                panel1.Visible = true;
                i = 0;
            }
            DataSource = dataSource;            
            timer1.Enabled = true;
            timer2.Enabled = true;
            FilesPathIn = filepathIn;
            FilesPathOut = filepathOut;
            FillDataGridView();
            if (currentRow == -1) Clear_Fields();
            else SetFieldswithData(currentRow);
            IqrarPurpose.SelectedIndex = IqrarType + 1;
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
            IqrarNumberPart = (dtbl.Rows.Count + 1).ToString();
            dataGridView1.Columns[0].Visible = false;
            sqlCon.Close();
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ApplicantSex.CheckState == CheckState.Unchecked && IqrarPurpose.Text == "غير متزوجة") 
            {
                personalNonPersonal.CheckState = CheckState.Unchecked;
                panel1.Visible = true;
            }

            if (IqrarPurpose.Text == "غير متزوج" || IqrarPurpose.Text == "غير متزوجة")
                personalNonPersonal.Visible = true;
            else
            {
                personalNonPersonal.Visible = false;
                panel1.Visible = false;
            }

            if (IqrarPurpose.Text == "إعالة أسرية") 
            {
                panel1.Visible = true;
                i = 0;

            }
        }

        private void SetFieldswithData(int Rowindex)
        {
            Rowindex--;
            ApplicantID = Convert.ToInt32(dataGridView1.Rows[Rowindex].Cells[0].Value.ToString());
            IqrarNo.Text = CurrentIqrarId = dataGridView1.Rows[Rowindex].Cells[1].Value.ToString();
            ApplicantName.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString().ToString();
            if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
            else if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
            PassIqama.Text = dataGridView1.Rows[Rowindex].Cells[4].Value.ToString().ToString();
            ApplicantIdoc.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString().ToString();
            IssuedSource.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString().ToString();
            IqrarPurpose.Text = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString().ToString();
            if (IqrarPurpose.Text == "إعالة") panel1.Visible = true; else panel1.Visible = false;
            AllFamilyMembers.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString().ToString();
            GregorianDate.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString().ToString();
            HijriDate.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString().ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[11].Value.ToString().ToString();
            if (dataGridView1.Rows[Rowindex].Cells[12].Value.ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                IqrarNo.Text = NextRelId;
            }
            else checkedViewed.CheckState = CheckState.Checked;
            AppType.Text = dataGridView1.Rows[Rowindex].Cells[13].Value.ToString();
            if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked; else AppType.CheckState = CheckState.Unchecked;

            if (AppType.CheckState == CheckState.Unchecked) { 
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
            labelArch.Visible = true;
            printOnly.Visible = true;
            SaveOnly.Visible = true;
            btnSavePrint.Text = "حفظ";
            btnSavePrint.Visible = false;
        }


        private void personalNonPersonal_CheckedChanged(object sender, EventArgs e)
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
            HijriDate.Text = Stringdate + "-" + month.ToString() + "-" + year.ToString();
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

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void AddChildren_Click_1(object sender, EventArgs e)
        {
            string newLine = Environment.NewLine;
            
            if (motherDaughter.Text == "ابني" || motherDaughter.Text == "والدي" || motherDaughter.Text == "شقيقي") titleFam = " حامل ";else titleFam = " حاملة ";
            DaughterMother[i] = motherDaughter.Text + " " + FamilyMebersName.Text;
            FamilySupport = DaughterMother[0] + titleFam + textBox5.Text + " رقم " + textBox3.Text + " إصدار " + textBox1.Text+"،";
            if (!Firstline) AllFamilyMembers.Text = (i + 1).ToString() + "- " + DaughterMother[i] + titleFam+ textBox5.Text + " رقم " + textBox3.Text + " إصدار " + textBox1.Text;
            else AllFamilyMembers.Text = AllFamilyMembers.Text + newLine + (i + 1).ToString() + "- " + DaughterMother[i] + titleFam + textBox5.Text + " رقم " + textBox3.Text + " إصدار " + textBox1.Text;
            Firstline = true;
            i++;
            FamilyMebersName.Clear();
        }

        private void ApplicantName_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void CreateWordFile()
        {
            string newLine = Environment.NewLine;
            Firstline = false;
            if (ApplicantSex.CheckState == CheckState.Unchecked)
            {

                labelName.ForeColor = Color.Black;
                labelName.Text = "مقدم الطلب:";
                route = FilesPathIn +"Igrar_SocialStatusM.docx";
            }
            else if (ApplicantSex.CheckState == CheckState.Checked)
            {
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
                route = FilesPathIn + "Igrar_SocialStatusF.docx";
            }
            //
            for (x = 0; x <= i || personalNonPersonal.CheckState == CheckState.Checked; x++)
            {

                string ActiveCopy;
                if (x == 0) ActiveCopy = FilesPathOut + ApplicantName.Text + NewFileName + ".docx";
                else ActiveCopy = FilesPathOut + ApplicantName.Text + NewFileName + x.ToString() + ".docx";

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

                    BookIqrarNo.Text = IqrarNo.Text;
                    Bookname1.Text = Bookname2.Text = ApplicantName.Text;
                    Bookigama.Text = ApplicantIdoc.Text;
                    BookvConsul.Text = AttendViceConsul.Text;
                    BookAppiIssSource.Text = IssuedSource.Text;
                    BookPassIqama.Text = PassIqama.Text;
                    BookGreData.Text = GregorianDate.Text;
                    BookHijriData.Text = HijriDate.Text;
                    if (IqrarPurpose.Text == "إعالة أسرية")
                    {
                        if (i == 1) BookPurpose.Text = "العائل الوحيد بعد الله عز وجل ل" + FamilySupport + "،";
                        else BookPurpose.Text = "العائل الوحيد بعد الله عز وجل لأسرتي المكونة من الآتي:" + newLine + AllFamilyMembers.Text + "." + newLine;
                        if (personalNonPersonal.CheckState == CheckState.Checked) BookFamilyName.Text = "ي";
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
                        if (ApplicantSex.CheckState == CheckState.Unchecked) BookPurpose.Text = "أقر بأنني متزوج، والله على ما أقول شهيد";
                        else BookPurpose.Text = "أقر بأنني متزوجة، والله على ما أقول شهيد";
                    }
                    else if (IqrarPurpose.SelectedIndex == 3)
                    {
                        if (ApplicantSex.CheckState == CheckState.Unchecked) BookPurpose.Text = "أقر بأنني غير متزوج، والله على ما أقول شهيد";
                        else BookPurpose.Text = "أقر بأنني غير متزوجة، والله على ما أقول شهيد";
                    }
                    else if (IqrarPurpose.SelectedIndex == 4)
                    {
                        if (ApplicantSex.CheckState == CheckState.Checked) BookPurpose.Text = "أقر بأنني أرملة وبأني لم اتزوج بعد وفاة زوجي، والله على ما أقول شهيد";
                        else MessageBox.Show("اختيار خاظئ لجنس مقدم الطلب");

                    }
                    else
                    {
                        BookPurpose.Text = IqrarPurpose.Text + "، والله على ما أقول شهيد";
                        if (personalNonPersonal.CheckState == CheckState.Checked) BookFamilyName.Text = " أقر بأني" + DaughterMother[x];
                        else BookFamilyName.Text = " إقر بأن " + DaughterMother[x];
                    }

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
                    oBDoc.Activate();
                    oBDoc.Save();
                    oBMicroWord.Visible = true;
                    if (personalNonPersonal.CheckState == CheckState.Checked) break;
               }
                else { 
                    MessageBox.Show("يرجى حذف الملف الموجودأولاً");
                    printOnly.Enabled = true;
                    btnSavePrint.Enabled = true;
                    i = 0;
                    break;
                }
            }
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
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
                if (btnSavePrint.Text == "حفظ وطباعة")
                {

                    if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                    else Viewed = "غير معالج";
                    SqlCommand sqlCmd = new SqlCommand("MultiAddorEdit", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");
                    sqlCmd.Parameters.AddWithValue("@DocID", IqrarNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", PassIqama.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", ApplicantIdoc.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", IssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@IqrarPurpose", IqrarPurpose.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@FamilyName", AllFamilyMembers.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text);
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", "");
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
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", PassIqama.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", ApplicantIdoc.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", IssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@IqrarPurpose", IqrarPurpose.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@FamilyName", AllFamilyMembers.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text);
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
                IqrarNo.Text = CurrentIqrarId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                ApplicantName.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString().ToString();
                if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
                else if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
                PassIqama.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().ToString();
                ApplicantIdoc.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString().ToString();
                IssuedSource.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString().ToString();
                IqrarPurpose.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString().ToString();
                if (IqrarPurpose.Text == "إعالة") panel1.Visible = true; else panel1.Visible = false;
                AllFamilyMembers.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString().ToString();
                GregorianDate.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString().ToString();
                HijriDate.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString().ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString().ToString();
                if (dataGridView1.CurrentRow.Cells[12].Value.ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;
                    IqrarNo.Text = NextRelId;
                }
                else checkedViewed.CheckState = CheckState.Checked;
                AppType.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked; else AppType.CheckState = CheckState.Unchecked;

                if (AppType.CheckState == CheckState.Unchecked) { 
                    mandoubVisibilty(); mandoubName.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();  
                }
                ConsulateEmployee.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
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
                printOnly.Visible = true;
                SaveOnly.Visible = true;
                btnSavePrint.Text = "حفظ";
                btnSavePrint.Visible = false;
            }
        }

        private void PassIqama_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (PassIqama.Text == "اقامة ") { ApplicantIdoc.Text = ""; } else ApplicantIdoc.Text = "P";
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
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            Search.Text = dlg.FileName;
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

        private void button5_Click(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            IqrarNo.Text = CurrentIqrarId;
            printOnly.Text = "طباعة";
            printOnly.Enabled = false;
            CreateWordFile();
            Clear_Fields();
        }

        private void btnSavePrint_Click(object sender, EventArgs e)
        {
            Save2DataBase();
            if (btnSavePrint.Text != "حفظ وطباعة") return;
            btnSavePrint.Text = "جاري المعالجة";
            btnSavePrint.Enabled = false;
            CreateWordFile();
            Clear_Fields();
        }

        private void ApplicantIdoc_TextChanged(object sender, EventArgs e)
        {

        }

        private void SaveOnly_Click(object sender, EventArgs e)
        {
            Save2DataBase();            
            Clear_Fields();
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
             ApplicantName.Text = IssuedSource.Text =  ApplicantIdoc.Text = PassIqama.Text =  "";
             AttendViceConsul.SelectedIndex = 2;
             PassIqama.SelectedIndex = 0;
             ApplicantIdoc.Text = "P";
             PassIqama.Text = "جواز سفر";
             mandoubName.Text = Search.Text = "";
             AppType.CheckState = CheckState.Checked;
             mandoubVisibilty();
            btnSavePrint.Visible = true;
            btnSavePrint.Text = "حفظ وطباعة";
            btnSavePrint.Visible = true;
            Comment.Clear();
            panel1.Visible = false;
            FillDataGridView();
            AttendViceConsul.SelectedIndex = 2;            
            IqrarNo.Text = IqrarStaticPart + IqrarNumberPart;
            i = 0;
            NewFileName = IqrarNumberPart + "_03";
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            labelArch.Visible = false;
            ArchivedSt.BackColor = Color.Red;
            printOnly.Enabled = true;
            printOnly.Visible = false;
            SaveOnly.Visible = false;
            IqrarPurpose.SelectedIndex = 0;
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if(textBox5.Text == "إقامة")  labelIqama.Text = "رقم الاقامة:";
            else labelIqama.Text = "رقم جواز السفر:";
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            GregorianDate.Text = DateTime.Now.ToString("dd-MM-yyyy");
        }
    }
}
