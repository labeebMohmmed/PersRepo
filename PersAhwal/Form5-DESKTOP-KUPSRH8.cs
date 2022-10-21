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
    public partial class Form5 : Form
    {
        public static string[] SponceDesc3 = new string[10];
        public static string[] SponceDoc3 = new string[10];
        public static string[] SponcerName3 = new string[10];
        public static string[] SponcePassIqama3 = new string[10];
        public static string[] ApplicantIdocNo3 = new string[10];
        public static string[] SponceIssueSource3 = new string[10];

        public static string[] DaughterMotheDocSource = new string[10];

        public static string route = "";
        public static int thirdPartyIndex = 0, x = 0, ProCase = 1;
        static string titleFam = "حاملة";
        static string titleFam3 = "حاملة";
        private bool fileloaded = false;
        string Viewed;
        string ConsulateEmpName;
        
        public static string ModelFileroute = "";
        String IqrarStaticPart = "ق س ج/160/05/";
        String IqrarNumberPart;
        static string DataSource;
        int ApplicantID = 0;
        string NewFileName;
        string PreFamilySponAppId = "";
        static public string FamilySupport;
        private string[] FamelyMember = new string[10];
        static public string title = "";
        string  PreRelatedID = "", NextRelId = "";
        string FilesPathIn, FilesPathOut;
        public Form5(int currentRow, string EmpName, string dataSource, string filepathIn, string filepathOut)
        {
            InitializeComponent();
            timer1.Enabled = true;
            timer2.Enabled = true;
            DataSource = dataSource;
            FilesPathIn = filepathIn;
            FilesPathOut = filepathOut;
            FillDataGridView();
            if (currentRow == -1) Clear_Fields();
            else SetFieldswithData(currentRow);            
            ConsulateEmpName = EmpName;
        }

        private void SetFieldswithData(int Rowindex)
        {
            ApplicantID = Convert.ToInt32(dataGridView1.Rows[Rowindex].Cells[0].Value.ToString());
            NextRelId = dataGridView1.Rows[Rowindex].Cells[1].Value.ToString();
            ApplicantName.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
            else if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
            ApplicantIdocType.Text = dataGridView1.Rows[Rowindex].Cells[4].Value.ToString();
            ApplicantIdocNo.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString();
            IssuedSource.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString();

            SponcerName.Text = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString();
            SponceDesc.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString();
            SponcePassIqama.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString();
            SponcedocNo.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString();
            SponceIssueSource.Text = dataGridView1.Rows[Rowindex].Cells[11].Value.ToString();

            GregorianDate.Text = dataGridView1.Rows[Rowindex].Cells[12].Value.ToString();
            HijriDate.Text = dataGridView1.Rows[Rowindex].Cells[13].Value.ToString();
            AllFamilyMembers.Text = dataGridView1.Rows[Rowindex].Cells[14].Value.ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[15].Value.ToString();

            if (dataGridView1.Rows[Rowindex].Cells[16].Value.ToString().ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                Iqrarid.Text = NextRelId;
            }
            else checkedViewed.CheckState = CheckState.Checked;
            AppType.Text = dataGridView1.Rows[Rowindex].Cells[17].Value.ToString();
            Employee.Text = dataGridView1.Rows[Rowindex].Cells[18].Value.ToString();
            if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked; else AppType.CheckState = CheckState.Unchecked;
            if (AppType.CheckState == CheckState.Unchecked)
            {
                mandoubVisibilty();
                mandoubName.Text = dataGridView1.Rows[Rowindex].Cells[19].Value.ToString();
            }

            PreRelatedID = dataGridView1.Rows[Rowindex].Cells[20].Value.ToString();
            SponserCase.SelectedIndex = Convert.ToInt32(dataGridView1.Rows[Rowindex].Cells[21].Value.ToString()) - 1;
            Comment.Text = dataGridView1.Rows[Rowindex].Cells[26].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[27].Value.ToString() != "غير مؤرشف")
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

        private void SponceDesc_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SponceDesc.SelectedIndex == 1 || SponceDesc.SelectedIndex == 3 || SponceDesc.SelectedIndex == 5 || SponceDesc.SelectedIndex == 7 || SponceDesc.SelectedIndex == 9)
                titleFam = " حامل "; 
            else titleFam = " حاملة ";

        }

        private void AddChildren_Click(object sender, EventArgs e)
        {
            string newLine = Environment.NewLine;
            if (ThirdPartyDesc.SelectedIndex == 1 || ThirdPartyDesc.SelectedIndex == 3 || ThirdPartyDesc.SelectedIndex == 5) titleFam3 = " حامل "; else titleFam3 = " حاملة ";
            SponceDesc3[thirdPartyIndex] = ThirdPartyDesc.Text;
            SponcerName3[thirdPartyIndex] = ThirdPartyName.Text;
            SponcePassIqama3[thirdPartyIndex] = ThirdPartyDocType.Text;
            ApplicantIdocNo3[thirdPartyIndex] = ThirdPartyDocNo.Text;
            SponceIssueSource3[thirdPartyIndex] = ThirdPartyDocIssueSource.Text;
            AllFamilyMembers.Text= SponceDesc3[x] + "/ " + SponcerName3[x] + " " + titleFam3 + " " + SponcePassIqama3[x] + " رقم " + ApplicantIdocNo3[x] + " إصدار " + SponceIssueSource3[x];

            thirdPartyIndex++;
            ThirdPartyName.Clear();
            ThirdPartyDesc.SelectedIndex = 0;
            ThirdPartyDocType.SelectedIndex = 0;
            ThirdPartyDocNo.Clear();
            ThirdPartyDocIssueSource.Clear();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SponserCase.Text == "نقل كفالة مقدم الطلب إلى كفالة طرف ثاني")
            {
                panel1.Visible = false;
                ProCase = 1;
            }
            else if (SponserCase.Text == "نقل كفالة طرف ثاني إلى كفالة مقدم الطلب")
            {
                panel1.Visible = false;
                ProCase = 2;
            }
            else
            {
                panel1.Visible = true;
                ProCase = 3;
            }

        }

        private void Review_Click(object sender, EventArgs e)
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

        private void button1_Click_1(object sender, EventArgs e)
        {
            
        }

        private void CreateWordFile(bool printOnly)
        {
            string Thirdpartymember = "";
            
            if (ApplicantSex.CheckState == CheckState.Unchecked)
            {
                labelName.ForeColor = Color.Black;
                labelName.Text = "مقدم الطلب:";
                route = FilesPathIn + "IqrarSponsershipM.docx";
            }
            else if (ApplicantSex.CheckState == CheckState.Checked)
            {
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
                route = FilesPathIn + "IqrarSponsershipF.docx";
            }

            for (x = 0; x < thirdPartyIndex || ProCase != 3; x++)
            {
                string ActiveCopy;
                ActiveCopy = FilesPathOut + ApplicantName.Text + NewFileName + ".docx";
                if (!File.Exists(ActiveCopy))
                { 


                System.IO.File.Copy(route, ActiveCopy);
                object oBMiss = System.Reflection.Missing.Value;
                Word.Application oBMicroWord = new Word.Application();
                object Routseparameter = ActiveCopy;
                Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

                object ParaGreData = "MarkGreData";
                object ParaHijriData = "MarkHijriData";
                object Paraname = "MarkApplicantName";
                object ParaAplicantDoc = "MarkAplicantDoc";
                object ParaAppiIssSource = "MarkAppIssSource";
                object ParaSponcer1 = "MarkSponcer1";
                object ParaSponcer2 = "MarkSponcer2";
                object ParaSponcerDesc = "MarkSponcerDesc";
                object ParaSponcerName = "MarkSponcerName";
                object ParaSponcerDoc = "MarkSponcerDoc";
                object ParavConsul = "MarkViseConsul";
                object ParaAuthorization = "MarkAuthorization";

                Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
                Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
                Word.Range Bookname = oBDoc.Bookmarks.get_Item(ref Paraname).Range;
                Word.Range BookAplicantDoc = oBDoc.Bookmarks.get_Item(ref ParaAplicantDoc).Range;
                Word.Range BookSponcer1 = oBDoc.Bookmarks.get_Item(ref ParaSponcer1).Range;
                Word.Range BookSponcer2 = oBDoc.Bookmarks.get_Item(ref ParaSponcer2).Range;
                Word.Range BookSponcerDesc = oBDoc.Bookmarks.get_Item(ref ParaSponcerDesc).Range;
                Word.Range BookSponcerName = oBDoc.Bookmarks.get_Item(ref ParaSponcerName).Range;
                Word.Range BookSponcerDoc = oBDoc.Bookmarks.get_Item(ref ParaSponcerDoc).Range;
                Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;
                Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;

                BookGreData.Text = GregorianDate.Text;
                BookHijriData.Text = HijriDate.Text;
                Bookname.Text = ApplicantName.Text;
                BookAplicantDoc.Text = ApplicantIdocType.Text + " رقم " + ApplicantIdocNo.Text + " إصدار " + IssuedSource.Text;
                if (ProCase == 1)
                {
                    BookSponcer1.Text = "كفالتي إلى كفالة";
                    BookSponcer2.Text = "";
                }
                else if (ProCase == 2)
                {
                    BookSponcer1.Text = "كفالة";
                    BookSponcer2.Text = "إلى كفالتي";
                }
                else
                {
                    BookSponcer1.Text = "كفالة";
                    BookSponcer2.Text = "إلى كفالة " + SponceDesc.Text + "/ " + SponcerName.Text + " " + titleFam + " " + SponcePassIqama.Text + " رقم " + ApplicantIdocNo.Text + " إصدار " + SponceIssueSource.Text; ;
                }
                if (ProCase == 1 || ProCase == 2)
                {

                    BookSponcerDesc.Text = SponceDesc.Text;
                    BookSponcerName.Text = SponcerName.Text;
                    BookSponcerDoc.Text = titleFam + " " + SponcePassIqama.Text + " رقم " + SponcedocNo.Text + " إصدار " + SponceIssueSource.Text;
                }
                else
                {

                    BookSponcerDesc.Text = SponceDesc3[x];
                    BookSponcerName.Text = SponcerName3[x];
                    BookSponcerDoc.Text = titleFam3 + " " + SponcePassIqama3[x] + " رقم " + ApplicantIdocNo3[x] + " إصدار " + SponceIssueSource3[x];
                    Thirdpartymember = SponceDesc3[x] + "/ " + SponcerName3[x] + " " + titleFam3 + " " + SponcePassIqama3[x] + " رقم " + ApplicantIdocNo3[x] + " إصدار " + SponceIssueSource3[x];

                }
                BookvConsul.Text = AttendViceConsul.Text;
                if (!printOnly)
                {
                    if (ProCase == 3) Save2DataBase(Thirdpartymember);
                    else Save2DataBase("");
                    
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

                object rangeGreData = BookGreData;
                object rangeHijriData = BookHijriData;
                object rangeName = Bookname;
                object rangeAplicantDoc = BookAplicantDoc;
                object rangeSponcer1 = BookSponcer1;
                object rangeSponcer2 = BookSponcer2;
                object rangeSponcerDesc = BookSponcerDesc;
                object rangeSponcerDoc = BookSponcerDoc;
                object rangeSponcerName = BookSponcerName;
                object rangevConsul = BookvConsul;
                object rangeAuthorization = BookAuthorization;

                oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
                oBDoc.Bookmarks.Add("MarkHijiData", ref rangeHijriData);
                oBDoc.Bookmarks.Add("MarkApplicantName", ref rangeName);
                oBDoc.Bookmarks.Add("MarkAplicantDoc", ref rangeAplicantDoc);
                oBDoc.Bookmarks.Add("MarkSponcer1", ref rangeSponcer1);
                oBDoc.Bookmarks.Add("MarkSponcer2", ref rangeSponcer2);
                oBDoc.Bookmarks.Add("MarkSponcerDesc", ref rangeSponcerDesc);
                oBDoc.Bookmarks.Add("MarkSponcerDoc", ref rangeSponcerDoc);
                oBDoc.Bookmarks.Add("MarkSponcerName", ref rangeSponcerName);
                oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);
                oBDoc.Bookmarks.Add("MarkAuthorization", ref rangeAuthorization);


                oBDoc.Activate();

                oBDoc.Save();
                oBMicroWord.Visible = true;
                if (ProCase != 3) break;
                     }
                    else
                    {
                        MessageBox.Show("يرجى حذف الملف الموجودأولاً");
                        btnprintOnly.Enabled = true;
                        btnSavePrint.Enabled = true;
                        thirdPartyIndex = 0;
                    }

                }
            
        }

        private void ApplicantIdocType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ApplicantIdocType.SelectedIndex == 0)
            {
                labeldoctype.Text = "رقم جواز السفر: ";
                ApplicantIdocNo.Text = "P";
            }
            else
            {
                ApplicantIdocNo.Text = ""; labeldoctype.Text = "رقم الإقامة:";
            }
        }

        private void SponcePassIqama_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SponcePassIqama.SelectedIndex == 0)
            {
                label2.Text = "رقم جواز السفر: ";
                SponcedocNo.Text = "P";
            }
            else {
                SponcedocNo.Text = ""; 
                label2.Text = "رقم الإقامة:"; 
            }
        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void AppType_CheckedChanged(object sender, EventArgs e)
        {
            mandoubVisibilty();
        }

        private void dataGridView1_DoubleClick_1(object sender, EventArgs e)
        {

            if (dataGridView1.CurrentRow.Index != -1)
            {
                ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                NextRelId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                ApplicantName.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
                else if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
                ApplicantIdocType.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                ApplicantIdocNo.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                IssuedSource.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();

                SponcerName.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                SponceDesc.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                SponcePassIqama.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                SponcedocNo.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                SponceIssueSource.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                
                GregorianDate.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                HijriDate.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                AllFamilyMembers.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();

                if (dataGridView1.CurrentRow.Cells[16].Value.ToString().ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;
                    Iqrarid.Text = NextRelId;
                }
                else checkedViewed.CheckState = CheckState.Checked;                
                AppType.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
                Employee.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
                if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked; else AppType.CheckState = CheckState.Unchecked;
                if (AppType.CheckState == CheckState.Unchecked)
                {
                    mandoubVisibilty(); 
                    mandoubName.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
                }

                PreRelatedID = dataGridView1.CurrentRow.Cells[20].Value.ToString(); 
                SponserCase.SelectedIndex = Convert.ToInt32(dataGridView1.CurrentRow.Cells[21].Value.ToString())-1;
                Comment.Text = dataGridView1.CurrentRow.Cells[26].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[27].Value.ToString() != "غير مؤرشف")
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

        private void printOnly_Click(object sender, EventArgs e)
        {
            btnSavePrint.Enabled = false;
            btnSavePrint.Text = "جاري المعالجة";
            CreateWordFile(true);
            btnSavePrint.Text = "حفظ وطباعة";
            btnSavePrint.Enabled = true;
        }

        private void SaveOnly_Click(object sender, EventArgs e)
        {
            
        }

        private void btnSavePrint_Click_1(object sender, EventArgs e)
        {
            btnSavePrint.Enabled = false;
            CreateWordFile(false);
            btnSavePrint.Text = "طباعة وحفظ";
            btnSavePrint.Enabled = true;
            Clear_Fields();
        }

        private void SearchDoc_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            SearchFile.Visible = true;
            SearchFile.Text = dlg.FileName;
            if (SearchFile.Text != "") fileloaded = true;
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

        private void printOnly_Click_1(object sender, EventArgs e)
        {
            btnprintOnly.Enabled = false;
            btnprintOnly.Text = "طباعة";
            CreateWordFile(true);
            btnprintOnly.Text = "حفظ وطباعة";
            btnprintOnly.Enabled = true;
            Clear_Fields();
        }

    

        private void SaveOnly_Click_2(object sender, EventArgs e)
        {
            Save2DataBase(AllFamilyMembers.Text);
            Clear_Fields();
        }

        private void ResetAll_Click(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void mandoubName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void OpenFile(int id, int fileNo)
        {
            string query;

            SqlConnection Con = new SqlConnection(DataSource);
            if (fileNo == 1)
            {
                query = "select Data1, Extension1,FileName1 from TableFamilySponApp where ID=@id";
            }
            else
            {
                query = "select Data2, Extension2,FileName2 from TableFamilySponApp where ID=@id";
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

        private void Save2DataBase( string FamilyThirdPart )
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
                SqlCommand sqlCmd = new SqlCommand("FamilySponAddorEdit", sqlCon);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                
                if (btnSavePrint.Text == "طباعة وحفظ")
                {
                    
                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");
                    sqlCmd.Parameters.AddWithValue("@DocID", Iqrarid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", ApplicantIdocType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", ApplicantIdocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", IssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartName", SponcerName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDesc", SponceDesc.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDocType", SponcePassIqama.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDocNo", SponcedocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDocSource", SponceIssueSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@ThirdPart", FamilyThirdPart.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedVisaApp", PreFamilySponAppId.Trim());
                    sqlCmd.Parameters.AddWithValue("@ProCase", ProCase.ToString().Trim());
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
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", ApplicantIdocType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", ApplicantIdocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", IssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartName", SponcerName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDesc", SponceDesc.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDocType", SponcePassIqama.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDocNo", SponcedocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDocSource", SponceIssueSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@ThirdPart", FamilyThirdPart.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@RelatedVisaApp", PreFamilySponAppId.Trim());
                    sqlCmd.Parameters.AddWithValue("@ProCase", ProCase.ToString().Trim());
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

        }

        private void Clear_Fields()
        {
            ApplicantName.Text = AllFamilyMembers.Text = ApplicantName.Text =  IssuedSource.Text = "";
            SponcerName.Text = ""; 
            SponcedocNo.Text = "";
            SponceIssueSource.Text = "";
            ThirdPartyName.Text = "";
            ThirdPartyDocNo.Text = "";
            ThirdPartyDocIssueSource.Text = "";
            AttendViceConsul.SelectedIndex = 2;            
            ApplicantIdocNo.Text = "P";
            SponserCase.SelectedIndex = 0;
            ApplicantIdocType.SelectedIndex = 0;
            panel1.Visible = false;
            ApplicantSex.CheckState = CheckState.Unchecked;
            SponceDesc.SelectedIndex = 0;
            ThirdPartyDesc.SelectedIndex = 0;
            ThirdPartyDocType.SelectedIndex = 0;
            FillDataGridView();
            Iqrarid.Text = IqrarStaticPart + IqrarNumberPart;
            mandoubName.Text = ListSearch.Text = "";
            AppType.CheckState = CheckState.Checked;
            mandoubVisibilty();
            btnprintOnly.Visible = false;
            btnSavePrint.Text = "طباعة وحفظ";
            btnSavePrint.Visible = true;
            SaveOnly.Visible = false;
            Comment.Text = "لا تعليق";
            
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            labelArch.Visible = false;
            ArchivedSt.BackColor = Color.Red;
            SearchFile.Visible = false;
            fileloaded = false;
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            Iqrarid.Text = IqrarStaticPart + IqrarNumberPart;
            thirdPartyIndex = 0;
            Employee.Text = ConsulateEmpName;

        }
        private void FillDataGridView()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("FamilySponViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ApplicantName", ListSearch.Text.Trim());
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            IqrarNumberPart = (dtbl.Rows.Count + 1).ToString();            
            sqlCon.Close();
            NewFileName = IqrarNumberPart + "_05";
        }
    }
}
