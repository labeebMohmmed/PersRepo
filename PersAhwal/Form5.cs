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
        public static string[] AllFamilyMemberList = new string[20];
        public static string AllFamilyList = "";
        public static string[] SponceDesc3 = new string[10];
        public static string[] SponceDoc3 = new string[10];
        public static string[] SponcerName3 = new string[10];
        public static string[] SponcePassIqama3 = new string[10];
        public static string[] ApplicantIdocNo3 = new string[10];
        public static string[] SponceIssueSource3 = new string[10];

        public static string[] DaughterMotheDocSource = new string[10];
        bool newData = true;
        public static string route = "";
        public static int thirdPartyIndex = 0, x = 0, ProCase = 1;
        static string titleFam = "حاملة";
        static string titleFam3 = "حاملة";
        private bool fileloaded = false;
        string Viewed;
        string ConsulateEmpName;

        public static string ModelFileroute = "";
        String IqrarNumberPart;
        static string DataSource;
        int ApplicantID = 0;
        string NewFileName;
        string PreFamilySponAppId = "";
        static public string FamilySupport;
        private string[] FamelyMember = new string[10];
        static public string title = "";
        string PreRelatedID = "", NextRelId = "";
        string FilesPathIn, FilesPathOut;
        string Jobposition;
        int ATVC = 0;
        string[] colIDs = new string[100];
        string AuthTitle = "نائب قنصل";
        public Form5(int Atvc, int currentRow, string EmpName, string dataSource, string filepathIn, string filepathOut, string jobposition, string gregorianDate, string hijriDate)
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
            Jobposition = jobposition;
            FillDataGridView();
            if (currentRow == -1) Clear_Fields();
            else SetFieldswithData(currentRow);
            colIDs[4] = ConsulateEmpName = EmpName;
            if (jobposition.Contains("قنصل"))
                btnEditID.Visible = deleteRow.Visible = true;
            else btnEditID.Visible = deleteRow.Visible = false;
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
        //private void OpenFileDoc(int id, int fileNo)
        //{
        //    string query;

        //    SqlConnection Con = new SqlConnection(DataSource);
        //    if (fileNo == 1)
        //    {
        //        query = "select Data1, Extension1,FileName1 from TableFamilySponApp  where ID=@id";
        //    }
        //    else if (fileNo == 2)
        //    {
        //        query = "select Data2, Extension2,FileName2 from TableFamilySponApp  where ID=@id";
        //    }
        //    else query = "select Data3, Extension3,FileName3 from TableFamilySponApp  where ID=@id";
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
            ApplicantID = Convert.ToInt32(dataGridView1.Rows[Rowindex].Cells[0].Value.ToString());
            NextRelId = dataGridView1.Rows[Rowindex].Cells[1].Value.ToString();
            مقدم_الطلب.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "ذكر") النوع.CheckState = CheckState.Unchecked;
            else if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "أنثى") النوع.CheckState = CheckState.Checked;
            نوع_الهوية.Text = dataGridView1.Rows[Rowindex].Cells[4].Value.ToString();
            رقم_الهوية.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString();
            مكان_الإصدار.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString();

            SponcerName.Text = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString();
            SponceDesc.Text = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString();
            SponcePassIqama.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString();
            SponcedocNo.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString();
            SponceIssueSource.Text = dataGridView1.Rows[Rowindex].Cells[11].Value.ToString();

            التاريخ_الميلادي.Text = dataGridView1.Rows[Rowindex].Cells[12].Value.ToString();
            التاريخ_الهجري.Text = dataGridView1.Rows[Rowindex].Cells[13].Value.ToString();
            AllFamilyMembers.Text = dataGridView1.Rows[Rowindex].Cells[14].Value.ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[15].Value.ToString();

            if (dataGridView1.Rows[Rowindex].Cells[16].Value.ToString().ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                
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
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            التاريخ_الميلادي.Text = DateTime.Now.ToString("MM-dd-yyyy");
            timer1.Enabled=false;
        }

        private void SponceDesc_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SponceDesc.SelectedIndex == 1 || SponceDesc.SelectedIndex == 3 || SponceDesc.SelectedIndex == 5 || SponceDesc.SelectedIndex == 7 || SponceDesc.SelectedIndex == 9)
                titleFam = " حامل ";
            else titleFam = " حاملة ";

        }

        private void AddChildren_Click_1(object sender, EventArgs e)
        {
            string newLine = Environment.NewLine;
            if (ThirdPartyDesc.SelectedIndex == 1 || ThirdPartyDesc.SelectedIndex == 3 || ThirdPartyDesc.SelectedIndex == 5 || ThirdPartyDesc.SelectedIndex == 7 || ThirdPartyDesc.SelectedIndex == 9) titleFam3 = " حامل "; else titleFam3 = " حاملة ";

            AllFamilyMemberList[x] = ThirdPartyDesc.Text + "/" + ThirdPartyName.Text + "/" + titleFam3 + "/" + ThirdPartyDocType.Text + "/" + ThirdPartyDocNo.Text + "/" + ThirdPartyDocIssueSource.Text;
            if (thirdPartyIndex == 0) AllFamilyList = AllFamilyMemberList[thirdPartyIndex];
            else AllFamilyList = AllFamilyList + "*" + AllFamilyMemberList[thirdPartyIndex];
            AllFamilyMembers.Text = Environment.NewLine + AllFamilyMembers.Text + AllFamilyMemberList[thirdPartyIndex];
            thirdPartyIndex++;
            ThirdPartyName.Clear();
            ThirdPartyDesc.SelectedIndex = 0;
            ThirdPartyDocType.SelectedIndex = 0;
            ThirdPartyDocNo.Clear();
            ThirdPartyDocIssueSource.Clear();
        }

        private void SponserCase_SelectedIndexChanged(object sender, EventArgs e)
        {

            ProCase = SponserCase.SelectedIndex + 1;
            if (ProCase >= 3) panel1.Visible = true;
            else panel1.Visible = false;
            if (ProCase == 3) label15.Text = "بيانات المراد نقل كفالتهم";
            else if (ProCase == 4) label15.Text = "بيانات المراد استقدامهم";
            //if (SponserCase.Text == "نقل كفالة مقدم الطلب إلى كفالة طرف ثاني")
            //{
            //    panel1.Visible = false;
            //    ProCase = 1;
            //}
            //else if (SponserCase.Text == "نقل كفالة طرف ثاني إلى كفالة مقدم الطلب")
            //{
            //    panel1.Visible = false;
            //    ProCase = 2;
            //}
            //else if(SponceDesc.Text == "نقل كفالة أحد مكفولي مقدم الطلب إلى كفالة طرف ثاني")
            //{
            //    panel1.Visible = true;
            //    ProCase = 3;
            //}

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
                mandoubLabel.Visible = panel3.Visible = false;
            }
            else
            {
                AppType.Text = "عن طريق أحد مندوبي القنصلية";
                mandoubName.Visible = true;
                mandoubLabel.Visible = panel3.Visible = true;
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
            if (ProCase == 3 && thirdPartyIndex == 0)
            {
                MessageBox.Show("يرجى إضافة المكفولين أولا");
                return;

            }
            else if (ProCase == 4 && thirdPartyIndex == 0)
            {
                MessageBox.Show("يرجى إضافة المستقدمين أولا");
                return;

            }
            string Thirdpartymember = "";
            string ReportName = DateTime.Now.ToString("mmss");
            if (النوع.CheckState == CheckState.Unchecked)
            {
                labelName.ForeColor = Color.Black;
                labelName.Text = "مقدم الطلب:";
                route = FilesPathIn + "IqrarSponsershipM.docx";
            }
            else if (النوع.CheckState == CheckState.Checked)
            {
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
                route = FilesPathIn + "IqrarSponsershipF.docx";
            }

            if (!printOnly)
            {
                if (ProCase == 3) Save2DataBase(AllFamilyList);
                else Save2DataBase("");
            }
            for (x = 0; x < thirdPartyIndex || ProCase != 3; x++)
            {

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
                    object Paraname = "MarkApplicantName";
                    object Paraname2 = "MarkApplicantName2";
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
                    Word.Range BookIqrarNo = oBDoc.Bookmarks.get_Item(ref ParaIqrarNo).Range;
                    Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
                    Word.Range Bookname = oBDoc.Bookmarks.get_Item(ref Paraname).Range;
                    Word.Range Bookname2 = oBDoc.Bookmarks.get_Item(ref Paraname2).Range;
                    Word.Range BookAplicantDoc = oBDoc.Bookmarks.get_Item(ref ParaAplicantDoc).Range;
                    Word.Range BookSponcer1 = oBDoc.Bookmarks.get_Item(ref ParaSponcer1).Range;
                    Word.Range BookSponcer2 = oBDoc.Bookmarks.get_Item(ref ParaSponcer2).Range;
                    Word.Range BookSponcerDesc = oBDoc.Bookmarks.get_Item(ref ParaSponcerDesc).Range;
                    Word.Range BookSponcerName = oBDoc.Bookmarks.get_Item(ref ParaSponcerName).Range;
                    Word.Range BookSponcerDoc = oBDoc.Bookmarks.get_Item(ref ParaSponcerDoc).Range;
                    Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;
                    Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;

                    BookIqrarNo.Text = colIDs[0] = Iqrarid.Text;
                    BookGreData.Text = التاريخ_الميلادي_off.Text;
                    colIDs[2] = التاريخ_الميلادي.Text;
                    BookHijriData.Text = التاريخ_الهجري.Text;
                    Bookname2.Text = Bookname.Text = colIDs[3] = مقدم_الطلب.Text;
                    colIDs[5] = AppType.Text;
                    colIDs[6] = mandoubName.Text;
                    BookAplicantDoc.Text = نوع_الهوية.Text + " رقم " + رقم_الهوية.Text + " إصدار " + مكان_الإصدار.Text;
                    if (SponceDesc.SelectedIndex == 1 || SponceDesc.SelectedIndex == 3 || SponceDesc.SelectedIndex == 5 || SponceDesc.SelectedIndex == 7 || SponceDesc.SelectedIndex == 9) titleFam = "حامل";
                    else titleFam = "حاملة";

                    if (ProCase == 1)
                    {
                        BookSponcer1.Text = " نقل كفالتي إلى كفالة";
                        BookSponcer2.Text = "";
                    }
                    else if (ProCase == 2)
                    {
                        BookSponcer1.Text = " نقل كفالة";
                        BookSponcer2.Text = "إلى كفالتي";
                    }
                    else if (ProCase == 3)
                    {
                        BookSponcer1.Text = " نقل كفالة";

                        if (SponcePassIqama.Text != "") BookSponcer2.Text = "إلى كفالة " + SponceDesc.Text + "/ " + SponcerName.Text + " " + titleFam + " " + SponcePassIqama.Text + " رقم " + SponcedocNo.Text + " إصدار " + SponceIssueSource.Text;
                        else BookSponcer2.Text = "إلى كفالة " + SponceDesc.Text + "/ " + SponcerName.Text;

                    }

                    else if (ProCase == 4)
                    {
                        BookSponcer1.Text = " استقدام";
                        if (SponcePassIqama.Text != "") BookSponcer2.Text = "على كفالة " + SponceDesc.Text + "/ " + SponcerName.Text + " " + titleFam + " " + SponcePassIqama.Text + " رقم " + SponcedocNo.Text + " إصدار " + SponceIssueSource.Text;
                        else BookSponcer2.Text = "على كفالة " + SponceDesc.Text + "/ " + SponcerName.Text;
                    }
                    if (ProCase == 1 || ProCase == 2)
                    {

                        BookSponcerDesc.Text = SponceDesc.Text;
                        BookSponcerName.Text = SponcerName.Text;
                        BookSponcerDoc.Text = titleFam + " " + SponcePassIqama.Text + " رقم " + SponcedocNo.Text + " إصدار " + SponceIssueSource.Text;
                    }
                    else
                    {
                        string[] familylist = new string[5];
                        familylist = AllFamilyMemberList[x].Split('/');
                        BookSponcerDesc.Text = familylist[0];
                        BookSponcerName.Text = " " + familylist[1];
                        if (familylist[4] != "") BookSponcerDoc.Text = familylist[2] + " " + familylist[3] + " رقم " + familylist[4] + " إصدار " + familylist[5];
                        else BookSponcerDoc.Text = "";
                        Thirdpartymember = AllFamilyMemberList[x];
                    }
                    BookvConsul.Text = AttendViceConsul.Text + Environment.NewLine + AuthTitle;


                    if (AppType.CheckState == CheckState.Checked)
                    {
                        if (النوع.CheckState == CheckState.Unchecked) BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text +" "+ AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد حضر للقنصلية ووقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                        if (النوع.CheckState == CheckState.Checked) BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + AuthTitle+"  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد حضرت للقنصلية ووقعت بتوقيعها على هذا الإقرار بعد تلاوته عليها وبعد أن فهمت مضمونه ومحتواه. ";
                    }
                    else
                    {
                        if (النوع.CheckState == CheckState.Unchecked)
                            BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد وقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                        if (النوع.CheckState == CheckState.Checked)
                            BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد وقعت بتوقيعها على هذا الإقرار بعد تلاوته عليها وبعد أن فهمت مضمونه ومحتواه. ";

                    }

                    object rangeIqrarNo = BookIqrarNo;
                    object rangeGreData = BookGreData;
                    object rangeHijriData = BookHijriData;
                    object rangeName = Bookname;
                    object rangeName2 = Bookname2;
                    object rangeAplicantDoc = BookAplicantDoc;
                    object rangeSponcer1 = BookSponcer1;
                    object rangeSponcer2 = BookSponcer2;
                    object rangeSponcerDesc = BookSponcerDesc;
                    object rangeSponcerDoc = BookSponcerDoc;
                    object rangeSponcerName = BookSponcerName;
                    object rangevConsul = BookvConsul;
                    object rangeAuthorization = BookAuthorization;

                    oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);
                    oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
                    oBDoc.Bookmarks.Add("MarkHijiData", ref rangeHijriData);
                    oBDoc.Bookmarks.Add("MarkApplicantName", ref rangeName);
                    oBDoc.Bookmarks.Add("MarkApplicantName2", ref rangeName2);
                    oBDoc.Bookmarks.Add("MarkAplicantDoc", ref rangeAplicantDoc);
                    oBDoc.Bookmarks.Add("MarkSponcer1", ref rangeSponcer1);
                    oBDoc.Bookmarks.Add("MarkSponcer2", ref rangeSponcer2);
                    oBDoc.Bookmarks.Add("MarkSponcerDesc", ref rangeSponcerDesc);
                    oBDoc.Bookmarks.Add("MarkSponcerDoc", ref rangeSponcerDoc);
                    oBDoc.Bookmarks.Add("MarkSponcerName", ref rangeSponcerName);
                    oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);
                    oBDoc.Bookmarks.Add("MarkAuthorization", ref rangeAuthorization);

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
                    //string pdfouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
                    oBDoc.SaveAs2(docxouput);
                    //oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
                    oBDoc.Close(false, oBMiss);
                    oBMicroWord.Quit(false, false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
                    System.Diagnostics.Process.Start(docxouput);
                    object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;

                    if (ProCase != 3) break;
                }
                else
                {
                    MessageBox.Show("يرجى حذف الملف الموجودأولاً");                    
                    thirdPartyIndex = 0;
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


        private void ApplicantIdocType_SelectedIndexChanged(object sender, EventArgs e)
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

        private void SponcePassIqama_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SponcePassIqama.SelectedIndex == 0)
            {
                label2.Text = "رقم جواز السفر: ";
                SponcedocNo.Text = "P";
            }
            else
            {
                SponcedocNo.Text = "";
                label2.Text = "رقم الإقامة:";
            }
        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void AppType_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void dataGridView1_DoubleClick_1(object sender, EventArgs e)
        {


        }

        private void SaveOnly_Click_1(object sender, EventArgs e)
        {
            btnSavePrint.Text = "جاري المعالجة";
            CreateWordFile(true);
            btnSavePrint.Enabled = true;
            this.Close();
        }

       

        private void btnSavePrint_Click(object sender, EventArgs e)
        {
            getTitle(DataSource, AttendViceConsul.Text); 
            
            CreateWordFile(false);
            
            var selectedOption = MessageBox.Show("", "حفظ وإنهاء المعاملة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                Clear_Fields();
            }
            this.Close();
        }

        private void SearchDoc_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            var selectRows = dataGridView1.SelectedRows;
            foreach (var row in selectRows)
            {
                int id = (int)((DataGridViewRow)row).Cells[0].Value;
                //OpenFile(id, 2);
                FillDatafromGenArch("data2", colIDs[1], "TableFamilySponApp");
            }
            if (ApplicantID != 0) FillDatafromGenArch("data2", colIDs[1], "TableFamilySponApp");
            //ApplicantID = 0;
        }

        private void btnprintOnly_Click(object sender, EventArgs e)
        {
            if (!checkGender(mainpanel, "مقدم_الطلب", "النوع"))
            {
                return;
            }
            else addNewAppNameInfo(مقدم_الطلب);
            if(SponcerName.Text != "") addNewAppNameInfo1(SponcerName);
            if (ThirdPartyName.Text != "") addNewAppNameInfo2(ThirdPartyName);
            CreateWordFile(true);
            this.Close();
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
        }private void addNewAppNameInfo1(TextBox textName)
        {
            //SponcedocNo, SponcePassIqama, , SponcerName.Text)
            string query = "insert into TableGenNames ([الاسم], رقم_الهوية,نوع_الهوية,مكان_الإصدار) values (@col1,@col2,@col6,@col7) ;SELECT @@IDENTITY as lastid";
            string id = checkExist(textName.Text);
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
            sqlCommand.Parameters.AddWithValue("@col1", SponcerName.Text);
            sqlCommand.Parameters.AddWithValue("@col2", SponcedocNo.Text);
            
            sqlCommand.Parameters.AddWithValue("@col6", SponcePassIqama.Text);
            sqlCommand.Parameters.AddWithValue("@col7", SponceIssueSource.Text);

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
                MessageBox.Show("addNewAppNameInfo1");
            }
        }private void addNewAppNameInfo2(TextBox textName)
        {

            string query = "insert into TableGenNames ([الاسم], رقم_الهوية,نوع_الهوية,مكان_الإصدار) values (@col1,@col2,@col6,@col7) ;SELECT @@IDENTITY as lastid";
            string id = checkExist(textName.Text);
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
            sqlCommand.Parameters.AddWithValue("@col1", ThirdPartyName.Text);
            sqlCommand.Parameters.AddWithValue("@col2", ThirdPartyDocNo.Text);
            
            sqlCommand.Parameters.AddWithValue("@col6", ThirdPartyDocType.Text);
            sqlCommand.Parameters.AddWithValue("@col7", ThirdPartyDocIssueSource.Text);

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
                MessageBox.Show("addNewAppNameInfo1");
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

        private void SaveOnly_Click_2(object sender, EventArgs e)
        {
            Save2DataBase(AllFamilyMembers.Text);
            this.Close();
            Clear_Fields();
        }

        private void ResetAll_Click_1(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void mandoubName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void deleteRow_Click_1(object sender, EventArgs e)
        {
            deleteRowsData(ApplicantID, "TableFamilySponApp", DataSource);
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

        private string loadRerNo(int id)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT DocID from TableFamilySponApp where ID=@ID", sqlCon);
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
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from TableFamilySponApp order by ID desc", sqlCon);
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                gridFill = true;
                dataGridView1.Visible = false;
                mainpanel.Visible = true;
                Iqrarid.Text = NextRelId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                colIDs[1] =  dataGridView1.CurrentRow.Cells[0].Value.ToString();
                colIDs[0] = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                AppType.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
                if (AppType.Text == "حضور مباشرة إلى القنصلية") 
                    AppType.CheckState = CheckState.Checked; 
                else AppType.CheckState = CheckState.Unchecked;
                //MessageBox.Show(AppType.Text);
                if (AppType.CheckState == CheckState.Unchecked)
                {
                    mandoubVisibilty();
                    mandoubName.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
                }

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
                    colIDs[7] = "new";
                   ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                    SponserCase.SelectedIndex = Convert.ToInt32(dataGridView1.CurrentRow.Cells[21].Value.ToString()) - 1;
                    //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 1);
                    FillDatafromGenArch("data1", colIDs[1], "TableFamilySponApp");
                    newData = false;
                    if (Jobposition.Contains("قنصل")) deleteRow.Visible = true;
                    gridFill = false;
                    return;
                }
                gridFill = false;
                colIDs[7] = "old";
                ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                if (Jobposition.Contains("قنصل")) deleteRow.Visible = true;
                
                مقدم_الطلب.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "ذكر") النوع.CheckState = CheckState.Unchecked;
                else if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "أنثى") النوع.CheckState = CheckState.Checked;
                نوع_الهوية.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                رقم_الهوية.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                مكان_الإصدار.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();

                SponcerName.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                SponceDesc.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                SponcePassIqama.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                SponcedocNo.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                SponceIssueSource.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();

                التاريخ_الميلادي.Text = "";
                التاريخ_الميلادي.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                التاريخ_الهجري.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                AllFamilyList = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                if (AllFamilyList != "")
                {
                    AllFamilyMemberList = AllFamilyList.Split('*');
                    thirdPartyIndex = AllFamilyMemberList.Length;
                    AllFamilyMembers.Text = AllFamilyList.Replace("*", Environment.NewLine);
                    AllFamilyMembers.Text = AllFamilyMembers.Text.Replace("/", " ");


                }
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();

                if (dataGridView1.CurrentRow.Cells[16].Value.ToString().ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;
                    
                }
                else checkedViewed.CheckState = CheckState.Checked;
                
                Employee.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
                
                PreRelatedID = dataGridView1.CurrentRow.Cells[20].Value.ToString();
                SponserCase.SelectedIndex = Convert.ToInt32(dataGridView1.CurrentRow.Cells[21].Value.ToString()) - 1;
                Comment.Text = dataGridView1.CurrentRow.Cells[22].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[23].Value.ToString() != "غير مؤرشف")
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
            }
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            
            autoCompleteTextBox1(مقدم_الطلب, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox1(SponcerName, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox1(ThirdPartyName, DataSource, "الاسم", "TableGenNames");
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

        private void ApplicantSex_CheckedChanged_1(object sender, EventArgs e)
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

        private void AppType_CheckedChanged_1(object sender, EventArgs e)
        {
            mandoubVisibilty();
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
                if (dataGridView1.Rows[i].Cells[27].Value.ToString() == "مؤرشف نهائي") dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;

                //else dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.;

            }
            //
        }

        private void Form5_Load_1(object sender, EventArgs e)
        {
            fileComboBox(mandoubName, DataSource, "MandoubNames", "TableListCombo"); 
            fileComboBox(نوع_الهوية, DataSource, "DocType", "TableListCombo");
            fileComboBox(SponcePassIqama, DataSource, "DocType", "TableListCombo");
            fileComboBox(ThirdPartyDocType, DataSource, "DocType", "TableListCombo");
            autoCompleteTextBox(مكان_الإصدار, DataSource, "SDNIssueSource", "TableListCombo");
            autoCompleteTextBox(SponceIssueSource, DataSource, "SDNIssueSource", "TableListCombo");
            autoCompleteTextBox(ThirdPartyDocIssueSource, DataSource, "SDNIssueSource", "TableListCombo");
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
                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    if (!string.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                    {
                        autoComplete.Add(dataRow[comlumnName].ToString());
                    }
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            var selectRows = dataGridView1.SelectedRows;
            foreach (var row in selectRows)
            {
                int id = (int)((DataGridViewRow)row).Cells[0].Value;
                //OpenFile(id, 1);
                FillDatafromGenArch("data1", colIDs[1], "TableFamilySponApp");
            }
            if (ApplicantID != 0) FillDatafromGenArch("data1", colIDs[1], "TableFamilySponApp"); //OpenFile(ApplicantID, 1);
            ApplicantID = 0;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var selectRows = dataGridView1.SelectedRows;
            foreach (var row in selectRows)
            {
                int id = (int)((DataGridViewRow)row).Cells[0].Value;
                //OpenFile(id, 2);
                FillDatafromGenArch("data2", colIDs[1], "TableFamilySponApp");
            }
            if (ApplicantID != 0) FillDatafromGenArch("data2", colIDs[1], "TableFamilySponApp"); //OpenFile(ApplicantID,2);
            //ApplicantID = 0;
        }

        void FillDatafromGenArch(string doc, string id, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource.Replace("AhwalDataBase", "ArchFilesDB"));
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
        private void Form5_FormClosed(object sender, FormClosedEventArgs e)
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

        private void التاريخ_الميلادي_TextChanged(object sender, EventArgs e)
        {
            try
            {
                التاريخ_الميلادي_off.Text = التاريخ_الميلادي.Text.Split('-')[1] + " - " + التاريخ_الميلادي.Text.Split('-')[0] + " - " + التاريخ_الميلادي.Text.Split('-')[2];
            }
            catch (Exception ex) { }
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
        public void getID(TextBox رقم_الهوية_1, ComboBox نوع_الهوية_1, TextBox مكان_الإصدار_1   , string name)
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
            foreach (DataRow row in dtbl.Rows)
            {
                رقم_الهوية_1.Text = row["رقم_الهوية"].ToString();
                نوع_الهوية_1.Text = row["نوع_الهوية"].ToString();
                مكان_الإصدار_1.Text = row["مكان_الإصدار"].ToString();
            }
        }

        private void SponcerName_TextChanged(object sender, EventArgs e)
        {
            getID(SponcedocNo, SponcePassIqama, SponceIssueSource, SponcerName.Text);
        }

        private void ThirdPartyName_TextChanged(object sender, EventArgs e)
        {
            getID(ThirdPartyDocNo, ThirdPartyDocType, ThirdPartyDocIssueSource, ThirdPartyName.Text);
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
        private void btnEditID_Click(object sender, EventArgs e)
        {
            if (btnEditID.Text == "إجراء")
            {
                btnEditID.Text = "تعديل";
                SqlConnection sqlCon = new SqlConnection(DataSource);
                SqlCommand sqlCmd = new SqlCommand("update TableFamilySponApp SET DocID = @DocID WHERE ID = @ID", sqlCon);
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

        //private void OpenFile(int id, int fileNo)
        //{
        //    string query;

        //    SqlConnection Con = new SqlConnection(DataSource);
        //    if (fileNo == 1)
        //    {
        //        query = "select Data1, Extension1,FileName1 from TableFamilySponApp where ID=@id";
        //    }
        //    else
        //    {
        //        query = "select Data2, Extension2,FileName2 from TableFamilySponApp where ID=@id";
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

        private void Save2DataBase(string FamilyThirdPart)
        {
            if (!checkGender(mainpanel, "مقدم_الطلب", "النوع"))
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
                SqlCommand sqlCmd = new SqlCommand("FamilySponAddorEdit", sqlCon);
                sqlCmd.CommandType = CommandType.StoredProcedure;

                if (btnSavePrint.Text == "طباعة وحفظ" && newData )
                {

                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");
                    sqlCmd.Parameters.AddWithValue("@DocID", Iqrarid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", مقدم_الطلب.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", نوع_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", رقم_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", مكان_الإصدار.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartName", SponcerName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDesc", SponceDesc.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDocType", SponcePassIqama.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDocNo", SponcedocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDocSource", SponceIssueSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
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
                    sqlCmd.Parameters.AddWithValue("@DocID", Iqrarid.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", مقدم_الطلب.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", نوع_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", رقم_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", مكان_الإصدار.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartName", SponcerName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDesc", SponceDesc.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDocType", SponcePassIqama.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDocNo", SponcedocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@SecPartDocSource", SponceIssueSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
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
            مقدم_الطلب.Text = AllFamilyMembers.Text = مقدم_الطلب.Text = مكان_الإصدار.Text = "";
            SponcerName.Text = "";
            SponcedocNo.Text = "";
            SponceIssueSource.Text = "";
            ThirdPartyName.Text = "";
            ThirdPartyDocNo.Text = "";
            ThirdPartyDocIssueSource.Text = "";
            AttendViceConsul.SelectedIndex = 2;
            رقم_الهوية.Text = "P";
            SponserCase.SelectedIndex = 0;
            نوع_الهوية.SelectedIndex = 0;
            panel1.Visible = false;
            النوع.CheckState = CheckState.Unchecked;
            SponceDesc.SelectedIndex = 0;
            ThirdPartyDesc.SelectedIndex = 0;
            ThirdPartyDocType.SelectedIndex = 0;
            FillDataGridView();
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            التاريخ_الميلادي.Text = DateTime.Now.ToString("dd-MM-yyyy");

            mandoubName.Text = ListSearch.Text = "";
            AppType.CheckState = CheckState.Checked;
            mandoubVisibilty();
            
            Comment.Text = "لا تعليق";
            newData = true;
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            labelArch.Visible = false;
            ArchivedSt.BackColor = Color.Red;
            SearchFile.Visible = false;
            fileloaded = false;
            dataGridView1.Visible = true;
            mainpanel.Visible = false;
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
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            IqrarNumberPart = loadRerNo(loadIDNo());
            sqlCon.Close();
            NewFileName = IqrarNumberPart + "_05";
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 250;
        }
    }
}
