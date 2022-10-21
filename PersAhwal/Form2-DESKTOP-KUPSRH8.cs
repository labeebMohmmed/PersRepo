using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Threading;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;

namespace PersAhwal
{
    public partial class Form2 : Form
    {
        int i = 0;
        string[] ChildFromDataBase;
        private static int childindex = 0;
        static public bool ApplicantSexStatus = false;
        public static string[] ChildName = new string[10];
        static bool[] Son_Daughter = new bool[10];
        static string ChildernDescription="";
        static string ChildDataBase="", Mentioned="";
        string Viewed;
        string ConsulateEmpName;
        public static string ModelFileroute = "";
        String IqrarStaticPart = "ق س ج/160/02/";
        String IqrarNumberPart;
        static string strDataSource;
        
        int ApplicantID = 0;
        bool fileloaded = false;
        string PreAppId = "", CurrentIqrarId = "", PreRelatedID="",NextRelId="";
        string  CurrentFileName;
        string FilesPathIn, FilesPathOut;
        public Form2(int currentRow, string EmpName, string dataSource, string filepathIn, string filepathOut)
        {
            InitializeComponent();
            timer1.Enabled = true;
            timer2.Enabled = true;
            strDataSource = dataSource;
            FilesPathIn = filepathIn;
            FilesPathOut = filepathOut;
            ConsulateEmpName = EmpName;
            FillDataGridView();
            if (currentRow == -1) Clear_Fields();
            else SetFieldswithData(currentRow);
        }

        private void Clear_Fields()
        {
            ApplicantName.Text = IssuedSource.Text = TravellingPurpo.Text = ApplicantIdoc.Text = ChildernDescription = PassIqama.Text = ChildernDescription = ChildNameDesView.Text = ChildrenName.Text = "";
            AttendViceConsul.SelectedIndex = 2;
            TravelDestin.SelectedIndex = 0;
            TravellerDescrib.SelectedIndex = 0;
            EmbassySource.SelectedIndex = 0;
            PassIqama.SelectedIndex = 0;
            familyJob.SelectedIndex = 0;
            familyJob.Visible = labeljob.Visible = false;
            ApplicantIdoc.Text = "P";
            Comment.Text = "لا تعليق";
            PassIqama.Text = "جواز سفر";
            childindex = 0;
            mandoubName.Text = Search.Text = "";
            AppType.CheckState = CheckState.Checked;
            mandoubVisibilty();
            printOnly.Visible = false;
            btnSavePrint.Text = "حفظ وطباعة";
            btnSavePrint.Enabled= true;
            Employee.Text = ConsulateEmpName;
            FillDataGridView();
            IqrarNo.Text = IqrarStaticPart + IqrarNumberPart;
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            labelArch.Visible = false;
            ArchivedSt.BackColor = Color.Red;
        }
        private void SetFieldswithData(int Rowindex)
        {
            Rowindex--;
            ApplicantID = Convert.ToInt32(dataGridView1.Rows[Rowindex].Cells[0].Value.ToString());
            CurrentIqrarId = dataGridView1.Rows[Rowindex].Cells[1].Value.ToString();
            ApplicantName.Text = dataGridView1.Rows[Rowindex].Cells[2].Value.ToString().ToString();
            if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
            else if (dataGridView1.Rows[Rowindex].Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
            PassIqama.Text = dataGridView1.Rows[Rowindex].Cells[4].Value.ToString().ToString();
            ApplicantIdoc.Text = dataGridView1.Rows[Rowindex].Cells[5].Value.ToString().ToString();
            IssuedSource.Text = dataGridView1.Rows[Rowindex].Cells[6].Value.ToString().ToString();
            ChildernDescription = dataGridView1.Rows[Rowindex].Cells[7].Value.ToString().ToString();
            string ChildrenList = dataGridView1.Rows[Rowindex].Cells[8].Value.ToString().ToString();
            if (ChildrenList.Contains("_"))
            {
                ChildFromDataBase = ChildrenList.Split('_');
                childindex = ChildFromDataBase.Length;
            }
            else
            {
                childindex = 1;
                ChildDataBase = ChildrenList;
            }

            for (int i = 1; i < childindex; i++)
            {
                ChildDataBase = ChildDataBase + "_" + ChildFromDataBase[i];
            }
            ChildNameDesView.Text = ChildernDescription + "/ " + ChildDataBase;
            textBox1.Text = ChildernDescription;
            textBox2.Text = ChildDataBase;
            EmbassySource.Text = dataGridView1.Rows[Rowindex].Cells[9].Value.ToString();
            GregorianDate.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString();
            HijriDate.Text = dataGridView1.Rows[Rowindex].Cells[11].Value.ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[12].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[16].Value.ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                IqrarNo.Text = NextRelId;
            }
            else checkedViewed.CheckState = CheckState.Checked;
            if (checkedViewed.CheckState == CheckState.Checked)
            {
                PreAppId = CurrentIqrarId;
            }
            else
            {
                PreAppId = "";
                IqrarNo.Text = CurrentIqrarId;
            }
            TravelDestin.Text = dataGridView1.Rows[Rowindex].Cells[14].Value.ToString().ToString();
            TravellingPurpo.Text = dataGridView1.Rows[Rowindex].Cells[15].Value.ToString().ToString();
            AppType.Text = dataGridView1.Rows[Rowindex].Cells[16].Value.ToString();
            if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked;
            else AppType.CheckState = CheckState.Unchecked;
            ConsulateEmployee.Text = dataGridView1.Rows[Rowindex].Cells[17].Value.ToString();
            if (AppType.CheckState == CheckState.Unchecked)
            {
                mandoubVisibilty(); mandoubName.Text = dataGridView1.Rows[Rowindex].Cells[18].Value.ToString();
            }
            Comment.Text = dataGridView1.Rows[Rowindex].Cells[24].Value.ToString();
            PreRelatedID = dataGridView1.Rows[Rowindex].Cells[19].Value.ToString();
            Comment.Text = dataGridView1.Rows[Rowindex].Cells[24].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[25].Value.ToString() != "غير مؤرشف")
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

            void FillDataGridView()
        {
            SqlConnection sqlCon = new SqlConnection(strDataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("TravViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ApplicantName", Search.Text.Trim());
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            CurrentFileName = IqrarNumberPart + "_02";
            IqrarNumberPart = (dtbl.Rows.Count + 1).ToString();
            sqlCon.Close();
        }
        private void Review_Click(object sender, EventArgs e)
        {
            
        }

        private void CreateWordFile()
        {
            ModelFileroute = FilesPathIn + "Igrar_TravM.docx";

            if (ApplicantSex.CheckState == CheckState.Checked)
            {
                ApplicantSexStatus = false;
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
                ModelFileroute = FilesPathIn + "Igrar_TravF.docx";
            }

            string CurrentCopy = FilesPathOut + ApplicantName.Text + CurrentFileName + ".docx";
            if (!File.Exists(CurrentCopy))
            {
            
                System.IO.File.Copy(ModelFileroute, CurrentCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();

            object objCurrentCopy = CurrentCopy;

            Word.Document oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);

            object ParaIqrarNo = "MarkIqrarNo";
            object ParaHijriData = "MarkHijriData";
            object ParaGreData = "MarkGreData";
            object Paraname1 = "MarkApplicantName";
            object Paraname2 = "MarkApplicantName2";
            object Paraigama = "MarkAppliigamaNo";
            object ParavConsul = "MarkViseConsul";            
            object ParaChildren = "MarkChildrenName";
            object ParaAppiIssSource = "MarkAppIssSource";
            object ParaMention = "MarkMention";
            object ParaMarkEmbassy = "MarkEmbassy";
            object ParaCountDestin = "MarkCountDestin";
            object ParaCountryDesc = "MarkCountryDesc";
            object ParaTravelPurpose = "MarkTravelPurpose";
            object ParaPassIqama = "MarkPassIqama";
            object ParaAuthorization = "MarkAuthorization";

            Word.Range BookIqrarNo = oBDoc.Bookmarks.get_Item(ref ParaIqrarNo).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;

            Word.Range Bookname1 = oBDoc.Bookmarks.get_Item(ref Paraname1).Range;
            Word.Range Bookname2 = oBDoc.Bookmarks.get_Item(ref Paraname2).Range;
            Word.Range Bookigama = oBDoc.Bookmarks.get_Item(ref Paraigama).Range;
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;           
            Word.Range BookChildren = oBDoc.Bookmarks.get_Item(ref ParaChildren).Range;
            Word.Range BookAppiIssSource = oBDoc.Bookmarks.get_Item(ref ParaAppiIssSource).Range;
            Word.Range BookMention = oBDoc.Bookmarks.get_Item(ref ParaMention).Range;
            Word.Range BookMarkEmbassy = oBDoc.Bookmarks.get_Item(ref ParaMarkEmbassy).Range;
            Word.Range BookCountDestin = oBDoc.Bookmarks.get_Item(ref ParaCountDestin).Range;
            Word.Range BookCountryDesc = oBDoc.Bookmarks.get_Item(ref ParaCountryDesc).Range;
            Word.Range BookTravelPurpose = oBDoc.Bookmarks.get_Item(ref ParaTravelPurpose).Range;
            Word.Range BookPassIqama = oBDoc.Bookmarks.get_Item(ref ParaPassIqama).Range;
            Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;

            BookIqrarNo.Text = IqrarNo.Text;
            BookHijriData.Text = HijriDate.Text;
            BookGreData.Text = GregorianDate.Text;
            BookAppiIssSource.Text = IssuedSource.Text;
            Bookname1.Text = Bookname2.Text = ApplicantName.Text;
            Bookigama.Text = ApplicantIdoc.Text;
            BookvConsul.Text = AttendViceConsul.Text;

            BookMarkEmbassy.Text = EmbassySource.Text;

            if (TravellerDescrib.Text == "ابناء فقط")
            {
                BookChildren.Text = ChildernDescription + "/" + ChildDataBase;
                BookMention.Text = Mentioned;

            }
            else if (TravellerDescrib.Text == "ابناء برفقة مرافق غير الزوجة")
            {
                BookChildren.Text = ChildernDescription + "/" + ChildDataBase + " " + "برفقة " + TravellerAttenDescrib.Text + "/ " + TravellerAttenName.Text;                
                BookMention.Text = Mentioned;

            }
            else if (TravellerDescrib.Text == "زوجة فقط")
            {
                BookChildren.Text = "زوجتي /" + WifeName.Text;
                BookMention.Text = "للمذكورة";
            }
            else if (TravellerDescrib.Text == "زوجة وابناء")
            {
                BookChildren.Text = ChildernDescription + "/" + ChildDataBase + " " + "برفقة زوجتي " + "/" + WifeName.Text;
                if (Mentioned == "ابنتي") {
                    BookMention.Text = "للمذكورتين";
                }
                else BookMention.Text = "للمذكورين";

            }

            BookCountDestin.Text = TravelDestin.Text;
            
            if (TravelDestin.Text == "المملكة العربية السعودية")
                BookCountryDesc.Text = "قدوم";
            else
                BookCountryDesc.Text = "سفر";
            if (TravellingPurpo.Text == "العمل")
            {
                BookTravelPurpose.Text = "العمل بمهنة" + " " + familyJob.Text;
            }
            else
            {
                BookTravelPurpose.Text = TravellingPurpo.Text;
            }

            BookPassIqama.Text = PassIqama.Text;
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



            object rangeHijriDate = BookHijriData;
            object rangeGreData = BookGreData;
            object rangeIqrarNo = BookIqrarNo;

            object rangeName1 = Bookname1;
            object rangeName2 = Bookname2;
            object rangeigama = Bookigama;
            object rangevConsul = BookvConsul;            
            object rangeChildren = BookChildren;
            object rangeAppiIssSource = BookAppiIssSource;
            object rangeMention = BookMention;
            object rangeMarkEmbassy = BookMarkEmbassy;
            object rangeCountDestin = BookCountDestin;
            object rangeCountryDesc = BookCountryDesc;
            object rangeTravelPurpose = BookTravelPurpose;
            object rangePassIqama = BookPassIqama;
            object rangeAuthorization = BookAuthorization;

            oBDoc.Bookmarks.Add("MarkHijriDate", ref rangeHijriDate);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);

            oBDoc.Bookmarks.Add("MarkApplicantName", ref rangeName1);
            oBDoc.Bookmarks.Add("MarkApplicantName2", ref rangeName2);
            oBDoc.Bookmarks.Add("MarkAppliigamaNo", ref rangeigama);
            oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);            
            oBDoc.Bookmarks.Add("MarkChildrenName", ref rangeChildren);
            oBDoc.Bookmarks.Add("MarkAppIssSource", ref rangeAppiIssSource);
            oBDoc.Bookmarks.Add("MarkMention", ref rangeMention);
            oBDoc.Bookmarks.Add("MarkMarkEmbassy", ref rangeMarkEmbassy);
            oBDoc.Bookmarks.Add("MarkCountDestin", ref rangeCountDestin);
            oBDoc.Bookmarks.Add("MarkCountryDesc", ref rangeCountryDesc);
            oBDoc.Bookmarks.Add("MarkTravelPurpose", ref rangeTravelPurpose);
            oBDoc.Bookmarks.Add("MarkPassIqama", ref rangePassIqama);
            oBDoc.Bookmarks.Add("MarkAuthorization", ref rangeAuthorization);
            

            oBDoc.Activate();
            oBDoc.Save();
            oBMicroWord.Visible = true; 
            }
            else
            {
                MessageBox.Show("يرجى حذف الملف الموجودأولاً");
                printOnly.Enabled = true;
                btnSavePrint.Enabled = true;
                i = 0;
            }
        }

        public bool fileIsOpen(string path)
        {
            System.IO.FileStream a = null;

            try
            {
                a = System.IO.File.Open(path,
                System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None);
                return false;
            }
            catch (System.IO.IOException ex)
            {

                return true;
            }

            finally
            {
                if (a != null)
                {
                    a.Close();
                    a.Dispose();
                }
            }
        }

        private void TravellerDescrib_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            if (TravellerDescrib.Text == "ابناء فقط")
            {
                ChildrenOnly();
                    
            }
            else if (TravellerDescrib.Text == "زوجة فقط")
            {
                WifeOnly();
            }
            else if(TravellerDescrib.Text == "زوجة وابناء")
            {
                wifeAndChildren();
            } if (TravellerDescrib.Text == "ابناء برفقة مرافق غير الزوجة") 
            {
                ChildrenWithAttend();
            }
        }

        private void ChildrenWithAttend()
        {
            Attendecheck.Checked = true;
            WifeName.Visible = false;
            labelwifeName.Visible = false;
            labelchildren.Visible = true;
            labelchildren.Visible = true;
            AddChildren.Visible = true;
            ChildNameDesView.Visible = true;
            childboygirls.Visible = true;
            ChildrenName.Visible = true;
            labelattendchildren.Visible = true;
            Attendecheck.Visible = true;
            TravellerAttenDescrib.Visible = true;
            labelattenddesc.Visible = true;
            TravellerAttenName.Visible = true;
            labelchildrenatten.Visible = true;
            groupBox1.Visible = true ;
        }

        private void wifeAndChildren()
        {
            WifeName.Visible = true;
            labelwifeName.Visible = true;
            labelchildren.Visible = true;
            labelchildren.Visible = true;
            AddChildren.Visible = true;
            ChildNameDesView.Visible = true;
            childboygirls.Visible = true;
            ChildrenName.Visible = true;
            labelattendchildren.Visible = true;
            Attendecheck.Visible = true;
            groupBox1.Visible = false;
            if (Attendecheck.CheckState == CheckState.Unchecked)
            {
                TravellerAttenDescrib.Visible = true;
                labelattenddesc.Visible = true;
                TravellerAttenName.Visible = true;
                labelchildrenatten.Visible = true;
            }
        }

        private void WifeOnly()
        {
            WifeName.Visible = true;
            labelwifeName.Visible = true;
            labelchildren.Visible = false;
            labelchildren.Visible = false;
            AddChildren.Visible = false;
            ChildNameDesView.Visible = false;
            childboygirls.Visible = false;
            ChildrenName.Visible = false;
            labelattendchildren.Visible = false;
            Attendecheck.Visible = false;
            TravellerAttenDescrib.Visible = false;
            labelattenddesc.Visible = false;
            TravellerAttenName.Visible = false;
            labelchildrenatten.Visible = false;
            groupBox1.Visible = false;
        }

        private void ChildrenOnly()
        {
            WifeName.Visible = false;
            labelwifeName.Visible = false;
            labelchildren.Visible = true;
            labelchildren.Visible = true;
            AddChildren.Visible = true;
            ChildNameDesView.Visible = true;
            childboygirls.Visible = true;
            ChildrenName.Visible = true;
            labelattendchildren.Visible = true;
            Attendecheck.Visible = true;
            Attendecheck.Checked = false;
            groupBox1.Visible = false;
            if (Attendecheck.CheckState == CheckState.Checked)
            {
                TravellerAttenDescrib.Visible = true;
                labelattenddesc.Visible = true;
                TravellerAttenName.Visible = true;
                labelchildrenatten.Visible = true;
            }
            else
            {

                TravellerAttenDescrib.Visible = false;
                labelattenddesc.Visible = false;
                TravellerAttenName.Visible = false;
                labelchildrenatten.Visible = false;
            }
        }

        private void Attendecheck_CheckedChanged(object sender, EventArgs e)
        {
            if (Attendecheck.CheckState == CheckState.Unchecked ){

                Attendecheck.Text = "لا يوجد";
                labelchildrenatten.Visible = false;
                TravellerAttenName.Visible = false;
                labelattenddesc.Visible = false;
                TravellerAttenDescrib.Visible = false;

            }
            else if (Attendecheck.CheckState == CheckState.Checked && TravellerDescrib.Text != "زوجة فقط")
            {
                Attendecheck.Text = "يوجد";
                labelchildrenatten.Visible = true;
                TravellerAttenName.Visible = true;
                labelattenddesc.Visible = true;
                TravellerAttenDescrib.Visible = true;

            }
        }

        private void TravellingPurpo_SelectedIndexChanged(object sender, EventArgs e)
        {
            
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
            int Ddiffer = HijriDateDifferment(strDataSource, true);
            int Mdiffer = HijriDateDifferment(strDataSource, false);
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
                ApplicantSex.Text = "أنثى";
                labelName.Text = "مقدمة الطلب:";

            }
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            if (printPreviewDialog1.ShowDialog() == DialogResult.OK) printDocument1.Print();
        }


        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            FillDataGridView();
        }


        private void AddChildren_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "")
            {


                string ChildrenList = textBox2.Text;
                ChildernDescription = textBox1.Text;
                if (ChildrenList.Contains("_"))
                {
                    ChildName = ChildrenList.Split('_');
                    childindex = ChildName.Length;

                    //News.Text = "Contains";
                }
                else
                {
                    childindex = 1;
                    ChildDataBase = ChildrenList;
                }

                for (int i = 1; i < childindex; i++)
                {
                    ChildDataBase = ChildDataBase + "_" + ChildName[i];
                }
                textBox1.Text = textBox2.Text = "";
            }
            ChildName[childindex] = ChildrenName.Text;
            if (childindex == 0)
            {
                if (childboygirls.CheckState == CheckState.Checked)
                {

                    ChildernDescription = "ابني";
                    Mentioned = "للمذكور";
                }
                else
                {

                    ChildernDescription = "ابنتي";
                    Mentioned = "للمذكورة";
                }
                ChildDataBase = ChildName[childindex];

            }
            else if (childindex == 1)
            {
                if (childboygirls.CheckState == CheckState.Checked && ChildernDescription == "ابني")
                {
                    ChildernDescription = "ابنيَّ";
                    Mentioned = "للمذكورين";
                }
                else if (childboygirls.CheckState == CheckState.Unchecked && ChildernDescription == "ابنتي")
                {
                    ChildernDescription = "ابنتيَّ";
                    Mentioned = "للمذكورتين";
                }                
                ChildDataBase = ChildDataBase + "_" + ChildName[1];
            }
            else if(childindex >= 2)
            {
                if (childboygirls.CheckState == CheckState.Unchecked && ChildernDescription == "ابنتيَّ")
                {
                    ChildernDescription = "بناتي";
                    Mentioned = "للذكورات";
                }
                else
                {
                    ChildernDescription = "أبنائي";
                    Mentioned = "للمذكورين";
                }
                ChildDataBase = ChildDataBase + "_" + ChildName[childindex];
            }
            for (int j = 1; j < childindex; j++)
            {
                ChildDataBase = ChildDataBase + "_" + ChildName[j];
            }
            if (ChildDataBase.Contains("_")) ChildNameDesView.Text = ChildernDescription + "/ " + ChildDataBase.Replace("_", " و");
            else ChildNameDesView.Text = ChildernDescription + "/ " + ChildDataBase;
            childindex++;
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
        private void OpenFile(int id, int fileNo)
        {
            string query;

            SqlConnection Con = new SqlConnection(strDataSource);
            if (fileNo == 1)
            {
                query = "select Data1, Extension1,FileName1 from TableTravIqrar where ID=@id";
            }
            else
            {
                query = "select Data2, Extension2,FileName2 from TableTravIqrar where ID=@id";
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

        private void dataGridView1_DoubleClick_1(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                CurrentIqrarId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                ApplicantName.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString().ToString();
                if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
                else if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
                PassIqama.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().ToString();
                ApplicantIdoc.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString().ToString();
                IssuedSource.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString().ToString();
                ChildernDescription = dataGridView1.CurrentRow.Cells[7].Value.ToString().ToString();
                string ChildrenList = dataGridView1.CurrentRow.Cells[8].Value.ToString().ToString();
                if (ChildrenList.Contains("_"))
                {
                    ChildFromDataBase = ChildrenList.Split('_');
                    childindex = ChildFromDataBase.Length;
                }
                else
                {
                    childindex = 1;
                    ChildDataBase = ChildrenList;
                }

                for (int i = 1; i < childindex; i++)
                {
                    ChildDataBase = ChildDataBase + "_" + ChildFromDataBase[i];
                }
                ChildNameDesView.Text = ChildernDescription + "/ " + ChildDataBase;
                textBox1.Text = ChildernDescription;
                textBox2.Text = ChildDataBase;
                EmbassySource.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                GregorianDate.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                HijriDate.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[16].Value.ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;
                    IqrarNo.Text = NextRelId;
                }
                else checkedViewed.CheckState = CheckState.Checked; 
                if (checkedViewed.CheckState == CheckState.Checked)
                {
                    PreAppId = CurrentIqrarId;
                }
                else {
                    PreAppId = "";
                    IqrarNo.Text = CurrentIqrarId;
                }
                TravelDestin.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString().ToString();
                TravellingPurpo.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString().ToString();
                AppType.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();
                if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked;
                else AppType.CheckState = CheckState.Unchecked;
                ConsulateEmployee.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
                if (AppType.CheckState == CheckState.Unchecked)
                {
                    mandoubVisibilty(); mandoubName.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
                }
                Comment.Text = dataGridView1.CurrentRow.Cells[24].Value.ToString();
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

                ArchivedSt.Visible = true;
                labelArch.Visible = true;
                printOnly.Visible = true;
                SaveOnly.Visible = true;
                btnSavePrint.Text = "حفظ";
                btnSavePrint.Visible = false;
            }
        }

        private void printOnly_Click(object sender, EventArgs e)
        {
            
        }

        private void Search_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnSavePrint_Click(object sender, EventArgs e)
        {
            Save2DataBase();
            btnSavePrint.Text = "جاري المعالجة";
            btnSavePrint.Enabled = false;
            CreateWordFile();
            Clear_Fields();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            printOnly.Text = "طباعة";
            printOnly.Enabled = false;
            CreateWordFile();
            Clear_Fields();
        }

        private void PassIqama_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (PassIqama.Text == "اقامة ") { ApplicantIdoc.Text = ""; } else ApplicantIdoc.Text = "P";
        }

        private void TravellingPurpo_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (TravellingPurpo.Text == "العمل")
            {
                labeljob.Visible = true;
                familyJob.Visible = true;
            }
            else
            {
                labeljob.Visible = false;
                familyJob.Visible = false;
            }
        }

        private void childboygirls_CheckedChanged(object sender, EventArgs e)
        {
            if (childboygirls.CheckState == CheckState.Checked) childboygirls.Text = "ابن";
            else childboygirls.Text = "ابنة";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void SaveOnly_Click(object sender, EventArgs e)
        {
            Save2DataBase();          
            Clear_Fields();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void Save2DataBase()
        {
            SqlConnection sqlCon = new SqlConnection(strDataSource);

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
                    SqlCommand sqlCmd = new SqlCommand("TravAddorEdit", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");
                    sqlCmd.Parameters.AddWithValue("@DocID", IqrarNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", PassIqama.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", ApplicantIdoc.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", IssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@ChildDesc", ChildernDescription.Trim());
                    sqlCmd.Parameters.AddWithValue("@ChildNames", ChildDataBase);
                    sqlCmd.Parameters.AddWithValue("@Embassy", EmbassySource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@CountDestin", TravelDestin.Text);
                    sqlCmd.Parameters.AddWithValue("@TravelPurpose", TravellingPurpo.Text);
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

                    sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
                    sqlCmd.ExecuteNonQuery();



                }
                else
                {
                    if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                    else Viewed = "غير معالج";

                    SqlCommand sqlCmd = new SqlCommand("TravAddorEdit", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                    sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                    sqlCmd.Parameters.AddWithValue("@DocID", IqrarNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AppName", ApplicantName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", PassIqama.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", ApplicantIdoc.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", IssuedSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@ChildDesc", ChildernDescription.Trim());
                    sqlCmd.Parameters.AddWithValue("@ChildNames", ChildDataBase.Trim());
                    sqlCmd.Parameters.AddWithValue("@Embassy", EmbassySource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text);
                    sqlCmd.Parameters.AddWithValue("@CountDestin", TravelDestin.Text);
                    sqlCmd.Parameters.AddWithValue("@TravelPurpose", TravellingPurpo.Text);
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

                    if (fileloaded)
                        sqlCmd.Parameters.AddWithValue("@ArchivedState", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    else sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");

                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
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


    }

}


