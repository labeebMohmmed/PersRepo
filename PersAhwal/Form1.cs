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
using System.Data.SqlClient;
using System.Configuration;
using System.Globalization;
using System.Threading;
using System.IO;
using Xceed.Document.NET;


//https://www.youtube.com/watch?v=-2UcDV4uUu8
//https://www.youtube.com/watch?v=QTWKUkiEqpQ
namespace PersAhwal
{
    public partial class Form1 : Form
    {

        private static int childindex = 0;
        string Viewed = "تمت المعالجة";
        int ApplicantID = 0;
        static public bool ApplicantSexStatus = false;
        public static string route = "";
        string Mentioned="", CurrentFileName;
        string ChildernDescription = "";
        string IqrarStaticPart = "ق س ج/160/01/";
        string IqrarNumberPart;
        string ChildDataBase="";
        string PreAppId = "", PreRelatedID,NextRelId;
        private static bool fileloaded=false;
        static string DataSource;        
        public static string[] ChildName = new string[10];
        public static string[] DataInterType = new string[3];
        public static string[] DataInterName = new string[3];
        string[] ChildFromDataBase ;
        string ConsulateEmpName;
        string FilesPathIn, FilesPathOut;

            
        public Form1(int currentRow, string EmpName, string dataSource, string filepathIn, string filepathOut)
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


        private void Clear_Fields() {
            ApplicantName.Text = IssuedSource.Text = ApplicantIdoc.Text = AttendViceConsul.Text = ChildernDescription  = PassIqama.Text = "";
            ChildernDescription = ChildNameDesView.Text  = ChildrenName.Text = ""; 
            AttendViceConsul.SelectedIndex = 2;
            
            Comment.Text = "لا تعليق";
            IssueDocSource.Text = "السودان";
            PassIqama.Text = "جواز سفر";
            ApplicantIdoc.Text = "P";
            childindex = 0;
            mandoubName.Text = ListSearch.Text = "";
            printOnly.Visible = false;
            btnSavePrint.Text = "طباعة وحفظ";
            btnSavePrint.Enabled = true;
            btnSavePrint.Visible = true;
            SaveOnly.Visible = false;
            Comment.Text = "لا تعليق";
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            labelArch.Visible = false;
            ArchivedSt.BackColor = Color.Red;
            SearchFile.Visible = false;
            fileloaded = false;
            CurrentFileName = IqrarNumberPart + "_01";
            IqrarNo.Text = IqrarStaticPart + IqrarNumberPart;
            childindex = 0;
            ChildernDescription = "";

            FillDataGridView();
        }

        private void label2_Click(object sender, EventArgs e)
        {
            labelName.ForeColor = Color.Black;
        }


        private void ApplicantSex_CheckedChanged_1(object sender, EventArgs e)
        {
            if (ApplicantSex.CheckState == CheckState.Unchecked) ApplicantSex.Text = "ذكر";
            else ApplicantSex.Text = "أنثى";
        }

        
       

        private void childboygirls_CheckedChanged(object sender, EventArgs e)
        {
            if (childboygirls.CheckState == CheckState.Checked) childboygirls.Text = "ابن";
            else childboygirls.Text = "ابنة";
        }

        
        private void CreateWordFile()
        {
            if (ApplicantSex.CheckState == CheckState.Unchecked)
            {
                ApplicantSexStatus = true;
                labelName.ForeColor = Color.Black;
                labelName.Text = "مقدم الطلب:";
                route = FilesPathIn + "IgrarDocM.docx";
            }
            else if (ApplicantSex.CheckState == CheckState.Checked)
            {
                ApplicantSexStatus = false;
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
                route = FilesPathIn + "IgrarDocF.docx";
            }

            string ActiveVopy = FilesPathOut + ApplicantName.Text + CurrentFileName + ".docx";
            System.IO.File.Copy(route, ActiveVopy);            
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveVopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);
            var document = Xceed.Words.NET.DocX.Load(ActiveVopy);
            var t = document.AddTable( 3, 3);
            t.Design = TableDesign.TableGrid;
            t.Alignment = Alignment.center;
            t.SetColumnWidth(2, 40);
            t.SetColumnWidth(1, 180);
            t.SetColumnWidth(0, 180);


            t.Rows[0].Cells[0].Paragraphs[0].Append("نوع المعاملة").FontSize(15d).Bold().Alignment = Alignment.center;
            t.Rows[0].Cells[1].Paragraphs[0].Append("اسم مقدم الطلب").FontSize(15d).Bold().Alignment = Alignment.center;
            t.Rows[0].Cells[2].Paragraphs[0].Append("الرقم").FontSize(15d).Bold().Alignment = Alignment.center;

            for (int x = 1; x <= 2; x++)
            {
                t.Rows[x].Cells[0].Paragraphs[0].Append("RetrievedTypeList[x - 1]").FontSize(15d).Direction = Direction.RightToLeft;
                t.Rows[x].Cells[1].Paragraphs[0].Append("RetrievedNameList[x - 1]").FontSize(15d).Direction = Direction.RightToLeft;
                t.Rows[x].Cells[2].Paragraphs[0].Append((x).ToString()).FontSize(15d).Direction = Direction.RightToLeft;
            }



            object ParaIqrarNo = "MarkIqrarNo";
            object Paraname = "MarkApplicantName";
            object Paraname2 = "MarkApplicantName2";
            object Paraigama = "MarkAppliigamaNo";
            object ParavConsul = "MarkViseConsul";
            object ParaAuthorization = "MarkAuthorization";
            object ParaChildDesc = "MarkChildDesc";
            object ParaChildren = "MarkChildrenName";
            object ParaAppiIssSource = "MarkAppIssSource";
            object ParaPassIqama = "MarkPassIqama";
            object ParaGreData = "MarkGreData";
            object ParaHijriData = "MarkHijriData";

            Word.Range BookIqrarNo = oBDoc.Bookmarks.get_Item(ref ParaIqrarNo).Range;
            Word.Range Bookname = oBDoc.Bookmarks.get_Item(ref Paraname).Range;
            Word.Range Bookname2 = oBDoc.Bookmarks.get_Item(ref Paraname2).Range;
            Word.Range Bookigama = oBDoc.Bookmarks.get_Item(ref Paraigama).Range;
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;
            Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;
            Word.Range BookChildDesc = oBDoc.Bookmarks.get_Item(ref ParaChildDesc).Range;
            Word.Range BookChildren = oBDoc.Bookmarks.get_Item(ref ParaChildren).Range;
            Word.Range BookAppiIssSource = oBDoc.Bookmarks.get_Item(ref ParaAppiIssSource).Range;
            Word.Range BookPassIqama = oBDoc.Bookmarks.get_Item(ref ParaPassIqama).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;

            BookIqrarNo.Text = IqrarNo.Text;
            
            Bookname.Text = Bookname2.Text = ApplicantName.Text;
            Bookigama.Text = ApplicantIdoc.Text;
            BookvConsul.Text =  AttendViceConsul.Text;
            if(AppType.CheckState == CheckState.Checked){
                if (ApplicantSex.CheckState == CheckState.Unchecked) BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " نائب قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد حضر للقنصلية ووقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                if (ApplicantSex.CheckState == CheckState.Checked) BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " نائب قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد حضرت للقنصلية ووقعت بتوقيعها على هذا الإقرار بعد تلاوتها عليها وبعد أن فهمت مضمونه ومحتواه. ";
            }else {
                if (ApplicantSex.CheckState == CheckState.Unchecked) BookAuthorization.Text = "المواطن المذكور أعلاه حضر ووقع بتوقيعه على هذا الإقرار أمام مندوب الجالية لدى القنصلية السيد/ " + mandoubName.Text + "، وذلك بموجب التفويض الممنوح له، ";
                if (ApplicantSex.CheckState == CheckState.Checked) BookAuthorization.Text = "المواطنة المذكورة أعلاه حضرت ووقعت بتوقيعها على هذا الإقرار أمام مندوب الجالية لدى القنصلية السيد/ " + mandoubName.Text + "، وذلك بموجب التفويض الممنوح له، ";
            }
            BookChildDesc.Text = ChildernDescription;
            BookChildren.Text = ChildDataBase;
            BookAppiIssSource.Text = IssuedSource.Text;
            BookPassIqama.Text = PassIqama.Text;
            BookGreData.Text = GregorianDate.Text;
            BookHijriData.Text = HijriDate.Text;


            object rangeIqrarNo = BookIqrarNo;
            object rangeName = Bookname;
            object rangeName2 = Bookname2;
            object rangeigama = Bookigama;
            object rangevConsul = BookvConsul;
            object rangeAuthorization = BookAuthorization;
            object rangeChildDesc = BookChildDesc;
            object rangeChildren = BookChildren;
            object rangeAppiIssSource = BookAppiIssSource;
            object rangePassIqama = BookPassIqama;
            object rangeGreData = BookGreData;
            object rangeHijriData = BookHijriData;



            oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);
            oBDoc.Bookmarks.Add("MarkApplicantName", ref rangeName);
            oBDoc.Bookmarks.Add("MarkApplicantName2", ref rangeName2);
            oBDoc.Bookmarks.Add("MarkAppliigamaNo", ref rangeigama);
            oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);
            oBDoc.Bookmarks.Add("MarkAuthorization", ref rangeAuthorization);
            oBDoc.Bookmarks.Add("MarkChildDesc", ref rangeChildDesc);
            oBDoc.Bookmarks.Add("MarkChildrenName", ref rangeChildren);
            oBDoc.Bookmarks.Add("MarkAppIssSource", ref rangeAppiIssSource);
            oBDoc.Bookmarks.Add("MarkPassIqama", ref rangePassIqama);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkHijiData", ref rangeHijriData);


            oBDoc.Activate();

            oBDoc.Save();
            oBMicroWord.Visible = true;
            if (IssueDocSource.Text != "السودان")
            {
                CreateMessageWord();
                Clear_Fields();
            }
        }

        private void CreateMessageWord()
        {
            string ActiveCopy;
            route = FilesPathIn + "MessageCap.docx";
            ActiveCopy = FilesPathOut + "Message" + ApplicantName.Text + CurrentFileName+ ".docx";
            if (!File.Exists(ActiveCopy))
            {
                System.IO.File.Copy(route, ActiveCopy);
                object oBMiss2 = System.Reflection.Missing.Value;
                Word.Application oBMicroWord2 = new Word.Application();



                Word.Document oBDoc2 = oBMicroWord2.Documents.Open(ActiveCopy, oBMiss2);

                Object ParacapitalMessage = "MarkcapitalMessage";
                Object ParaMApplicantName = "MarkApplicantName";
                Object ParaMassageIqrarNo = "MarkMassageIqrarNo";
                Object ParaApliSex = "MarkApliSex";
                Object ParaHijriDate = "MarkHijriDate";
                Object ParaDateGre = "MarkDateGre";
                Object ParaGregorDate2 = "MarkGregorDate2";

                Word.Range BookMApplicantName = oBDoc2.Bookmarks.get_Item(ref ParaMApplicantName).Range;
                Word.Range BookcapitalMessage = oBDoc2.Bookmarks.get_Item(ref ParacapitalMessage).Range;
                Word.Range BookMassageIqrarNo = oBDoc2.Bookmarks.get_Item(ref ParaMassageIqrarNo).Range;
                Word.Range BookApliSex = oBDoc2.Bookmarks.get_Item(ref ParaApliSex).Range;
                Word.Range BookDateGre = oBDoc2.Bookmarks.get_Item(ref ParaDateGre).Range;
                Word.Range BookHijriDate = oBDoc2.Bookmarks.get_Item(ref ParaHijriDate).Range;
                Word.Range BookGregorDate2 = oBDoc2.Bookmarks.get_Item(ref ParaGregorDate2).Range;


                BookMApplicantName.Text = ApplicantName.Text;
                BookcapitalMessage.Text = IssueDocSource.Text;
                BookMassageIqrarNo.Text = IqrarNo.Text;
                if (ApplicantSex.CheckState == CheckState.Unchecked)
                    BookApliSex.Text = "المواطن";
                else BookApliSex.Text = "المواطنة";
                BookGregorDate2.Text = BookDateGre.Text = GregorianDate.Text;
                BookHijriDate.Text = HijriDate.Text;
                object rangeMApplicantName = BookMApplicantName;
                object rangecapitalMessage = BookcapitalMessage;
                object rangeMassageIqrarNo = BookMassageIqrarNo;
                object rangeApliSex = BookApliSex;
                object rangeDateGre = BookDateGre;
                object rangeHijriDate = BookHijriDate;
                object rangeGregorDate2 = BookGregorDate2;

                oBDoc2.Bookmarks.Add("MarkApplicantName", ref rangeMApplicantName);
                oBDoc2.Bookmarks.Add("MarkcapitalMessage", ref rangecapitalMessage);
                oBDoc2.Bookmarks.Add("MarkMassageIqrarNo", ref rangeMassageIqrarNo);
                oBDoc2.Bookmarks.Add("MarkApliSex", ref rangeApliSex);
                oBDoc2.Bookmarks.Add("MarkDateGre", ref rangeDateGre);
                oBDoc2.Bookmarks.Add("MarkGregorDate2", ref rangeGregorDate2);
                oBDoc2.Bookmarks.Add("MarkHijiData", ref rangeHijriDate);
                oBDoc2.Activate();
                oBDoc2.Save();
                oBMicroWord2.Visible = true;

            }

            else
            {
                MessageBox.Show("يرجى حذف الملف الموجودأولاً");
                printOnly.Enabled = true;
                btnSavePrint.Enabled = true;

            }
}
        
        private void Save2DataBase()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);

            string AppGender;
            if (ApplicantSex.CheckState == CheckState.Unchecked) AppGender = "ذكر"; else AppGender = "أنثى";
            if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                if (btnSavePrint.Text == "طباعة وحفظ")
                {                    
                    if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                    else Viewed = "غير معالج";
                    SqlCommand sqlCmd = new SqlCommand("DocAddorEdit", sqlCon);
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
                    sqlCmd.Parameters.AddWithValue("@ChildNames", ChildDataBase.Trim());
                    sqlCmd.Parameters.AddWithValue("@ChildDocIssueSource", IssueDocSource.Text.Trim());
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
                    if (ListSearch.Text != "") filePath2 = ListSearch.Text;
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
                    if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                    else Viewed = "غير معالج";
                    SqlCommand sqlCmd = new SqlCommand("DocAddorEdit", sqlCon);
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
                    sqlCmd.Parameters.AddWithValue("@ChildDocIssueSource", IssueDocSource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim()+" "+ DateTime.Now.ToString("hh:mm"));                    
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
                    if (ListSearch.Text != "") filePath2 = ListSearch.Text;
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
            
            sqlCon.Close();
            FillDataGridView();
        }
    

        private void timer1_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new HijriCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            int Ddiffer = HijriDateDifferment(DataSource,true);
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

        private void timer2_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            GregorianDate.Text = DateTime.Now.ToString("dd-MM-yyyy");
           
        }
        private void PassIqama_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (PassIqama.Text == "جواز سفر")
            {
                ApplicantIdoc.Text = "P";
                labeldoctype.Text = "رقم جواز السفر: ";
            }
            else if (PassIqama.Text == "إقامة")
            {               
                ApplicantIdoc.Text = "";
                labeldoctype.Text = "رقم الاقامة:";
            }
            else {
                labeldoctype.Text = "رقم وطني";
                ApplicantIdoc.Text = "";
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        void FillDataGridView()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("DocViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ApplicantName", ListSearch.Text.Trim());
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGrid.DataSource = dtbl;
            IqrarNumberPart = (dtbl.Rows.Count + 1).ToString();
            CurrentFileName= IqrarNumberPart + "_1";
            sqlCon.Close();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                FillDataGridView();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {            
            if (textBox1.Text != "" && textBox2.Text != "") {
                
                string ChildrenList = textBox2.Text;
                ChildernDescription = textBox1.Text;
                if (ChildrenList.Contains("_"))
                {
                    ChildName = ChildrenList.Split('_');
                    childindex = ChildName.Length;

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
                childindex=1;
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
                childindex=2;
            }
            else if (childindex == 2)
            {
                if (childboygirls.CheckState == CheckState.Unchecked && ChildernDescription == "ابنتيَّ")
                {
                    ChildernDescription = "بناتي";
                    Mentioned = "للمذكورات";
                }
                else
                {
                    ChildernDescription = "أبنائي";
                    Mentioned = "للمذكورين";
                }

                ChildDataBase = ChildDataBase + "_" + ChildName[2];
                childindex=3;
            }
            else {
                if (childboygirls.CheckState == CheckState.Unchecked && ChildernDescription == "بناتي")
                {
                    ChildernDescription = "بناتي";
                    Mentioned = "للمذكورات";
                }
                else
                {
                    ChildernDescription = "أبنائي";
                    Mentioned = "للمذكورين";
                }
                ChildDataBase = ChildDataBase + "_" + ChildName[childindex];
                Mentioned = "للمذكورين";
                childindex++;
            }


            

            if(ChildDataBase.Contains("_")) ChildNameDesView.Text = ChildernDescription + "/ " + ChildDataBase.Replace("_", " و");
            else ChildNameDesView.Text = ChildernDescription + "/ " + ChildDataBase;
            
            ChildrenName.Clear();
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


        private void SetFieldswithData(int Rowindex)
        {
            Rowindex--;
            
            ApplicantID = Convert.ToInt32(dataGrid.Rows[Rowindex].Cells[0].Value.ToString());
            IqrarNo.Text = dataGrid.Rows[Rowindex].Cells[1].Value.ToString();
            ApplicantName.Text = dataGrid.Rows[Rowindex].Cells[2].Value.ToString();
            if (dataGrid.Rows[Rowindex].Cells[3].Value.ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
            else if (dataGrid.Rows[Rowindex].Cells[3].Value.ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
            PassIqama.Text = dataGrid.Rows[Rowindex].Cells[4].Value.ToString();
            ApplicantIdoc.Text = dataGrid.Rows[Rowindex].Cells[5].Value.ToString();
            IssuedSource.Text = dataGrid.Rows[Rowindex].Cells[6].Value.ToString();
            ChildernDescription = dataGrid.Rows[Rowindex].Cells[7].Value.ToString();
            string ChildrenList = dataGrid.Rows[Rowindex].Cells[8].Value.ToString();
            if (ChildrenList.Contains("_"))
            {
                ChildFromDataBase = ChildrenList.Split('_');
                childindex = ChildFromDataBase.Length;
                ChildDataBase = ChildFromDataBase[0];
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
            IssueDocSource.Text = dataGrid.Rows[Rowindex].Cells[9].Value.ToString();
            GregorianDate.Text = dataGrid.Rows[Rowindex].Cells[10].Value.ToString();
            HijriDate.Text = dataGrid.Rows[Rowindex].Cells[11].Value.ToString();
            AttendViceConsul.Text = dataGrid.Rows[Rowindex].Cells[12].Value.ToString();

            if (dataGrid.Rows[Rowindex].Cells[13].Value.ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                IqrarNo.Text = NextRelId;
            }
            else checkedViewed.CheckState = CheckState.Checked;
            AppType.Text = dataGrid.Rows[Rowindex].Cells[14].Value.ToString();

            if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked;
            else AppType.CheckState = CheckState.Unchecked;
            ConsulateEmployee.Text = dataGrid.Rows[Rowindex].Cells[15].Value.ToString();
            if (AppType.CheckState == CheckState.Unchecked)
            {
                mandoubVisibilty(); mandoubName.Text = dataGrid.Rows[Rowindex].Cells[16].Value.ToString();
            }
            PreRelatedID = dataGrid.Rows[Rowindex].Cells[17].Value.ToString();
            Comment.Text = dataGrid.Rows[Rowindex].Cells[22].Value.ToString();
            if (dataGrid.Rows[Rowindex].Cells[23].Value.ToString() != "غير مؤرشف")
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

        private void button6_Click(object sender, EventArgs e)
        {
            FillDataGridView();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            
        }

        private void SearchDoc_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            SearchFile.Visible = true;
            SearchFile.Text = dlg.FileName;
            if (SearchFile.Text != "") fileloaded = true;
        }


        private void button7_Click(object sender, EventArgs e)
        {
            var selectRows = dataGrid.SelectedRows;
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

            SqlConnection Con = new SqlConnection(DataSource);
            if (fileNo == 1)
            {
                query = "select Data1, Extension1,FileName1 from TableDocIqrar where ID=@id";
            }
            else
            {
                query = "select Data2, Extension2,FileName2 from TableDocIqrar where ID=@id";
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
        private void button3_Click_1(object sender, EventArgs e)
        {
            var selectRows = dataGrid.SelectedRows;
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
            if (dataGrid.CurrentRow.Index != -1) {
                ApplicantID = Convert.ToInt32(dataGrid.CurrentRow.Cells[0].Value.ToString());
                IqrarNo.Text = dataGrid.CurrentRow.Cells[1].Value.ToString();
                ApplicantName.Text = dataGrid.CurrentRow.Cells[2].Value.ToString();
                if (dataGrid.CurrentRow.Cells[3].Value.ToString() == "ذكر") ApplicantSex.CheckState = CheckState.Unchecked;
                else if (dataGrid.CurrentRow.Cells[3].Value.ToString() == "أنثى") ApplicantSex.CheckState = CheckState.Checked;
                PassIqama.Text = dataGrid.CurrentRow.Cells[4].Value.ToString();
                ApplicantIdoc.Text = dataGrid.CurrentRow.Cells[5].Value.ToString();
                IssuedSource.Text = dataGrid.CurrentRow.Cells[6].Value.ToString();
                ChildernDescription = dataGrid.CurrentRow.Cells[7].Value.ToString();
                string ChildrenList = dataGrid.CurrentRow.Cells[8].Value.ToString();
                if (ChildrenList.Contains("_"))
                {
                    ChildFromDataBase = ChildrenList.Split('_');
                    childindex = ChildFromDataBase.Length;
                    ChildDataBase = ChildFromDataBase[0];
                }
                else { childindex = 1;
                    ChildDataBase = ChildrenList;
                }
                
                for (int i = 1; i < childindex; i++)
                {
                    ChildDataBase = ChildDataBase + "_" + ChildFromDataBase[i];
                }
                ChildNameDesView.Text = ChildernDescription + "/ " + ChildDataBase;
                textBox1.Text = ChildernDescription;
                textBox2.Text = ChildDataBase;
                IssueDocSource.Text = dataGrid.CurrentRow.Cells[9].Value.ToString();
                GregorianDate.Text = dataGrid.CurrentRow.Cells[10].Value.ToString();
                HijriDate.Text = dataGrid.CurrentRow.Cells[11].Value.ToString();
                AttendViceConsul.Text = dataGrid.CurrentRow.Cells[12].Value.ToString();

                if (dataGrid.CurrentRow.Cells[13].Value.ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;
                    IqrarNo.Text = NextRelId;
                }
                else checkedViewed.CheckState = CheckState.Checked;
                AppType.Text = dataGrid.CurrentRow.Cells[14].Value.ToString();

                if (AppType.Text == "حضور مباشرة إلى القنصلية") AppType.CheckState = CheckState.Checked;
                else AppType.CheckState = CheckState.Unchecked;
                ConsulateEmployee.Text = dataGrid.CurrentRow.Cells[15].Value.ToString();
                if (AppType.CheckState == CheckState.Unchecked)
                {
                    mandoubVisibilty(); mandoubName.Text = dataGrid.CurrentRow.Cells[16].Value.ToString();
                }
                PreRelatedID = dataGrid.CurrentRow.Cells[17].Value.ToString();
                Comment.Text = dataGrid.CurrentRow.Cells[22].Value.ToString();
                if (dataGrid.CurrentRow.Cells[23].Value.ToString() != "غير مؤرشف")
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

        private void IssueDocSource_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnSavePrint_Click(object sender, EventArgs e)
        {
            Save2DataBase();
            btnSavePrint.Text = "جاري المعالجة";
            btnSavePrint.Enabled = false; 
            CreateWordFile();
            if (IssueDocSource.Text != "السودان")
                Clear_Fields();

        }

        private void ResetAll_Click(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void printOnly_Click(object sender, EventArgs e)
        {
            printOnly.Text = "طباعة";
            printOnly.Enabled = false;
            CreateWordFile();
            Clear_Fields();
        }

        private void SaveOnly_Click(object sender, EventArgs e)
        {
            Save2DataBase();            
            Clear_Fields();
        }

        private int HijriDateDifferment(string source, bool daymonth)
        {
            int differment = 0;
            string query;
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                if(daymonth) query = "select hijriday from TableSettings";
                else query = "select hijrimonth from TableSettings";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    if(daymonth) differment = Convert.ToInt32(reader["hijriday"].ToString());
                    else differment = Convert.ToInt32(reader["hijrimonth"].ToString());
                }

                saConn.Close();
            }
            return differment;
        }
    }

}
