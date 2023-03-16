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
using DocumentFormat.OpenXml.Office2010.Excel;
using Color = System.Drawing.Color;
using System.Security.AccessControl;
using DocumentFormat.OpenXml.Vml.Spreadsheet;

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
        static string ChildernDescription = "";
        static string ChildDataBase = "", Mentioned = "";
        string Viewed;
        int MessageDocNo = 0;
        string routeDoc;
        string ConsulateEmpName;
        public static string ModelFileroute = "";
        
        String IqrarNumberPart;
        bool SaveEdit = false;
        static string DataSource;
        int rowIndexTodelete = 0;
        int ApplicantID = 0;
        bool fileloaded = false;
        string PreAppId = "", CurrentIqrarId = "", PreRelatedID = "", NextRelId = "";
        string CurrentFileName;
        string FilesPathIn, FilesPathOut;
        string UserJobposition;
        int ATVC = 0;
        static string[] colIDs = new string[100];
        string GregorianDate = "";
        string HijriDate = "";
        bool gridFill = true;
        bool gridFillauto = true;
        string AuthTitle = "نائب قنصل";
        public Form2(int atvc,int currentRow, string EmpName, string dataSource, string filepathIn, string filepathOut, string jobposition, string gregorianDate, string hijriDate)
        {
            InitializeComponent();
            //timer1.Enabled = true;
            //timer2.Enabled = true;

            التاريخ_الميلادي.Text =  GregorianDate =gregorianDate;
            التاريخ_الهجري.Text = HijriDate = hijriDate;
            ATVC = atvc;
            DataSource = dataSource;
            FilesPathIn = filepathIn + @"\";
            FilesPathOut = filepathOut;
            colIDs[4] = ConsulateEmpName = EmpName;
            UserJobposition = jobposition;
            FillDataGridView();
            if (currentRow == -1) Clear_Fields();
            else SetFieldswithData(currentRow);
            
            if (jobposition.Contains("قنصل"))
                btnEditID.Visible = button5.Visible = true;
            else 
                btnEditID.Visible = button5.Visible = false;
            getTitle(DataSource, EmpName);
        }
        private void ColorFulGrid9()
        {
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            dataGridView1.BackgroundColor = Color.White;

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {

                if (dataGridView1.Rows[i].Cells["تاريخ_الميلاد"].Value.ToString() != "" && dataGridView1.Rows[i].Cells["المهنة"].Value.ToString() != "" && dataGridView1.Rows[i].Cells[26].Value.ToString() == "مؤرشف نهائي") 
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                else if (dataGridView1.Rows[i].Cells["تاريخ_الميلاد"].Value.ToString() == "" || dataGridView1.Rows[i].Cells["المهنة"].Value.ToString() == "")
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightCoral;

                }
                else dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;
            }
        }


        private void loadMessageNo()
        {
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select MessageNo  from TableSettings where ID=@id", Con);
            sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();

            if (reader.Read())
            {
                MessageDocNo = Convert.ToInt32(reader["MessageNo"].ToString());
            }

            Con.Close();


        }

        private void Clear_Fields()
        {
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();

            مقدم_الطلب.Text = مكان_الإصدار.Text = TravellingPurpo.Text = رقم_الهوية.Text = ChildernDescription = نوع_الهوية.Text = ChildernDescription = ChildNameDesView.Text = ChildrenName.Text = "";
            AttendViceConsul.SelectedIndex = 2;
            TravelDestin.SelectedIndex = 0;
            TravellerDescrib.SelectedIndex = 0;
            EmbassySource.SelectedIndex = 26;
            
            familyJob.SelectedIndex = 0;
            familyJob.Visible = labeljob.Visible = false;
            رقم_الهوية.Text = "P";
            Comment.Text = "لا تعليق";
            نوع_الهوية.Text = "جواز سفر";
            childindex = 0;
            personUnderPro.Text = mandoubName.Text = Search.Text = "";
            AppType.CheckState = CheckState.Checked;
            mandoubVisibilty();            
            btnSavePrint.Text = "حفظ وطباعة";
            btnSavePrint.Enabled = true;
            ConsulateEmployee.Text = ConsulateEmpName;
            FillDataGridView();
            ArchivedSt.Text = "غير مؤرشف";
            ArchivedSt.Visible = false;
            
            ArchivedSt.BackColor = Color.Red;
            
            ProcedureType.SelectedIndex = 0;

            ChildernDescription = Mentioned = ChildDataBase = textBox2.Text = textBox1.Text = "";
            EmbassySource.SelectedIndex = 26;
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
            التاريخ_الميلادي.Text = dataGridView1.Rows[Rowindex].Cells[10].Value.ToString();
            التاريخ_الهجري.Text = dataGridView1.Rows[Rowindex].Cells[11].Value.ToString();
            AttendViceConsul.Text = dataGridView1.Rows[Rowindex].Cells[12].Value.ToString();
            if (dataGridView1.Rows[Rowindex].Cells[16].Value.ToString() == "غير معالج")
            {
                checkedViewed.CheckState = CheckState.Unchecked;
                
            }
            else checkedViewed.CheckState = CheckState.Checked;
            if (checkedViewed.CheckState == CheckState.Checked)
            {
                PreAppId = CurrentIqrarId;
            }
            else
            {
                PreAppId = "";
                
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
        }

        void FillDataGridView()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("TravViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ApplicantName", Search.Text.Trim());
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            CurrentFileName = IqrarNumberPart + "_02";
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 200;
            IqrarNumberPart = loadRerNo(loadIDNo());
            sqlCon.Close();
            ColorFulGrid9();
        }
        private void Review_Click(object sender, EventArgs e)
        {

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
        private string loadRerNo(int id)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT DocID from TableTravIqrar where ID=@ID", sqlCon);
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
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from TableTravIqrar order by ID desc", sqlCon);
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

        private void CreateWordFileDoc()
        {
            string docxouput = FilesPathOut + مقدم_الطلب.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilesPathOut + مقدم_الطلب.Text + DateTime.Now.ToString("ssmm") + ".pdf";

            string ReportName = DateTime.Now.ToString("mmss");
            if (النوع.CheckState == CheckState.Unchecked)
            {
                ApplicantSexStatus = true;
                labelName.ForeColor = Color.Black;
                labelName.Text = "مقدم الطلب:";
                routeDoc = FilesPathIn + "IgrarDocM.docx";
            }
            else if (النوع.CheckState == CheckState.Checked)
            {
                ApplicantSexStatus = false;
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
                routeDoc = FilesPathIn + "IgrarDocF.docx";
            }

            string CurrentCopy = FilesPathOut + مقدم_الطلب.Text + ReportName + ".docx";
            if (!File.Exists(CurrentCopy))
            {
                System.IO.File.Copy(routeDoc, CurrentCopy);
                object oBMiss = System.Reflection.Missing.Value;
                Word.Application oBMicroWord = new Word.Application();

                object objCurrentCopy = CurrentCopy;

                Word.Document oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);

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

                BookIqrarNo.Text = colIDs[0] = IqrarNo.Text;

                Bookname.Text = Bookname2.Text = colIDs[3] = مقدم_الطلب.Text;
                colIDs[5] = AppType.Text;
                colIDs[6] = mandoubName.Text;
                Bookigama.Text = رقم_الهوية.Text;
                BookvConsul.Text = AttendViceConsul.Text + Environment.NewLine + AuthTitle;
                if (AppType.CheckState == CheckState.Checked)
                {
                    
                    if (النوع.CheckState == CheckState.Unchecked) 
                        BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد حضر للقنصلية ووقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                    if (النوع.CheckState == CheckState.Checked) 
                        BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد حضرت للقنصلية ووقعت بتوقيعها على هذا الإقرار بعد تلاوته عليها وبعد أن فهمت مضمونه ومحتواه. ";
                }
                else
                {
                    if (النوع.CheckState == CheckState.Unchecked) 
                        BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد وقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                    if (النوع.CheckState == CheckState.Checked) 
                        BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد وقعت بتوقيعها على هذا الإقرار بعد تلاوته عليها وبعد أن فهمت مضمونه ومحتواه. ";
                }
                BookChildDesc.Text = ChildernDescription;
                int listid = ChildDataBase.Length;
                string[] strlist = new string[4];
                if (ChildNameDesView.Text.Contains("/"))
                {
                    if (ChildNameDesView.Text.Split('/')[1].Trim().Contains("_"))
                    {
                        strlist = ChildNameDesView.Text.Split('/')[1].Trim().Split('_');
                        string chlidrenlist;
                        chlidrenlist = strlist[0];
                        for (int a = 1; a < strlist.Length; a++) chlidrenlist = chlidrenlist + " و" + strlist[a];
                        BookChildren.Text = chlidrenlist.Replace("_", " و");
                    }
                    else
                        BookChildren.Text = ChildNameDesView.Text.Split('/')[1].Trim().Replace("_"," و");
                }
                
                BookAppiIssSource.Text = مكان_الإصدار.Text;
                BookPassIqama.Text = نوع_الهوية.Text;
                BookGreData.Text = التاريخ_الميلادي_off.Text;
                colIDs[2] = التاريخ_الميلادي.Text;
                BookHijriData.Text = التاريخ_الهجري.Text;


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

                oBDoc.Activate();

                //oBDoc.Save();


                //oBMicroWord.Visible = true;





                oBDoc.SaveAs2(docxouput);
                oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
                oBDoc.Close(false, oBMiss);
                oBMicroWord.Quit(false, false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
                System.Diagnostics.Process.Start(pdfouput);
                object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;

            }

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

        private void CreateMessageWord(string MessageNo)
        {
            string ActiveCopy;
            string ReportName = DateTime.Now.ToString("mmss");
            routeDoc = FilesPathIn + "MessageCap.docx";
            loadMessageNo();
            ActiveCopy = FilesPathOut + "Message" + مقدم_الطلب.Text + ReportName + ".docx";
            if (!File.Exists(ActiveCopy))
            {
                System.IO.File.Copy(routeDoc, ActiveCopy);
                object oBMiss2 = System.Reflection.Missing.Value;
                Word.Application oBMicroWord2 = new Word.Application();



                Word.Document oBDoc2 = oBMicroWord2.Documents.Open(ActiveCopy, oBMiss2);

                Object ParacapitalMessage = "MarkcapitalMessage";
                Object ParaMApplicantName = "MarkApplicantName";
                Object ParaMassageIqrarNo = "MarkMassageIqrarNo";
                Object ParaMassageTitle = "MarkMassageTitle";
                Object ParaMassageNo = "MarkMassageNo";
                Object ParaApliSex = "MarkApliSex";
                Object ParaHijriDate = "MarkHijriDate";
                Object ParaDateGre = "MarkDateGre";
                Object ParaGregorDate2 = "MarkGregorDate2";
                Object ParaViseConsul1 = "MarkViseConsul1";

                Word.Range BookMApplicantName = oBDoc2.Bookmarks.get_Item(ref ParaMApplicantName).Range;
                Word.Range BookcapitalMessage = oBDoc2.Bookmarks.get_Item(ref ParacapitalMessage).Range;
                Word.Range BookMassageIqrarNo = oBDoc2.Bookmarks.get_Item(ref ParaMassageIqrarNo).Range;
                Word.Range BookMassageNo = oBDoc2.Bookmarks.get_Item(ref ParaMassageNo).Range;
                Word.Range BookApliSex = oBDoc2.Bookmarks.get_Item(ref ParaApliSex).Range;
                Word.Range BookDateGre = oBDoc2.Bookmarks.get_Item(ref ParaDateGre).Range;
                Word.Range BookHijriDate = oBDoc2.Bookmarks.get_Item(ref ParaHijriDate).Range;
                Word.Range BookGregorDate2 = oBDoc2.Bookmarks.get_Item(ref ParaGregorDate2).Range;
                Word.Range BookMassageTitle = oBDoc2.Bookmarks.get_Item(ref ParaMassageTitle).Range;
                Word.Range BookViseConsul1 = oBDoc2.Bookmarks.get_Item(ref ParaViseConsul1).Range;

                BookMApplicantName.Text = مقدم_الطلب.Text;
                BookcapitalMessage.Text = EmbassySource.Text;
                BookMassageNo.Text = MessageNo + (MessageDocNo + 1).ToString();
                BookMassageIqrarNo.Text = IqrarNo.Text;
                if (النوع.CheckState == CheckState.Unchecked)
                    BookApliSex.Text = "المواطن";
                else BookApliSex.Text = "المواطنة";
                BookGregorDate2.Text = BookDateGre.Text = التاريخ_الميلادي_off.Text;
                BookHijriDate.Text = التاريخ_الهجري.Text;

                switch (ProcedureType.SelectedIndex)
                {
                    case 0:
                        BookMassageTitle.Text = " إقراراً باستخراج وثائق ثبوتية ";

                        break;
                    case 1:
                        BookMassageTitle.Text = " إقراراً بعدم الممانعة من السفر ";

                        break;
                    case 2:
                        BookMassageTitle.Text = " إقراراً باستخراج وثائق ثبوتية وإقراراً بعدم الممانعة من السفر ";
                        break;
                }
                BookViseConsul1.Text = AttendViceConsul.Text;

                object rangeViseConsul1 = BookViseConsul1;
                object rangeMApplicantName = BookMApplicantName;
                object rangecapitalMessage = BookcapitalMessage;
                object rangeMassageIqrarNo = BookMassageIqrarNo;
                object rangeMassageNo = BookMassageNo;
                object rangeApliSex = BookApliSex;
                object rangeDateGre = BookDateGre;
                object rangeHijriDate = BookHijriDate;
                object rangeGregorDate2 = BookGregorDate2;
                object rangeMassageTitle = BookMassageTitle;

                oBDoc2.Bookmarks.Add("MarkViseConsul1", ref rangeViseConsul1);
                oBDoc2.Bookmarks.Add("MarkApplicantName", ref rangeMApplicantName);
                oBDoc2.Bookmarks.Add("MarkcapitalMessage", ref rangecapitalMessage);
                oBDoc2.Bookmarks.Add("MarkMassageIqrarNo", ref rangeMassageIqrarNo);
                oBDoc2.Bookmarks.Add("MarkMassageNo", ref rangeMassageNo);
                oBDoc2.Bookmarks.Add("MarkApliSex", ref rangeApliSex);
                oBDoc2.Bookmarks.Add("MarkDateGre", ref rangeDateGre);
                oBDoc2.Bookmarks.Add("MarkGregorDate2", ref rangeGregorDate2);
                oBDoc2.Bookmarks.Add("MarkHijiData", ref rangeHijriDate);
                oBDoc2.Bookmarks.Add("MarkMassageTitle", ref rangeMassageTitle);

                //oBDoc2.Activate();
                //oBDoc2.Save();
                //oBMicroWord2.Visible = true;


                string docxouput = FilesPathOut + مقدم_الطلب.Text + DateTime.Now.ToString("ssmm") + ".docx";
                string pdfouput = FilesPathOut + مقدم_الطلب.Text + DateTime.Now.ToString("ssmm") + ".pdf";
                oBDoc2.SaveAs2(docxouput);
                oBDoc2.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
                oBDoc2.Close(false);
                oBMicroWord2.Quit(false, false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord2);
                System.Diagnostics.Process.Start(pdfouput);
                object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;


                NewMessageNo();
            }

            else
            {
                MessageBox.Show("يرجى حذف الملف الموجودأولاً");
                
                btnSavePrint.Enabled = true;

            }
        }


        private void NewMessageNo()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    SqlCommand sqlCmd = new SqlCommand("UPDATE TableSettings SET MessageNo=@MessageNo WHERE ID=@ID", sqlCon);
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@ID", 1);
                    sqlCmd.Parameters.AddWithValue("@MessageNo", MessageDocNo + 1);
                    sqlCmd.ExecuteNonQuery();
                    sqlCon.Close();

                }

                catch (Exception ex)
                {
                    MessageBox.Show("الوصول لقاعدة البيانات غير متاح");
                }
                finally
                {
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

                if(combbox.Items.Count > 0) combbox.SelectedIndex = 0;
            }
        }
        private void CreateWordFile(bool caseDoc)
        {
            string ReportName = DateTime.Now.ToString("mmss");
            ModelFileroute = FilesPathIn + "Igrar_TravM.docx";

            if (النوع.CheckState == CheckState.Checked)
            {
                ApplicantSexStatus = false;
                labelName.Text = "مقدمة الطلب:";
                labelName.ForeColor = Color.Black;
                ModelFileroute = FilesPathIn + "Igrar_TravF.docx";
            }

            string CurrentCopy = FilesPathOut + مقدم_الطلب.Text + ReportName + ".docx";
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
                object ParaDocType = "MarkDocType";
                object ParaEmbassyFrom = "MarkEmbassyFrom";

                

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
                Word.Range BookDocType = oBDoc.Bookmarks.get_Item(ref ParaDocType).Range;
                Word.Range BookEmbassyFrom = oBDoc.Bookmarks.get_Item(ref ParaEmbassyFrom).Range;

                if (caseDoc)
                    BookDocType.Text = "إقرار";
                else
                    BookDocType.Text = "إقرار مشفوع باليمين";
                BookIqrarNo.Text = IqrarNo.Text;
                BookHijriData.Text = التاريخ_الهجري.Text;
                BookGreData.Text = التاريخ_الميلادي_off.Text;
                BookAppiIssSource.Text = مكان_الإصدار.Text;
                Bookname1.Text = Bookname2.Text = colIDs[3] = مقدم_الطلب.Text;
                Bookigama.Text = رقم_الهوية.Text;
                BookvConsul.Text = AttendViceConsul.Text + Environment.NewLine + AuthTitle; 
                BookMention.Text = Mentioned;
                BookMarkEmbassy.Text = EmbassySource.Text;
                colIDs[5] = AppType.Text;
                colIDs[6] = mandoubName.Text;
                if (TravellerDescrib.Text == "ابناء فقط")
                {
                    int listid = ChildDataBase.Length;
                    string[] strlist = new string[4];

                    if (ChildDataBase.Contains("_"))
                    {
                        strlist = ChildDataBase.Split('_');
                        string chlidrenlist;
                        chlidrenlist = strlist[0];
                        for (int a = 1; a < strlist.Length; a++) chlidrenlist = chlidrenlist + " و" + strlist[a];
                        BookChildren.Text = ChildernDescription + "/" + chlidrenlist.Replace("_", " و"); ;
                    }
                    else
                        BookChildren.Text = ChildernDescription + "/" + ChildDataBase.Replace("_", " و"); ;




                }
                else if (TravellerDescrib.Text == "ابناء برفقة مرافق غير الزوجة")
                {
                    int listid = ChildDataBase.Length;
                    string[] strlist = new string[4];
                    if (ChildNameDesView.Text.Contains("/"))
                    {

                        if (ChildNameDesView.Text.Split('/')[1].Trim().Contains("_"))
                        {
                            strlist = ChildNameDesView.Text.Split('/')[1].Trim().Split('_');
                            string chlidrenlist;
                            chlidrenlist = strlist[0];
                            for (int a = 1; a < strlist.Length; a++) 
                                chlidrenlist = chlidrenlist + " و" + strlist[a];

                            BookChildren.Text = ChildernDescription + "/" + chlidrenlist.Replace("_", " و") + " " + "برفقة " + TravellerAttenDescrib.Text + "/ " + TravellerAttenName.Text;
                        }
                        else

                            BookChildren.Text = ChildernDescription + "/" + ChildNameDesView.Text.Split('/')[1].Trim().Replace("_", " و") + " " + "برفقة " + TravellerAttenDescrib.Text + "/ " + TravellerAttenName.Text;
                    }

                }
                else if (TravellerDescrib.Text == "زوجة فقط")
                {
                    BookChildren.Text = comboPersonUnderPro.Text + " /" + personUnderPro.Text;
                    BookMention.Text = "لمذكورة";
                }
                else if (TravellerDescrib.Text == "آخرين")
                {
                    BookChildren.Text = "زوجتي /" + personUnderPro.Text;
                    BookMention.Text = "لمذكورة";
                }
                else if (TravellerDescrib.Text == "زوجة وابناء")
                {
                    
                    string[] strlist = new string[4];
                    if (ChildNameDesView.Text.Contains("/"))
                    {
                        int listid = ChildNameDesView.Text.Split('/')[1].Trim().Length;

                        if (ChildNameDesView.Text.Split('/')[1].Trim().Contains("_"))
                        {
                            strlist = ChildNameDesView.Text.Split('/')[1].Trim().Split('_');
                            string chlidrenlist;
                            chlidrenlist = strlist[0];
                            for (int a = 1; a < strlist.Length; a++) chlidrenlist = chlidrenlist + " و" + strlist[a];
                            BookChildren.Text = ChildernDescription + "/" + chlidrenlist.Replace("_"," و") + " " + "برفقة زوجتي " + "/" + personUnderPro.Text;
                        }
                        else

                            BookChildren.Text = ChildernDescription + "/" + ChildDataBase.Replace("_", " و") + " " + "برفقة زوجتي " + "/" + personUnderPro.Text;
                    }
                    
                    if (Mentioned == "ابنتي")
                    {
                        BookMention.Text = "لمذكورتين";
                    }
                    else BookMention.Text = "لمذكورين";

                }

                BookCountDestin.Text = TravelDestin.Text;

                if (TravelDestin.Text == "المملكة العربية السعودية")
                    BookCountryDesc.Text = "قدوم";
                else
                    BookCountryDesc.Text = "سفر";
                if (TravellingPurpo.Text == "العمل")
                {
                    if (caseDoc)
                        BookTravelPurpose.Text = "الإقامة معي";
                    else
                        BookTravelPurpose.Text = "العمل بمهنة" + " " + familyJob.Text;
                }
                else
                {
                    BookTravelPurpose.Text = TravellingPurpo.Text;
                }

                BookPassIqama.Text = نوع_الهوية.Text;
                if (AppType.CheckState == CheckState.Checked)
                {
                    if (النوع.CheckState == CheckState.Unchecked) BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle+ "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد حضر للقنصلية ووقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                    if (النوع.CheckState == CheckState.Checked) BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle+ "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد حضرت للقنصلية ووقعت بتوقيعها على هذا الإقرار بعد تلاوتها عليها وبعد أن فهمت مضمونه ومحتواه. ";
                }
                else
                {

                    if (النوع.CheckState == CheckState.Unchecked)
                        BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد وقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                    if (النوع.CheckState == CheckState.Checked)
                        BookAuthorization.Text = "أشهد أنا/" + AttendViceConsul.Text + " " + AuthTitle + "  بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد وقعت بتوقيعها على هذا الإقرار بعد تلاوته عليها وبعد أن فهمت مضمونه ومحتواه. ";

                }

                BookEmbassyFrom.Text = "سفارة " + TravelDestin.Text;

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
                object rangeDocType = BookDocType;
                object rangeEmbassyFrom = BookEmbassyFrom;

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
                oBDoc.Bookmarks.Add("MarkDocType", ref rangeDocType);
                oBDoc.Bookmarks.Add("MarkEmbassyFrom", ref rangeEmbassyFrom);
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

                oBDoc.Activate();
                //oBDoc.Save();

                //oBMicroWord.Visible = true;
                string docxouput = FilesPathOut + مقدم_الطلب.Text + DateTime.Now.ToString("ssmm") + ".docx";
                string pdfouput = FilesPathOut + مقدم_الطلب.Text + DateTime.Now.ToString("ssmm") + ".pdf";
                oBDoc.SaveAs2(docxouput);
                oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
                oBDoc.Close(false, oBMiss);
                oBMicroWord.Quit(false, false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
                System.Diagnostics.Process.Start(pdfouput);
                object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
            }
            else
            {
                MessageBox.Show("يرجى حذف الملف الموجودأولاً");                
                btnSavePrint.Enabled = true;
                i = 0;
            }
       
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
                //MessageBox.Show(allList[i] +" - "+text[i - 1]);
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

        private void TravellerDescrib_SelectedIndexChanged_1(object sender, EventArgs e)
        {

            if (TravellerDescrib.Text == "ابناء فقط")
            {
                ChildrenOnly();

            }
            else if (TravellerDescrib.Text == "زوجة فقط")
            {
                WifeOnly();
            }
            else if (TravellerDescrib.Text == "زوجة وابناء")
            {
                wifeAndChildren();
            }
            if (TravellerDescrib.Text == "ابناء برفقة مرافق غير الزوجة")
            {
                ChildrenWithAttend();
            }
        }

        private void ChildrenWithAttend()
        {
            Attendecheck.Checked = true;
            personUnderPro.Visible = false;
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
            groupBox1.Visible = true;
        }

        private void wifeAndChildren()
        {
            personUnderPro.Visible = true;
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
            personUnderPro.Visible = true;
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
            personUnderPro.Visible = false;
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

        private void Attendecheck_CheckedChanged_1(object sender, EventArgs e)
        {
            if (Attendecheck.CheckState == CheckState.Unchecked)
            {

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
            التاريخ_الميلادي.Text = DateTime.Now.ToString("MM-dd-yyyy");
            timer2.Enabled = false;
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
        
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            if (printPreviewDialog1.ShowDialog() == DialogResult.OK) printDocument1.Print();
        }


        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            FillDataGridView();
        }


        private void AddChildren_Click_1(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "")
            {

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
                    Mentioned = "لمذكور";
                }
                else
                {

                    ChildernDescription = "ابنتي";
                    Mentioned = "لمذكورة";
                }
                ChildDataBase = ChildName[childindex];
                childindex = 1;
            }
            else if (childindex == 1)
            {
                if (childboygirls.CheckState == CheckState.Checked && ChildernDescription == "ابني")
                {
                    ChildernDescription = "ابنيَّ";
                    Mentioned = "لمذكورين";
                }
                else if (childboygirls.CheckState == CheckState.Unchecked && ChildernDescription == "ابنتي")
                {
                    ChildernDescription = "ابنتيَّ";
                    Mentioned = "لمذكورتين";
                }
                else
                {
                    ChildernDescription = "ابنائي";
                    Mentioned = "لمذكورين";
                }
                ChildDataBase = ChildDataBase + "_" + ChildName[1];
                childindex = 2;
            }
            else if (childindex == 2)
            {
                if (childboygirls.CheckState == CheckState.Unchecked && ChildernDescription == "ابنتيَّ")
                {
                    ChildernDescription = "بناتي";
                    Mentioned = "لمذكورات";
                }
                else
                {
                    ChildernDescription = "أبنائي";
                    Mentioned = "لمذكورين";
                }

                ChildDataBase = ChildDataBase + "_" + ChildName[2];
                childindex = 3;
            }
            else
            {
                if (childboygirls.CheckState == CheckState.Unchecked && ChildernDescription == "بناتي")
                {
                    ChildernDescription = "بناتي";
                    Mentioned = "لمذكورات";
                }
                else
                {
                    ChildernDescription = "أبنائي";
                    Mentioned = "لمذكورين";
                }
                ChildDataBase = ChildDataBase + "_" + ChildName[childindex];
                Mentioned = "لمذكورين";
                childindex++;

            }




            if (ChildDataBase.Contains("_")) ChildNameDesView.Text = ChildernDescription + "/ " + ChildDataBase.Replace("_", " و");
            else ChildNameDesView.Text = ChildernDescription + "/ " + ChildDataBase;

            ChildrenName.Clear();
            //if (textBox1.Text != "" && textBox2.Text != "")
            //{


            //    string ChildrenList = textBox2.Text;
            //    ChildernDescription = textBox1.Text;
            //    if (ChildrenList.Contains("_"))
            //    {
            //        ChildName = ChildrenList.Split('_');
            //        childindex = ChildName.Length;

            //        //News.Text = "Contains";
            //    }
            //    else
            //    {
            //        childindex = 1;
            //        ChildDataBase = ChildrenList;
            //    }

            //    for (int i = 1; i < childindex; i++)
            //    {
            //        ChildDataBase = ChildDataBase + "_" + ChildName[i];
            //    }
            //    textBox1.Text = textBox2.Text = "";
            //}
            //ChildName[childindex] = ChildrenName.Text;
            //if (childindex == 0)
            //{
            //    if (childboygirls.CheckState == CheckState.Checked)
            //    {

            //        ChildernDescription = "ابني";
            //        Mentioned = "للمذكور";
            //    }
            //    else
            //    {

            //        ChildernDescription = "ابنتي";
            //        Mentioned = "للمذكورة";
            //    }
            //    ChildDataBase = ChildName[childindex];

            //}
            //else if (childindex == 1)
            //{
            //    if (childboygirls.CheckState == CheckState.Checked && ChildernDescription == "ابني")
            //    {
            //        ChildernDescription = "ابنيَّ";
            //        Mentioned = "للمذكورين";
            //    }
            //    else if (childboygirls.CheckState == CheckState.Unchecked && ChildernDescription == "ابنتي")
            //    {
            //        ChildernDescription = "ابنتيَّ";
            //        Mentioned = "للمذكورتين";
            //    }                
            //    ChildDataBase = ChildDataBase + "_" + ChildName[1];
            //}
            //else if(childindex >= 2)
            //{
            //    if (childboygirls.CheckState == CheckState.Unchecked && ChildernDescription == "ابنتيَّ")
            //    {
            //        ChildernDescription = "بناتي";
            //        Mentioned = "للذكورات";
            //    }
            //    else
            //    {
            //        ChildernDescription = "أبنائي";
            //        Mentioned = "للمذكورين";
            //    }
            //    ChildDataBase = ChildDataBase + "_" + ChildName[childindex];
            //}
            //for (int j = 1; j < childindex; j++)
            //{
            //    ChildDataBase = ChildDataBase + "_" + ChildName[j];
            //}
            //if (ChildDataBase.Contains("_")) ChildNameDesView.Text = ChildernDescription + "/ " + ChildDataBase.Replace("_", " و");
            //else ChildNameDesView.Text = ChildernDescription + "/ " + ChildDataBase;
            //childindex++;
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
                mandoubName.Text = "";
                mandoubName.Visible = false;
                mandoubLabel.Visible = panel1.Visible = false;
            }
            else
            {
                AppType.Text = "عن طريق أحد مندوبي القنصلية";                    
                mandoubName.Visible = true;
                mandoubLabel.Visible = panel1.Visible = true;
            }
        }

        private void SearchDoc_Click_1(object sender, EventArgs e)
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
        private void OpenFile(int id, int fileNo)
        {
            string query;

            SqlConnection Con = new SqlConnection(DataSource);
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

        private void dataGridView1_DoubleClick_1(object sender, EventArgs e)
        {

        }

        private void printOnly_Click(object sender, EventArgs e)
        {

        }

        private void Search_TextChanged(object sender, EventArgs e)
        {

        }




        private int loadIDNo(string table)
        {


            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from " + table + " order by ID desc", sqlCon);
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





        private void btnSavePrint_Click_1(object sender, EventArgs e)
        {
            getTitle(DataSource, AttendViceConsul.Text);
            التاريخ_الميلادي.Text = GregorianDate ;
            التاريخ_الهجري.Text = HijriDate ;

            if (!checkGender(PanelMain, "مقدم_الطلب", "النوع"))
            {
                return;
            }
            else addNewAppNameInfo(مقدم_الطلب);
            
            if (تاريخ_الميلاد.Text == "")
            {
                MessageBox.Show("يرجى إضافة تاريخ ميلاد مقدم الطلب"); return;
            }
            if (المهنة.Text == "")
            {
                MessageBox.Show("يرجى إختيار مهنة مقدم الطلب"); return;
            }
            //IqrarNo.Text = "ق س ج/160/02/" + loadRerNo(loadIDNo("TableTravIqrar"));
            btnSavePrint.Text = "جاري المعالجة";
            btnSavePrint.Enabled = false;
            Save2DataBase(SaveEdit);

            if (ProcedureType.SelectedIndex == 0 || ProcedureType.SelectedIndex == 2 || ProcedureType.SelectedIndex == 3 || ProcedureType.SelectedIndex == 4) {
                CreateWordFileDoc();

                
            }
            if (ProcedureType.SelectedIndex != 0)
            {
                CreateWordFile(false);
                CreateWordFile(true);

                
            }
            colIDs[2] = التاريخ_الميلادي.Text;
            colIDs[3] = مقدم_الطلب.Text;
            colIDs[5] = AppType.Text;
            colIDs[6] = mandoubName.Text;
            addarchives(colIDs);
            if (EmbassySource.Text != "الخرطوم") CreateMessageWord("ق س ج/80/" + DateTime.Now.Year.ToString().Replace("20", "") + "/" + "02" + "/");
            this.Close();   

            //Clear_Fields();
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

        private void PassIqama_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (نوع_الهوية.Text == "جواز سفر ")
            {
                رقم_الهوية.Text = "P0";
                autoCompleteTextBox(مكان_الإصدار, DataSource, "SDNIssueSource", "TableListCombo");
            }
            else if (نوع_الهوية.Text == "اقامة ")
            {
                رقم_الهوية.Text = "";
                autoCompleteTextBox(مكان_الإصدار, DataSource, "KSAIssureSource", "TableListCombo");
            }
            else
            {
                رقم_الهوية.Text = "";
            }
        }

        private void TravellingPurpo_SelectedIndexChanged_2(object sender, EventArgs e)
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

        private void Form2_Load(object sender, EventArgs e)
        {
            
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
        
        //private void OpenFileDoc(int id, int fileNo)
        //{
        //    string query;

        //    SqlConnection Con = new SqlConnection(DataSource);
        //    if (fileNo == 1)
        //    {
        //        query = "select Data1, Extension1,FileName1 from TableTravIqrar  where ID=@id";
        //    }
        //    else if (fileNo == 2)
        //    {
        //        query = "select Data2, Extension2,FileName2 from TableTravIqrar  where ID=@id";
        //    }
        //    else query = "select Data3, Extension3,FileName3 from TableTravIqrar  where ID=@id";
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                dataGridView1.Visible = false;
                PanelFiles.Visible = true;
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
                AppType.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
                
                mandoubName.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
                //MessageBox.Show(mandoubName.Text); 
                if (AppType.Text == "حضور مباشرة إلى القنصلية")
                {
                    AppType.CheckState = CheckState.Checked;
                    mandoubName.Text = "";
                }
                else AppType.CheckState = CheckState.Unchecked;
                if (dataGridView1.CurrentRow.Cells[2].Value.ToString() == "")
                {
                    SaveEdit = false;
                    colIDs[7] = "new";
                    IqrarNo.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                    //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 1);
                    FillDatafromGenArch("data1", dataGridView1.CurrentRow.Cells[0].Value.ToString(), "TableTravIqrar");
                    if (UserJobposition.Contains("قنصل")) deleteRow.Visible = true;
                    gridFill = false;
                    return;
                }
                gridFill = false;
                colIDs[7] = "old";
                ChildDataBase = ""; 
                ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                if (UserJobposition.Contains("قنصل")) deleteRow.Visible = true;
                IqrarNo.Text = CurrentIqrarId = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                مقدم_الطلب.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString().ToString();
                if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "ذكر") النوع.CheckState = CheckState.Unchecked;
                else if (dataGridView1.CurrentRow.Cells[3].Value.ToString().ToString() == "أنثى") النوع.CheckState = CheckState.Checked;
                نوع_الهوية.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().ToString();
                رقم_الهوية.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString().ToString();
                مكان_الإصدار.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString().ToString();
                string[] nameMention = new string[2];
                if (dataGridView1.CurrentRow.Cells[7].Value.ToString().Contains("_"))
                {
                    nameMention = dataGridView1.CurrentRow.Cells[7].Value.ToString().Split('_');
                    ChildernDescription = nameMention[0];
                    Mentioned = nameMention[1];
                }
                else ChildernDescription = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                string ChildrenList = dataGridView1.CurrentRow.Cells[8].Value.ToString().ToString();
                if (dataGridView1.CurrentRow.Cells[8].Value.ToString().Contains("_"))
                {
                    ChildFromDataBase = ChildrenList.Split('_');
                    childindex = ChildFromDataBase.Length;
                }
                else
                {
                    childindex = 1;
                    ChildDataBase = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                }

                //for (int i = 1; i < childindex; i++)
                //{
                //    ChildDataBase = ChildDataBase + "_" + ChildFromDataBase[i];
                //}
                ChildDataBase = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                ChildNameDesView.Text = ChildernDescription + "/ " + ChildDataBase;
                textBox1.Text = ChildernDescription;
                textBox2.Text = ChildDataBase;
                personUnderPro.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                EmbassySource.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                التاريخ_الميلادي.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                التاريخ_الهجري.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[14].Value.ToString() == "غير معالج")
                {
                    checkedViewed.CheckState = CheckState.Unchecked;
                    //IqrarNo.Text = NextRelId;
                }
                else checkedViewed.CheckState = CheckState.Checked;
                if (checkedViewed.CheckState == CheckState.Checked)
                {
                    PreAppId = CurrentIqrarId;
                }
                else
                {
                    PreAppId = "";
                    // IqrarNo.Text = CurrentIqrarId;
                }
                TravelDestin.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString().ToString();
                string[] str = new string[3];
                TravellingPurpo.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();
                if (dataGridView1.CurrentRow.Cells[16].Value.ToString().ToString().Contains("_"))
                {
                    str = dataGridView1.CurrentRow.Cells[16].Value.ToString().ToString().Split('_');
                    TravellingPurpo.Text = str[0];
                    TravellerDescrib.Text = str[1];
                    familyJob.Text = str[2];
                    if (str[1] == "زوجة فقط" || str[1] == "زوجة وابناء")
                    {
                        personUnderPro.Visible = true;
                        labelwifeName.Visible = true;
                    }
                }


                
                

                ConsulateEmployee.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
                PreRelatedID = dataGridView1.CurrentRow.Cells[20].Value.ToString();
                Comment.Text = dataGridView1.CurrentRow.Cells[25].Value.ToString();

                if (dataGridView1.CurrentRow.Cells[26].Value.ToString() != "غير مؤرشف")
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
                ProcedureType.SelectedIndex = Convert.ToInt32(dataGridView1.CurrentRow.Cells[27].Value.ToString());
                TravellerAttenName.Text = dataGridView1.CurrentRow.Cells[28].Value.ToString();
                TravellerAttenDescrib.Text = dataGridView1.CurrentRow.Cells[29].Value.ToString();
                المهنة.Text = dataGridView1.CurrentRow.Cells["المهنة"].Value.ToString();
                تاريخ_الميلاد.Text = dataGridView1.CurrentRow.Cells["تاريخ_الميلاد"].Value.ToString();
                ArchivedSt.Visible = true;
                
                SaveEdit = false;

            }
        }

        private void deleteRow_Click_1(object sender, EventArgs e)
        {
            deleteRowsData(ApplicantID, "TableTravIqrar", DataSource);
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

        private void ProcedureType_SelectedIndexChanged(object sender, EventArgs e)
        {

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
                النوع.Text = "أنثى";
                labelName.Text = "مقدمة الطلب:";

            }
        }

        private void mandoubName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void AttendViceConsul_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void printOnly_Click_1(object sender, EventArgs e)
        {
            if (ProcedureType.SelectedIndex == 0 || ProcedureType.SelectedIndex == 2 || ProcedureType.SelectedIndex == 3 || ProcedureType.SelectedIndex == 4)
            {
                CreateWordFileDoc();
            }
            if (ProcedureType.SelectedIndex != 0)
            {
                CreateWordFile(false);
                CreateWordFile(true);
            }
            if (EmbassySource.Text != "الخرطوم") CreateMessageWord("ق س ج/80/" + التاريخ_الميلادي.Text.Split('-')[2].Replace("20", "") + "/02/");
            Clear_Fields();            
        }

       

        private void button2_Click_1(object sender, EventArgs e)
        {

        }

        private void ProcedureType_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '" + ListSearch.Text + "%'";
            dataGridView1.DataSource = bs; 
            ColorFulGrid9();
        }

        private void ListSearch_TextChanged(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '" + ListSearch.Text + "%'";
            dataGridView1.DataSource = bs;
            ColorFulGrid9();
        }

        private void Search_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void Form2_Load_1(object sender, EventArgs e)
        {
            FillDataGridView();
            
            
            fileComboBox(نوع_الهوية, DataSource, "DocType", "TableListCombo");
            autoCompleteTextBox(مكان_الإصدار, DataSource, "SDNIssueSource", "TableListCombo");
            autoCompleteTextBox(المهنة, DataSource, "jobs", "TableListCombo");
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
        private void button1_Click_1(object sender, EventArgs e)
        {
            ChildNameDesView.Enabled = true;
        }

        private void btnListView_Click(object sender, EventArgs e)
        {
            Clear_Fields(); 
            FillDataGridView();
            if (dataGridView1.Visible)
            {
                dataGridView1.Visible = false;
                PanelFiles.Visible = true;
                PanelMain.Visible = true;
            }
            else
            {
                dataGridView1.Visible = true;
                PanelFiles.Visible = false;
                PanelMain.Visible = false;
            }
        }

        private void SearchDoc_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            Search.Text = dlg.FileName;
        }

        private void btnFile1_Click(object sender, EventArgs e)
        {
            //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 1);
            FillDatafromGenArch("data1", dataGridView1.CurrentRow.Cells[0].Value.ToString(), "TableTravIqrar");
        }

        private void btnFile2_Click(object sender, EventArgs e)
        {
            //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 2);
            FillDatafromGenArch("data2", dataGridView1.CurrentRow.Cells[0].Value.ToString(), "TableTravIqrar");
        }

        private void btnFile3_Click(object sender, EventArgs e)
        {
            //OpenFileDoc(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()), 3);
        }

        private void PanelMain_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
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
        string lastInput2 = "";
        private void تاريخ_الميلاد_TextChanged(object sender, EventArgs e)
        {            
            if (تاريخ_الميلاد.Text.Length == 10 && !gridFillauto)
            {
                int month = Convert.ToInt32(SpecificDigit(تاريخ_الميلاد.Text, 1, 2));
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //تاريخ_الميلاد.Text = "";
                    تاريخ_الميلاد.Text = SpecificDigit(تاريخ_الميلاد.Text, 3, 10);
                    return;
                }
                gridFillauto = false;
            }
            
            if (تاريخ_الميلاد.Text.Length == 11)
            {
                تاريخ_الميلاد.Text = lastInput2; return;
            }
            if (تاريخ_الميلاد.Text.Length == 10) return;
            if (تاريخ_الميلاد.Text.Length == 4) تاريخ_الميلاد.Text = "-" + تاريخ_الميلاد.Text;
            else if (تاريخ_الميلاد.Text.Length == 7) تاريخ_الميلاد.Text = "-" + تاريخ_الميلاد.Text;
            lastInput2 = تاريخ_الميلاد.Text;
        }
        private string SpecificDigit(string text, int Firstdigits, int Lastdigits)
        {
            char[] characters = text.ToCharArray();
            string firstNchar = "";
            int z = 0;
            for (int x = Firstdigits - 1; x < Lastdigits && x < text.Length; x++)
            {
                firstNchar = firstNchar + characters[x];

            }
            return firstNchar;
        }
        private void label36_Click(object sender, EventArgs e)
        {

        }

        private void التاريخ_ValueChanged(object sender, EventArgs e)
        {
            
        }

        private void المهنة_TextChanged(object sender, EventArgs e)
        {

        }

        private void label37_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            if (!checkGender(PanelMain, "مقدم_الطلب", "النوع"))
            {
                return;
            }
            
            if (تاريخ_الميلاد.Text == "")
            {
                MessageBox.Show("يرجى إضافة تاريخ ميلاد مقدم الطلب"); return;
            }
            if (المهنة.Text == "")
            {
                MessageBox.Show("يرجى إختيار مهنة مقدم الطلب"); return;
            }
            //IqrarNo.Text = "ق س ج/160/02/" + loadRerNo(loadIDNo("TableTravIqrar"));
            btnSavePrint.Text = "جاري المعالجة";
            btnSavePrint.Enabled = false;
            Save2DataBase(SaveEdit);
            colIDs[2] = التاريخ_الميلادي.Text;
            colIDs[3] = مقدم_الطلب.Text;
            colIDs[5] = AppType.Text;
            colIDs[6] = mandoubName.Text;

            addarchives(colIDs);
            if (EmbassySource.Text != "الخرطوم") CreateMessageWord("ق س ج/80/" + DateTime.Now.Year.ToString().Replace("20", "") + "/" + "02" + "/");

            Clear_Fields();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (تاريخ_الميلاد.Text.Length != 10)
            {
                MessageBox.Show("يرجى إدخال تاريخ ميلاد مقدم الطلب أولا");
                return;
            }
            
            updateGenName(ApplicantID.ToString(), تاريخ_الميلاد.Text, المهنة.Text, DataSource);
            تاريخ_الميلاد.Text= المهنة.Text = "";
            btnListView.PerformClick();
        }
        private void updateGenName(string idDoc, string birth, string job, string source)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string query = "update TableTravIqrar set تاريخ_الميلاد=N'" + birth + "',المهنة=N'" + job + "' where ID = '" + idDoc + "'";
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void تاريخ_الميلاد_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void التاريخ_الميلادي_TextChanged(object sender, EventArgs e)
        {
            التاريخ_الميلادي_off.Text = التاريخ_الميلادي.Text.Split('-')[1] + " - " + التاريخ_الميلادي.Text.Split('-')[0] + " - " + التاريخ_الميلادي.Text.Split('-')[2];
        }

        private void مقدم_الطلب_TextChanged(object sender, EventArgs e)
        {
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
                gridFillauto = true;
                رقم_الهوية_1.Text = row["رقم_الهوية"].ToString();
                نوع_الهوية_1.Text = row["نوع_الهوية"].ToString();
                مكان_الإصدار_1.Text = row["مكان_الإصدار"].ToString();
                المهنة_1.Text = row["المهنة"].ToString();
                تاريخ_الميلاد_1.Text = row["تاريخ_الميلاد"].ToString();
                النوع_1.Text = row["النوع"].ToString();
                return;
            }
            //MessageBox.Show(رقم_الهوية_1.Text);
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
            txtEditID1.Visible = txtEditID2.Visible = true;
            if (btnEditID.Text == "إجراء")
            {
                btnEditID.Text = "تعديل";

                SqlConnection sqlCon = new SqlConnection(DataSource);
                SqlCommand sqlCmd = new SqlCommand("update TableTravIqrar SET DocID = @DocID WHERE ID = @ID", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                sqlCmd.Parameters.AddWithValue("@DocID", txtEditID1.Text + txtEditID2.Text);
                sqlCmd.ExecuteNonQuery();
                sqlCon.Close();
            }
            else
                btnEditID.Text = "إجراء";
        }

        private void deleteRow_Click(object sender, EventArgs e)
        {
            var selectedOption = MessageBox.Show("", "تأكيد عملية الحذف", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {

                deleteRowsData(ApplicantID, "TableTravIqrar", DataSource);
                deleteRow.Visible = false;
                FillDataGridView();
                dataGridView1.Visible = false;
                PanelFiles.Visible = true;
                PanelMain.Visible = true;
            }
        }

       
        private void childboygirls_CheckedChanged_1(object sender, EventArgs e)
        {
            if (childboygirls.CheckState == CheckState.Checked) childboygirls.Text = "ابن";
            else childboygirls.Text = "ابنة";
        }

        

   

       

        private void button1_Click(object sender, EventArgs e)
        {
            Clear_Fields();
        }

        private void Save2DataBase(bool newData)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);

            string AppGender;
            if (النوع.CheckState == CheckState.Unchecked) AppGender = "ذكر";
            else AppGender = "أنثى";
            if (AppType.CheckState == CheckState.Checked)
            {
                mandoubName.Text = "";
            }
                try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                if (newData)
                {
                    
                    if (checkedViewed.CheckState == CheckState.Checked) Viewed = "تمت المعالجة بواسطة " + ConsulateEmpName;
                    else Viewed = "غير معالج";
                    SqlCommand sqlCmd = new SqlCommand("TravAddorEdit", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@ID", 0);
                    sqlCmd.Parameters.AddWithValue("@mode", "Add");
                    sqlCmd.Parameters.AddWithValue("@DocID", "ق س ج/80/" + DateTime.Now.Year.ToString().Replace("20", "") + "/02/"  +loadRerNo(loadIDNo()));
                    sqlCmd.Parameters.AddWithValue("@AppName", مقدم_الطلب.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", نوع_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", رقم_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", مكان_الإصدار.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@ChildDesc", ChildernDescription.Trim() + "_" + Mentioned);
                    if (ChildNameDesView.Text.Contains("/"))
                        sqlCmd.Parameters.AddWithValue("@ChildNames", ChildNameDesView.Text.Split('/')[1].Trim());
                    else
                        sqlCmd.Parameters.AddWithValue("@ChildNames", ChildNameDesView.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Embassy", EmbassySource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@CountDestin", TravelDestin.Text);
                    sqlCmd.Parameters.AddWithValue("@TravelPurpose", TravellingPurpo.Text + "_" + TravellerDescrib.Text + "_" + familyJob.Text);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text);
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", "");
                    

                    sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@WifeName", personUnderPro.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@ProType", ProcedureType.SelectedIndex);
                    sqlCmd.Parameters.AddWithValue("@AttendName", TravellerAttenName.Text);
                    sqlCmd.Parameters.AddWithValue("@AttendDesc", TravellerAttenDescrib.Text);
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
                    sqlCmd.Parameters.AddWithValue("@AppName", مقدم_الطلب.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Gender", AppGender.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocType", نوع_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocNo", رقم_الهوية.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DocSource", مكان_الإصدار.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@ChildDesc", ChildernDescription.Trim() + "_" + Mentioned);
                    if (ChildNameDesView.Text.Contains("/"))
                        sqlCmd.Parameters.AddWithValue("@ChildNames", ChildNameDesView.Text.Split('/')[1].Trim());
                    else
                        sqlCmd.Parameters.AddWithValue("@ChildNames", ChildNameDesView.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Embassy", EmbassySource.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@GriDate", التاريخ_الميلادي.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Hijri", التاريخ_الهجري.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@AtteVicCo", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@Viewed", Viewed);
                    sqlCmd.Parameters.AddWithValue("@DataInterType", AppType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@DataInterName", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@DataMandoubName", mandoubName.Text);
                    sqlCmd.Parameters.AddWithValue("@CountDestin", TravelDestin.Text);
                    sqlCmd.Parameters.AddWithValue("@TravelPurpose", TravellingPurpo.Text + "_" + TravellerDescrib.Text + "_" + familyJob.Text);
                    sqlCmd.Parameters.AddWithValue("@RelatedApp", PreAppId.Trim());
                    sqlCmd.Parameters.AddWithValue("@المهنة", المهنة.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@تاريخ_الميلاد", تاريخ_الميلاد.Text.Trim());
                    
                    

                    if (fileloaded)
                        sqlCmd.Parameters.AddWithValue("@ArchivedState", ConsulateEmpName.Trim() + " " + DateTime.Now.ToString("hh:mm"));
                    else sqlCmd.Parameters.AddWithValue("@ArchivedState", "غير مؤرشف");
                    sqlCmd.Parameters.AddWithValue("@Comment", Comment.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@WifeName", personUnderPro.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@ProType", ProcedureType.SelectedIndex);
                    sqlCmd.Parameters.AddWithValue("@AttendName", TravellerAttenName.Text);
                    sqlCmd.Parameters.AddWithValue("@AttendDesc", TravellerAttenDescrib.Text);
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


