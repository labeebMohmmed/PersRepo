using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Configuration;
using System.Globalization;
using System.Threading;
using System.IO;
using System.Drawing.Printing;
using System.Drawing.Imaging;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Xceed.Words.NET;
using System.Diagnostics;
using Xceed.Document.NET;
using System.Globalization;
using System.Threading;
using System;
using System.Runtime.InteropServices;
using WIA;
using PersAhwal;
using System.Net;
using Image = System.Drawing.Image;
using Microsoft.Reporting.WinForms;
using Microsoft.Office.Core;

namespace PersAhwal
{
    public partial class NoteVerbal : Form
    {
        string[] colIDs = new string[100]; 
        string JobPosition;
        string DataSource;
        string FilespathIn;
        string FilespathOut;
        string EmpName;
        int Atvc = 0;
        string PrimariFiles = @"D:\PrimariFiles\";
        bool ArchiveState = false;
        string MessNoPart = "";
        DeviceInfo AvailableScanner = null;
        string[] PathImage = new string[100];
        string rowCount = "";
        int imagecount = 0;
        string ActiveCopy = "";
        string GregorianDate = "";
        string GregorianDate1 = "";
        string CurrentFile1 = "";
        string CurrentFile2 = "";
        bool finalArch = false;
        string docIDNumber = "";
        string archstat = "";
        int ID = 0;
        string GreDate = "";
        string actionDone = "";
        string DocNoPro = "";
        int ResponceIndex = 1;
        string responces = "";
        string reference = "";
        string recommon = "";
        string currentNo = "";
        string HijriDate = "";
        string CurrentBtnName = "";
        bool grdiFill = false;
        bool newData = false;
        string[,] preffix = new string[20, 20];
        int contextID = 0;
        bool oldOne = false;
        string حالة_الأجراء = "";
//        string fileNo = "104";
        int VCIndex = 0;
        bool ModifyPermit = true;
        string colList = "";
        string refFile = "";
        bool fileUpdate = false;
        public NoteVerbal(bool modifyPermit,int vcIndex,string greDate, string hijriDate,string jobPosition, string dataSource,string filespathIn, string filespathOut, string empName, int atvc, bool archiveState)
        {
            InitializeComponent();
            Console.WriteLine(1);
            ModifyPermit = modifyPermit;
            HijriDate = hijriDate;
            GreDate = greDate;
            JobPosition = jobPosition;
            DataSource = dataSource;
            VCIndex = vcIndex;
            FilespathIn = filespathIn;
            FilespathOut = filespathOut;
            colIDs[4] = labEmp.Text = EmpName = empName;
            colIDs[5] = المصدر.Text;
            colIDs[6] = "";
            colIDs[7] = "new";
            Atvc = atvc;
            ArchiveState = archiveState;
            الإجراء_الذي_تم.Visible = true;
            combAction.SelectedIndex = 0;
           
            if (!Directory.Exists(PrimariFiles))
            {
                string appFileName = Environment.GetCommandLineArgs()[0];
                string directory = Path.GetDirectoryName(appFileName);
                directory = directory + @"\";
                PrimariFiles = directory + @"PrimariFiles\";
            }
            Suffex_preffixList();
            colList = checkColumnName("TableMessages");
            
            if (ArchiveState)
            {
                
                DocIDGenerator();
                ID = 0;
                //checkID();
                btnArch.Visible = loadPic.Visible = true;
                panelMain.Visible = true;
                labLnfo.Visible = dataGridView1.Visible = false;
            }
            else
            {
                رقم_معاملة_القسم.Enabled = true;
                panelMain.Visible = false;
                labLnfo.Visible = dataGridView1.Visible = true;
                timer2.Enabled = true;
            }
            if (الحالة.Text == "مجهول")
            {
                المصدر.Text = "إختر نوع المصدر";
                المصدر.Enabled = تاريخ_الاستلام.Enabled = الإجراء_الذي_تم.Enabled = الموضوع.Enabled = تاريخ_الإصدار.Enabled = رقم_معاملة_المصدر.Enabled = true;
                المهنة.Enabled = false;
                المهنة.Text = "مهنة غير نظامية";
            }
            else
            {
                المهنة.Text = "";
                المصدر.Text = المنطقة.Text;
                المصدر.Enabled = تاريخ_الاستلام.Enabled = الإجراء_الذي_تم.Enabled = الموضوع.Enabled = تاريخ_الإصدار.Enabled = رقم_معاملة_المصدر.Enabled = false;
                المهنة.Enabled = true;
            }
            Console.WriteLine("VCIndex " + VCIndex);
            ListSearch.Select();
            fileComboBoxAVC(combAttendVC, DataSource, "ArabicAttendVC", "TableListCombo");
            if (combAttendVC.Items.Count >= VCIndex) 
                combAttendVC.SelectedIndex = VCIndexData();
            //FillDataGridDocs("ارشفة_المستندات", "المستندات_الأولية");
            //FillDataGridDocs("مذكرات", "مذكرات_من_فرع_وزارة_الخارجية_السعودية");
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
            //MessageBox.Show(strList);
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            Console.WriteLine(strList);
            SqlCommand sqlCommand = new SqlCommand("insert into archives values (" + strList + ")", sqlConnection);
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
        private string countMonths(string date1)
        {
            if (!GreDate.Contains("-")) return "0";
            int yearS = Convert.ToInt32(date1.Split('-')[0]);
            Console.WriteLine("yearS " + yearS .ToString() );
            int monthS = Convert.ToInt32(date1.Split('-')[1]);
            Console.WriteLine("monthS " + monthS.ToString() );
            int yearE = Convert.ToInt32(GreDate.Split('-')[2]);
            Console.WriteLine("yearE " + yearE.ToString());
            int monthE = Convert.ToInt32(GreDate.Split('-')[1]);
            Console.WriteLine("monthE " + monthE.ToString() );
            int months = 1;
            int month = monthS + 1;
            int m = 12;
            for (int year = yearS; year <= yearE; year++)
            {
                if (year == yearE) m = monthE - 1;
                for (; month <= m; month++)
                {

                    months++;
                    Console.WriteLine(month.ToString() + " - " + months.ToString());
                }

                month = 1;
            }


            return months.ToString();
        }
        private void speaclNormalLetters()
        {

            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilespathIn + @"\قائمة وثائق السفر.docx";
            ActiveCopy = PrimariFiles + "Docx" + ReportName + ".docx";
            System.IO.File.Copy(route, ActiveCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

            object ParaGreData = "MarkGreData";//     
            object ParaFileNo = "MarkFileNo";//
            object ParavConsul = "MarkViseConsul";  //note verbal لتعرب       


            Word.Range BookFileNo = oBDoc.Bookmarks.get_Item(ref ParaFileNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;


            BookFileNo.Text = رقم_الملف.Text;

            BookGreData.Text = dateTimeTo.Text;
            BookvConsul.Text = combAttendVC.Text;

            object rangeFileNo = BookFileNo;
            object rangeGreData = BookGreData;
            object rangevConsul = BookvConsul;
            int indexNo = 1;
            Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
            for (int x = 0; x < dataGridView1.RowCount -1; x++)
            {
                string name = dataGridView1.Rows[x].Cells[7].Value.ToString().Split('_')[0];
                string arrestNo = "الرقم غير مدرج"; 
                if(dataGridView1.Rows[x].Cells[15].Value.ToString()!= "") 
                    arrestNo = dataGridView1.Rows[x].Cells[15].Value.ToString();
                
                if (name != "" && arrestNo != "")
                {
                    

                    try
                    {
                        table.Rows.Add();
                        table.Rows[x + 2].Cells[1].Range.Text = indexNo.ToString() + ".";
                        table.Rows[x + 2].Cells[2].Range.Text = name;
                        table.Rows[x + 2].Cells[3].Range.Text = arrestNo;
                        indexNo++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(indexNo.ToString()  + name + arrestNo);
                        
                    }
                }

            }
            
            oBDoc.Bookmarks.Add("MarkFileNo", ref rangeFileNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);

            string docxouput = FilespathOut +"_"+رقم_الملف.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilespathOut + "_" + رقم_الملف.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(pdfouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
        }

        private string MessageDocx(string messageNo,string ActiveCopy)
        {
            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilespathIn + @"\برقية رد بشأن وثيقة.docx";
            System.IO.File.Copy(route, ActiveCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

            object ParaDestin = "MarkDestin";//     
            object ParaBody = "MarkBody";//     
            object ParaMassageNo = "MarkMassageNo";//     
            object ParaGreData = "MarkGreData";//     
            object ParaHijriDate = "MarkHijriDate";//     
            object ParavConsul = "MarkViseConsul";  //note verbal لتعرب       

            Word.Range BookDestin = oBDoc.Bookmarks.get_Item(ref ParaDestin).Range;
            Word.Range BookBody = oBDoc.Bookmarks.get_Item(ref ParaBody).Range;
            Word.Range BookMassageNo = oBDoc.Bookmarks.get_Item(ref ParaMassageNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookHijriDate = oBDoc.Bookmarks.get_Item(ref ParaHijriDate).Range;
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;


            BookDestin.Text = مستقبل_المكاتبة.Text;// comboDes.Text;
            BookBody.Text = نص_المكاتبة.Text; // txtRecom.Text;
            BookMassageNo.Text = messageNo;
            BookGreData.Text = GreDate;
            BookHijriDate.Text = HijriDate;
            BookvConsul.Text = combAttendVC.Text;
            
            
            object rangeDestin = BookDestin;
            object rangeBody = BookBody;
            object rangeMassageNo = BookMassageNo;
            object rangeGreData = BookGreData;
            object rangeHijriDate = BookHijriDate;
            object rangevConsul = BookvConsul;
            
            oBDoc.Bookmarks.Add("MarkBody", ref rangeBody);
            oBDoc.Bookmarks.Add("MarkDestin", ref rangeDestin);
            oBDoc.Bookmarks.Add("MarkMassageNo", ref rangeMassageNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkHijriDate", ref rangeHijriDate);
            oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);

            string docxouput = FilespathOut + "_" + رقم_الملف.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilespathOut + "_" + رقم_الملف.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(pdfouput);
            File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
            return pdfouput;
        }

        private string NoteVerbalRespond(string messageNo,string ActiveCopy)
        {
            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilespathIn + @"\مذكرة رد بشأن وثيقة.docx";            
            System.IO.File.Copy(route, ActiveCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

            object ParaBody = "MarkBody";//     
            object ParaMassageNo = "MarkMassageNo";//     
            object ParaGreData = "MarkGreData";//     
            object ParaHijriDate = "MarkHijriDate";//                
            
            Word.Range BookBody = oBDoc.Bookmarks.get_Item(ref ParaBody).Range;
            Word.Range BookMassageNo = oBDoc.Bookmarks.get_Item(ref ParaMassageNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookHijriDate = oBDoc.Bookmarks.get_Item(ref ParaHijriDate).Range;

            BookMassageNo.Text = messageNo;
            BookGreData.Text = GreDate;
            BookHijriDate.Text = HijriDate;
            BookBody.Text = Environment.NewLine + actionSum.Text + Environment.NewLine ;

            object rangeBody = BookBody;
            object rangeMassageNo = BookMassageNo;
            object rangeGreData = BookGreData;
            object rangeHijriDate = BookHijriDate;

            oBDoc.Bookmarks.Add("MarkBody", ref rangeBody);
            oBDoc.Bookmarks.Add("MarkMassageNo", ref rangeMassageNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkHijriDate", ref rangeHijriDate);

            string docxouput = FilespathOut + "_" + رقم_الملف.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilespathOut + "_" + رقم_الملف.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(pdfouput);
            File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
            return pdfouput;
        }

        private void createList(string ActiveCopy, string fileNo)
        {

            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilespathIn + @"\قائمة التعليقات.docx";
            System.IO.File.Copy(route, ActiveCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

            object ParaGreData = "MarkGreData";
            object ParaGreData1 = "MarkGreDa1";
            object ParaHijriData = "MarkHijriData";
            object ParaFileNo = "MarkFileNo";
            object ParaDest = "MarkDest";
            object ParaPurpose = "MarkPurpose";//message لعناية ... note verbal فارغ
            object ParaPurposeText = "MarkPurposeText";
            object ParaPurText = "MarkPurText";
            object ParaIndivNo = "MarkIndivNo";
            object ParavConsul = "MarkViseConsul";  //note verbal لتعرب       


            Word.Range BookDestin = oBDoc.Bookmarks.get_Item(ref ParaDest).Range;
            Word.Range BookIndivNo = oBDoc.Bookmarks.get_Item(ref ParaIndivNo).Range;
            Word.Range BookFileNo = oBDoc.Bookmarks.get_Item(ref ParaFileNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookGreData1 = oBDoc.Bookmarks.get_Item(ref ParaGreData1).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            Word.Range BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;
            Word.Range BookPurposeText = oBDoc.Bookmarks.get_Item(ref ParaPurposeText).Range;
            Word.Range BookPurText = oBDoc.Bookmarks.get_Item(ref ParaPurText).Range;
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;

            BookDestin.Text = " مدير إدارة الوافدين";




            BookGreData1.Text = GreDate;
            BookGreData.Text = GreDate;
            BookHijriData.Text = HijriDate;
            BookPurText.Text = (dataGridView1.RowCount - 1).ToString();
            BookPurposeText.Text = fileNo;
            BookvConsul.Text = combAttendVC.Text;

            object rangeDesin = BookDestin;
            object rangeIndivNo = BookIndivNo;
            object rangePurpose = BookPurpose;
            object rangeFileNo = BookFileNo;
            object rangeGreData = BookGreData;
            object rangeGreData1 = BookGreData1;
            object rangeHijriData = BookHijriData;
            object rangePurposeText = BookPurposeText;
            object rangePurText = BookPurText;
            object rangevConsul = BookvConsul;
            int indexNo = 0;
            Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
            for (int x = 0; x < dataGridView1.RowCount -1; x++)
            {
                string AppNames = dataGridView1.Rows[x].Cells[7].Value.ToString();
                string arrestNo = dataGridView1.Rows[x].Cells[15].Value.ToString();
                string comment = dataGridView1.Rows[x].Cells[8].Value.ToString();
                string responces = dataGridView1.Rows[x].Cells[18].Value.ToString();
                string country= dataGridView1.Rows[x].Cells[22].Value.ToString();
                string requestedDoc = dataGridView1.Rows[x].Cells[23].Value.ToString();

                if(responces.Contains("ينتمي"))
                    responces = responces +  country;
                
                else if (responces == "الإجراء قد تم" && requestedDoc != "")
                    responces = responces + " ب" + requestedDoc;

                else if (requestedDoc != "")
                    responces = responces + Environment.NewLine + "(" + requestedDoc + ")";
                 

                if (!string.IsNullOrEmpty(AppNames))
                {
                    table.Rows.Add();
                    indexNo++;
                    table.Rows[x + 2].Cells[1].Range.Text = (x + 1).ToString() + ".";
                    table.Rows[x + 2].Cells[2].Range.Text = AppNames.Split('_')[0];
                    table.Rows[x + 2].Cells[3].Range.Text = arrestNo;
                    table.Rows[x + 2].Cells[4].Range.Text = responces;                    
                }

            }

            BookPurText.Text = BookFileNo.Text = indexNo.ToString() + fileNo;
            BookIndivNo.Text = indexNo.ToString();

            oBDoc.Bookmarks.Add("MarkDest", ref rangeDesin);
            oBDoc.Bookmarks.Add("MarkIndivNo", ref rangeIndivNo);
            oBDoc.Bookmarks.Add("MarkPurpose", ref rangePurpose);
            oBDoc.Bookmarks.Add("MarkFileNo", ref rangeFileNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkGreDa1", ref rangeGreData1);
            oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);
            oBDoc.Bookmarks.Add("MarkPurposeText", ref rangePurposeText);
            oBDoc.Bookmarks.Add("MarkPurText", ref rangePurText);
            oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);

            string docxouput = FilespathOut + رقم_الملف.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilespathOut + رقم_الملف.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(pdfouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
        }

        private void TranvelDoc()
        {

            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilespathIn + @"\الوثيقة.docx";
            ActiveCopy = PrimariFiles + "Docx" + ReportName + ".docx";
            System.IO.File.Copy(route, ActiveCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

            object ParaName = "MarkName";//     
                


            Word.Range BookName = oBDoc.Bookmarks.get_Item(ref ParaName).Range;


            BookName.Text = اسم_المواطن_موضوع_الإجراء.Text;


            object rangeName = BookName;
            
            

            oBDoc.Bookmarks.Add("MarkName", ref rangeName);
            
            string docxouput = FilespathOut + "_" + رقم_الملف.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilespathOut + "_" + رقم_الملف.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(pdfouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
        }

        private void UpdateState(int id, string text, string col, string table)
        {
            // MessageBox.Show(id.ToString() +"-"+ table + "-"+column + "-"+text);
            //string qurey = "update "+table+" set "+ column + "=@"+ column + " where ID=@id";
            string qurey = "update "+table+" set " + col + "=@" + col+ " where ID=@id";
            //MessageBox.Show(qurey);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@" + col, text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName + " order by " + comlumnName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;


                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                if (comlumnName.Contains("ArabCountries")) combbox.Items.Add("جمهورية السودان");
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

        private void fileComboBoxAVC(ComboBox combbox, string source, string comlumnName, string tableName)
        {

            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;


                try
                {
                    cmd.ExecuteNonQuery();
                    DataTable table = new DataTable();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                    dataAdapter.Fill(table);

                    foreach (DataRow dataRow in table.Rows)
                    {
                        if (dataRow[comlumnName].ToString() != "")
                        {
                            bool found = false;
                            for (int x = 0; x < combbox.Items.Count; x++)
                            {
                                if (combbox.Items[x].ToString() == dataRow[comlumnName].ToString()) found = true;
                            }
                            if (!found) combbox.Items.Add(dataRow[comlumnName].ToString());
                        }
                    }
                }
                catch (Exception ex) { }
                saConn.Close();
            }
        }


        private void fillFileData(string source, string fileNo)
        {
            المكاتبات.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select نوع_المستند from TableGeneralArch where رقم_معاملة_القسم=@رقم_معاملة_القسم ";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;

                cmd.Parameters.AddWithValue("@رقم_معاملة_القسم", fileNo);
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                
                foreach (DataRow dataRow in table.Rows)
                {
                    if (!string.IsNullOrEmpty(dataRow["نوع_المستند"].ToString()))
                    {
                        المكاتبات.Items.Add(dataRow["نوع_المستند"].ToString());
                    }
                }
                saConn.Close();
            }
        }
        private void DocIDGenerator()
        {
            رقم_معاملة_القسم.Enabled = false;
            rowCount = loadRerNo(loadIDNo());
            if (rowCount == "0") rowCount = "1";
                رقم_معاملة_القسم.Text = rowCount;
            MessNoPart = "ق س ج/80/" + DateTime.Now.Year.ToString().Replace("20", "") + "/14/" + rowCount;
            docIDNumber = "ق س ج/80/22/14/" + رقم_معاملة_القسم.Text;
        }

        private int loadIDNo()
        {


            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from TableMessages order by ID desc", sqlCon);
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
        private string loadRerNo(int id)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT رقم_معاملة_القسم from TableMessages where ID=@ID", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ID", id);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string rowCnt = "0";

            foreach (DataRow row in dtbl.Rows)
            {
                if (row["رقم_معاملة_القسم"].ToString() != "")
                {
                    rowCnt = (Convert.ToInt32(row["رقم_معاملة_القسم"].ToString().Split('/')[4]) + 1).ToString();
                }
                
            }
            return rowCnt;

        }
        private string checkColumnName(string table)
        {
            string colList = "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS " + table, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            int id = 0;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (!dataRow["COLUMN_NAME"].ToString().Contains("Data"))
                {
                    if (id == 0) 
                        colList = " " + dataRow["COLUMN_NAME"].ToString();
                    else 
                        colList = colList + "," + dataRow["COLUMN_NAME"].ToString() + " ";
                    id++;
                }
            }
            return colList;
        }

        public void FillDataGridView(string colList)
        {
            //MessageBox.Show(colList);
            Console.WriteLine(DataSource);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select "+ colList+ " from TableMessages order by responce DESC", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            
            dataGridView1.DataSource = dtbl;

            rowCount = dataGridView1.Rows.Count.ToString();
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["رقم_معاملة_القسم"].Width = 130;
            dataGridView1.Columns["المصدر"].Width = 130;
            dataGridView1.Columns["الموضوع"].Width = 150;
            dataGridView1.Columns["اسم_المواطن_موضوع_الإجراء"].Width = 150;
            sqlCon.Close();
            //MessageBox.Show("reference");

            //for (int x = 0; x < dataGridView1.RowCount - 1; x++)
            //////for (int x = 0; x < 10; x++)
            //{
            //    int id = Convert.ToInt32(dataGridView1.Rows[x].Cells[0].Value.ToString());
            //    string filePath = "";
            //    docIDNumber = dataGridView1.Rows[x].Cells[6].Value.ToString();
            //    string reference = dataGridView1.Rows[x].Cells["reference"].Value.ToString();
            //    string recommendation = dataGridView1.Rows[x].Cells["recommendation"].Value.ToString();
            //    //UpdateState(id, "", "reference", "TableMessages");
            //    //UpdateState(id, "", "recommendation", "TableMessages");
            //    //UpdateState(id, "", "recommendation");
            //    ////MessageBox.Show(filePath);
            //    //if (reference == "")
            //    //{

            //    //    UpdateState(id, "تم", "reference");
            //    //}

            //    //UpdateState(id, "", "recommendation", "TableMessages");
            //    //UpdateState(id, "", "reference", "TableMessages");
            //    //if (reference == "")
            //    //{
            //    //    filePath = MoveFileDoc(id, 1, PrimariFiles, x, docIDNumber);
            //    //    UpdateState(id, "تم", "reference", "TableMessages");
            //    //    //insertDoc(DataSource, filePath, docIDNumber, "ارشفة_المستندات");

            //    //}
            //    //if (recommendation == "")
            //    //{
            //    //    filePath = MoveFileDoc(id, 2, PrimariFiles, x, docIDNumber);
            //    //    UpdateState(id, "تم", "recommendation", "TableMessages");
            //    //    //insertDoc(DataSource, filePath, docIDNumber, "ارشفة_المستندات");

            //    //}
            //    //filePath = "";
            //    //filePath = MoveFileDoc(id, 1, PrimariFiles, x);

            //    //if (filePath != "" && recommendation == "")
            //    //{
            //    //    insertDoc(DataSource, filePath, docIDNumber, "ارشفة_المستندات");
            //    //    File.Delete(filePath);
            //    //    UpdateState(id, "تم", "recommendation");
            //    //}
            //    Console.WriteLine(dataGridView1.Rows[x].Cells[0].Value.ToString());
            //    filePath = "";
            //}
        }

        public void FillDataGridDocs(string dataType,string dataType1)
        {
            //MessageBox.Show(colList);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select ID,نوع_المستند from TableGeneralArch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);

           
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                int id = Convert.ToInt32(dataRow["ID"].ToString());
                if (dataRow["نوع_المستند"].ToString() == dataType)
                {
                    UpdateState(id, dataType1, "نوع_المستند", "TableGeneralArch");
                }
                Console.WriteLine(id);
            }
        }

        public void itemsCombo(ComboBox comboBox, string itemName, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select "+ itemName+" from "+ table+" group by " + itemName, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            comboBox.Items.Clear();
            comboBox.Items.Add("الجميع");
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow[itemName].ToString() != "")
                {
                    comboBox.Items.Add(dataRow[itemName].ToString());
                }
            }
            sqlCon.Close();

        }

        private void Colorcomment1(int index)
        {
            //24
            int alldata = 0;
            int determinedData = 0;
            int finishData = 0;
            int nonSudanese = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;
                string nationality = dataGridView1.Rows[i].Cells[index].Value.ToString();

                if (!nationality.Contains("سودان") && nationality != "" && nationality != "غير معروف")
                {
                    //MessageBox.Show(nationality);
                    nonSudanese++;
                }

                if (nationality != "غير معروف" && dataGridView1.Rows[i].Cells[14].Value.ToString() != "تم الإجراء")
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    determinedData++;
                    
                }
                if (dataGridView1.Rows[i].Cells[14].Value.ToString() == "تم الإجراء")
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Green;
                    finishData++;
                }
                alldata++;
            }
            timer2.Enabled = false;
            labLnfo.Text = "العدد الكلي (" + (alldata).ToString()+ ") معاملة، تم إنهاء عدد ("+ finishData.ToString()+ ") معاملة، وتم تحديد موقف القنصلية مبدئيا في عدد (" + determinedData.ToString() + ") معاملة، وتبقى عدد (" +(alldata - determinedData - finishData).ToString()+ ") معاملة في الانتظار" +" تبين ان عدد ( " + nonSudanese.ToString() + ") غير سودانيين";
            //
        }
        private void btnAuth_Click(object sender, EventArgs e)
        {
            loadPic.Enabled = btnArchBasic.Visible = btnArch.Enabled = false;
            btnArchBasic.Visible = true;
            try

            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.

                    var ScanerItem = device.Items[1]; // select the scanner.


                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);

                    PathImage[imagecount] = PrimariFiles + "ScanImg" + rowCount + (imagecount).ToString() + ".jpg";


                    if (File.Exists(PathImage[imagecount]))
                    {
                        File.Delete(PathImage[imagecount]);
                    }
                    imgFile.SaveFile(PathImage[imagecount]);
                    pictureBox1.ImageLocation = PathImage[imagecount];
                    //panel1.Visible = false;
                    imagecount++;
                    panelPic.Visible = true;
                    panelPic.BringToFront();
                }
                else
                {

                    MessageBox.Show("لا يوجد جهاز ماسح متصل");
                }

            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
            loadPic.BackColor = btnArch.BackColor = System.Drawing.Color.LightGreen;
            loadPic.Text = btnArch.Text = "اضافة مستند آخر (" + imagecount.ToString() + ")";
            btnArch.Size = btnArch.Size = new System.Drawing.Size(150, 37);
            loadPic.Size = btnArch.Size = new System.Drawing.Size(150, 37);
            loadPic.Location = new System.Drawing.Point(889, 561);
            btnArch.Location = new System.Drawing.Point(889, 520);
            loadPic.Enabled = btnArchBasic.Visible = btnArch.Enabled = true;
            
        }
        private void loadScanner()
        {
            try
            {
                var deviceManager = new DeviceManager();



                for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++) // Loop Through the get List Of Devices.
                {
                    if (deviceManager.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType) // Skip device If it is not a scanner
                    {
                        continue;
                    }
                    AvailableScanner = deviceManager.DeviceInfos[i];
                    break;
                }


            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try


            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.

                    var ScanerItem = device.Items[1]; // select the scanner.


                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);

                    PathImage[imagecount - 1] = PrimariFiles + "ScanImg" + rowCount + (imagecount - 1).ToString() + ".jpg";


                    if (File.Exists(PathImage[imagecount - 1]))
                    {
                        File.Delete(PathImage[imagecount - 1]);
                    }
                    imgFile.SaveFile(PathImage[imagecount - 1]);
                    pictureBox1.ImageLocation = PathImage[imagecount - 1];
                    //panel1.Visible = false;

                }
                else
                {

                    MessageBox.Show("لا يوجد جهاز ماسح متصل");
                }

            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void loadPic_Click(object sender, EventArgs e)
        {
            btnArchBasic.Visible = true; 
            string fileName = loadDocxFile();
            if (fileName != "")
            {
                pictureBox1.ImageLocation = PathImage[imagecount] = fileName;
                imagecount++;
                btnArch.BackColor = System.Drawing.Color.LightGreen;
                loadPic.Text = btnArch.Text = "اضافة مستند آخر (" + imagecount.ToString() + ")";
                loadPic.Size = btnArch.Size = new System.Drawing.Size(150, 37);
                loadPic.Location = new System.Drawing.Point(889, 561);
                btnArch.Location = new System.Drawing.Point(889, 520);
                panelPic.Visible = true;
                panelPic.BringToFront();
            }
        }

        private void reLoadPic_Click(object sender, EventArgs e)
        {
            string fileName = loadDocxFile();
            if (fileName != "")
            {
                pictureBox1.ImageLocation = PathImage[imagecount - 1] = fileName;
            }
        }
        private string loadDocxFile()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return dlg.FileName;
            }
            return "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!ModifyPermit) {
                MessageBox.Show("حساب الموظف غير مخول بإجراء تعديلات بنافذة " + this.Text + Environment.NewLine + "يرجى التواصل  مع مدير النظام");
                return;
            }
            if (ArchiveState)
            {
                DocIDGenerator();
            }
            if (رقم_معاملة_القسم.Text == "") return;
            if (picVerify.Visible && newData)
            {
                Console.WriteLine("checkDataInfo");
                int id = checkDataInfo(false);
                if (picVerify.Visible)
                {
                    var selectedOption = MessageBox.Show("المتابعة؟", "يوجد إجراء سابق متطابق مع رقم الايقاف", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (selectedOption == DialogResult.No)
                    {
                        return;
                    }
                }
            }
            btnArchBasic.Enabled = false;
            CreatePic(rowCount, PathImage);
            this.Close();
            
            
        }

        private void UpdateFileList(int iD, string fileNo, string fileDest, string individualNo, string conToWafidDate, string consToWafidNo, string wafidToWorkDate, string wafidToWorkNo, string workToWafidDate, string workToWafidNo, string workNotesNo, string workFinsihed)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableFiles (FileNo, FileDest, IndividualNo, ConToWafidDate, ConsToWafidNo, WafidToWorkDate, WafidToWorkNo, WorkToWafidDate, WorkToWafidNo, WorkNotesNo, WorkFinsihed) values(@FileNo, @FileDest, @IndividualNo, @ConToWafidDate, @ConsToWafidNo, @WafidToWorkDate, @WafidToWorkNo, @WorkToWafidDate, @WorkToWafidNo, @WorkNotesNo, @WorkFinsihed)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            if (iD != 0)
                sqlCmd.Parameters.AddWithValue("@ID", iD);
            sqlCmd.Parameters.AddWithValue("@FileNo", fileNo);
            sqlCmd.Parameters.AddWithValue("@FileDest", fileDest);
            sqlCmd.Parameters.AddWithValue("@IndividualNo", individualNo);
            sqlCmd.Parameters.AddWithValue("@ConToWafidDate", conToWafidDate);
            sqlCmd.Parameters.AddWithValue("@ConsToWafidNo", consToWafidNo);
            sqlCmd.Parameters.AddWithValue("@WafidToWorkDate", wafidToWorkDate);
            sqlCmd.Parameters.AddWithValue("@WafidToWorkNo", wafidToWorkNo);
            sqlCmd.Parameters.AddWithValue("@WorkToWafidDate", workToWafidDate);
            sqlCmd.Parameters.AddWithValue("@WorkToWafidNo", workToWafidNo);
            sqlCmd.Parameters.AddWithValue("@WorkNotesNo", workNotesNo);
            sqlCmd.Parameters.AddWithValue("@WorkFinsihed", workFinsihed);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
            

        }

        private void CreatePic(string reportName, string[] location)
        {
            docIDNumber = "ق س ج/80/22/14/" + رقم_معاملة_القسم.Text;
            string docDirction = "-واردة";
            int docid = NormalAddEdit(DataSource, docIDNumber, ID);
            if (docid == 0) return;
            //MessageBox.Show(docid.ToString());
            if (responce.SelectedIndex == 2 || responce.SelectedIndex == 5) finished();
            if (fileUpdate && (responce.SelectedIndex == 1 || responce.SelectedIndex == 2 || responce.SelectedIndex == 3)) {
                var selectedOption = MessageBox.Show("", "المستندا صادر عن القنصلية العامة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                {
                    docDirction = "-صادرة";
                }
            }
            fileUpdate = true;
            colIDs[2] = GreDate;
            colIDs[3] = اسم_المواطن_موضوع_الإجراء.Text;
            string filename = "";
            switch (combAction.SelectedIndex)
            {
                case 0:
                    ID = 0;
                    for (int x = 0; x < imagecount; x++)
                    {
                        if (location[x] != "")
                        {
                            using (Stream stream = File.OpenRead(location[x]))
                            {
                                byte[] buffer1 = new byte[stream.Length];
                                stream.Read(buffer1, 0, buffer1.Length);
                                var fileinfo1 = new FileInfo(location[x]);
                                string extn1 = fileinfo1.Extension;
                                string DocName1 = fileinfo1.Name;
                                insertDoc(docid.ToString(), GregorianDate, labEmp.Text, DataSource, extn1, DocName1, docIDNumber, "مستندات" + docDirction, buffer1);
                            }
                        }
                    }
                    break;
                case 1:
                    for (int x = 0; x < imagecount; x++)
                    {
                        if (location[x] != "")
                        {
                            using (Stream stream = File.OpenRead(location[x]))
                            {
                                byte[] buffer1 = new byte[stream.Length];
                                stream.Read(buffer1, 0, buffer1.Length);
                                var fileinfo1 = new FileInfo(location[x]);
                                string extn1 = fileinfo1.Extension;
                                string DocName1 = fileinfo1.Name;
                                //insertDoc(DataSource, extn1, DocName1, docIDNumber, "مستندات" + docDirction, buffer1);
                            }
                        }
                    }
                    break;

                
                case 2:
                    ActiveCopy = PrimariFiles + "Docx" + DateTime.Now.ToString("mmss") + ".docx";
                    while (File.Exists(ActiveCopy))
                    {
                         ActiveCopy = PrimariFiles + "Docx" + DateTime.Now.ToString("mmss") + ".docx";
                    }                    
                    currentNo = docIDNumber + "/" + CurrentBtnName;
                     filename = MessageDocx(currentNo, ActiveCopy);
                    if (filename != "")
                    {
                        using (Stream stream = File.OpenRead(filename))
                        {
                            byte[] buffer1 = new byte[stream.Length];
                            stream.Read(buffer1, 0, buffer1.Length);
                            var fileinfo1 = new FileInfo(filename);
                            string extn1 = fileinfo1.Extension;
                            string DocName1 = fileinfo1.Name;
                            //insertDoc(DataSource, extn1, DocName1, docIDNumber, "برقيات" + docDirction, buffer1);
                        }
                    }
                    addarchives(colIDs);
                    break;
                case 3:
                    currentNo = docIDNumber + "/" + CurrentBtnName;
                    ActiveCopy = PrimariFiles + "Docx" + DateTime.Now.ToString("mmss") + ".docx";
                    while (File.Exists(ActiveCopy))
                    {
                        ActiveCopy = PrimariFiles + "Docx" + DateTime.Now.ToString("mmss") + ".docx";
                    }
                    filename = NoteVerbalRespond(currentNo, ActiveCopy);
                    if (filename != "")
                    {
                        using (Stream stream = File.OpenRead(filename))
                        {
                            byte[] buffer1 = new byte[stream.Length];
                            stream.Read(buffer1, 0, buffer1.Length);
                            var fileinfo1 = new FileInfo(filename);
                            string extn1 = fileinfo1.Extension;
                            string DocName1 = fileinfo1.Name;
                            //insertDoc(DataSource, extn1, DocName1, docIDNumber, "مذكرات" + docDirction, buffer1);
                        }
                    }
                    addarchives(colIDs);
                    break;


                case 4:
                    ActiveCopy = PrimariFiles + "Docx" + DateTime.Now.ToString("mmss") + ".docx";
                    while (File.Exists(ActiveCopy))
                    {
                        ActiveCopy = PrimariFiles + "Docx" + DateTime.Now.ToString("mmss") + ".docx";
                    }
                    createList(ActiveCopy, رقم_الملف.Text);
                    break;
            }
            imagecount = 0;
        }


        //private void FinalDataArch(string dataSource, string documentID)
        //{
        //    if(documentID == "") { return; }
        //    SqlConnection sqlCon = new SqlConnection(dataSource);
        //    if (sqlCon.State == ConnectionState.Closed)
        //        sqlCon.Open();
        //    string query2 = "UPDATE TableMessages SET رقم_معاملة_المصدر = @رقم_معاملة_المصدر, المصدر = @المصدر, تاريخ_الإصدار = @تاريخ_الإصدار, تاريخ_الاستلام = @تاريخ_الاستلام, الموضوع = @الموضوع, رقم_معاملة_القسم = @رقم_معاملة_القسم, الإجراء_الذي_تم = @الإجراء_الذي_تم, مسوؤل_الملف = @مسوؤل_الملف, الموظف = @الموظف, حالة_الأرشفة = @حالة_الأرشفة,اسم_المواطن_موضوع_الإجراء=@اسم_المواطن_موضوع_الإجراء,رقم_الايقاف=@رقم_الايقاف,المنطقة=@المنطقة,رقم_الملف=@رقم_الملف,actionSum=@actionSum,حالة_الأجراء=@حالة_الأجراء,الحالة=@الحالة,country=@country WHERE رقم_معاملة_القسم = @رقم_معاملة_القسم";
        //    SqlCommand sqlCmd = new SqlCommand(query2, sqlCon);
        //    sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", documentID);
        //    sqlCmd.Parameters.AddWithValue("@المصدر", المصدر.Text);
        //    sqlCmd.Parameters.AddWithValue("@تاريخ_الإصدار", تاريخ_الإصدار.Text);
        //    sqlCmd.Parameters.AddWithValue("@تاريخ_الاستلام", تاريخ_الاستلام.Text);
        //    sqlCmd.Parameters.AddWithValue("@الموضوع", الموضوع.Text);
        //    sqlCmd.Parameters.AddWithValue("@رقم_معاملة_المصدر", رقم_معاملة_المصدر.Text);
        //    sqlCmd.Parameters.AddWithValue("@الإجراء_الذي_تم", txtAction1.Text + Environment.NewLine + "تمت الارشفة  بواسطة  " + labEmp.Text + "-" + GregorianDate + Environment.NewLine + "--------------------------");
        //    sqlCmd.Parameters.AddWithValue("@مسوؤل_الملف", مسوؤل_الملف.Text + Environment.NewLine + الإجراء_الذي_تم.Text);
        //    sqlCmd.Parameters.AddWithValue("@actionSum", responce.Text);
        //    sqlCmd.Parameters.AddWithValue("@الموظف", labEmp.Text);
        //    sqlCmd.Parameters.AddWithValue("@المنطقة", المنطقة.Text);
        //    sqlCmd.Parameters.AddWithValue("@اسم_المواطن_موضوع_الإجراء", اسم_المواطن_موضوع_الإجراء.Text);
        //    sqlCmd.Parameters.AddWithValue("@الحالة", الحالة.Text);
        //    sqlCmd.Parameters.AddWithValue("@حالة_الأرشفة", "مؤرشف مبدئيا");
        //    sqlCmd.Parameters.AddWithValue("@حالة_الأجراء", "قيد الإجراء");
        //    sqlCmd.Parameters.AddWithValue("@رقم_الايقاف", رقم_الايقاف.Text);
        //    sqlCmd.Parameters.AddWithValue("@رقم_الملف", fileNo);
        //    sqlCmd.Parameters.AddWithValue("@country", country.Text);
           
        //    sqlCmd.ExecuteNonQuery();
        //    sqlCon.Close();
        //}

        //private void addDocx(string dataSource, string filePath, string documentID, string docType)
        //{
        //    if (documentID == "") { return; }
        //    SqlConnection sqlCon = new SqlConnection(dataSource);
        //    if (sqlCon.State == ConnectionState.Closed)
        //        sqlCon.Open();
        //    string query = "UPDATE TableMessages SET ارشفة_المستندات=@ارشفة_المستندات,Data1=@Data1,Extension1=@Extension1 WHERE رقم_معاملة_القسم = @رقم_معاملة_القسم";
        //    if(docType.Contains("مذكرات")) 
        //        query = "UPDATE TableMessages SET مذكرات=@مذكرات,Data2=@Data2,Extension2=@Extension2 WHERE رقم_معاملة_القسم = @رقم_معاملة_القسم"; 
        //    if(docType.Contains("وثيقة_السفر")) 
        //        query = "UPDATE TableMessages SET وثيقة_السفر=@وثيقة_السفر,Data3=@Data3,Extension3=@Extension3 WHERE رقم_معاملة_القسم = @رقم_معاملة_القسم"; 
        //    if(docType.Contains("برقيات")) 
        //        query = "UPDATE TableMessages SET برقيات=@برقيات,Data3=@Data3,Extension3=@Extension3 WHERE رقم_معاملة_القسم = @رقم_معاملة_القسم"; 
        //    SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
        //    sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", documentID);
        //    using (Stream stream = File.OpenRead(filePath))
        //    {
        //        byte[] buffer1 = new byte[stream.Length];
        //        stream.Read(buffer1, 0, buffer1.Length);
        //        var fileinfo1 = new FileInfo(filePath);
        //        string extn1 = fileinfo1.Extension;
        //        string DocName1 = fileinfo1.Name;
        //        if (docType.Contains("مذكرات"))
        //        {
        //            sqlCmd.Parameters.Add("@Data2", SqlDbType.VarBinary).Value = buffer1;
        //            sqlCmd.Parameters.Add("@Extension2", SqlDbType.Char).Value = extn1;
        //        }
        //        else if (docType.Contains("ارشفة_المستندات")) { 
        //            sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
        //            sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
        //        }
        //        else {
        //            sqlCmd.Parameters.Add("@Data3", SqlDbType.VarBinary).Value = buffer1;
        //            sqlCmd.Parameters.Add("@Extension3", SqlDbType.Char).Value = extn1;
        //        }
        //        sqlCmd.Parameters.Add(docType, SqlDbType.NVarChar).Value = DocName1;
        //    }
        //    sqlCmd.ExecuteNonQuery();
        //    sqlCon.Close();
        //}

        //private void updateReportEntry(string dataSource, string filePath, string authNo)
        //{

        //    SqlConnection sqlCon = new SqlConnection(dataSource);
        //    if (sqlCon.State == ConnectionState.Closed)
        //        sqlCon.Open();
        //    string query2 = "UPDATE TableMessages SET رقم_معاملة_المصدر = @رقم_معاملة_المصدر, المصدر = @المصدر, تاريخ_الإصدار = @تاريخ_الإصدار, تاريخ_الاستلام = @تاريخ_الاستلام, الموضوع = @الموضوع, رقم_معاملة_القسم = @رقم_معاملة_القسم, الإجراء_الذي_تم = @الإجراء_الذي_تم, مسوؤل_الملف = @مسوؤل_الملف, الموظف = @الموظف, حالة_الأرشفة = @حالة_الأرشفة,اسم_المواطن_موضوع_الإجراء=@اسم_المواطن_موضوع_الإجراء,رقم_الايقاف=@رقم_الايقاف,المنطقة=@المنطقة,رقم_الملف=@رقم_الملف,actionSum=@actionSum,Data1=@Data1,Extension1=@Extension1,ارشفة_المستندات=@ارشفة_المستندات,حالة_الأجراء=@حالة_الأجراء,رقم_الايقاف=@رقم_الايقاف,رقم_الملف=@رقم_الملف,الحالة=@الحالة,country=@country WHERE رقم_معاملة_القسم = @رقم_معاملة_القسم";
        //    SqlCommand sqlCmd = new SqlCommand(query2, sqlCon);
        //    sqlCmd.CommandType = CommandType.Text;
        //    sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", authNo);
        //    sqlCmd.Parameters.AddWithValue("@المصدر", المصدر.Text);
        //    sqlCmd.Parameters.AddWithValue("@تاريخ_الإصدار", تاريخ_الإصدار.Text);
        //    sqlCmd.Parameters.AddWithValue("@تاريخ_الاستلام", تاريخ_الاستلام.Text);
        //    sqlCmd.Parameters.AddWithValue("@الموضوع", الموضوع.Text);
        //    sqlCmd.Parameters.AddWithValue("@رقم_معاملة_المصدر", رقم_معاملة_المصدر.Text);
        //    sqlCmd.Parameters.AddWithValue("@الإجراء_الذي_تم", txtAction1.Text + Environment.NewLine + "تمت الارشفة  بواسطة  " + labEmp.Text + "-" + GregorianDate + Environment.NewLine + "--------------------------");
        //    sqlCmd.Parameters.AddWithValue("@مسوؤل_الملف", مسوؤل_الملف.Text + Environment.NewLine + الإجراء_الذي_تم.Text);
        //    sqlCmd.Parameters.AddWithValue("@actionSum", responce.Text);
        //    sqlCmd.Parameters.AddWithValue("@الموظف", labEmp.Text);
        //    sqlCmd.Parameters.AddWithValue("@المنطقة", المنطقة.Text);
        //    sqlCmd.Parameters.AddWithValue("@اسم_المواطن_موضوع_الإجراء", اسم_المواطن_موضوع_الإجراء.Text);
        //    sqlCmd.Parameters.AddWithValue("@الحالة",  الحالة.Text);
        //    sqlCmd.Parameters.AddWithValue("@حالة_الأرشفة", "مؤرشف مبدئيا");
        //    sqlCmd.Parameters.AddWithValue("@حالة_الأجراء", "قيد الإجراء");
        //    sqlCmd.Parameters.AddWithValue("@رقم_الايقاف", رقم_الايقاف.Text);
        //    sqlCmd.Parameters.AddWithValue("@رقم_الملف", fileNo);
        //    sqlCmd.Parameters.AddWithValue("@country", country.Text);
        //    if (filePath != "")
        //    {

        //        using (Stream stream = File.OpenRead(filePath))
        //        {
        //            byte[] buffer1 = new byte[stream.Length];
        //            stream.Read(buffer1, 0, buffer1.Length);
        //            var fileinfo1 = new FileInfo(filePath);
        //            string extn1 = fileinfo1.Extension;
        //            string DocName1 = fileinfo1.Name;
        //            sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
        //            sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
        //            sqlCmd.Parameters.Add("@ارشفة_المستندات", SqlDbType.NVarChar).Value = DocName1;

        //        }
        //    }
        //    sqlCmd.ExecuteNonQuery();

        //    sqlCon.Close();
        //}

        private string OpenFile1(string documenNo)
        {
            string str = "";
            string query;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "SELECT ID,رقم_معاملة_القسم,Data1, Extension1,مذكرات,الموضوع,الموظف,مسوؤل_الملف,الإجراء_الذي_تم,تاريخ_الاستلام,تاريخ_الإصدار,المصدر,اسم_المواطن_موضوع_الإجراء,ارشفة_المستندات,رقم_الايقاف from TableMessages where رقم_معاملة_القسم=@رقم_معاملة_القسم";
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@رقم_معاملة_القسم", SqlDbType.NVarChar).Value = documenNo;

            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                
                string FileName1 = reader["ارشفة_المستندات"].ToString();


                var ext = reader["Extension1"].ToString();
                CurrentFile1 = PrimariFiles + FileName1.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                var Data = (byte[])reader["Data1"];
                File.WriteAllBytes(CurrentFile1, Data);
            }
            Con.Close();
            return str;
        }
        private string OpenFile2(string documenNo)
        {
            string str = "";
            string query;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "SELECT ID,رقم_معاملة_القسم,Data2, Extension2,مذكرات,الموضوع,الموظف,مسوؤل_الملف,الإجراء_الذي_تم,تاريخ_الاستلام,تاريخ_الإصدار,المصدر,اسم_المواطن_موضوع_الإجراء,ارشفة_المستندات,رقم_الايقاف,الحالة from TableMessages where رقم_معاملة_القسم=@رقم_معاملة_القسم";
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@رقم_معاملة_القسم", SqlDbType.NVarChar).Value = documenNo;
            
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                DocNoPro = reader["رقم_معاملة_القسم"].ToString();

                 رقم_الايقاف.Text = reader["رقم_الايقاف"].ToString();
               الموضوع.Text = reader["الموضوع"].ToString();
                المصدر.Text = reader["المصدر"].ToString();
                تاريخ_الإصدار.Text = reader["تاريخ_الإصدار"].ToString();
                تاريخ_الاستلام.Text = reader["تاريخ_الاستلام"].ToString();
                الإجراء_الذي_تم.Text = reader["الإجراء_الذي_تم"].ToString();
                مسوؤل_الملف.Text = reader["مسوؤل_الملف"].ToString();
                الحالة.Text = reader["الحالة"].ToString();
                //labEmp.Text = reader["الموظف"].ToString();
                string strList = reader["اسم_المواطن_موضوع_الإجراء"].ToString();
                if (strList.Contains("_"))
                {
                    اسم_المواطن_موضوع_الإجراء.Text = strList.Split('_')[0];
                    المهنة.Text = strList.Split('_')[2];
                    if(strList.Split('_')[1] == "مجهول")
                        الحالة.Checked = true;
                    else
                        الحالة.Checked = false;
                }
                else اسم_المواطن_موضوع_الإجراء.Text = strList;

                string FileName = reader["مذكرات"].ToString();                 
                var ext = reader["Extension2"].ToString();
                CurrentFile2 = PrimariFiles + FileName.Replace(ext, DateTime.Now.ToString("mmss")) + ext;                
                var Data = (byte[])reader["Data2"];
                File.WriteAllBytes(CurrentFile2, Data);
                
            }
            Con.Close();
            return str;
        }

        private bool OpenFileDoc(string id, string table, string dataType)
        {
            string query = "select Data1, Extension1,المستند from "+ table+ "  where رقم_معاملة_القسم=@رقم_معاملة_القسم and نوع_المستند=@نوع_المستند";

            SqlConnection Con = new SqlConnection(DataSource);             
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@رقم_معاملة_القسم", SqlDbType.NVarChar).Value = id;
            sqlCmd1.Parameters.Add("@نوع_المستند", SqlDbType.NVarChar).Value = dataType;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                 var name = reader["المستند"].ToString();
                    if (reader["المستند"].ToString() == "") return false;
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                return true;
            }
            Con.Close();
            return false;
        }

        private string MoveFileDoc(int id, int fileNo, string NewFileName, int index, string messNo)
        {
            
            string query = "select تاريخ_المعاملة,الموظف,Data1, Extension1,ارشفة_المستندات from TableMessages  where ID=@id";

            SqlConnection Con = new SqlConnection(DataSource);


            if (fileNo == 2)
            {
                query = "select تاريخ_المعاملة,الموظف,Data2, Extension2,مذكرات from TableMessages  where ID=@id";
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
                    string date = reader["تاريخ_المعاملة"].ToString();
                    string emp = reader["الموظف"].ToString();
                    var name = reader["ارشفة_المستندات"].ToString();
                    if (reader["ارشفة_المستندات"].ToString() == "") return "";
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                     //NewFileName = NewFileName + name.Replace(ext, index.ToString()) + ext;

                    insertDoc(id.ToString(), date, emp,DataSource, ext, name, messNo, "ارشفة_المستندات", Data);

                    //File.WriteAllBytes(NewFileName, Data);
                    //System.Diagnostics.Process.Start(NewFileName);
                }
                else if (fileNo == 2)
                {
                    string date = reader["تاريخ_المعاملة"].ToString();
                    string emp = reader["الموظف"].ToString();
                    var name = reader["مذكرات"].ToString();
                    if (reader["مذكرات"].ToString() == "") return "";
                    var Data = (byte[])reader["Data2"];
                    var ext = reader["Extension2"].ToString();
                    insertDoc(id.ToString(), date, emp, DataSource, ext, name, messNo, "مذكرات", Data);
                    //NewFileName = NewFileName + name.Replace(ext, index.ToString()) + ext;
                    //File.WriteAllBytes(NewFileName, Data);
                    //System.Diagnostics.Process.Start(NewFileName);
                }
                          
            }
            //Con.Close();
            return NewFileName;
        }
        //private void NewReportEntry(string dataSource, string filePath, string messNo)
        //{
        //    if (messNo == "") return;

        //    string query = "INSERT INTO TableMessages (رقم_معاملة_المصدر,المصدر,تاريخ_الإصدار,تاريخ_الاستلام,الموضوع,رقم_معاملة_القسم,الإجراء_الذي_تم,مسوؤل_الملف,الموظف,حالة_الأرشفة,اسم_المواطن_موضوع_الإجراء,حالة_الأجراء,رقم_الايقاف,المنطقة,رقم_الملف,actionSum,الحالة,country) values(@رقم_معاملة_المصدر,@المصدر,@تاريخ_الإصدار,@تاريخ_الاستلام,@الموضوع,@رقم_معاملة_القسم,@الإجراء_الذي_تم,@مسوؤل_الملف,@الموظف,@حالة_الأرشفة,@اسم_المواطن_موضوع_الإجراء,@حالة_الأجراء,@رقم_الايقاف,@المنطقة,@رقم_الملف,@actionSum,@الحالة,@country)";
        //    SqlConnection sqlCon = new SqlConnection(dataSource);
        //    if (sqlCon.State == ConnectionState.Closed)
        //        sqlCon.Open();
        //    SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
        //    sqlCmd.CommandType = CommandType.Text;
        //    sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", messNo);
        //    sqlCmd.Parameters.AddWithValue("@المصدر", المصدر.Text);
        //    sqlCmd.Parameters.AddWithValue("@تاريخ_الإصدار", تاريخ_الإصدار.Text);
        //    sqlCmd.Parameters.AddWithValue("@تاريخ_الاستلام", تاريخ_الاستلام.Text);
        //    sqlCmd.Parameters.AddWithValue("@الموضوع", الموضوع.Text);
        //    sqlCmd.Parameters.AddWithValue("@رقم_معاملة_المصدر", رقم_معاملة_المصدر.Text);
        //    sqlCmd.Parameters.AddWithValue("@الإجراء_الذي_تم", txtAction1.Text + Environment.NewLine + "تمت الارشفة  بواسطة  " + labEmp.Text + "-" + GregorianDate + Environment.NewLine + "--------------------------"); 
        //    sqlCmd.Parameters.AddWithValue("@مسوؤل_الملف", مسوؤل_الملف.Text);
        //    sqlCmd.Parameters.AddWithValue("@actionSum", responce.Text);
        //    sqlCmd.Parameters.AddWithValue("@الموظف", labEmp.Text);
        //    sqlCmd.Parameters.AddWithValue("@المنطقة", المنطقة.Text);
        //    sqlCmd.Parameters.AddWithValue("@اسم_المواطن_موضوع_الإجراء", اسم_المواطن_موضوع_الإجراء.Text);
        //    sqlCmd.Parameters.AddWithValue("@الحالة", الحالة.Text);
        //    sqlCmd.Parameters.AddWithValue("@حالة_الأرشفة", "مؤرشف مبدئيا");
        //    sqlCmd.Parameters.AddWithValue("@حالة_الأجراء", "قيد الإجراء");
        //    sqlCmd.Parameters.AddWithValue("@رقم_الايقاف", رقم_الايقاف.Text);
        //    sqlCmd.Parameters.AddWithValue("@رقم_الملف", fileNo);
        //    sqlCmd.Parameters.AddWithValue("@country", country.Text);
            
        //    sqlCmd.ExecuteNonQuery();

        //    sqlCon.Close();
        //}

        private void insertDoc(string id, string date,string employee, string dataSource, string extn1, string DocName1, string messNo,string docType, byte[] buffer1)
        {
            //GregorianDate1
            string query = "INSERT INTO TableGeneralArch (Data1,Extension1,نوع_المستند,رقم_معاملة_القسم,المستند,الموظف,التاريخ,رقم_المرجع) values (@Data1,@Extension1,@نوع_المستند,@رقم_معاملة_القسم,@المستند,@الموظف,@التاريخ,@رقم_المرجع)";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", messNo);
            sqlCmd.Parameters.AddWithValue("@نوع_المستند", docType);
            sqlCmd.Parameters.AddWithValue("@الموظف", employee);
            sqlCmd.Parameters.AddWithValue("@التاريخ", date);
            sqlCmd.Parameters.AddWithValue("@رقم_المرجع", id);

            sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
            sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
            sqlCmd.Parameters.Add("@المستند", SqlDbType.NVarChar).Value = DocName1;


            //if (filePath != "")
            //{

            //    using (Stream stream = File.OpenRead(filePath))
            //    {
            //        byte[] buffer1 = new byte[stream.Length];
            //        stream.Read(buffer1, 0, buffer1.Length);
            //        var fileinfo1 = new FileInfo(filePath);
            //        string extn1 = fileinfo1.Extension;
            //        string DocName1 = fileinfo1.Name;
            //        sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
            //        sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
            //        sqlCmd.Parameters.Add("@المستند", SqlDbType.NVarChar).Value = DocName1;

            //    }
            //}
            sqlCmd.ExecuteNonQuery();

            sqlCon.Close();
        }


        private int NormalAddEdit(string dataSource, string messNo, int id)
        {
            if (!رقم_الملف.Text.All(char.IsDigit)) { MessageBox.Show("يرج اختيار رقم ملف صالح"); return 0; }
            string query1 = "INSERT INTO TableMessages (رقم_معاملة_المصدر,المصدر,تاريخ_الإصدار,تاريخ_الاستلام,الموضوع,رقم_معاملة_القسم,الإجراء_الذي_تم,مسوؤل_الملف,الموظف,حالة_الأرشفة,اسم_المواطن_موضوع_الإجراء,رقم_الايقاف,المنطقة,تاريخ_المعاملة,رقم_الملف,actionSum,الحالة,country,تاريخ_الاجراء,reference,النوع,المهنة) values(@رقم_معاملة_المصدر,@المصدر,@تاريخ_الإصدار,@تاريخ_الاستلام,@الموضوع,@رقم_معاملة_القسم,@الإجراء_الذي_تم,@مسوؤل_الملف,@الموظف,@حالة_الأرشفة,@اسم_المواطن_موضوع_الإجراء,@رقم_الايقاف,@المنطقة,@تاريخ_المعاملة,@رقم_الملف,@actionSum,@الحالة,@country,@تاريخ_الاجراء,@reference,@النوع,@المهنة);SELECT @@IDENTITY as lastid";
            string query2 = "UPDATE TableMessages SET رقم_معاملة_المصدر = @رقم_معاملة_المصدر, المصدر = @المصدر, تاريخ_الإصدار = @تاريخ_الإصدار, تاريخ_الاستلام = @تاريخ_الاستلام, الموضوع = @الموضوع, رقم_معاملة_القسم = @رقم_معاملة_القسم, الإجراء_الذي_تم = @الإجراء_الذي_تم, مسوؤل_الملف = @مسوؤل_الملف, الموظف = @الموظف, حالة_الأرشفة = @حالة_الأرشفة,اسم_المواطن_موضوع_الإجراء=@اسم_المواطن_موضوع_الإجراء,رقم_الايقاف=@رقم_الايقاف,المنطقة=@المنطقة,رقم_الملف=@رقم_الملف,actionSum=@actionSum,الحالة=@الحالة,country=@country,responce=@responce ,reference=@reference,النوع=@النوع,المهنة=@المهنة WHERE ID = @ID";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();

            if (id != 0)
            {
                SqlCommand sqlCmd = new SqlCommand(query2, sqlCon);
                sqlCmd.CommandType = CommandType.Text;

                sqlCmd.Parameters.AddWithValue("@ID", id);
                sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", messNo);
                sqlCmd.Parameters.AddWithValue("@المصدر", المصدر.Text);
                sqlCmd.Parameters.AddWithValue("@تاريخ_الإصدار", تاريخ_الإصدار.Text);
                sqlCmd.Parameters.AddWithValue("@تاريخ_الاستلام", تاريخ_الاستلام.Text);
                sqlCmd.Parameters.AddWithValue("@الموضوع", الموضوع.Text);
                sqlCmd.Parameters.AddWithValue("@رقم_معاملة_المصدر", رقم_معاملة_المصدر.Text);
                sqlCmd.Parameters.AddWithValue("@الإجراء_الذي_تم", txtAction1.Text + Environment.NewLine + "تم التعديل  بواسطة  " + labEmp.Text + "-" + GregorianDate + Environment.NewLine + "--------------------------" + Environment.NewLine + الإجراء_الذي_تم.Text);
                sqlCmd.Parameters.AddWithValue("@مسوؤل_الملف", مسوؤل_الملف.Text);
                sqlCmd.Parameters.AddWithValue("@الموظف", labEmp.Text);
                sqlCmd.Parameters.AddWithValue("@اسم_المواطن_موضوع_الإجراء", اسم_المواطن_موضوع_الإجراء.Text);
                sqlCmd.Parameters.AddWithValue("@الحالة", الحالة.Text);
                sqlCmd.Parameters.AddWithValue("@حالة_الأرشفة", "مؤرشف مبدئيا");
                sqlCmd.Parameters.AddWithValue("@المنطقة", المنطقة.Text);
                sqlCmd.Parameters.AddWithValue("@رقم_الايقاف", رقم_الايقاف.Text);
                sqlCmd.Parameters.AddWithValue("@رقم_الملف", رقم_الملف.Text);
                sqlCmd.Parameters.AddWithValue("@actionSum", actionSum.Text);
                sqlCmd.Parameters.AddWithValue("@النوع", النوع.Text);
                sqlCmd.Parameters.AddWithValue("@المهنة", المهنة.Text);

                sqlCmd.Parameters.AddWithValue("@reference", رقم_الملف.Text);
                sqlCmd.Parameters.AddWithValue("@responce", responce.Text);

                sqlCmd.Parameters.AddWithValue("@country", country.Text);
                sqlCmd.ExecuteNonQuery();
                return id;
            }
            else
            {
                SqlCommand sqlCmd = new SqlCommand(query1, sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@ID", 1);
                sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", messNo);
                sqlCmd.Parameters.AddWithValue("@المصدر", المصدر.Text);
                sqlCmd.Parameters.AddWithValue("@تاريخ_الإصدار", تاريخ_الإصدار.Text);
                sqlCmd.Parameters.AddWithValue("@تاريخ_الاستلام", تاريخ_الاستلام.Text);
                sqlCmd.Parameters.AddWithValue("@الموضوع", الموضوع.Text);
                sqlCmd.Parameters.AddWithValue("@رقم_معاملة_المصدر", رقم_معاملة_المصدر.Text);
                sqlCmd.Parameters.AddWithValue("@الإجراء_الذي_تم", txtAction1.Text + Environment.NewLine + "تمت الاضافة  بواسطة  " + labEmp.Text + "-" + GregorianDate + Environment.NewLine + "--------------------------");
                sqlCmd.Parameters.AddWithValue("@مسوؤل_الملف", مسوؤل_الملف.Text);
                sqlCmd.Parameters.AddWithValue("@الموظف", labEmp.Text);
                if (اسم_المواطن_موضوع_الإجراء.Text.Contains("_"))
                {
                    sqlCmd.Parameters.AddWithValue("@اسم_المواطن_موضوع_الإجراء", اسم_المواطن_موضوع_الإجراء.Text.Split('_')[0]);
                    
                }
                else sqlCmd.Parameters.AddWithValue("@اسم_المواطن_موضوع_الإجراء", اسم_المواطن_موضوع_الإجراء.Text);
                sqlCmd.Parameters.AddWithValue("@النوع", النوع.Text);
                sqlCmd.Parameters.AddWithValue("@المهنة", المهنة.Text);

                sqlCmd.Parameters.AddWithValue("@الحالة", الحالة.Text);
                sqlCmd.Parameters.AddWithValue("@حالة_الأرشفة", "مؤرشف مبدئيا");
                sqlCmd.Parameters.AddWithValue("@المنطقة", المنطقة.Text);
                sqlCmd.Parameters.AddWithValue("@رقم_الايقاف", رقم_الايقاف.Text);
                sqlCmd.Parameters.AddWithValue("@رقم_الملف", رقم_الملف.Text);
                sqlCmd.Parameters.AddWithValue("@تاريخ_المعاملة", dateTimeTo.Text);
                if (!reference.Contains(رقم_الملف.Text) || reference == "" || reference=="تم")                
                    sqlCmd.Parameters.AddWithValue("@reference", reference + "_"+ رقم_الملف.Text);
                else 
                    sqlCmd.Parameters.AddWithValue("@reference", رقم_الملف.Text);

                sqlCmd.Parameters.AddWithValue("@actionSum", actionSum.Text);

                sqlCmd.Parameters.AddWithValue("@country", country.Text);
                sqlCmd.Parameters.AddWithValue("@تاريخ_الاجراء", GregorianDate);
                //sqlCmd.ExecuteNonQuery();
                var reader = sqlCmd.ExecuteReader();
                if (reader.Read())
                {

                    return Convert.ToInt32(reader["lastid"].ToString());
                }
            }
            
            sqlCon.Close();
            return 0;
        }
        


        private void timer1_Tick(object sender, EventArgs e)
        {
            //if (combAction.SelectedIndex >= 7) {
            //    panelRespond.BringToFront();                
            //}else 
            //    panelPic.BringToFront();

            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            GregorianDate = DateTime.Now.ToString("MM-dd-yyyy");
            GregorianDate1 = DateTime.Now.ToString("MM-dd-yyyy hh: mm");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            finalArch = false;
            if (رقم_معاملة_القسم.Text.Length <= 0)
            {
                MessageBox.Show("يرجى كتابة الرقم المرجعي كاملا");
                return;
            }
            LoadDocs();

        }

        private void LoadDocs()
        {
            docIDNumber = "ق س ج/80/22/14/" + رقم_معاملة_القسم.Text;
            OpenFile2(docIDNumber);
            OpenFile1(docIDNumber);

            if (DocNoPro != "")
            {
                btnArch.Visible = loadPic.Visible = true;
                if (CurrentFile1 != "")
                {
                    حالة_الأرشفة.Text = حالة_الأرشفة.Text = "مؤرشف أوليا بالإسام أعلاه ";

                    btnArch.Visible = true;
                }

                else if (CurrentFile2 != "")
                {
                    حالة_الأرشفة.Text = "مؤرشف نهائيا بالاسم أعلاه ";
                    btnArch.Visible = true;

                }
                else return;
                
            }
        }

        private string loadName(string documenNo)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT اسم_المواطن_موضوع_الإجراء from TableMessages where رقم_معاملة_القسم=@رقم_معاملة_القسم", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_معاملة_القسم", documenNo);

            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string rowCnt = "معاملة غير موجودة";

            foreach (DataRow row in dtbl.Rows)
            {
                rowCnt = row["اسم_المواطن_موضوع_الإجراء"].ToString();
            }
            return rowCnt;

        }

        private bool checkExist(string documenNo, string TableFiles)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * from "+ TableFiles+" where FileNo = @FileNo", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@FileNo", documenNo);

            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count == 0)
                return false;
            else return true;

        }

        private bool checkFileExist(string documenNo, string TableFiles)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT FileName1 from " + TableFiles + " where FileNo = @FileNo", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@FileNo", documenNo);

            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow row in dtbl.Rows)
            {
                if (row["FileName1"].ToString() != "") return true;
                 
            }

            return false;
        }

        private void NoteVerbal_Load(object sender, EventArgs e)
        {
            //autoCompleteTextBox(txtFileLocation, DataSource, "مسوؤل_الملف", "TableMessages");
            Console.WriteLine(6);
            autoCompleteTextBox(المهنة, DataSource, "jobs", "TableListCombo"); 
            autoCompleteTextBox(actionSum, DataSource, "actionSum", "TableMessages"); 
            
            //fileComboBox(comboRespo, DataSource, "TextTitle", "TableAddContextAffair", "2"); comboRespo
            fileComboBox(country, DataSource, "ArabCountries", "TableListCombo");
            
            itemsCombo(رقم_الملف, "رقم_الملف", "TableMessages");
            FillDataGridView(colList);
            if (رقم_الملف.Items.Count > 0) رقم_الملف.SelectedIndex = رقم_الملف.Items.Count - 1;
            Console.WriteLine(7);
        }
        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName, string div)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + ",division,createBox from " + tableName;
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
                        if (dataRow[comlumnName].ToString() != "" && dataRow["division"].ToString() == div)
                        {
                            combbox.Items.Add(dataRow[comlumnName].ToString());

                            if (dataRow["createBox"].ToString() == "قائمة المستندات:")
                            {
                                actionSum.Size = new System.Drawing.Size(419, 274);
                                
                                actionSum.Location = new System.Drawing.Point(3, 127);
                            }
                            else if (dataRow["createBox"].ToString() == "الجنسية:")
                            {
                                actionSum.Size = new System.Drawing.Size(419, 362);
                                actionSum.Location = new System.Drawing.Point(3, 39);
                            }
                            else
                            {
                                actionSum.Size = new System.Drawing.Size(419, 362);
                                actionSum.Location = new System.Drawing.Point(3, 39);
                            }                            
                            
                            }
                    }
                }
                saConn.Close();
            }
        }
        

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            actionDone = "";
            if (dataGridView1.CurrentRow.Index != -1)
            {
                رقم_معاملة_القسم.Text = dataGridView1.CurrentRow.Cells["رقم_معاملة_القسم"].Value.ToString().Split('/')[4];
                fillFileData(DataSource, dataGridView1.CurrentRow.Cells["رقم_معاملة_القسم"].Value.ToString());
                ID = Convert.ToInt32(dataGridView1.CurrentRow.Cells["ID"].Value.ToString());
                colIDs[0] = dataGridView1.CurrentRow.Cells["رقم_معاملة_القسم"].Value.ToString();
                colIDs[1] = dataGridView1.CurrentRow.Cells["ID"].Value.ToString();
                if (JobPosition.Contains("قنصل"))
                {
                    responce.Size = new System.Drawing.Size(352, 35);
                    comboProc.Visible = true;
                    الإجراء_الذي_تم.Enabled = btndelete.Visible = true;                    
                }
                else
                    comboProc.Visible = btndelete.Visible = false;

                اسم_المواطن_موضوع_الإجراء.Text = dataGridView1.CurrentRow.Cells["اسم_المواطن_موضوع_الإجراء"].Value.ToString();
                if (اسم_المواطن_موضوع_الإجراء.Text == "")
                {
                    newData = true;
                    OpenFileDoc(dataGridView1.CurrentRow.Cells["رقم_معاملة_القسم"].Value.ToString(), "TableMessages", "المستندات_الأولية");
                    labLnfo.Visible = dataGridView1.Visible = false;
                    panelMain.Visible = true;
                    return;
                } 

                newData = false;


                رقم_معاملة_المصدر.Text = dataGridView1.CurrentRow.Cells["رقم_معاملة_المصدر"].Value.ToString();
                المصدر.Text = dataGridView1.CurrentRow.Cells["المصدر"].Value.ToString();
                تاريخ_الإصدار.Text = dataGridView1.CurrentRow.Cells["تاريخ_الإصدار"].Value.ToString();
                تاريخ_الاستلام.Text = dataGridView1.CurrentRow.Cells["تاريخ_الاستلام"].Value.ToString();
                الموضوع.Text = dataGridView1.CurrentRow.Cells["الموضوع"].Value.ToString();
                النوع.Text = dataGridView1.CurrentRow.Cells["النوع"].Value.ToString();
                رقم_معاملة_القسم.Enabled= false;

                if (النوع.Text == "ذكر")
                {
                    النوع.CheckState = CheckState.Checked;
                }
                else
                {
                    النوع.CheckState = CheckState.Unchecked;
                }
                الإجراء_الذي_تم.Text = dataGridView1.CurrentRow.Cells["الإجراء_الذي_تم"].Value.ToString();
                الإجراء_الذي_تم.Visible = true;
                مسوؤل_الملف.Text = dataGridView1.CurrentRow.Cells["مسوؤل_الملف"].Value.ToString();
                حالة_الأرشفة.Text = dataGridView1.CurrentRow.Cells["حالة_الأرشفة"].Value.ToString();
                حالة_الأجراء = dataGridView1.CurrentRow.Cells["حالة_الأجراء"].Value.ToString();
                رقم_الايقاف.Text = dataGridView1.CurrentRow.Cells["رقم_الايقاف"].Value.ToString();
                المنطقة.Text = dataGridView1.CurrentRow.Cells["المنطقة"].Value.ToString();
                responce.Text = dataGridView1.CurrentRow.Cells["responce"].Value.ToString();
                reference = dataGridView1.CurrentRow.Cells["reference"].Value.ToString();

                if (reference.Contains("_"))
                {
                    refFile = "الموقوف تم ايادعه برقم الملف: " + reference.Split('_')[0];
                    for (int x = 1; x < reference.Split('_').Length; x++) 
                    {
                        refFile = refFile + Environment.NewLine + "الموقوف تمت إحالته إلى الملف بالرقم: " + reference.Split('_')[x];
                    }
                    if (refFile != "")
                        Panel_PaintLabel(refFile);

                }
                    recommon = dataGridView1.CurrentRow.Cells["recommendation"].Value.ToString();
                الحالة.Text = dataGridView1.CurrentRow.Cells["الحالة"].Value.ToString();
                المهنة.Text = dataGridView1.CurrentRow.Cells["المهنة"].Value.ToString();
                if (الحالة.Text == "مجهول")
                    الحالة.Checked = true;
                else
                    الحالة.Checked = false;
                country.Text = dataGridView1.CurrentRow.Cells["country"].Value.ToString();
                actionSum.Text = dataGridView1.CurrentRow.Cells["actionSum"].Value.ToString();
                labLnfo.Visible = dataGridView1.Visible = false;
                panelMain.Visible = true;
                رقم_الملف.Text = dataGridView1.CurrentRow.Cells["رقم_الملف"].Value.ToString();


                string الملف_الارشفة = dataGridView1.CurrentRow.Cells["الملف_الارشفة"].Value.ToString();
                foreach (Control control in panelArch.Controls)
                {
                    control.Visible = false;
                    control.Name = "لاغي";
                }
                if (الملف_الارشفة.Contains("*"))
                {
                    string[] fileList = الملف_الارشفة.Split('*');
                    for (int x = 0; x < fileList.Length; x++)
                    {
                        string btnText = "الموقوف مذكور بالملف بالرقم (" + fileList[x].Split('_')[0] + ") بالحالة رقم (" + fileList[x].Split('_')[2] + ") : " + fileList[x].Split('_')[1];
                        Panel_Paint(fileList[x].Split('_')[0], btnText);
                    }
                }
                else if (الملف_الارشفة.Contains("_"))
                {
                    string btnText = "الموقوف مذكور بالملف بالرقم (" + الملف_الارشفة.Split('_')[0] + ") بالحالة رقم (" + الملف_الارشفة.Split('_')[2] + ") : " + الملف_الارشفة.Split('_')[1];
                    
                    Panel_Paint(الملف_الارشفة.Split('_')[0], btnText);
                }
                panelArch.Visible = true;
                panelArch.BringToFront();                
            }
        }

        private void Panel_Paint(string index,string text)
        {
            Button button = new Button();
                button.Dock = System.Windows.Forms.DockStyle.Top;
            button.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            button.Location = new System.Drawing.Point(0, 0);
            button.Name = index;
            button.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            button.Size = new System.Drawing.Size(166, 42);
            button.TabIndex = 740;
            button.Text = text;
            button.UseVisualStyleBackColor = true;
            button.Click += new System.EventHandler(this.button_Click);
            if(!checkFileExist(index, "TableFiles"))
                button.Enabled = false;
            panelArch.Controls.Add(button);
            ResponceIndex++;
            
        }

        private void Panel_PaintLabel( string text)
        {
            int lines = text.Split('\n').Length;
            Label label= new Label();
            label.Dock = System.Windows.Forms.DockStyle.Top;
            label.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label.Location = new System.Drawing.Point(0, 0);
            label.Name = "label";
            label.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            label.Size = new System.Drawing.Size(166, 42 * lines);
            label.TabIndex = 740;
            label.Text = text;
            

            //Label label1 = new Label();
            //label1.Dock = System.Windows.Forms.DockStyle.Top;
            //label1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            //label1.Location = new System.Drawing.Point(0, 0);
            //label1.Name = "label";
            //label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            //label1.Size = new System.Drawing.Size(166, 42);
            //label1.TabIndex = 740;
            //label1.Text = "";

            //panelArch.Controls.Add(label1);
            panelArch.Controls.Add(label);
            

        }
        private void getDocInfo(string fileNo)
        {
            string query;
            string NewFileName = "";
            SqlConnection Con = new SqlConnection(DataSource);
            query = "select Data1, Extension1,FileName1 from TableFiles  where FileNo=@FileNo";
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@FileNo", SqlDbType.NVarChar).Value = fileNo;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                var name = reader["FileName1"].ToString();
                if (string.IsNullOrEmpty(name)) return ;
                var ext = reader["Extension1"].ToString();
                if (string.IsNullOrEmpty(ext)) return ;
                var Data = (byte[])reader["Data1"];
                NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                File.WriteAllBytes(NewFileName, Data);
                System.Diagnostics.Process.Start(NewFileName);
            }
            Con.Close();
        }
        private void button_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            getDocInfo(button.Name);
        }
        
        private void btnFinished_Click(object sender, EventArgs e)
        {
            finished();
            panelMain.Visible = false;
            labLnfo.Visible = dataGridView1.Visible = true;
            timer2.Enabled = true;
        }

        private void finished()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableMessages SET حالة_الأجراء=@حالة_الأجراء WHERE رقم_معاملة_القسم=@رقم_معاملة_القسم", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@حالة_الأجراء", "تم الإجراء");
            sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", "ق س ج/80/22/14/" + رقم_معاملة_القسم.Text);
            sqlCmd.ExecuteNonQuery();

            sqlCon.Close();
        }

        private void btnUnderprocess_Click(object sender, EventArgs e)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableMessages SET حالة_الأجراء=@حالة_الأجراء WHERE رقم_معاملة_القسم=@رقم_معاملة_القسم", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@حالة_الأجراء", "قيد الإجراء");
            sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", "ق س ج/80/22/14/" + رقم_معاملة_القسم.Text);
            sqlCmd.ExecuteNonQuery();

            sqlCon.Close();
            panelMain.Visible = false;
            labLnfo.Visible = dataGridView1.Visible = true;
            timer2.Enabled = true;
        }

        private void btndelete_Click(object sender, EventArgs e)
        {
            var selectedOption = MessageBox.Show("", "تأكيد عملية الحذف", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                deleteRowsData(ID, "TableMessages", DataSource);
                FillDataGridView(colList);
                panelMain.Visible = false;
                labLnfo.Visible = dataGridView1.Visible = true;
                timer2.Enabled = true;
            }
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
        }
        private void getText(TextBox textBox, ComboBox comboBox, Button button)
        {
            using (SqlConnection saConn = new SqlConnection(DataSource))
            {
                saConn.Open();
                string query = "select ID,TextModel,TextTitle,createBox from TableAddContextAffair";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                foreach (DataRow dataRow in table.Rows)
                {
                    if (!string.IsNullOrEmpty(dataRow["TextTitle"].ToString()))
                    {
                        if (dataRow["TextTitle"].ToString() == comboBox.Text.Trim())
                        {
                            contextID = Convert.ToInt32(dataRow["ID"].ToString());
                            button.Text = "تعديل";
                            textBox.Text = dataRow["TextModel"].ToString();
                            if (responce == comboBox)
                            {
                                if (dataRow["createBox"].ToString() == "قائمة المستندات:")
                                {
                                    actionSum.Size = new System.Drawing.Size(419, 274);
                                    actionSum.Location = new System.Drawing.Point(3, 127);
                                }
                                else if (dataRow["createBox"].ToString() == "الجنسية:")
                                {
                                    actionSum.Size = new System.Drawing.Size(419, 362);
                                    actionSum.Location = new System.Drawing.Point(3, 39);
                                }
                                else {
                                    actionSum.Size = new System.Drawing.Size(419, 362);
                                    actionSum.Location = new System.Drawing.Point(3, 39);
                                }
                            }
                            for (int x = 0; x < 10; x++)
                                textBox.Text = SuffPrefReplacements(textBox.Text);
                            return;
                        }

                    }
                }
                saConn.Close();
            }
        }

        private string SuffPrefReplacements(string text)
        {
            int index = 0;
            string title = "";
            if (النوع.CheckState == CheckState.Unchecked)
            {
                title = "ة";
                index = 1;
            }
            
                
            if (text.Contains("#1"))
                return text.Replace("#1", رقم_معاملة_المصدر.Text);
            if (text.Contains("#2"))
                return text.Replace("#2", تاريخ_الإصدار.Text);
            if (text.Contains("#3"))
                return text.Replace("#3", اسم_المواطن_موضوع_الإجراء.Text);
            if (text.Contains("#4"))
                return text.Replace("#4", رقم_الايقاف.Text);
            if (text.Contains("#5"))
                return text.Replace("#5", preffix[index, 2]);
            if (text.Contains("#7"))
                return text.Replace("#7", Environment.NewLine);
            if (text.Contains("$$$"))
                return text.Replace("$$$", preffix[index, 7]);
            if (text.Contains("###"))
                return text.Replace("###", preffix[index, 1]);
            if (text.Contains("***"))
                return text.Replace("***", preffix[index, 0]);
            if (text.Contains("%%%"))
                return text.Replace("%%%", preffix[index, 2]);
            if (text.Contains("@@@"))
                return text.Replace("@@@", preffix[index, 8]);
            else return text;
        }
        private void Suffex_preffixList()
        {

            preffix[0, 0] = "";//***
            preffix[1, 0] = "ت";
            preffix[2, 0] = "ا";
            preffix[3, 0] = "تا";
            preffix[4, 0] = "ن";
            preffix[5, 0] = "وا";

            preffix[0, 1] = "ه";//###
            preffix[1, 1] = "ها";
            preffix[2, 1] = "هما";
            preffix[3, 1] = "هما";
            preffix[4, 1] = "هن";
            preffix[5, 1] = "هم";

            preffix[0, 2] = "";//%%%
            preffix[1, 2] = "ة";
            preffix[2, 2] = "ان";
            preffix[3, 2] = "تان";
            preffix[4, 2] = "ات";
            preffix[5, 2] = "ون";

            preffix[0, 3] = "";//#5
            preffix[1, 3] = "ة";
            preffix[2, 3] = "ين";
            preffix[3, 3] = "تين";
            preffix[4, 3] = "ات";
            preffix[5, 3] = "ين";



            preffix[0, 4] = "ت";//#*#
            preffix[1, 4] = "";

            preffix[0, 5] = "التي";//#1
            preffix[1, 5] = "الذي";


            preffix[0, 7] = "يكون";//$$$
            preffix[1, 7] = "تكون";
            preffix[2, 7] = "يكونا";
            preffix[3, 7] = "تكونا";
            preffix[4, 7] = "يكن";
            preffix[5, 7] = "يكونو";

            preffix[0, 8] = "يقدم";//@@@
            preffix[1, 8] = "تقدم";
            preffix[2, 8] = "يقدما";
            preffix[3, 8] = "تقدما";
            preffix[4, 8] = "يقدمن";
            preffix[5, 8] = "يقدموا";
        }
        private void button4_Click(object sender, EventArgs e)
        {
            //رقم_الملف.Text = "";
            FillDataGridView(colList);
            if (dataGridView1.Visible == true)
            {
                labLnfo.Visible = dataGridView1.Visible = false;
                panelMain.Visible = true;
            }
            else {
                labLnfo.Visible = dataGridView1.Visible = true;
                panelMain.Visible = false;
                timer2.Enabled = true;
            }
        }

        private void autoCompleteTextBox(TextBox combbox, string source, string comlumnName, string tableName)
        {
            AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                DataTable Textboxtable = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(Textboxtable);

                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    if (dataRow[comlumnName].ToString() != "")
                    {
                        autoComplete.Add(dataRow[comlumnName].ToString());
                        //MessageBox.Show(dataRow[comlumnName].ToString());
                    }
                }
                combbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                combbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                combbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }

        
        private void timer2_Tick(object sender, EventArgs e)
        {
            //if (dataGridView1.Visible) 
            //    Colorcomment1(22);
        }
        string lastInput2 = "";
        private void txtIssueDate_TextChanged(object sender, EventArgs e)
        {
            if (تاريخ_الإصدار.Text.Length == 11)
            {
                تاريخ_الإصدار.Text = lastInput2; return;
            }
            if (تاريخ_الإصدار.Text.Length == 10) return;
            if (تاريخ_الإصدار.Text.Length == 4) تاريخ_الإصدار.Text = "-"+تاريخ_الإصدار.Text ;
            else if (تاريخ_الإصدار.Text.Length == 7) تاريخ_الإصدار.Text = "-"+تاريخ_الإصدار.Text;
            lastInput2 = تاريخ_الإصدار.Text;
        }
        string lastInput1 = "";
        private void txtReceiveDate_TextChanged(object sender, EventArgs e)
        {
            if (تاريخ_الاستلام.Text.Length == 11)
            {
                 تاريخ_الاستلام.Text = lastInput1; return;
            }
            if (تاريخ_الاستلام.Text.Length == 4) تاريخ_الاستلام.Text = "-"+تاريخ_الاستلام.Text;
            else if (تاريخ_الاستلام.Text.Length == 7) تاريخ_الاستلام.Text = "-"+تاريخ_الاستلام.Text;
            lastInput1 = تاريخ_الاستلام.Text;       
        }

        private void ProType_CheckedChanged(object sender, EventArgs e)
        {
            if (الحالة.Checked)
            {
                المصدر.Text = "إختر نوع المصدر";
                الحالة.Text = "مجهول";
                الإجراء_الذي_تم.Enabled = txtAction1.Enabled = المصدر.Enabled = تاريخ_الاستلام.Enabled = الإجراء_الذي_تم.Enabled = الموضوع.Enabled = تاريخ_الإصدار.Enabled = رقم_معاملة_المصدر.Enabled = true;
                المهنة.Enabled = false;
                المهنة.Text = "مهنة غير نظامية";
            }
            else
            {

                المهنة.Text = "";
                الحالة.Text = "مقيم نظامي";
                مسوؤل_الملف.Text = "مدير إدارة الجوازات والسجل المدني";
                المصدر.Text = المنطقة.Text;
                الإجراء_الذي_تم.Enabled = txtAction1.Enabled = المصدر.Enabled = تاريخ_الاستلام.Enabled = الإجراء_الذي_تم.Enabled = الموضوع.Enabled = تاريخ_الإصدار.Enabled = رقم_معاملة_المصدر.Enabled = false;
                المهنة.Enabled = true;
            }

        }
        
        private void combFileNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!رقم_الملف.Text.All(char.IsDigit) || رقم_الملف.Text == "الجميع")
            {
                SqlConnection sqlCon = new SqlConnection(DataSource);
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("select " + colList + " from TableMessages order by responce asc, actionSum", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);
                dataGridView1.DataSource = dtbl;
                rowCount = dataGridView1.Rows.Count.ToString();

                dataGridView1.Columns["ID"].Visible = false;
                dataGridView1.Columns["رقم_معاملة_القسم"].Width = 130;
                dataGridView1.Columns["المصدر"].Width = 130;
                dataGridView1.Columns["الموضوع"].Width = 150;
                dataGridView1.Columns["اسم_المواطن_موضوع_الإجراء"].Width = 150;

                sqlCon.Close();
            }
            else
            {
                //BindingSource bs = new BindingSource();
                //bs.DataSource = dataGridView1.DataSource;
                //bs.Filter = dataGridView1.Columns[17].HeaderText.ToString() + " LIKE '" + combFileNo.Text + "%'";
                //dataGridView1.DataSource = bs;
                //dataGridView1.Sort(dataGridView1.Columns["responce"], System.ComponentModel.ListSortDirection.Ascending);

                SqlConnection sqlCon = new SqlConnection(DataSource);
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("select "+ colList+ " from TableMessages where رقم_الملف = " + رقم_الملف.Text + " order by responce asc, actionSum", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);                
                dataGridView1.DataSource = dtbl;
                rowCount = dataGridView1.Rows.Count.ToString();

                dataGridView1.Columns["ID"].Visible = false;
                dataGridView1.Columns["رقم_معاملة_القسم"].Width = 130;
                dataGridView1.Columns["المصدر"].Width = 130;
                dataGridView1.Columns["الموضوع"].Width = 150;
                dataGridView1.Columns["اسم_المواطن_موضوع_الإجراء"].Width = 150;

                sqlCon.Close();

                button1.Text = "طباعة القائمة";
            }
            Colorcomment1(22);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            button1.Enabled = false;
            ActiveCopy = PrimariFiles + "Docx1" + DateTime.Now.ToString("mmss") + ".docx";
            string fileNo = "NV" + رقم_الملف.Text + (dataGridView1.RowCount - 1).ToString();
            var selectionMessage = MessageBox.Show("","إضافة الملف إلى قائمة ملخص الملفات؟", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (selectionMessage == DialogResult.OK)
            {
                if (!checkExist(fileNo, "TableFiles"))
                    UpdateFileList(0, fileNo, "الوافدين الشميسي", (dataGridView1.RowCount - 1).ToString(), GregorianDate, (dataGridView1.RowCount - 1).ToString() + رقم_الملف.Text, "", "", "", "", "", "");
            }
            createTables(fileNo);
            //if(combFileNo.SelectedIndex >= 0)
            //speaclNormalLetters();
            button1.Text = "رقم الملف";
            button1.Enabled = true;
            }

        private int VCIndexData()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT VCIndesx FROM TableSettings", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            int index = 0;
            foreach (DataRow dataRow in table.Rows)
            {
                if (!string.IsNullOrEmpty(dataRow["VCIndesx"].ToString()))
                {
                    index = Convert.ToInt32(dataRow["VCIndesx"].ToString());
                }
            }
            return index;
        }


        private void NoteVerbal_FormClosed(object sender, FormClosedEventArgs e)
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

        private void SearchFile_TextChanged(object sender, EventArgs e)
        {
            
            if (!ListSearch.Text.All(char.IsDigit) && ListSearch.Text.Length > 0 && !ListSearch.Text.Contains("/"))
            {
                Console.WriteLine("File No 5555555" + 1);
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns["اسم_المواطن_موضوع_الإجراء"].HeaderText.ToString() + " LIKE '%" + ListSearch.Text + "%'";
                dataGridView1.DataSource = bs;
            }
            else if (ListSearch.Text.All(char.IsDigit) && ListSearch.Text.Length > 0)
            {
                Console.WriteLine("File No 11111" + 1);
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns["رقم_الايقاف"].HeaderText.ToString() + " LIKE '%" + ListSearch.Text.Trim() + "%'";
                dataGridView1.DataSource = bs;
            }
            else if (ListSearch.Text.Contains("/") && ListSearch.Text.Length > 0)
            {
                Console.WriteLine("File No 77777" + 1);
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                string strList = "ق س ج/80/22/14" + ListSearch.Text;
                bs.Filter = dataGridView1.Columns["رقم_معاملة_القسم"].HeaderText.ToString() + " LIKE '%" + strList + "%'";
                dataGridView1.DataSource = bs;
            }
            else if(ListSearch.Text.Length == 0)
            {
                Console.WriteLine("File No 88888" + 1);
                FillDataGridView(colList);
            }

            //if (dataGridView1.RowCount == 1 && ListSearch.Text.StartsWith("9"))
            //{
            //    Console.WriteLine("File No 99999" + 3);
            //    FillDataGridView();
            //}
        }

        

        private void txtResponce_Click(object sender, EventArgs e)
        {
            
        }
        
        private void updateText(int id, string text1, string text2, string text3, string text4)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAddContextAffair SET TextTitle=@TextTitle,TextModel=@TextModel,division=@division,createBox=@createBox, WHERE ID=@id", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@TextTitle", text1);
            sqlCmd.Parameters.AddWithValue("@TextModel", text2);
            sqlCmd.Parameters.AddWithValue("@division", text3);
            sqlCmd.Parameters.AddWithValue("@createBox", text4);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

       
        private void comboResponce_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            
            
        }

        private void picVerify_Click(object sender, EventArgs e)
        {
            checkDataInfo(true);
        }
        
        private void selectedID(int index)
        {
            ID = Convert.ToInt32(dataGridView1.Rows[index].Cells[0].Value.ToString());

            رقم_معاملة_المصدر.Text = dataGridView1.Rows[index].Cells[6].Value.ToString();
            المصدر.Text = dataGridView1.Rows[index].Cells[2].Value.ToString();
            تاريخ_الإصدار.Text = dataGridView1.Rows[index].Cells[3].Value.ToString();
            تاريخ_الاستلام.Text = dataGridView1.Rows[index].Cells[4].Value.ToString();
            الموضوع.Text = dataGridView1.Rows[index].Cells[5].Value.ToString();
            string[] str = dataGridView1.Rows[index].Cells[1].Value.ToString().Split('/');
            رقم_معاملة_القسم.Text = str[4];
            LoadDocs();
            رقم_معاملة_القسم.Enabled = false;
            string strList = dataGridView1.Rows[index].Cells[7].Value.ToString();
            if (strList.Contains("_"))
            {
                اسم_المواطن_موضوع_الإجراء.Text = strList.Split('_')[0];
                المهنة.Text = strList.Split('_')[2];
                if (strList.Split('_')[1] == "مجهول")
                    الحالة.Checked = true;
                else
                    الحالة.Checked = false;
            }
            else اسم_المواطن_موضوع_الإجراء.Text = strList;

            الإجراء_الذي_تم.Text = dataGridView1.Rows[index].Cells[8].Value.ToString();
            الإجراء_الذي_تم.Visible = true;
            مسوؤل_الملف.Text = dataGridView1.Rows[index].Cells[10].Value.ToString();
            حالة_الأرشفة.Text = dataGridView1.Rows[index].Cells[13].Value.ToString();
            رقم_الايقاف.Text = dataGridView1.Rows[index].Cells[15].Value.ToString();
            المنطقة.Text = dataGridView1.Rows[index].Cells[16].Value.ToString();
            رقم_الملف.Text = dataGridView1.Rows[index].Cells[17].Value.ToString();
            responces = dataGridView1.Rows[index].Cells[18].Value.ToString();
            reference = dataGridView1.Rows[index].Cells[19].Value.ToString();
            recommon = dataGridView1.Rows[index].Cells[20].Value.ToString();
            
            
            if (responces.Contains("*"))
            {
                string[] respList = responces.Split('*');
                for (int x = 0; x < respList.Length; x++)
                {
                    Panel_Paint(respList[x].Split('_')[1], respList[x].Split('_')[0] + " بالرقم " + respList[x].Split('_')[1]);
                }
                
            }
            else if (responces.Contains("_"))
            {
                Panel_Paint("1", responces.Split('_')[0] + " بالرقم " + responces.Split('_')[1]);
                
            }

            
            labLnfo.Visible = dataGridView1.Visible = false;
            panelMain.Visible = true;
        }
       
        private int checkDataInfo(bool show)
        {
            grdiFill = true;
            picVerify.Visible = true;
            picVerified.Visible = false;
            for (int z = 0; z < dataGridView1.RowCount - 1; z++)
            {
                if (dataGridView1.Rows[z].Cells[15].Value.ToString() == رقم_الايقاف.Text.Trim())
                {
                    if (show)
                    {
                        var selectedOption = MessageBox.Show("عرض؟", "رقم الايقاف تطابق مع إجرا سابق", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (selectedOption == DialogResult.Yes)
                        {
                           // selectedID(z);
                        }
                    }

                    return Convert.ToInt32(dataGridView1.Rows[z].Cells[0].Value.ToString());
                }
            }
            return -1;
        }

        private void checkID()
        {
            FillDataGridView(colList);
            string docID = "ق س ج/80/22/14/" + رقم_معاملة_القسم.Text;
            for (int z = 0; z < dataGridView1.RowCount - 1; z++)
            {
                if (dataGridView1.Rows[z].Cells[1].Value.ToString() == docID)
                {
                    MessageBox.Show("رقم المعاملة متطابق موجود مسبقا يرجى اعادة الإجراء");
                    this.Close();
                }
            }
            
        }

        private void checkSex_CheckedChanged(object sender, EventArgs e)
        {
            if (النوع.Checked) 
                النوع.Text = "ذكر";
            else 
                النوع.Text = "أنثى";
        }

        
        private void addText(int id, string text1, string text2, string text3, string text4)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO  TableAddContextAffair (TextTitle,TextModel,division,createBox) values(@TextTitle,@TextModel,@division,@createBox)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@TextTitle", text1);
            sqlCmd.Parameters.AddWithValue("@TextModel", text2);
            sqlCmd.Parameters.AddWithValue("@division", text3);
            sqlCmd.Parameters.AddWithValue("@createBox", text4);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        

        private void panel3_MouseEnter(object sender, EventArgs e)
        {
            
        }

        private void panel3_MouseLeave(object sender, EventArgs e)
        {
            
        }

        private void txtSubject_MouseEnter(object sender, EventArgs e)
        {
            
        }

        private void txtSubject_MouseLeave(object sender, EventArgs e)
        {
            
        }

        private void txtAction1_MouseEnter(object sender, EventArgs e)
        {
            
        }

        private void txtAction1_MouseLeave(object sender, EventArgs e)
        {
            
        }

        private void comboRespo_MouseEnter(object sender, EventArgs e)
        {
           
        }

        private void comboRespo_MouseLeave(object sender, EventArgs e)
        {
                    
        }

        private void comboRef_MouseEnter(object sender, EventArgs e)
        {
            
        }

        private void comboRef_MouseLeave(object sender, EventArgs e)
        {
                      
        }

        private void comboRecom_MouseEnter(object sender, EventArgs e)
        {
            
        }

        private void comboRecom_MouseLeave(object sender, EventArgs e)
        {
           
        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {
            panel3.Size = new System.Drawing.Size(441, 374);
            panel3.BringToFront();
            pictureBox11.Visible = false;
            pictureBox13.Visible = true;
        }

        private void pictureBox13_Click(object sender, EventArgs e)
        {
            panel3.Size = new System.Drawing.Size(441, 38);
            panel3.BringToFront();
            pictureBox11.Visible = true;
            pictureBox13.Visible = false;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            panel2.Size = new System.Drawing.Size(441, 331);
            panel2.BringToFront();
            pictureBox2.Visible = false;
            pictureBox6.Visible = true;
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            panel2.Size = new System.Drawing.Size(441, 38);
            panel2.BringToFront();
            pictureBox2.Visible = true;
            pictureBox6.Visible = false;

        }

       
        private void CombSource_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (CombSource.SelectedIndex == 1)
            //    comboDes.Visible = true;
            //else
            //    comboDes.Visible = false;
        }

        private void comboDes_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (ResponceIndex != 1)
            {
                var selectedOption = MessageBox.Show("", "تحرير مكاتبة جديدة", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                {
                    
                    currentNo = docIDNumber + "/" + ResponceIndex.ToString();
                }
                else if (selectedOption == DialogResult.No)
                {
                    currentNo = docIDNumber + "/" + CurrentBtnName;
                    
                }
            }
        }

        private void combAction_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combAction.SelectedIndex == 2)
            {
                panelIssueDoc.Visible = true;
                panelRespond.Visible = false;
                combAction.Size = new System.Drawing.Size(650, 537);
                combAction.Location = new System.Drawing.Point(17, 57);
                panelIssueDoc.BringToFront();
            }
            else if (combAction.SelectedIndex == 4)
            {
                combAction.Size = new System.Drawing.Size(650, 581);
                combAction.Location = new System.Drawing.Point(17, 13);
                panelIssueDoc.Visible = true;
                panelRespond.Visible = false;
                panelIssueDoc.BringToFront();

            }else panelIssueDoc.SendToBack();
        }

        private void createTables(string fileNo)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select responce,count(responce) as allData from TableMessages where responce <> '' and رقم_الملف='" + رقم_الملف.Text + "'group by responce", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string route = FilespathIn + @"\قوائم التعليقات.docx";
            string ReportName = DateTime.Now.ToString("mmss");
            ActiveCopy = FilespathOut + @"\DocxTableList" + ReportName + ".docx";
            System.IO.File.Copy(route, ActiveCopy);
            using (var document = DocX.Load(ActiveCopy))
            {
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
                string strHeader = "التاريخ :" + GregorianDate + " م" + "                                                          " + "الموافق: " + HijriDate + "هـ" + Environment.NewLine;
                document.InsertParagraph(strHeader)
                .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                .FontSize(20d)
                .Alignment = Alignment.center;
                string fileInfo = Environment.NewLine + " كشف رقم )" + رقم_الملف.Text + ":( الذي يحتوى على بيانات عدد )" + (dataGridView1.RowCount-1).ToString() + ":( موقوف";
                document.InsertParagraph(fileInfo)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(21d).UnderlineStyle(UnderlineStyle.singleLine)
                    .Alignment = Alignment.center;
                int table = 1;
                string title = "";
                foreach (DataRow row in dtbl.Rows)
                {
                    int rowsCount = Convert.ToInt32(row["allData"].ToString());
                    var t = document.AddTable(rowsCount + 1, 4);
                    t.Design = TableDesign.TableGrid;
                    t.Alignment = Alignment.center;
                    int count = 0;
                    t.SetColumnWidth(0, 240);
                    t.SetColumnWidth(1, 95);
                    t.SetColumnWidth(2, 140);
                    t.SetColumnWidth(3, 40);

                    t.Rows[0].Cells[1].Paragraphs[0].Append("رقم الموقوف").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[2].Paragraphs[0].Append("الاسم").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[3].Paragraphs[0].Append("الرقم").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    for (int x = 0; x < dataGridView1.RowCount - 1; x++)
                    {
                        int id = Convert.ToInt32(dataGridView1.Rows[x].Cells[0].Value.ToString());
                        Console.WriteLine(dataGridView1.RowCount.ToString());
                        string responces = dataGridView1.Rows[x].Cells["responce"].Value.ToString().Trim();
                        responces = responces.TrimStart();
                        responces = responces.TrimEnd();
                        if (row["responce"].ToString().Trim() == responces.Trim())
                        {
                            string AppNames = dataGridView1.Rows[x].Cells["اسم_المواطن_موضوع_الإجراء"].Value.ToString();
                            string arrestNo = dataGridView1.Rows[x].Cells["رقم_الايقاف"].Value.ToString();
                            string country = dataGridView1.Rows[x].Cells["country"].Value.ToString();
                            string requestedDoc = dataGridView1.Rows[x].Cells["actionSum"].Value.ToString();
                            string fileArchGrid = dataGridView1.Rows[x].Cells["الملف_الارشفة"].Value.ToString();
                            string fileArchNew = fileNo + "_" + responces + "_" + x.ToString();
                            if (responces.Contains("ينتمي"))
                            {
                                responces = country;
                                title = "قائمة الموقوفين الذين تبين بأن لهم جنسيات غير سودانية";
                                //t.Rows[0].Cells[0].Paragraphs[0].Append("الدولة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                            }
                            else if (responces == "الإجراء قد تم" && requestedDoc.Contains("غادر"))
                            {
                                responces = requestedDoc;
                                title = "قائمة الموقوفين الذين تم إنهاء وقد غادرو المملكة العربية السعودية";
                               // t.Rows[0].Cells[0].Paragraphs[0].Append("الإجراء الذي تم").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                            }
                            else if (responces == "الإجراء قد تم" && requestedDoc != "")
                            {
                                responces = requestedDoc;
                                title = "قائمة الموقوفين الذين تم إنهاء معاملاتهم من جانب القنصلية العامة";
                               // t.Rows[0].Cells[0].Paragraphs[0].Append("الإجراء الذي تم").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                            }

                            else if (responces.Contains("التحري"))
                            {
                                responces = "";
                                title = "قائمة الموقوفين الذين ما زالت معاملاتهم قيد التحري";
                                //t.Rows[0].Cells[0].Paragraphs[0].Append("الملاحظة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                            }
                            else if (responces.Contains("يقدم مستندات "))
                            {
                                responces = "";
                                title = "قائمة الموقوفين الذين تم التحري معهم مرات عدة ولم يقدموا مستندات تثبيت هويتهم السودانية";
                                //t.Rows[0].Cells[0].Paragraphs[0].Append("الملاحظة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                            }
                            else if (responces.Contains("مطلوب مستندات"))
                            {
                                responces = requestedDoc;
                                title = "قائمة المستندات المطلوبة من الموقفين لاثبات هويتهم السودانية";
                                //t.Rows[0].Cells[0].Paragraphs[0].Append("المستندات المطلوبة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                            }

                            t.Rows[count + 1].Cells[0].Paragraphs[0].Append(responces).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                            t.Rows[count + 1].Cells[1].Paragraphs[0].Append(arrestNo).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                            t.Rows[count + 1].Cells[2].Paragraphs[0].Append(AppNames).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                            t.Rows[count + 1].Cells[3].Paragraphs[0].Append((count + 1).ToString() + ".").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                            count++;
                        }
                    }

                    if (title == "قائمة الموقوفين الذين تبين بأن لهم جنسيات غير سودانية")
                    {
                        t.Rows[0].Cells[0].Paragraphs[0].Append("الدولة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    }
                    else if (title == "قائمة الموقوفين الذين تم إنهاء معاملاتهم من جانب القkصنلية العامة")
                    {
                        t.Rows[0].Cells[0].Paragraphs[0].Append("الإجراء الذي تم").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    }
                    else if (title == "قائمة الموقوفين الذين تم إنهاء وقد غادرو المملكة العربية السعودية")
                    {
                        t.Rows[0].Cells[0].Paragraphs[0].Append("المجهولين الذين قد غادرو المملكة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    }

                    else if (title == "قائمة الموقوفين الذين ما زالت معاملاتهم قيد التحري")
                    {
                        t.Rows[0].Cells[0].Paragraphs[0].Append("الملاحظة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    }
                    else if (title == "قائمة الموقوفين الذين تم التحري معهم مرات عدة ولم يقدموا مستندات تثبيت هويتهم السودانية")
                    {
                        t.Rows[0].Cells[0].Paragraphs[0].Append("الملاحظة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    }
                    else if (title == "قائمة المستندات المطلوبة من الموقفين لاثبات هويتهم السودانية")
                    {
                        t.Rows[0].Cells[0].Paragraphs[0].Append("المستندات المطلوبة").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    }

                    string tableTitle = Environment.NewLine + " جدول رقم)" + table.ToString() + ":(" + " " + title;
                    document.InsertParagraph(tableTitle)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(20d).UnderlineStyle(UnderlineStyle.singleLine)
                    .Alignment = Alignment.center;
                    var p = document.InsertParagraph(Environment.NewLine);
                    p.InsertTableAfterSelf(t);
                    table++;
                }
                document.Save();                
            }
            Process.Start("WINWORD.EXE", ActiveCopy);
            this.Close();
        }

        private void comboRespo_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboRespo.Text == "لم يقدم مستندات رسمية تثبت جنسيته السودانية")
            //{
            //    if (checkSex.Text != "ذكر")
            //    {
            //        comboRespo.Text = comboRespo.Text.Replace("يقدم", "تقدم");
            //        comboRespo.Text = comboRespo.Text.Replace("جنسيته", "جنسيتها");
            //    }

            //}
            
            actionSum.Multiline = false;
            actionSum.Enabled = true;
            if (responce.SelectedIndex == 0)
            {
                country.Enabled = false;
            }

            else if (responce.SelectedIndex == 1)
            {
                country.Enabled = true;
            }
            else if (responce.SelectedIndex == 2) {
                country.SelectedIndex = 0;
            }
           
        }

        private void ListSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != (char)13) return;
                actionDone = "";
            if (dataGridView1.Rows.Count > 1)
            {
                رقم_معاملة_القسم.Text = dataGridView1.Rows[0].Cells["رقم_معاملة_القسم"].Value.ToString().Split('/')[4];
                fillFileData(DataSource, dataGridView1.Rows[0].Cells["رقم_معاملة_القسم"].Value.ToString());
                ID = Convert.ToInt32(dataGridView1.Rows[0].Cells["ID"].Value.ToString());
                colIDs[0] = dataGridView1.Rows[0].Cells["رقم_معاملة_القسم"].Value.ToString();
                colIDs[1] = dataGridView1.Rows[0].Cells["ID"].Value.ToString();
                if (JobPosition.Contains("قنصل"))
                {
                    responce.Size = new System.Drawing.Size(352, 35);
                    comboProc.Visible = true;
                    الإجراء_الذي_تم.Enabled = btndelete.Visible = true;
                }
                else
                    comboProc.Visible = btndelete.Visible = false;

                اسم_المواطن_موضوع_الإجراء.Text = dataGridView1.Rows[0].Cells["اسم_المواطن_موضوع_الإجراء"].Value.ToString();
                if (اسم_المواطن_موضوع_الإجراء.Text == "")
                {
                    newData = true;
                    OpenFileDoc(dataGridView1.Rows[0].Cells["رقم_معاملة_القسم"].Value.ToString(), "TableMessages", "المستندات_الأولية");
                    labLnfo.Visible = dataGridView1.Visible = false;
                    panelMain.Visible = true;
                    return;
                }

                newData = false;


                رقم_معاملة_المصدر.Text = dataGridView1.Rows[0].Cells["رقم_معاملة_المصدر"].Value.ToString();
                المصدر.Text = dataGridView1.Rows[0].Cells["المصدر"].Value.ToString();
                تاريخ_الإصدار.Text = dataGridView1.Rows[0].Cells["تاريخ_الإصدار"].Value.ToString();
                تاريخ_الاستلام.Text = dataGridView1.Rows[0].Cells["تاريخ_الاستلام"].Value.ToString();
                الموضوع.Text = dataGridView1.Rows[0].Cells["الموضوع"].Value.ToString();
                النوع.Text = dataGridView1.Rows[0].Cells["النوع"].Value.ToString();
                رقم_معاملة_القسم.Enabled = false;

                if (النوع.Text == "ذكر")
                {
                    النوع.CheckState = CheckState.Checked;
                }
                else
                {
                    النوع.CheckState = CheckState.Unchecked;
                }
                الإجراء_الذي_تم.Text = dataGridView1.Rows[0].Cells["الإجراء_الذي_تم"].Value.ToString();
                الإجراء_الذي_تم.Visible = true;
                مسوؤل_الملف.Text = dataGridView1.Rows[0].Cells["مسوؤل_الملف"].Value.ToString();
                حالة_الأرشفة.Text = dataGridView1.Rows[0].Cells["حالة_الأرشفة"].Value.ToString();
                حالة_الأجراء = dataGridView1.Rows[0].Cells["حالة_الأجراء"].Value.ToString();
                رقم_الايقاف.Text = dataGridView1.Rows[0].Cells["رقم_الايقاف"].Value.ToString();
                المنطقة.Text = dataGridView1.Rows[0].Cells["المنطقة"].Value.ToString();
                responce.Text = dataGridView1.Rows[0].Cells["responce"].Value.ToString();
                reference = dataGridView1.Rows[0].Cells["reference"].Value.ToString();

                if (reference.Contains("_"))
                {
                    refFile = "الموقوف تم ايادعه برقم الملف: " + reference.Split('_')[0];
                    for (int x = 1; x < reference.Split('_').Length; x++)
                    {
                        refFile = refFile + Environment.NewLine + "الموقوف تمت إحالته إلى الملف بالرقم: " + reference.Split('_')[x];
                    }
                    if (refFile != "")
                        Panel_PaintLabel(refFile);

                }
                recommon = dataGridView1.Rows[0].Cells["recommendation"].Value.ToString();
                الحالة.Text = dataGridView1.Rows[0].Cells["الحالة"].Value.ToString();
                المهنة.Text = dataGridView1.Rows[0].Cells["المهنة"].Value.ToString();
                if (الحالة.Text == "مجهول")
                    الحالة.Checked = true;
                else
                    الحالة.Checked = false;
                country.Text = dataGridView1.Rows[0].Cells["country"].Value.ToString();
                actionSum.Text = dataGridView1.Rows[0].Cells["actionSum"].Value.ToString();
                labLnfo.Visible = dataGridView1.Visible = false;
                panelMain.Visible = true;
                رقم_الملف.Text = dataGridView1.Rows[0].Cells["رقم_الملف"].Value.ToString();


                string الملف_الارشفة = dataGridView1.Rows[0].Cells["الملف_الارشفة"].Value.ToString();
                foreach (Control control in panelArch.Controls)
                {
                    control.Visible = false;
                    control.Name = "لاغي";
                }
                if (الملف_الارشفة.Contains("*"))
                {
                    string[] fileList = الملف_الارشفة.Split('*');
                    for (int x = 0; x < fileList.Length; x++)
                    {
                        string btnText = "الموقوف مذكور بالملف بالرقم (" + fileList[x].Split('_')[0] + ") بالحالة رقم (" + fileList[x].Split('_')[2] + ") : " + fileList[x].Split('_')[1];
                        Panel_Paint(fileList[x].Split('_')[0], btnText);
                    }
                }
                else if (الملف_الارشفة.Contains("_"))
                {
                    string btnText = "الموقوف مذكور بالملف بالرقم (" + الملف_الارشفة.Split('_')[0] + ") بالحالة رقم (" + الملف_الارشفة.Split('_')[2] + ") : " + الملف_الارشفة.Split('_')[1];

                    Panel_Paint(الملف_الارشفة.Split('_')[0], btnText);
                }
                panelArch.Visible = true;
                panelArch.BringToFront();
                رقم_الملف.Select();
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboProc.SelectedIndex == 1)
            {
                finished();
                panelMain.Visible = false;
                labLnfo.Visible = dataGridView1.Visible = true;
                timer2.Enabled = true;
            }
            else if (comboProc.SelectedIndex == 0)
            {
                SqlConnection sqlCon = new SqlConnection(DataSource);
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlCommand sqlCmd = new SqlCommand("UPDATE TableMessages SET حالة_الأجراء=@حالة_الأجراء WHERE رقم_معاملة_القسم=@رقم_معاملة_القسم", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@حالة_الأجراء", "قيد الإجراء");
                sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", "ق س ج/80/22/14/" + رقم_معاملة_القسم.Text);
                sqlCmd.ExecuteNonQuery();

                sqlCon.Close();
                panelMain.Visible = false;
                labLnfo.Visible = dataGridView1.Visible = true;
                timer2.Enabled = true;
            }
        }

        

        private void المكاتبات_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            docIDNumber = "ق س ج/80/22/14/" + رقم_معاملة_القسم.Text;
            المكاتبات.Enabled = false;
            if (!OpenFileDoc(docIDNumber, "TableGeneralArch", المكاتبات.Text))
            {

                var selectedOption = MessageBox.Show("", "؟docx إضافة ملف بصيغة", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                {
                    
                    OpenFileDialog dlg = new OpenFileDialog();
                    //dlg.ShowDialog();
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string filePath = dlg.FileName;
                        var fileinfo1 = new FileInfo(filePath);
                        if (fileinfo1.Extension == ".docx")
                        {
                            //insertDoc(DataSource, filePath, docIDNumber, "ارشفة_المستندات");
                            //addDocx(DataSource, filePath, docIDNumber, "@ارشفة_المستندات");
                        }
                        else MessageBox.Show("يجب ان يكون الملف بصيغة docx فقط");
                    }

                }
            }
            المكاتبات.Enabled = true;

        }

        private void actionSum_TextChanged(object sender, EventArgs e)
        {
            if(actionSum.Text.Contains("اصدار وثيقة سفر إضطرارية"))
                actionSum.Size = new System.Drawing.Size(444, 38);
        }

        private void رقم_الملف_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                fileUpdate = false;
                btnArchBasic.PerformClick();
                
            }
        }

        private void actionSum_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                MessageBox.Show("done");
                fileUpdate = false;
                btnArchBasic.PerformClick();

            }
        }
    }
}

