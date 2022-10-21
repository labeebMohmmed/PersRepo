using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Threading;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Xceed.Words.NET;
using System.Diagnostics;
using Xceed.Document.NET;
using System.Globalization;
using System.Threading;

using System;
using System.Runtime.InteropServices;

namespace PersAhwal
{
    //https://www.youtube.com/watch?v=sHJVusS5Qz0
    //https://doc.xceed.com/xceed-document-libraries-for-net/Code_Snippets.html



    public partial class MainForm : Form
    {
        string DataSource, GregorianDate, HijriDate;
        string quorterS, quorterE;
        private SqlConnection sqlCon;
        static string EmployeeName, Jobposition;        
        int   totalrows = 0, totalRowDuration =0;
        int startofNextWeek;
        string FilespathIn, FilespathOut;
        int V = 0, A = 0;
        int i = 0;
        private static bool NewSettings = false, FirstDate = false, LastDate = false;
        private static int TableIndex = -1, IDNo = -1, IDVANo = -1;
        static string[] query = new string[10];

        static string[] AppNameV = new string[10];
        static string[] AppNameA = new string[10];
        static int[] IDA = new int [10];
        static int[] IDV = new int[10];
        static int[] DocA = new int[10];
        static int[] DocV = new int[10]; 
        
        static string[] queryVA = new string[10];
        static string[] DuratioListquery = new string[10];

        string applicanVA = "", ViewedVA = "", ArchivedStateVA = "";

        static string[] queryDateList = new string[10];
        static string[] querydatabase = new string[10];
        static string[] RetrievedNameList = new string[100];
        static string[] RetrievedTypeList = new string[100];
        static int [] Months= new int [12];
        static int[,] duartionReport= new int[5,10];
        static string[] DocType = new string[10];
        static string[] DocTypeVA = new string[10];
        static int[] IDNoList = new int[100];
        static string[] AppNames = new string[100];
        static string route;
        int[] weekSum = new int[5];
        int[] weekSum1 = new int[5];
        int[] weekSum2 = new int[5];
        int[] weekSum3 = new int[5];
        int[] weekSum4 = new int[5];
        int[] weekSum5 = new int[5];
        string Model, Output, ServerIP, Login, Pass, Database;
        int Hiday, Himonth;
        public MainForm(string Employee, string jobposition, string dataSource, string filepathIn, string filepathOut)
        {
            InitializeComponent();
            timer1.Enabled = true;
            timer2.Enabled = true;
            EmployeeName = Employee;
            DataSource = dataSource;
            FilespathIn = filepathIn;
            FilespathOut = filepathOut;
            Jobposition = jobposition;
            ConsulateEmployee.Text = EmployeeName;
            TablesList();
            
            ClearFileds();
            
            sqlCon = new SqlConnection(DataSource);
            loadSettings(DataSource, false, false, false, false) ;

        }

        private void TablesList()
        {
            query[0] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableDocIqrar where DocID=@DocID";
            query[1] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableTravIqrar where DocID=@DocID";
            query[2] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableMultiIqrar where DocID=@DocID";
            query[3] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableVisaApp where DocID=@DocID";
            query[4] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableFamilySponApp where DocID=@DocID";
            query[5] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableForensicApp where DocID=@DocID";
            query[6] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableTRName where DocID=@DocID";
            query[7] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableStudent where DocID=@DocID";
            query[8] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableMarriage where DocID=@DocID";
            query[9] = "select ID,AppName,GriDate,Viewed,DataMandoubName,ArchivedState,FileName1,FileName2  from TableFreeForm where DocID=@DocID";

            
            queryVA[0] = "select ID,AppName,Viewed,ArchivedState  from TableDocIqrar";
            queryVA[1] = "select ID,AppName,Viewed,ArchivedState  from TableTravIqrar";
            queryVA[2] = "select ID,AppName,Viewed,ArchivedState  from TableMultiIqrar";
            queryVA[3] = "select ID,AppName,Viewed,ArchivedState  from TableVisaApp";
            queryVA[4] = "select ID,AppName,Viewed,ArchivedState  from TableFamilySponApp";
            queryVA[5] = "select ID,AppName,Viewed,ArchivedState  from TableForensicApp";
            queryVA[6] = "select ID,AppName,Viewed,ArchivedState  from TableTRName";
            queryVA[7] = "select ID,AppName,Viewed,ArchivedState  from TableStudent";
            queryVA[8] = "select ID,AppName,Viewed,ArchivedState  from TableMarriage";
            queryVA[9] = "select ID,AppName,Viewed,ArchivedState  from TableFreeForm";
            

            DuratioListquery[0] = "select ID  from TableDocIqrar where GriDate=@GriDate";
            DuratioListquery[1] = "select ID  from TableTravIqrar where GriDate=@GriDate";
            DuratioListquery[2] = "select ID  from TableMultiIqrar where GriDate=@GriDate";
            DuratioListquery[3] = "select ID  from TableVisaApp where GriDate=@GriDate";
            DuratioListquery[4] = "select ID  from TableFamilySponApp where GriDate=@GriDate";
            DuratioListquery[5] = "select ID  from TableForensicApp where GriDate=@GriDate";
            DuratioListquery[6] = "select ID  from TableTRName where GriDate=@GriDate";
            DuratioListquery[7] = "select ID  from TableStudent where GriDate=@GriDate";
            DuratioListquery[8] = "select ID  from TableMarriage where GriDate=@GriDate";
            DuratioListquery[9] = "select ID  from TableFreeForm where GriDate=@GriDate";

            queryDateList[0] = "select AppName from TableDocIqrar where GriDate=@GriDate";
            queryDateList[1] = "select AppName from TableTravIqrar where GriDate=@GriDate";
            queryDateList[2] = "select AppName,IqrarPurpose from TableMultiIqrar where GriDate=@GriDate";
            queryDateList[3] = "select AppName from TableTRName where GriDate=@GriDate";
            queryDateList[4] = "select AppName,SpecType from TableFreeForm where GriDate=@GriDate";

            querydatabase[0] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableDocIqrar where ID=@id";
            querydatabase[1] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableTravIqrar where ID=@id";
            querydatabase[2] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableMultiIqrar where ID=@id";
            querydatabase[3] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableVisaApp where ID=@id";
            querydatabase[4] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableFamilySponApp where ID=@id";
            querydatabase[5] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableForensicApp where ID=@id";
            querydatabase[6] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableTRName where ID=@id";
            querydatabase[7] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableStudent where ID=@id";
            querydatabase[8] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableMarriage where ID=@id";
            querydatabase[9] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableFreeForm where ID=@id";

            DocTypeVA[0] = "إقرار باستخراج أوراق ثبوتية";
            DocTypeVA[1] = "إقرار سفر اسرة";
            DocTypeVA[2] = "إقرار متعدد";
            DocTypeVA[3] = "تأشيرة سفر";
            DocTypeVA[4] = "إقرار كفالة عائلة";
            DocTypeVA[5] = "افادة لمن يهمه الامر";
            DocTypeVA[6] = "إقرار بمطابقة اسمين";
            DocTypeVA[7] = "شهادة لمن يهمه الامر";
            DocTypeVA[8] = "شهادة لمن يهمه الامر";
            DocTypeVA[9] = "إقرار عام";


            DocType[0] = "إقرار باستخراج أوراق ثبوتية";
            DocType[1] = "إقرار بالموافقة على سفر الابناء";
            
            DocType[3] = "إقرار بمطابقة إسمين";

            


        }

        private int daysOfMonth(int month, int year)
        {
            Months[0] = 31;

            if (year % 400 == 0 || (year % 4 == 0 && year % 100 != 0))
                Months[1] = 29;
            else Months[1] = 28;

            Months[2] = 31;

            Months[3] = 30;
            Months[4] = 31;
            Months[5] = 30;

            Months[6] = 31;
            Months[7] = 31;
            Months[8] = 30;

            Months[9] = 31;
            Months[10] = 30;
            Months[11] = 31;

            return Months[month];
        }

        private void ClearFileds()
        {
            
            TableIndex = -1;
            IDNo = -1;
            ReportType.SelectedIndex = 0;           
            SearchPanel.Visible = false;
            ReportPanel.Visible = false;            
            AttendViceConsul.SelectedIndex = 2;
            if (Jobposition.Contains("قنصل")) {
                picadd.Visible = true;
                labelmonth.Visible = true;
                picremove.Visible = true;
                labeldate.Visible = true;
                picaddmonth.Visible = true;
                pictremovemonth.Visible = true;
            }
        }

        
        int DailyList(string dateFrom)
        {
            
            totalrows = 0;
               SqlConnection sqlCon = new SqlConnection(DataSource);
            DataTable dtbl = new DataTable();
            if (sqlCon.State == ConnectionState.Closed)
                {
                sqlCon.Open();
                for (TableIndex = 0; TableIndex < 5; TableIndex++)
                    {
                                          
                        SqlDataAdapter sqlDa = new SqlDataAdapter(queryDateList[TableIndex], sqlCon);
                        sqlDa.SelectCommand.CommandType = CommandType.Text;
                        sqlDa.SelectCommand.Parameters.AddWithValue("@GriDate", dateFrom);
                        sqlDa.Fill(dtbl);
                        dataGridView2.DataSource = dtbl;
                        for (; totalrows < dtbl.Rows.Count; totalrows++) {
                        if(TableIndex == 2 || TableIndex == 4) RetrievedTypeList[totalrows] = dataGridView2.Rows[totalrows].Cells[1].Value.ToString();
                        else RetrievedTypeList[totalrows] = DocType[TableIndex];
                            RetrievedNameList[totalrows] = dataGridView2.Rows[totalrows].Cells[0].Value.ToString();
                        }
                    }                    
                }
            sqlCon.Close();
            return totalrows;
        }

        int DailyList(string dateFrom, string dateTo)
        {
            totalRowDuration = 0;
            int w = 0;
            string CurrentDate="", Currentmonth = "", CurrentDay="";
            DateTime datetimeS = dateTimeFrom.Value;
            DateTime datetimeE = dateTimeTo.Value;
            if (datetimeS > datetimeE) {
                string datetime = dateFrom;
                dateFrom = dateTo;
                dateTo = datetime;

            }
            
            string[] YearMonthDayS = dateFrom.Split('-');
            int yearS, monthS, dateS;
            yearS = Convert.ToInt16(YearMonthDayS[0]);
            monthS = Convert.ToInt16(YearMonthDayS[1]);
            dateS = Convert.ToInt16(YearMonthDayS[2]);
            DateTime dateValue = new DateTime(yearS, monthS,dateS);


            int dayeofWeek = ((int)dateValue.DayOfWeek);
            
            if (dayeofWeek == 0) { startofNextWeek = dateS + 7; }
            else if (dayeofWeek == 1) { startofNextWeek = dateS + 6; }
            else if (dayeofWeek == 2) { startofNextWeek = dateS + 5; }
            else if (dayeofWeek == 3) { startofNextWeek = dateS + 4; }
            else if (dayeofWeek == 4) { startofNextWeek = dateS + 3; }
            else if (dayeofWeek == 5) { startofNextWeek = dateS + 2; }
            else if (dayeofWeek == 6) { startofNextWeek = dateS + 1; }


            string[] YearMonthDayE = dateTo.Split('-');
            int yearE, monthE, dateE;
            yearE = Convert.ToInt16(YearMonthDayE[0]);
            monthE = Convert.ToInt16(YearMonthDayE[1]);
            dateE = Convert.ToInt16(YearMonthDayE[2]);

            SqlConnection sqlCon = new SqlConnection(DataSource);
            DataTable dtbl = new DataTable();
            if (sqlCon.State == ConnectionState.Closed)
            {
                sqlCon.Open();
                
                for (int y = yearS; y <= yearE; y++) {
                    for (int m = monthS; m <= monthE && m <= 12; m++) {
                        int d;
                        for ( d = dateS; d <= dateE && d <= daysOfMonth(m, y); d++)
                        {
                            if (m < 10) Currentmonth = "0" + m.ToString();
                            else Currentmonth = m.ToString();
                            if (d < 10) CurrentDay = "0" + d.ToString();
                            else CurrentDay = d.ToString();
                            CurrentDate = CurrentDay + "-" + Currentmonth + "-" + y.ToString();

                            for (TableIndex = 0; TableIndex < 10; TableIndex++)
                            {
                                SqlDataAdapter sqlDa = new SqlDataAdapter(DuratioListquery[TableIndex], sqlCon);
                                sqlDa.SelectCommand.CommandType = CommandType.Text;
                                sqlDa.SelectCommand.Parameters.AddWithValue("@GriDate", CurrentDate);
                                sqlDa.Fill(dtbl);
                                dataGridView2.DataSource = dtbl;                                
                                duartionReport[w, TableIndex] = dtbl.Rows.Count;
                                totalRowDuration = 1;
                            }
                        }
                        if (d == startofNextWeek)
                        {
                            w = 0;
                            startofNextWeek += 7;
                        }
                    }
                }               
            }
            sqlCon.Close();
            
            return totalRowDuration;
        }

        private void CreateDailyReport(int rows, string date)
        {
            route = FilespathIn + "DailyReport.docx";
            string ActiveCopy = FilespathOut + "DailyReport"+ date+".docx";
            if (!File.Exists(ActiveCopy))
            {
                System.IO.File.Copy(route, ActiveCopy);
                using (var document = DocX.Load(ActiveCopy))
                {
                    System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                    InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);

                    string strHeader = "الرقم: " + ReportNo.Text + "     " + "التاريخ:" + GregorianDate + " م" + "     " + "الموافق: " + HijriDate + "هـ" + Environment.NewLine;
                    document.InsertParagraph(strHeader)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(18d)
                    .Alignment = Alignment.center;
                    string MessageDir = "من: سوداني -  جـــدة" + Environment.NewLine + "الى: سوداني - الخرطوم" + Environment.NewLine + "لعناية السيد/ مدير إدارة التوثيق"
                        + Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ"
                        + Environment.NewLine + "بالإشارة إلى برقيتكم بالرقم: و خ/توثيق/97 بتاريخ 23/04/2014 م بشأن إصدار راجعة الإقرارات، نفيدكم باعتماد القنصلية العامة للمعاملات الصادرة طرفها للمذكورين بالجدول أدناه"
                        + " بتاريخ " + dateTimeFrom.Text;
                    document.InsertParagraph(MessageDir)
                        .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                        .FontSize(20d)
                        .Direction = Direction.RightToLeft;

                    var t = document.AddTable(rows + 1, 3);
                    t.Design = TableDesign.TableGrid;
                    t.Alignment = Alignment.center;
                    t.SetColumnWidth(2, 40);
                    t.SetColumnWidth(1, 180);
                    t.SetColumnWidth(0, 180);


                    t.Rows[0].Cells[0].Paragraphs[0].Append("نوع المعاملة").FontSize(15d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[1].Paragraphs[0].Append("اسم مقدم الطلب").FontSize(15d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[2].Paragraphs[0].Append("الرقم").FontSize(15d).Bold().Alignment = Alignment.center;

                    for (int x = 1; x <= rows; x++)
                    {
                        t.Rows[x].Cells[0].Paragraphs[0].Append(RetrievedTypeList[x - 1]).FontSize(15d).Direction = Direction.RightToLeft;
                        t.Rows[x].Cells[1].Paragraphs[0].Append(RetrievedNameList[x - 1]).FontSize(15d).Direction = Direction.RightToLeft;
                        t.Rows[x].Cells[2].Paragraphs[0].Append((x).ToString()).FontSize(15d).Direction = Direction.RightToLeft;
                    }

                    

                    var p = document.InsertParagraph(Environment.NewLine);
                    p.InsertTableAfterSelf(t);
                    
                    string strAttvCo = Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ"+ Environment.NewLine + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + AttendViceConsul.Text + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + "ع/ القنصل العام بالإنابة";
                    var AttvCo = document.InsertParagraph(strAttvCo)
                        .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                        .FontSize(20d)
                        .Bold()
                        .Alignment = Alignment.center;
                   

                    document.Save();
                    Process.Start("WINWORD.EXE", ActiveCopy);
                    
                }
            }
            else
            {
                MessageBox.Show("يرجى حذف الملف الموجودأولاً");
                PrintReport.Enabled = true;
                
            }

        }



        private void CreateDurationReport()
        {
            

            using (DocX document = DocX.Load(FilespathIn+ "DailyDurationReport.docx"))
            {
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
                string strHeader = "الرقم: " + ReportNo.Text + "     " + "التاريخ:" + GregorianDate + " م" + "     " + "الموافق: " + HijriDate + "هـ";
                document.InsertParagraph(strHeader)
                .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                .FontSize(18d)
                .Alignment = Alignment.center;


                var t = document.AddTable(7, 6);
                t.Design = TableDesign.TableGrid;
                t.Alignment = Alignment.center;
                
                t.SetColumnWidth(5, 100);
                t.SetColumnWidth(4, 70);
                t.SetColumnWidth(3, 70);
                t.SetColumnWidth(2, 100);
                t.SetColumnWidth(1, 70);
                t.SetColumnWidth(0, 90);                

                querydatabase[0] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableDocIqrar where ID=@id";
                querydatabase[1] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableTravIqrar where ID=@id";
                querydatabase[2] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableMultiIqrar where ID=@id";
                querydatabase[3] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableVisaApp where ID=@id";
                querydatabase[4] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableFamilySponApp where ID=@id";
                querydatabase[5] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableForensicApp where ID=@id";
                querydatabase[6] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableTRName where ID=@id";
                querydatabase[7] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableStudent where ID=@id";
                querydatabase[8] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableMarriage where ID=@id";
                querydatabase[9] = "select Data1, Extension1,FileName1,Data2,Extension2,FileName2  from TableFreeForm where ID=@id";

                t.Rows[0].Cells[0].Paragraphs[0].Append("مجموع المعاملات").FontSize(15d).Bold().Alignment = Alignment.center;
                t.Rows[0].Cells[1].Paragraphs[0].Append("شهادات").FontSize(15d).Bold().Alignment = Alignment.center; //7 8
                t.Rows[0].Cells[2].Paragraphs[0].Append("افادات").FontSize(15d).Bold().Alignment = Alignment.center;//5
                t.Rows[0].Cells[3].Paragraphs[0].Append("مخاطبة لتأشيرة").FontSize(15d).Bold().Alignment = Alignment.center; //3
                t.Rows[0].Cells[4].Paragraphs[0].Append("إقرارات").FontSize(15d).Bold().Alignment = Alignment.center;//0 1 2 4 6 9
                t.Rows[0].Cells[5].Paragraphs[0].Append("الاسبوع").FontSize(15d).Bold().Alignment = Alignment.center;
                
                for (int w = 0; w < 5; w++)
                {
                    weekSum[w] = duartionReport[w, 0] + duartionReport[w, 1] + duartionReport[w, 2] + duartionReport[w, 3] + duartionReport[w, 4] + duartionReport[w, 5] + duartionReport[w, 6] + duartionReport[w, 7] + duartionReport[w, 8] + duartionReport[w, 9];
                    weekSum1[w] = duartionReport[w, 0] + duartionReport[w, 1] + duartionReport[w, 2] + duartionReport[w, 4] + duartionReport[w, 6] + duartionReport[w, 9];
                    weekSum2[w] = duartionReport[w, 3];
                    weekSum3[w] = duartionReport[w, 5];
                    weekSum4[w] = duartionReport[w, 0] + duartionReport[w, 1];
                    
                        t.Rows[w + 1].Cells[0].Paragraphs[0].Append(weekSum[w].ToString()).FontSize(15d).Bold().Alignment = Alignment.center;
                        t.Rows[w + 1].Cells[1].Paragraphs[0].Append((weekSum4[w]).ToString()).FontSize(15d).Bold().Alignment = Alignment.center; //7 8
                        t.Rows[w + 1].Cells[2].Paragraphs[0].Append(weekSum3[w].ToString()).FontSize(15d).Bold().Alignment = Alignment.center;//5
                        t.Rows[w + 1].Cells[3].Paragraphs[0].Append(weekSum2[w].ToString()).FontSize(15d).Bold().Alignment = Alignment.center; //3
                        t.Rows[w + 1].Cells[4].Paragraphs[0].Append((weekSum1[w]).ToString()).FontSize(15d).Bold().Alignment = Alignment.center;//0 1 2 4 6 9
                        t.Rows[w + 1].Cells[5].Paragraphs[0].Append("الاسبوع " + weekorder(w)).FontSize(15d).Bold().Alignment = Alignment.center;
                }

                int FullSum = 0;
                int FullSum1 = 0;
                int FullSum2 = 0;
                int FullSum3 = 0;
                int FullSum4 = 0;
                for (int w = 0; w < 5; w++) { 
                    FullSum += weekSum[w];
                    FullSum1 += weekSum1[w];
                    FullSum2 += weekSum2[w];
                    FullSum3 += weekSum3[w];
                    FullSum4 += weekSum4[w];
                }

                t.Rows[6].Cells[0].Paragraphs[0].Append(FullSum.ToString()).FontSize(15d).Bold().Alignment = Alignment.center;
                t.Rows[6].Cells[1].Paragraphs[0].Append(FullSum4.ToString()).FontSize(15d).Bold().Alignment = Alignment.center; //7 8
                t.Rows[6].Cells[2].Paragraphs[0].Append(FullSum3.ToString()).FontSize(15d).Bold().Alignment = Alignment.center;//5
                t.Rows[6].Cells[3].Paragraphs[0].Append(FullSum2.ToString()).FontSize(15d).Bold().Alignment = Alignment.center; //3
                t.Rows[6].Cells[4].Paragraphs[0].Append(FullSum1.ToString()).FontSize(15d).Bold().Alignment = Alignment.center;//0 1 2 4 6 9
                t.Rows[6].Cells[5].Paragraphs[0].Append(" المجموع" ).FontSize(15d).Bold().Alignment = Alignment.center;
                
                var p = document.InsertParagraph(Environment.NewLine);                
                p.InsertTableAfterSelf(t);
                string strAttvCo = Environment.NewLine + Environment.NewLine + AttendViceConsul.Text + Environment.NewLine + "ع/ القنصل العام بالإنابة";
                var AttvCo = document.InsertParagraph(strAttvCo)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(20d)
                    .Bold()
                    .Alignment = Alignment.left;

                document.Save();
                Process.Start("WINWORD.EXE", FilespathIn + "DailyDurationReport.docx");
            }

        }

        private string weekorder(int week)
        {
            switch (week) {
                case 0:
                    return "الأول";
                    
                case 1:
                    return "الثاني";
                    

                case 2:
                    return "الثالث";
                    

                case 3:
                    return "الرابع";
                    

                case 4:
                    return "الخامس";

                default:
                    return "";
                    
            }
        }

        void FillDataGridView()
        {
            
            SqlConnection sqlCon = new SqlConnection(DataSource);

            if (sqlCon.State == ConnectionState.Closed)
                
            if (txtSearch.Text != "")
            {
                for (TableIndex = 0; TableIndex < 10; TableIndex++)
                {
                    sqlCon.Open();
                    SqlCommand sqlCmd1 = new SqlCommand(query[TableIndex], sqlCon);
                    sqlCmd1.Parameters.Add("@DocID", SqlDbType.NVarChar).Value = txtSearch.Text;
                    var reader = sqlCmd1.ExecuteReader();
                        
                        if (reader.Read())
                    {
                            
                            IDNo = Convert.ToInt32(reader["ID"].ToString());                            
                            applicant.Text = reader["AppName"].ToString();
                            date.Text = reader["GriDate"].ToString();
                            string viewSt = reader["Viewed"].ToString();
                            string filename1 = reader["FileName1"].ToString();
                            string filename2 = reader["FileName2"].ToString();
                            if (filename1 == "text1.txt") Arch1.Visible = false;
                            if (filename2 == "text2.txt") Arch2.Visible = false;

                            if (viewSt == "غير معالج")
                            {
                                ProcessedSt.CheckState = CheckState.Unchecked;
                                ProcessedSt.Text = "غير معالج";
                                ProcessedSt.BackColor = Color.Red;
                            }
                            else
                            {
                                ProcessedSt.CheckState = CheckState.Checked;
                                ProcessedSt.Text = viewSt;
                                ProcessedSt.BackColor = Color.Green;
                            }

                            string mandoub = reader["DataMandoubName"].ToString();
                            if (mandoub != "")
                                Apptype.Text = "بواسطة مندوب القنصلية " + mandoub;
                            else Apptype.Text = "حضور مباشرة إلى القنصلية";


                            if (reader["ArchivedState"].ToString() != "غير مؤرشف")
                            {
                                ArchiveSt.CheckState = CheckState.Checked;
                                ArchiveSt.Text = "مؤرشف";
                                ArchiveSt.BackColor = Color.Green;
                            }
                            else
                            {
                                ArchiveSt.CheckState = CheckState.Unchecked;
                                ArchiveSt.Text = "غير مؤرشف";
                                ArchiveSt.BackColor = Color.Red;
                            }
                        
                        SearchPanel.Height = 296;
                        break;
                    }
                    else SearchPanel.Height = 40;
                        sqlCon.Close();
                }
            }
            
        }


        private void txtSearch_TextChanged(object sender, EventArgs e)
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


        private void btnSearch_Click(object sender, EventArgs e)
        {
            
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            FormDataBase formDataBase = new FormDataBase(DataSource,FilespathIn, FilespathOut);
            formDataBase.Show();
        }

        private void txtSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13) DetecedForm.PerformClick();
                
        }


        

        private void txtSearch_TextChanged_1(object sender, EventArgs e)
        {
            FillDataGridView();
        }

        private void txtSearch_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13) DetecedForm.PerformClick();

        }

        

        private void Arch1_Click(object sender, EventArgs e)
        {
            OpenFile(IDNo, 1);
        }

        private void OpenFile(int id, int fileNo)
        {
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand(querydatabase[TableIndex], Con);
            sqlCmd1.Parameters.Add("@ID", SqlDbType.Int).Value = id;
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

        private void Arch2_Click(object sender, EventArgs e)
        {
            OpenFile(IDNo, 2);
        }


        private int HijriDateDifferment(string source)
        {
            int differment=0;
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select hijriModification from TableSettings";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                     differment = Convert.ToInt32(reader["hijriModification"].ToString());

                    labeldate.Text = differment.ToString();
                }
                
                saConn.Close();
            }
            return differment;
        }


        private void ReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (ReportType.SelectedIndex)
            {
                case 0:
                    ReportNo.Enabled = true;
                    AttendViceConsul.Visible = true;
                    btnattendedVC.Enabled = true;
                    btnReportNo.Enabled = true;
                    ReportPanel.Height = 205;
                    break;
                case 1:
                    yearReport.Visible = false;
                    button24.Visible = false;
                    button24.Enabled = false;
                    button28.Visible = false;
                    dateTimeFrom.Visible = false;
                    dateTimeTo.Visible = false;
                    if (DailyList(GregorianDate) > 0)
                    {
                        
                        PrintReport.Enabled = true;
                        PrintReport.Visible = true;
                        ReportPanel.Height = 205;
                    }
                    else
                    {
                        PrintReport.Enabled = false;
                        PrintReport.Visible = false;
                        ReportNo.Enabled = true;
                        AttendViceConsul.Visible = true;
                        btnattendedVC.Enabled = true;
                        btnReportNo.Enabled = true;
                        ReportPanel.Height = 40;
                        MessageBox.Show("لا يوجد قائمة بالتاريخ المحدد");
                    }
                    break;
                case 2:
                    yearReport.Visible = false;
                    button24.Text = "يوم:";
                    button24.Enabled = true;
                    button24.Visible = true;
                    button28.Visible = false;
                    dateTimeFrom.Visible = true;
                    dateTimeTo.Visible = false;
                    dateTimeFrom.Width = 288;
                    btnattendedVC.Enabled = true;
                    btnReportNo.Enabled = true;
                    ReportPanel.Height = 205;
                    break;
                case 3:
                    yearReport.Visible = false;
                    button24.Text = "من:";
                    dateTimeFrom.Width = 113;
                    button24.Visible = true;
                    button28.Visible = true;
                    dateTimeFrom.Visible = true;
                    dateTimeTo.Visible = true;
                    ReportPanel.Height = 205;
                    break;
                case 4:
                    button24.Text = "السنة:";
                    quorterS = "-01-01";
                    quorterE = "-03-31";
                    button24.Enabled = true;
                    button24.Visible = true; 
                    yearReport.Visible = true;
                    ReportPanel.Height = 205;
                    button28.Visible = false;
                    dateTimeFrom.Visible = false;
                    dateTimeTo.Visible = false;
                    break;
                case 5:
                    button24.Text = "السنة:";
                    quorterS = "-04-01";
                    quorterE = "-06-30";
                    button24.Enabled = true;
                    yearReport.Visible = true;
                    button24.Visible = true;
                    button28.Visible = false;
                    dateTimeFrom.Visible = false;
                    dateTimeTo.Visible = false;
                    ReportPanel.Height = 205;
                    break;
                case 6:
                    button24.Text = "السنة:";
                    quorterS = "-07-01";
                    quorterE = "-09-30";
                    button24.Enabled = true;
                    yearReport.Visible = true;
                    button24.Visible = true;
                    button28.Visible = false;
                    dateTimeFrom.Visible = false;
                    dateTimeTo.Visible = false;
                    ReportPanel.Height = 205;
                    break;
                case 7:
                    button24.Text = "السنة:";
                    quorterS = "-10-01";
                    quorterE = "-12-31";
                    button24.Enabled = true;
                    yearReport.Visible = true;
                    button24.Visible = true;
                    button28.Visible = false;
                    dateTimeFrom.Visible = false;
                    dateTimeTo.Visible = false;
                    ReportPanel.Height = 205;
                    break;

            }
        }


        private void PrintReport_Click_1(object sender, EventArgs e)
        {
            PrintReport.Enabled = false;
            PrintReport.Text = "تجري عملية الطباعة";
            if (totalrows > 0) {
                CreateDailyReport(totalrows,dateTimeFrom.Text);
                totalrows = 0;
            }
            if (totalRowDuration == 1) { CreateDurationReport(); totalRowDuration = 0; }
            PrintReport.Text = "طباعة التقرير";
            PrintReport.Enabled = false;
            PrintReport.Visible = false;
            ReportPanel.Height = 36;
        }

        private void IqrarBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            

            if (IqrarBox.SelectedIndex >= 0 && IqrarBox.SelectedIndex <= 7) {
                Form3 form3 = new Form3(IDNo, IqrarBox.SelectedIndex, EmployeeName, DataSource, FilespathIn, FilespathOut);
                form3.ShowDialog();
            }

            else if (IqrarBox.SelectedIndex == 8)
            {
                Form5 form5 = new Form5(IDNo, EmployeeName, DataSource, FilespathIn, FilespathOut);
                form5.ShowDialog();
            }
            else if (IqrarBox.SelectedIndex == 9)
            {
                Form1 form1 = new Form1(IDNo, EmployeeName, DataSource, FilespathIn, FilespathOut);
                form1.ShowDialog();
            }

            else if (IqrarBox.SelectedIndex == 10)
            {
                Form2 form2 = new Form2(IDNo, EmployeeName, DataSource, FilespathIn, FilespathOut);
                form2.ShowDialog();
            }            

            else if (IqrarBox.SelectedIndex == 11)
            {
                Form7 form7 = new Form7(IDNo, EmployeeName, DataSource, FilespathIn, FilespathOut);
                form7.ShowDialog();
            }

            else if (IqrarBox.SelectedIndex == 12)
            {
                Form10 form10 = new Form10(IDNo, 0, EmployeeName, DataSource, FilespathIn, FilespathOut);
                form10.ShowDialog();
            }

        }

        private void IfadaBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (IfadaBox.SelectedIndex == 0)
            {
                Form6 form6 = new Form6(-1, EmployeeName, DataSource, FilespathIn, FilespathOut);
                form6.ShowDialog();
            }
            else if (IfadaBox.SelectedIndex == 1)
            {
                Form8 form8 = new Form8(-1, EmployeeName, DataSource, FilespathIn, FilespathOut);
                form8.ShowDialog();
            }
            else if (IfadaBox.SelectedIndex == 2)
            {
                Form10 form10 = new Form10(-1, 2, EmployeeName, DataSource, FilespathIn, FilespathOut);
                form10.ShowDialog();
            }
        }

        private void ShehadaBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ShehadaBox.SelectedIndex == 0)
            {
                Form9 form9 = new Form9(-1, EmployeeName, DataSource, FilespathIn, FilespathOut);
                form9.ShowDialog();
            }
            else if (IfadaBox.SelectedIndex == 1)
            {
                Form10 form10 = new Form10(-1, 3, EmployeeName, DataSource, FilespathIn, FilespathOut);
                form10.ShowDialog();
            }
            else if (ShehadaBox.SelectedIndex == 2)
            {
                MessageBox.Show("غير مدرجة حتى الآن");
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (VisaBox.SelectedIndex == 0)
            {
                Form4 form4 = new Form4(-1, EmployeeName, DataSource, FilespathIn, FilespathOut);
                form4.ShowDialog();
            }
            if (VisaBox.SelectedIndex == 1)
            {
                MessageBox.Show("غير مدرجة حتى الآن");
            }

        }

        private void btnSearch_Click_1(object sender, EventArgs e)
        {
            FillDataGridView();
        }

       

        private void button2_Click(object sender, EventArgs e)
        {
            if (SearchPanel.Visible == false)
            {
                SearchPanel.Visible = true;
                
                ReportPanel.Visible = false;
            }
            else SearchPanel.Visible = false;
        }

      

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {
           
        }

        private void button4_Click_1(object sender, EventArgs e)
        {

        }

       

        private void button3_Click(object sender, EventArgs e)
        {
            SearchPanel.Visible = false;
            
            if (ReportPanel.Visible == false)
            {
                ReportPanel.Visible = true;
                
            }
            else ReportPanel.Visible = false;
        }
        
        private void ViewArchShow(int Buttons, int Doc, int ID, string AppName)
        {
            labelVA.Text = labelVA.Text + Buttons.ToString() + " - " + AppName + Environment.NewLine + DocTypeVA[Doc] + Environment.NewLine;
        }
        
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            labelVA.Text = "";
            for (int x = 1; x < A; x++)
            {
                ViewArchShow(x, DocA[x], IDA[x], AppNameA[x]);
            }
            if (A <= 1) MessageBox.Show("لا توجد معاملات غير مؤرشفة");
        }

        private void pictureBox2_Click_1(object sender, EventArgs e)
        {
            labelVA.Text = "";
            for (int x = 1; x < V; x++)
            {
                ViewArchShow(x, DocV[x], IDV[x], AppNameV[x]);
            }
            if (V <= 1) MessageBox.Show("لا توجد معاملات غير معالجة");
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void picremove_Click(object sender, EventArgs e)
        {
            
            loadSettings(DataSource,false,true,false,false);
            
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            
            loadSettings(DataSource, false,false,true,true);
            
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            
            loadSettings(DataSource, false,false,false,true);
            
        }

        private void picadd_Click(object sender, EventArgs e)
        {
            
            loadSettings(DataSource,true,true,false,false);           
            
        }
        private void DetecedForm_Click(object sender, EventArgs e)
        {
            GoToForm(TableIndex-1, IDNo);
            ClearFileds();
        }

        

        private void loadSettings(string dataSource, bool day, bool daychange, bool month, bool monthchange)
        {
            SqlConnection Con = new SqlConnection(dataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select Modelfilespath,TempOutput,ServerName,Serverlogin,ServerPass,serverDatabase,hijriday,hijrimonth  from TableSettings where ID=@id", Con);
            sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
            if (Con.State == ConnectionState.Closed)
                try
                {
                    Con.Open();

                    var reader = sqlCmd1.ExecuteReader();

                    if (reader.Read())
                    {                        
                        Model = reader["Modelfilespath"].ToString();
                        Output = reader["TempOutput"].ToString();
                        ServerIP = reader["ServerName"].ToString();
                        Login = reader["Serverlogin"].ToString();
                        Pass = reader["ServerPass"].ToString();
                        Database = reader["serverDatabase"].ToString();
                        Hiday = Convert.ToInt32(reader["hijriday"].ToString());
                        Himonth = Convert.ToInt32(reader["hijrimonth"].ToString());

                        if (daychange) 
                        {
                            if (day) Hiday++; 
                            else Hiday--; 
                        }
                        if (monthchange)
                        {
                            if (month) Himonth++; 
                            else Himonth--;
                        }
                    }
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    Con.Close();
                    Con.Open();
                    SqlCommand sqlCmd = new SqlCommand("SettingsAddorEdit", Con);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@ID", 1);
                    sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                    sqlCmd.Parameters.AddWithValue("@Modelfilespath", Model);
                    sqlCmd.Parameters.AddWithValue("@TempOutput", Output);
                    sqlCmd.Parameters.AddWithValue("@ServerName", ServerIP);
                    sqlCmd.Parameters.AddWithValue("@Serverlogin", Login);
                    sqlCmd.Parameters.AddWithValue("@ServerPass", Pass);
                    sqlCmd.Parameters.AddWithValue("@serverDatabase", Database);
                    sqlCmd.Parameters.AddWithValue("@hijriday", Hiday);
                    sqlCmd.Parameters.AddWithValue("@hijrimonth", Himonth);
                    labeldate.Text = "فرق الشهر الهجري " + Hiday.ToString();
                    labelmonth.Text = "فرق الشهر الهجري " + Himonth.ToString();
                    sqlCmd.ExecuteNonQuery();
                    Con.Close();
                }
        }

        private void Aprove_Click_1(object sender, EventArgs e)
        {
            SignUp signUp = new SignUp(EmployeeName, Jobposition, DataSource);
            signUp.Show();
        }

        private void flowLayoutPanel1_MouseMove(object sender, MouseEventArgs e)
        {
            i = 0;
        }

        

        private void pictureBox1_Click_2(object sender, EventArgs e)
        {
            
        }


        private void buttonClick(object sender, EventArgs e)
        {
            Button button = sender as Button;
        }
        
        private void timer3_Tick(object sender, EventArgs e)
        {                        
            i++;
            V = A = 1;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            
            
            if (sqlCon.State == ConnectionState.Closed)
            {
                sqlCon.Open();
                for (TableIndex = 0; TableIndex < 1; TableIndex++)
                {

                    SqlDataAdapter sqlDa = new SqlDataAdapter(queryVA[TableIndex], sqlCon);
                    sqlDa.SelectCommand.CommandType = CommandType.Text;
                    DataTable dtbl = new DataTable();
                    sqlDa.Fill(dtbl);
                    dataGridView4.DataSource = dtbl;
                    //ID,AppName,Viewed,ArchivedState  from TableDocIqrar";

                    for (int x =0; x < dtbl.Rows.Count; x++)
                    {
                        if (dataGridView4.Rows[x].Cells[2].Value.ToString() == "غير معالج") {
                            DocV[V] = TableIndex;
                            IDV[V] = Convert.ToInt32(dataGridView4.Rows[x].Cells[0].Value.ToString());
                            AppNameV[V] = dataGridView4.Rows[x].Cells[1].Value.ToString();
                            V++;
                        }
                        if (dataGridView4.Rows[x].Cells[2].Value.ToString() != "غير معالج" && dataGridView4.Rows[x].Cells[3].Value.ToString() == "غير مؤرشف")
                        {
                            DocA[A] = TableIndex;
                            IDA[A] = Convert.ToInt32(dataGridView4.Rows[x].Cells[0].Value.ToString());
                            AppNameA[A] = dataGridView4.Rows[x].Cells[1].Value.ToString();
                            A++;
                        }

                    }
                }
            }
            if(A>1) labelarch.BackColor = Color.Red;else labelarch.BackColor = Color.Green;
            if (V>1) labelPro.BackColor = Color.Red;else labelPro.BackColor = Color.Green;
            sqlCon.Close();

        }

        
        private void button4_Click(object sender, EventArgs e)
        {
           
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string from =  yearReport.Text.Trim()  + quorterS;
            string to =  yearReport.Text.Trim() + quorterE;
            
            int rows = DailyList(from, to);
            if (rows > 0)
            {                

                PrintReport.Enabled = true;
                PrintReport.Visible = true;
                ReportPanel.Height = 205;
                
            }
            else
            {
                PrintReport.Enabled = false;
                PrintReport.Visible = false;
                ReportPanel.Height = 38;
                MessageBox.Show("لا يوجد قائمة بالتاريخ المحدد");
            }
        }

        private void dateTimeTo_ValueChanged(object sender, EventArgs e)
        {
            
            LastDate = true;
            if (FirstDate)
            {
                int rows = DailyList(dateTimeFrom.Text, dateTimeTo.Text);
                if (rows > 0)
                {
                    PrintReport.Enabled = true;
                    PrintReport.Visible = true;
                    ReportPanel.Height = 205;
                }
                else
                {
                    PrintReport.Enabled = false;
                    PrintReport.Visible = false;
                    ReportPanel.Height = 40;
                    MessageBox.Show("لا يوجد قائمة بالتاريخ المحدد");
                }
            }
        }

        private void dateTimeFrom_ValueChanged(object sender, EventArgs e)
        {
            FirstDate = true;
            if (LastDate) {
                int rows = DailyList(dateTimeFrom.Text, dateTimeTo.Text);
                if (rows > 0)
                {
                    PrintReport.Enabled = true;
                    PrintReport.Visible = true;
                    ReportPanel.Height = 205;
                }
                else
                {
                    PrintReport.Enabled = false;
                    PrintReport.Visible = false;
                    ReportPanel.Height = 40;
                    MessageBox.Show("لا يوجد قائمة بالتاريخ المحدد");
                }
            }
            string Currentmonth = "", CurrentDay = "";
            int year, month, date, m=0, d=0;
            DateTime datetime = dateTimeFrom.Value; 
            string[] YearMonthDayS = dateTimeFrom.Text.Split('-');
            year = Convert.ToInt16(YearMonthDayS[0]);
            m = Convert.ToInt16(YearMonthDayS[1]);
            d = Convert.ToInt16(YearMonthDayS[2]);
            

            if (m < 10) Currentmonth = "0" + m.ToString();
            else Currentmonth = m.ToString();
            if (d < 10) CurrentDay = "0" + d.ToString();
            else CurrentDay = d.ToString();
            string selecteddate = CurrentDay.ToString() + "-" + Currentmonth.ToString() + "-" + year.ToString();
            if (ReportType.SelectedIndex == 2 && DailyList(selecteddate)>0)
            {
                PrintReport.Enabled = true;
                PrintReport.Visible = true;
                ReportPanel.Height = 205;
            }
            else
            {
                PrintReport.Enabled = false;
                PrintReport.Visible = false;
                ReportPanel.Height = 40;
                MessageBox.Show("لا يوجد قائمة بالتاريخ المحدد");
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new HijriCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            HijriDate = DateTime.Now.ToString("dd-MM-yyyy");
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            GregorianDate = DateTime.Now.ToString("dd-MM-yyyy");
        }

        private void txtModel_TextChanged(object sender, EventArgs e)
        {

        }

        private void GoToForm(int indexNo, int locaIDNo)
        {            
            switch (indexNo)
            {
                case 0:
                    Form1 form1 = new Form1(locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut);
                    form1.ShowDialog();
                    break;
                case 1:
                    Form2 form2 = new Form2(locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut);
                    form2.ShowDialog();
                    break;
                case 2:
                    Form3 form3 = new Form3(locaIDNo, 0, EmployeeName, DataSource, FilespathIn, FilespathOut);
                    form3.ShowDialog();
                    break;
                case 3:
                    Form4 form4 = new Form4(locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut);
                    form4.ShowDialog();
                    break;
                case 4:
                    Form5 form5 = new Form5(locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut);
                    form5.ShowDialog();
                    break;
                case 5:
                    Form6 form6 = new Form6(locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut);
                    form6.ShowDialog();
                    break;
                case 6:
                    Form7 form7 = new Form7(locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut);
                    form7.ShowDialog();
                    break;
                case 7:
                    Form8 form8 = new Form8(locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut);
                    form8.ShowDialog();
                    break;
                case 8:
                    Form9 form9 = new Form9(locaIDNo, EmployeeName, DataSource, FilespathIn, FilespathOut);
                    form9.ShowDialog();
                    break;
                case 9:
                    Form10 form10 = new Form10(locaIDNo, 0, EmployeeName, DataSource, FilespathIn, FilespathOut);
                    form10.ShowDialog();
                    break;
                default:
                    break;
            }
        }
    }

}

