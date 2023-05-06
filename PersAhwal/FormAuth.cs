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
using OfficeOpenXml;
using Xceed.Document.NET;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Net;
using Xceed.Words.NET;
using System.Diagnostics;
using WIA;
using System.Diagnostics.Contracts;
using static System.Net.WebRequestMethods;
using File = System.IO.File;
using DocumentFormat.OpenXml.Office2010.Excel;
using Color = System.Drawing.Color;
using Chart = System.Windows.Forms.DataVisualization.Charting.Chart;
using ChartArea = System.Windows.Forms.DataVisualization.Charting.ChartArea;
using Legend = System.Windows.Forms.DataVisualization.Charting.Legend;
using Microsoft.Office.Core;
using static Azure.Core.HttpHeader;
using Aspose.Words.Settings;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.Security.AccessControl;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using System.Text.RegularExpressions;
using System.Data.SqlTypes;
using SautinSoft.Document;
using Path = System.IO.Path;
using OfficeOpenXml.Drawing.Controls;
using System.Windows.Forms.DataVisualization.Charting;

namespace PersAhwal
{
    public partial class FormAuth : Form
    {
        string archStat = "";
        string removedPersonName = "";
        string removedStat = "";
        string autheticatingOthes = "";
        string AuthenticName = "";
        bool colored = false;
        bool ArchData = true;
        string AuthNoPart1 = "ق س ج/160/12/";
        string AuthNoPart2 = "";
        public string rowCount = "";
        public bool NewAuth = false;
        int intID = -1;
        string archFile = @"D:\ArchiveFiles\";
        string FilespathIn, FilespathOut;
        bool timerColor = true;
        bool timer = true;
        bool steadyGrid = false;
        public Delegate DataMovePage;
        string[] dataSum = new string[50];
        string dataSource = "Data Source=192.168.100.100,49170;Network Library=DBMSSOCN;Initial Catalog=AhwalDataBase;User ID=ConsJeddahAdmin;Password=DataBC0nsJ49170";
        public Form11 ParentData { get; set; }
        bool fileloaded = false;
        string[] colIDs = new string[100];
        string[] allList;
        string DataSource = "";
        string updateAll = "";
        int addNameIndex = 0;
        int addAuthticIndex = 0;
        int currentPanelIndex = 0;
        int InvalidControl = 0;
        string EmpName = "";
        static string[,] preffix = new string[10, 20];
        string strRights = "";
        string strRightList = "";
        string ColName = "";
        string ColRight = "Col";
        Word.Document oBDoc;
        object oBMiss;
        Word.Application oBMicroWord;
        DataTable checkboxdt;
        int Nobox = 0, LastID = 0, LastTabIndex = 0;
        string LastCol = "";
        static string[] Text_statis = new string[5];
        string spacialCharacter = "";
        static int[] statistic = new int[100];
        static int[] staticIndex = new int[100];
        static int[] times = new int[100];
        bool ShowNewApp = false;
        string[] dataGrid = new string[50];
        string StrSpecPur = "";
        string TextModel = "";
        string starText = "";
        string starTextSub = "";
        int LegaceyIndex = 0;
        string LegaceyItem = "";
        string LegaceyPreStr = "";
        int idIndex = -1;
        string[] txtComboOptions = new string[5] { "", "", "", "", "" };
        string[] txtCheckOptions = new string[5] { "", "", "", "", "" };
        string lastInput1 = "";
        string lastInput2 = "";
        string lastInput3 = "";
        string lastInput4 = "";
        string lastInput5 = "";
        string[] foundList;
        bool test = false;
        int ButtonInfoIndex = 0;

        string legaceyAuthInfo = "";
        string archState = "new";
        string Jobposition = "";
        //public static string[] BirthName = new string[10];
        //public static string[] BirthPlace = new string[10];
        //public static string[] BirthDate = new string[10];
        //public static string[] BirthMother = new string[10];
        //public static string[] BirthDecs = new string[10];
        public string Mentioned = "باسمي";
        int idShow = 0;
        public string specialDataSum = "";
        bool addMade = false;
        string GreDate;
        string HijriDate;
        bool changeDetected = false;
        bool[] editsMade = new bool[2] { false, false };
        string oldText = "";
        string startText = "";
        bool notFiled = true;
        string[] charac = new string[20];
        string controlName = "";
        string IBAN = "";
        string[] itemsicheck1 = new string[5];
        int Atvc = 0;
        bool notAllowed = true;
        bool notAllowed1 = true;
        bool notAllowed2 = true;
        bool LibtnAdd1Vis = false;
        int MessageDocNo = 0;
        int onBehalfIndex = 0;
        bool proType1 = false;
        int autoFillIndex = 0;
        bool gridFill = true;
        string getSexIndex = "1";
        Chart chart1;
        Legend legend1;
        ChartArea chartArea1;
        int chartAreas1 = 0;
        bool holdData1 = false;
        int borderDashStyle1 = 0;
        string AuthTitle = "نائب قنصل";
        string AuthTitleLast = "نائب قنصل";
        string authother = "";
        string removeAuthother = "";
        string lastSentence = "";
        string[] txtReviewList;
        string[] txtReviewListStar;
        string[] txtRightList;
        int txtRevIndex = 0;
        int txtRevStarIndex = 0;
        int txtRigIndex = 0;
        int txtReviewListIndex = 0;
        int txtReviewStarIndex = 1;
        int txtRightListIndex = 0;
        string txtReviewLast = "";
        string قائمة_الحقوقLast = "";
        string selectTable = "";
        int starTextSubIndex = 0;
        int starTextIndex = 0;
        public FormAuth(int atvc, int rowid, string AuthNo, string datasource, string filespathIn, string filespathOut, string empName, string jobposition, string greDate, string hijriDate, bool testItems)
        {
            InitializeComponent();
            FilespathIn = filespathIn;
            FilespathOut = filespathOut + @"\";
            test = testItems;
            Atvc = atvc;
            //MessageBox.Show(Atvc.ToString());
            DataSource = datasource;
            fillSamplesCodes(DataSource);
            EmpName = empName;
            Jobposition = jobposition;
            التاريخ_الميلادي.Text = GreDate = greDate;
            التاريخ_الهجري.Text = HijriDate = hijriDate;
            genPreperations();
            FillDataGridView(DataSource);
            getMaxRange(DataSource);
            اسم_الموظف.Text = EmpName;

            //fillChart();
            chart1 = new Chart();
            legend1 = new Legend();
            chartArea1 = new ChartArea();
            //backgroundWorker2.RunWorkerAsync();
            //dataSourceWrite(FilespathOut + "autoDocs.txt", "No");
            //FindAndReplace(@"D:\ArchiveFiles\aa195648.docx", "إجراءات التنازل وتحويل السجل في إسمه", false);
        }
        private void fillSamplesCodes(string source)
        {
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select * from Tablechar";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;

                try
                {
                    cmd.ExecuteNonQuery();

                    DataTable table = new DataTable();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                    dataAdapter.Fill(table);
                    dataGridView2.DataSource = table;
                }
                catch (Exception ex) { }
                saConn.Close();
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
                AuthTitleLast = Environment.NewLine + AuthTitle;
            }
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
        public void PROCEGenNames()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { }
            SqlDataAdapter sqlDa = new SqlDataAdapter("PROCEGenNames", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);

        }
        public static void FindAndReplace(string loadPath, string text, bool remove)
        {
            DocumentCore dc = DocumentCore.Load(loadPath);
            Regex regex = new Regex(@text, RegexOptions.IgnoreCase);
            foreach (ContentRange item in dc.Content.Find(regex).Reverse())
            {
                if (remove)
                    item.Replace("", new CharacterFormat() { FontName = "Traditional Arabic", Size = 19.0, Bold = true });
                else item.Replace(text, new CharacterFormat() { FontName = "Traditional Arabic", Size = 19.0, Bold = true });
            }

            dc.Save(loadPath, SaveOptions.DocxDefault);
            System.Diagnostics.Process.Start(loadPath);
        }
        private void definColumn(FlowLayoutPanel panel)
        {
            foreach (System.Windows.Forms.Control control in panel.Controls)
            {
                if ((control is TextBox || control is ComboBox || control is CheckBox) && !control.Name.Contains("ff"))
                {
                    if (!checkColumnName(control.Name, DataSource))
                    {
                        CreateColumn(control.Name, DataSource);
                    }
                }
            }
        }

        private bool checkColumnName(string colNo, string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableAuth", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(dataRow["COLUMN_NAME"].ToString()))
                {
                    if (dataRow["COLUMN_NAME"].ToString() == colNo.Replace(" ", "_"))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private void CreateColumn(string Columnname, string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("alter table TableAuth add " + Columnname.Replace(" ", "_") + " nvarchar(500)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        public void boxesPreparations()
        {
            صفة_مقدم_الطلب_off.SelectedIndex = Appcases(النوع, addNameIndex);
            صفة_الموكل_off.SelectedIndex = Appcases(جنس_الموكَّل, addAuthticIndex);
            Console.WriteLine("boxesPreparations " + addNameIndex + صفة_مقدم_الطلب_off.SelectedIndex + addAuthticIndex + صفة_الموكل_off.SelectedIndex);
            
            
            if (addNameIndex == 1)
            {
                if (نوع_التوكيل.Text.Contains("ورثة"))
                {
                    نص_مقدم_الطلب1.Text = "أنا المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + "/ " + مقدم_الطلب.Text + "، المقيم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " بالمملكة العربية السعودية حامل" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " إصدار " + مكان_الإصدار.Text + "، وبكامل قوا" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " العقلية وبطوع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " واختيار" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " وحالت" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " المعتبرة شرعا وقانونا";
                    if (إجراء_التوكيل.Text != "إقرار بالتنازل")
                        legaceyAuthInfo = createAuthPart1(true);
                    else legaceyAuthInfo = createAuthPart1(false);
                }
                else
                {
                    if (إجراء_التوكيل.Text == "إقرار بالتنازل")
                        نص_مقدم_الطلب1.Text = "أنا المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + "/ " + مقدم_الطلب.Text + "، المقيم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " بالمملكة العربية السعودية حامل" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " إصدار " + مكان_الإصدار.Text + "، وبكامل قوا" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " العقلية وبطوع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " واختيار" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " وحالت" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " المعتبرة شرعاً وقانوناً، بهذا فقد تنازل" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 1] + " تنازلا نهائيا " + createAuthPart1(false);
                    else
                        نص_مقدم_الطلب1.Text = "أنا المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + "/ " + مقدم_الطلب.Text + "، المقيم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " بالمملكة العربية السعودية حامل" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " إصدار " + مكان_الإصدار.Text + "، وبكامل قوا" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " العقلية وبطوع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " واختيار" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " وحالت" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " المعتبرة شرعاً وقانوناً، بهذا فقد أوكل" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 1] + createAuthPart1(true);

                }
                
            }
            else if (addNameIndex > 1)
            {
                if (نوع_التوكيل.Text.Contains("ورثة"))
                {
                    نص_مقدم_الطلب0.Text = "نحن المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " الموقع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " ";
                    نص_مقدم_الطلب1.Text = "، والمقيم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " بالمملكة العربية السعودية، وبكامل قوا" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " العقلية وبطوع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " واختيار" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " وحالت" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " المعتبرة شرعا قانونا";
                    legaceyAuthInfo = createAuthPart1(true);
                }
                else
                {
                    نص_مقدم_الطلب0.Text = "نحن المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " الموقع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " ";
                    if (إجراء_التوكيل.Text.Contains("تنازل"))
                        نص_مقدم_الطلب1.Text = "، والمقيم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " بالمملكة العربية السعودية، وبكامل قوا" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " العقلية وبطوع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " واختيار" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " وحالت" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " المعتبرة شرعاً وقانوناً، بهذا فقد تنازل" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 1] + createAuthPart1(true);
                    else نص_مقدم_الطلب1.Text = "، والمقيم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " بالمملكة العربية السعودية، وبكامل قوا" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " العقلية وبطوع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " واختيار" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " وحالت" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " المعتبرة شرعاً وقانوناً، بهذا فقد أوكل" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 1] + createAuthPart1(true);

                }                
            }
            addMade = false;

            موقع_التوكيل1.Text = موقع_التوكيل.Text;
            if (!مقدم_الطلب.Text.Contains("_")) توقيع_مقدم_الطلب.Text = مقدم_الطلب.Text;
            else توقيع_مقدم_الطلب.Text = "";
        }


        public void panelFill(Control control)
        {
            if (allList is null) return;
            for (int col = 0; col < allList.Length; col++)
            {
                if (control.Name.Replace("V", "") == allList[col])
                {
                    if (dataGridView1.CurrentRow.Cells[allList[col]].Value.ToString() != "")
                    {
                        control.Text = dataGridView1.CurrentRow.Cells[allList[col]].Value.ToString();
                    }

                }
                else if (control.Name == allList[col])
                {
                    if (dataGridView1.CurrentRow.Cells[allList[col]].Value.ToString() != "")
                    {
                        control.Text = dataGridView1.CurrentRow.Cells[allList[col]].Value.ToString();
                    }

                }
            }
        }
        public void genPreperations()
        {
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            string[] forbidCol = new string[20] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", };

            //definColumn(panelapplicationInfo);
            //definColumn(Panelapp);
            //definColumn(panelAuthRights);
            //definColumn(finalPanel);

            forbidCol[0] = "المعالجة";
            forbidCol[1] = "ارشفة_المستندات";
            forbidCol[2] = "المكاتبة_النهائية";
            forbidCol[3] = "specialData";
            forbidCol[4] = "المكاتبات_الملغية";
            forbidCol[5] = "توكيل_مرجعي";
            forbidCol[6] = "رقم_هاتف1";
            forbidCol[7] = "sms";
            forbidCol[8] = "ID";
            forbidCol[9] = "Extension1";
            forbidCol[10] = "Extension2";
            forbidCol[11] = "Extension3";
            forbidCol[12] = "DocxData";

            charac[0] = "$$$";
            charac[1] = "&&&";
            charac[2] = "^^^";
            charac[3] = "###";
            charac[4] = "***";
            charac[5] = "%&%";
            charac[6] = "#$#";
            charac[7] = "&^&";
            charac[8] = "$$&";
            صفة_الموكل_off.SelectedIndex = 0;
            label36.Text = "الموظف:" + EmpName;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            dataGridView1.BackgroundColor = Color.White;
            allList = getColList("TableAuth", forbidCol);
            PanelDataGrid.Size = new System.Drawing.Size(1318, 600);
            PanelDataGrid.Location = new System.Drawing.Point(12, 38);
            //
            PanelDataGrid.BringToFront();
            //
            Suffex_preffixList();
            if (Jobposition.Contains("قنصل"))
            {
                btnDelete.Visible = true;
                //allowedEdit.Enabled = true;
            }
            اسم_المندوب.Text = "";
        }

        private void ColorFulGrid9()
        {

            int genAuth = 0;
            int arch = 0;
            int unDesc = 0;
            int inComb = 0;
            int i = 0;
            for (; i < dataGridView1.Rows.Count - 1; i++)
            {
                //dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;

                if (dataGridView1.Rows[i].Cells[2].Value.ToString() == "")
                {
                    inComb++;
                }
                if (dataGridView1.Rows[i].Cells["طريقة_الطلب"].Value.ToString().Contains("مندوب"))
                {
                    // timerColor = false;
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightPink;

                }
                if (dataGridView1.Rows[i].Cells["حالة_الارشفة"].Value.ToString() == "مؤرشف نهائي")
                {

                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    arch++;
                }

                if (dataGridView1.Rows[i].Cells["نوع_التوكيل"].Value.ToString() == "طلاق" && (dataGridView1.Rows[i].Cells["تاريخ_الميلاد"].Value.ToString() == "" || dataGridView1.Rows[i].Cells["المهنة"].Value.ToString() == ""))
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightCoral;

                }
            }
            labDescribed.Text = "عدد (" + i.ToString() + ") معاملة .. عدد (" + inComb.ToString() + ") غير مكتمل.. والمؤرشف منها عدد (" + arch.ToString() + ")...";

        }

        private string checkList(Panel panel, string[] List, string table)
        {
            string updateValues = "";

            foundList = new string[List.Length];
            for (int f = 0; f < List.Length; f++)
                foundList[f] = "";

            int found = 0;
            foreach (Control control in panel.Controls)
            {
                if (control is TextBox || control is ComboBox)
                    for (int col = 0; col < List.Length; col++)
                        if (control.Name == List[col] && control.Name != "رقم_المرجع_المرتبط_off")
                        {
                            foundList[found] = control.Name;
                            if (found == 0)
                            {
                                updateValues = control.Name + "=@" + control.Name;
                            }
                            else
                            {
                                updateValues = updateValues + "," + control.Name + "=@" + control.Name;
                            }
                            found++;
                        }
            }
            return updateAll = "UPDATE " + table + " SET " + updateValues + " where ID = @id";
        }

        private string checkList(FlowLayoutPanel panel, string[] List, string table)
        {
            string updateValues = "";

            foundList = new string[List.Length];
            for (int f = 0; f < List.Length; f++)
                foundList[f] = "";

            int found = 0;
            foreach (Control control in panel.Controls)
            {
                string name = control.Name;
                if (panel.Name == "PanelItemsboxes")
                    name = name.Replace("V", "");
                if (control is TextBox || control is ComboBox || control is CheckBox)
                    for (int col = 0; col < List.Length; col++)
                        if (name == List[col] && control.Name != "رقم_المرجع_المرتبط_off")
                        {
                            foundList[found] = name;
                            if (found == 0)
                            {
                                updateValues = name + "=@" + name;
                            }
                            else
                            {
                                updateValues = updateValues + "," + name + "=@" + name;
                            }
                            found++;
                        }
            }
            return updateAll = "UPDATE " + table + " SET " + updateValues + " where ID = @id";
        }

        private bool checkEmpty(FlowLayoutPanel panel)
        {
            foreach (Control control in panel.Controls)
            {
                if (!control.Visible) continue;
                if ((control.Name == "Vitext2" || control.Name == "Vitext3" || control.Name == "Vitext4" || control.Name == "Vitext5" || control.Name == "Vitext1") && ButtonInfoIndex != 0)
                    continue;
                if (control is TextBox || control is ComboBox)
                {
                    if (control.Text == "" || (control.Text == "P0" && !control.Name.Contains("هوية_الموكل")) || control.Text.Contains("إختر"))
                    {
                        control.BackColor = System.Drawing.Color.MistyRose;
                        if (panel.Name == "Panelapp") { panel.Height = 130 * addNameIndex; }
                        MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل " );
                        return false;
                    }
                }
            }
            return true;
        }

        private string numbersToLetters(string num) {
            switch (num) {
                case "0":
                    return " الأول";
                    break;
                case "1":
                    return " الثاني";
                    break;
                case "2":
                    return " الثالي";
                    break;
                case "3":
                    return " الرابع";
                    break;
                case "4":
                    return " الخامس";
                    break;
                case "5":
                    return " السادس";
                    break;
                case "6":
                    return " السابع";
                    break;
                case "7":
                    return " الثامن";
                    break;
                case "8":
                    return " التاسع";
                    break;
                case "9":
                    return " العاشر";
                    break;
               
                                      
            }
            return "";
        }

        private bool save2DataBase(FlowLayoutPanel panel)
        {

            string query = checkList(panel, allList, "TableAuth");
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@id", intID);
            bool cont = true;
            for (int i = 0; i < foundList.Length; i++)
            {
                if (foundList[i] == "تعليق")
                {
                    sqlCommand.Parameters.AddWithValue("@" + foundList[i], تعليق.Text);
                }
                else
                    foreach (Control control in panel.Controls)
                    {
                        string name = control.Name;
                        if (control is Label || control is Button || control is PictureBox || control.Name.Contains("_off") ) continue;

                        if (panel.Name == "PanelItemsboxes")
                            name = name.Replace("V", "");
                        if (name == foundList[i])
                        {
                            
                            Console.WriteLine(foundList[i] + " - " + control.Text);
                            sqlCommand.Parameters.AddWithValue("@" + foundList[i], control.Text);
                            break;
                        }
                    }
            }
            try
            {
                sqlCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(query);
            }
            return true;
        }
        //private bool save2DataBase(FlowLayoutPanel panel)
        //{

        //    string query = checkList(panel, allList, "TableAuth");
        //    SqlConnection sqlConnection = new SqlConnection(DataSource);
        //    if (sqlConnection.State == ConnectionState.Closed)
        //        sqlConnection.Open();
        //    SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
        //    sqlCommand.CommandType = CommandType.Text;
        //    sqlCommand.Parameters.AddWithValue("@id", intID);
        //    bool cont = true;
        //    for (int i = 0; i < foundList.Length; i++)
        //    {
        //        if (foundList[i] == "تعليق")
        //        {
        //            sqlCommand.Parameters.AddWithValue("@" + foundList[i], تعليق.Text);
        //        }
        //        else
        //            foreach (Control control in panel.Controls)
        //            {
        //                string name = control.Name;
        //                if (control is Label || control is Button || control is PictureBox || control.Name == "رقم_المرجع_المرتبط_off") continue;

        //                if (panel.Name == "PanelItemsboxes")
        //                    name = name.Replace("V", "");
        //                if (name == foundList[i])
        //                {
        //                    //Console.WriteLine(panel.Name + " - " + name);
        //                    if (control.Name == "اسم_المندوب" && control.Visible && !control.Text.Contains("-"))
        //                    {
        //                        control.BackColor = System.Drawing.Color.MistyRose;
        //                        MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل اسم المندوب ومنطقة التغطية مفصولين");
        //                        return false;
        //                    }
        //                    if ((control is TextBox && control.Text == "") || (control is ComboBox && control.Text.Contains("إختر")))
        //                        foreach (Control Econtrol in panel.Controls)
        //                        {
        //                            if ((Econtrol is TextBox && control.Text == "") || (Econtrol is ComboBox && Econtrol.Text.Contains("ختر")))
        //                                if (panel.Name != "PanelItemsboxes" || (Econtrol.Name != control.Name && Econtrol.Name.Contains(control.Name)) || Econtrol.Name == "اسم_المندوب")
        //                                {
        //                                    //MessageBox.Show(Econtrol.Name + " - " + control.Name);
        //                                    if (control.Name == "اسم_المندوب" && control.Visible)
        //                                    {
        //                                        //
        //                                        control.BackColor = System.Drawing.Color.MistyRose;
        //                                        MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل اسم_المندوب ");
        //                                        return false;
        //                                    }

        //                                    else if (!Econtrol.Name.Contains("هوية_الموكل") && control.Name != "اسم_المندوب" && control.Name != "txtRev")
        //                                    {
        //                                        Econtrol.BackColor = System.Drawing.Color.MistyRose;
        //                                        if (panel.Name == "Panelapp") { panel.Height = 130 * addNameIndex; }
        //                                        else if (panel.Name == "PanelAuthPers") { panel.Height = 90 * addAuthticIndex; }

        //                                        MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل " + control.Name.Replace("_", " ") + control.Text);
        //                                        return false;

        //                                    }
        //                                }
        //                                else if (panel.Name == "PanelItemsboxes")
        //                                {



        //                                    if (control.Visible && ButtonInfoIndex == 0)
        //                                    {
        //                                        control.BackColor = System.Drawing.Color.MistyRose;
        //                                        MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل غير المكتمل");
        //                                        return false;
        //                                    }
        //                                    else if (control.Visible && ButtonInfoIndex == 0)
        //                                    {
        //                                        if (!(Vitext1.Text == "" && Vitext2.Text == "" && Vitext3.Text == "" && Vitext4.Text == "" && Vicombo1.Text != ""))

        //                                        {
        //                                            MessageBox.Show("لا يمكن المتابعة يرجى تكملة بيانات الحقول غير المكتملة");
        //                                            return false;
        //                                        }
        //                                        //control.BackColor = System.Drawing.Color.MistyRose;

        //                                    }
        //                                }
        //                        }
        //                    //if (نوع_التوكيل.Text == "شهادة ميلاد")
        //                    //{
        //                    //    Vitext1.Text = BirthName[0];
        //                    //    Vitext2.Text = BirthPlace[0];
        //                    //    Vitext3.Text = BirthDate[0];
        //                    //    Vitext4.Text = BirthMother[0];
        //                    //    for (int x = 1; x < birthindex; x++)
        //                    //    {
        //                    //        Vitext1.Text = Vitext1.Text + "_" + BirthName[x];
        //                    //        Vitext2.Text = Vitext2.Text + "_" + BirthPlace[x];
        //                    //        Vitext3.Text = Vitext3.Text + "_" + BirthDate[x];
        //                    //        Vitext4.Text = Vitext4.Text + "_" + BirthMother[x];
        //                    //    }
        //                    //}
        //                    //if(control.Name == "طريقة_الطلب")
        //                    //MessageBox.Show(query);
        //                    Console.WriteLine(foundList[i] + " - " + control.Text);
        //                    sqlCommand.Parameters.AddWithValue("@" + foundList[i], control.Text);
        //                    break;
        //                }
        //            }
        //    }


        //    try
        //    {
        //        //MessageBox.Show(الاعدادات.Text);
        //        sqlCommand.ExecuteNonQuery();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(query);
        //    }
        //    return true;
        //}
        private void fillConInfo()
        {

        }

        private string[] getColList(string table, string[] forbidCol)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('" + table + "') and  name <> 'ID' and name not like 'Data%'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            string[] allList = new string[dtbl.Rows.Count];
            for (int col = 0; col < dtbl.Rows.Count; col++)
                allList[col] = "";

            int i = 0;
            string insertItems = "";
            string insertValues = "";
            string updateValues = "";
            foreach (DataRow row in dtbl.Rows)
            {
                bool forbidden = false;
                for (int f = 0; forbidCol[f] != ""; f++)
                    if (row["name"].ToString() == forbidCol[f])
                    {
                        forbidden = true;
                        break;
                    }
                if (!forbidden)
                {
                    Console.WriteLine(row["name"].ToString());
                    //MessageBox.Show(row["name"].ToString());
                    allList[i] = row["name"].ToString();
                    if (i == 0)
                    {
                        insertItems = row["name"].ToString();
                        insertValues = "@" + row["name"].ToString();
                        updateValues = row["name"].ToString() + "=@" + row["name"].ToString();
                    }
                    else
                    {
                        insertItems = insertItems + "," + row["name"].ToString();
                        insertValues = insertValues + "," + "@" + row["name"].ToString();
                        updateValues = updateValues + "," + row["name"].ToString() + "=@" + row["name"].ToString();
                    }
                    i++;
                }
            }
            updateAll = "UPDATE " + table + " SET " + updateValues + " where ID = @id";
            return allList;

        }
        private string[] getColList(string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('" + table + "') and  name <> 'ID' and name not like 'Data%'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            string[] allList = new string[dtbl.Rows.Count];
            for (int col = 0; col < dtbl.Rows.Count; col++)
                allList[col] = "";

            int i = 0;
            string insertItems = "";
            string insertValues = "";
            string updateValues = "";
            foreach (DataRow row in dtbl.Rows)
            {
                allList[i] = row["name"].ToString();
                if (i == 0)
                {
                    insertItems = row["name"].ToString();
                    insertValues = "@" + row["name"].ToString();
                    updateValues = row["name"].ToString() + "=@" + row["name"].ToString();
                }
                else
                {
                    insertItems = insertItems + "," + row["name"].ToString();
                    insertValues = insertValues + "," + "@" + row["name"].ToString();
                    updateValues = updateValues + "," + row["name"].ToString() + "=@" + row["name"].ToString();
                }
                i++;
            }
            updateAll = "UPDATE " + table + " SET " + updateValues + " where ID = @id";
            return allList;

        }

        public void addAuthenticPerson(string name, string sex, string nationality, string docNo)
        {
            // 
            // label42
            // 
            Label labelauthName = new Label();
            labelauthName.AutoSize = true;
            labelauthName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labelauthName.Location = new System.Drawing.Point(685, 0);
            labelauthName.Name = "label42" + addAuthticIndex + ".";
            labelauthName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labelauthName.Size = new System.Drawing.Size(68, 27);
            labelauthName.TabIndex = 433;
            labelauthName.Text = "اسم الموكَّل:";
            // 
            // txtAuthPerson1
            // 
            TextBox txtAuthPerson = new TextBox();
            txtAuthPerson.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            txtAuthPerson.Location = new System.Drawing.Point(419, 3);
            txtAuthPerson.Name = "الموكَّل_" + addAuthticIndex + ".";
            txtAuthPerson.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            txtAuthPerson.Size = new System.Drawing.Size(260, 35);
            txtAuthPerson.TabIndex = 432;
            txtAuthPerson.TextChanged += new System.EventHandler(this.txtAuthPerson1_TextChanged);
            txtAuthPerson.Text = name;
            // 
            // labeltitle7
            // 
            Label labelauthSex = new Label();
            labelauthSex.AutoSize = true;
            labelauthSex.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labelauthSex.Location = new System.Drawing.Point(373, 0);
            labelauthSex.Name = "labeltitle7_" + addAuthticIndex + ".";
            labelauthSex.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labelauthSex.Size = new System.Drawing.Size(40, 27);
            labelauthSex.TabIndex = 491;
            labelauthSex.Text = "النوع:";

            // 
            // txtAuthPersonsex1
            // 
            CheckBox txtAuthPersonsex = new CheckBox();
            txtAuthPersonsex.AutoSize = true;
            txtAuthPersonsex.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            txtAuthPersonsex.Location = new System.Drawing.Point(318, 3);
            txtAuthPersonsex.Name = "جنس_الموكَّل_" + addAuthticIndex + ".";
            txtAuthPersonsex.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            txtAuthPersonsex.Size = new System.Drawing.Size(49, 31);
            txtAuthPersonsex.TabIndex = 492;
            txtAuthPersonsex.Text = sex;
            if (txtAuthPersonsex.Text == "ذكر")
                txtAuthPersonsex.Checked = true;
            else txtAuthPersonsex.Checked = false;
            txtAuthPersonsex.UseVisualStyleBackColor = true;
            txtAuthPersonsex.CheckedChanged += new System.EventHandler(txtAuthPersonsex1_CheckedChanged);
            // 
            // combTitle7
            // 
            //this.combTitle7.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            //this.combTitle7.FormattingEnabled = true;
            //this.combTitle7.Items.AddRange(new object[] {
            //"Mr",
            //"Mrs",
            //"Miss",
            //"Madam"});
            //this.combTitle7.Location = new System.Drawing.Point(297, 3);
            //this.combTitle7.Name = "combTitle7";
            //this.combTitle7.Size = new System.Drawing.Size(15, 35);
            //this.combTitle7.TabIndex = 550;
            //this.combTitle7.Visible = false;
            // 
            // label3
            //
            Label labelauthNation = new Label();
            labelauthNation.AutoSize = true;
            labelauthNation.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labelauthNation.Location = new System.Drawing.Point(235, 0);
            labelauthNation.Name = "label3";
            labelauthNation.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labelauthNation.Size = new System.Drawing.Size(56, 27);
            labelauthNation.TabIndex = 530;
            labelauthNation.Text = "الجنسية:";
            // 
            // nantionality1
            // 
            ComboBox nantionality = new ComboBox();
            nantionality.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            nantionality.FormattingEnabled = true;
            nantionality.Items.AddRange(new object[] {
            "سوداني الجنسية",
            "سعودي الجنسية",
            "أخرى"});
            nantionality.Location = new System.Drawing.Point(55, 3);
            nantionality.Name = "جنسية_الموكل_" + addAuthticIndex + ".";
            nantionality.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            nantionality.Size = new System.Drawing.Size(174, 35);
            nantionality.TabIndex = 533;
            nantionality.Text = nationality;
            nantionality.TextChanged += new System.EventHandler(this.nantionalityID_TextChanged);
            // 
            // label9
            // 
            Label labelauthDocNo = new Label();
            labelauthDocNo.AutoSize = true;
            labelauthDocNo.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labelauthDocNo.Location = new System.Drawing.Point(667, 41);
            labelauthDocNo.Name = "label9";
            labelauthDocNo.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labelauthDocNo.Size = new System.Drawing.Size(86, 27);
            labelauthDocNo.TabIndex = 531;
            labelauthDocNo.Text = "الهوية/الاقامة:";
            // 
            // nantionalityID1
            // 
            TextBox authIDNo = new TextBox();
            authIDNo.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            authIDNo.Location = new System.Drawing.Point(175, 44);
            authIDNo.Name = "هوية_الموكل_" + addAuthticIndex + ".";
            authIDNo.RightToLeft = System.Windows.Forms.RightToLeft.No;
            authIDNo.Size = new System.Drawing.Size(486, 35);
            authIDNo.TabIndex = 532;
            authIDNo.Tag = "pass";
            authIDNo.Text = docNo;
            authIDNo.TextChanged += new System.EventHandler(this.authIDNo_TextChanged);
            authIDNo.MouseClick += new System.Windows.Forms.MouseEventHandler(this.DocAuthNo_MouseClick);
            // 
            // pictureBox11
            // 
            PictureBox addAuthPic = new PictureBox();
            addAuthPic.Image = global::PersAhwal.Properties.Resources.add;
            addAuthPic.Location = new System.Drawing.Point(115, 44);
            addAuthPic.Name = "pictureBox11_" + addAuthticIndex;
            addAuthPic.Size = new System.Drawing.Size(54, 35);
            addAuthPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            addAuthPic.TabIndex = 440;
            addAuthPic.TabStop = false;
            addAuthPic.Click += new System.EventHandler(addAuthPic_Click);
            // 
            // pictureBox13
            // 
            PictureBox removeAuthPic = new PictureBox();
            removeAuthPic.Image = global::PersAhwal.Properties.Resources.remove;
            removeAuthPic.Location = new System.Drawing.Point(55, 44);
            removeAuthPic.Name = "pictureBox13_" + addAuthticIndex;
            removeAuthPic.Size = new System.Drawing.Size(54, 35);
            removeAuthPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            removeAuthPic.TabIndex = 490;
            removeAuthPic.TabStop = false;
            removeAuthPic.Click += new System.EventHandler(removeAuthPic_Click);
            autoCompleteBulk(txtAuthPerson, DataSource, "الاسم", "TableGenNames");
            PanelAuthPers.Controls.Add(labelauthName);
            PanelAuthPers.Controls.Add(txtAuthPerson);
            PanelAuthPers.Controls.Add(labelauthSex);
            PanelAuthPers.Controls.Add(txtAuthPersonsex);
            PanelAuthPers.Controls.Add(labelauthNation);
            PanelAuthPers.Controls.Add(nantionality);
            PanelAuthPers.Controls.Add(labelauthDocNo);
            PanelAuthPers.Controls.Add(authIDNo);
            PanelAuthPers.Controls.Add(addAuthPic);
            PanelAuthPers.Controls.Add(removeAuthPic);

            addAuthticIndex++;
            //autoCompleteTextBox(txtAuthPerson, DataSource, "الاسم", "TableGenNames");
            //صفة_الموكل_off.SelectedIndex = Appcases(جنس_الموكَّل, addAuthticIndex);
            Console.WriteLine(addNameIndex + صفة_مقدم_الطلب_off.SelectedIndex + addAuthticIndex + صفة_الموكل_off.SelectedIndex);
        }

        private void txtAuthPersonsex1_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox checkBox = (CheckBox)sender;
            if (checkBox.CheckState == CheckState.Unchecked) checkBox.Text = "أنثى";
            else checkBox.Text = "ذكر";
            //صفة_الموكل_off.SelectedIndex = Appcases(جنس_الموكَّل, addAuthticIndex);
            checkChanged(جنس_الموكَّل, PanelAuthPers);
        }
        private void nantionalityID_TextChanged(object sender, EventArgs e)
        {
            checkChanged(جنسية_الموكل, PanelAuthPers);
        }
        private void txtAuthPerson1_TextChanged(object sender, EventArgs e)
        {
            checkChanged(الموكَّل, PanelAuthPers);
            if (gridFill) return;
            TextBox textBox = (TextBox)sender;
            string TextID = textBox.Name.Split('_')[1].Replace(".", "");
            string[] text = getID(textBox);
            fillFirstInfo(PanelAuthPers, "", text[5], text[1], text[0], text[2], "العربية", text[3], text[4], TextID);
        }
        private void authIDNo_TextChanged(object sender, EventArgs e)
        {
            checkChanged(هوية_الموكل, PanelAuthPers);
        }
        private void addAuthPic_Click(object sender, EventArgs e)
        {
            addAuthenticPerson("", "ذكر", "سوداني الجنسية", "P0");
            btnPanelAuthPers.Height = PanelAuthPers.Height = 90 * addAuthticIndex;
            checkChanged(جنس_الموكَّل, PanelAuthPers);
            checkChanged(جنسية_الموكل, PanelAuthPers);
            checkChanged(الموكَّل, PanelAuthPers);
            checkChanged(هوية_الموكل, PanelAuthPers);
        }

        private void removeAuthPic_Click(object sender, EventArgs e)
        {
            PictureBox pictureBox = (PictureBox)sender;
            //MessageBox.Show(pictureBox.Name); 
            string rowID = pictureBox.Name.Split('_')[1];

            foreach (Control control in PanelAuthPers.Controls)
            {
                if (control.Visible && control.Name.Contains("_" + rowID) && control.Name.Contains("."))
                {
                    control.Visible = false;
                    control.Name = "unvalid_" + InvalidControl.ToString();
                    InvalidControl++;
                }
            }
            if (addAuthticIndex > 0)
            {
                addAuthticIndex--;
                btnPanelAuthPers.Height = PanelAuthPers.Height = 90 * addAuthticIndex;
            }
            else
            {
                PanelAuthPers.Height = 90;
                addAuthenticPerson("", "ذكر", "سوداني الجنسية", "P0");
            }

            //صفة_الموكل_off.SelectedIndex = Appcases(جنس_الموكَّل, addAuthticIndex);
            Console.WriteLine(addNameIndex + صفة_مقدم_الطلب_off.SelectedIndex + addAuthticIndex + صفة_الموكل_off.SelectedIndex);
            checkChanged(جنس_الموكَّل, PanelAuthPers);
            checkChanged(جنسية_الموكل, PanelAuthPers);
            checkChanged(الموكَّل, PanelAuthPers);
            checkChanged(هوية_الموكل, PanelAuthPers);
        }

        public void addNames(string name, string sex, string docType, string docNo, string docIssue, string language, string job, string age)
        {
            Label labelName = new Label();
            labelName.AutoSize = true;
            labelName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labelName.Location = new System.Drawing.Point(673, 0);
            labelName.Name = "labelName_" + addNameIndex + ".";
            labelName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labelName.Size = new System.Drawing.Size(80, 27);
            labelName.TabIndex = 94;
            labelName.Text = "مقدم الطلب:";

            // 
            // AppName1
            // 
            TextBox AppName = new TextBox();
            AppName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            AppName.Location = new System.Drawing.Point(413, 3);
            AppName.Name = "مقدم_الطلب_" + addNameIndex + ".";
            AppName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            AppName.Size = new System.Drawing.Size(254, 35);
            AppName.TabIndex = 93;
            AppName.Text = name;

            AppName.TextChanged += new System.EventHandler(this.AppName_TextChanged);
            AppName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(AppName_KeyPress);

            // 
            // labeltitle1
            // 
            Label labeltitle1 = new Label();
            labeltitle1.AutoSize = true;
            labeltitle1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labeltitle1.Location = new System.Drawing.Point(367, 0);
            labeltitle1.Name = "labeltitle1_" + addNameIndex + ".";
            labeltitle1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labeltitle1.Size = new System.Drawing.Size(40, 27);
            labeltitle1.TabIndex = 176;
            labeltitle1.Text = "النوع:";
            // 
            // checkSexType1
            // 
            CheckBox checkSexType = new CheckBox();
            checkSexType.AutoSize = true;
            checkSexType.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            checkSexType.Location = new System.Drawing.Point(312, 3);
            checkSexType.Name = "النوع_" + addNameIndex + ".";
            checkSexType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            checkSexType.Size = new System.Drawing.Size(49, 31);
            checkSexType.TabIndex = 177;
            checkSexType.Text = sex;
            if (checkSexType.Text == "ذكر")
                checkSexType.Checked = true;
            else checkSexType.Checked = false;
            checkSexType.UseVisualStyleBackColor = true;
            checkSexType.CheckedChanged += new System.EventHandler(this.sexCheckedChanged);
            // 
            // combTitle1
            // 
            //ComboBox combTitle = new ComboBox();
            //combTitle.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            //combTitle.FormattingEnabled = true;
            //combTitle.Items.AddRange(new object[] {
            //"Mr",
            //"Mrs",
            //"Miss",
            //"Madam"});
            //combTitle.Location = new System.Drawing.Point(291, 3);
            //combTitle.Name = "النوع_الانجليزية_" + addNameIndex + ".";
            //combTitle.Size = new System.Drawing.Size(15, 35);
            //combTitle.TabIndex = 189;
            //combTitle.Visible = false;
            //combTitle.Text = sex;
            //if (language == "العربية")
            //{
            //    checkSexType.Visible = true;
            //    combTitle.Visible = false;
            //}
            //else
            //{
            //    checkSexType.Visible = false;
            //    combTitle.Visible = true;
            //}

            // 
            // label4
            // 
            Label labeldocType = new Label();
            labeldocType.AutoSize = true;
            labeldocType.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labeldocType.Location = new System.Drawing.Point(167, 0);
            labeldocType.Name = "label4_" + addNameIndex + ".";
            labeldocType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labeldocType.Size = new System.Drawing.Size(118, 27);
            labeldocType.TabIndex = 117;
            labeldocType.Text = "نوع اثبات الشخصية:";
            // 
            // DocType1
            // 
            ComboBox DocType = new ComboBox();
            DocType.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            DocType.FormattingEnabled = true;
            DocType.Items.AddRange(new object[] {
            "جواز سفر",
            "إقامة",
            "رقم وطني",
            "بطاقة قومية"});
            DocType.Location = new System.Drawing.Point(12, 3);
            DocType.Name = "نوع_الهوية_" + addNameIndex + ".";
            DocType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            DocType.Size = new System.Drawing.Size(149, 35);
            DocType.TabIndex = 122;
            DocType.Text = docType;
            DocType.TextChanged += new System.EventHandler(this.DocType_TextChanged);
            // 
            // labeldoctype1
            // 
            Label labeldocNo = new Label();
            labeldocNo.AutoSize = true;
            labeldocNo.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labeldocNo.Location = new System.Drawing.Point(653, 41);
            labeldocNo.Name = "labeldoctype1_" + addNameIndex + ".";
            labeldocNo.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labeldocNo.Size = new System.Drawing.Size(100, 27);
            labeldocNo.TabIndex = 119;
            labeldocNo.Text = "رقم جواز السفر: ";
            // 
            // DocNo1
            // 
            TextBox DocNo = new TextBox();
            DocNo.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            DocNo.Location = new System.Drawing.Point(464, 44);
            DocNo.Name = "رقم_الهوية_" + addNameIndex + ".";
            DocNo.RightToLeft = System.Windows.Forms.RightToLeft.No;
            DocNo.Size = new System.Drawing.Size(120, 35);
            DocNo.TabIndex = 120;
            DocNo.Tag = "pass";
            DocNo.Text = docNo;
            DocNo.TextChanged += new System.EventHandler(this.DocNo_TextChanged);
            DocNo.MouseClick += new System.Windows.Forms.MouseEventHandler(this.DocNo_MouseClick);

            // 
            // label7
            // 
            Label labeldocIssue = new Label();
            labeldocIssue.AutoSize = true;
            labeldocIssue.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labeldocIssue.Location = new System.Drawing.Point(371, 41);
            labeldocIssue.Name = "label7_" + addNameIndex + ".";
            labeldocIssue.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labeldocIssue.Size = new System.Drawing.Size(87, 27);
            labeldocIssue.TabIndex = 121;
            labeldocIssue.Text = "مكان الإصدار:";
            // 
            // DocIssue1
            // 
            TextBox DocIssue = new TextBox();
            DocIssue.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            DocIssue.Location = new System.Drawing.Point(152, 44);
            DocIssue.Name = "مكان_الإصدار_" + addNameIndex + ".";
            DocIssue.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            DocIssue.Size = new System.Drawing.Size(210, 35);
            DocIssue.TabIndex = 118;
            DocIssue.Text = docIssue;
            DocIssue.TextChanged += new System.EventHandler(this.DocIssue_TextChanged);
            // 
            // addName1
            //
            PictureBox addName = new PictureBox();
            addName.Image = global::PersAhwal.Properties.Resources.add;
            addName.Location = new System.Drawing.Point(92, 44);
            addName.Name = "addName_" + addNameIndex + ".";
            addName.Size = new System.Drawing.Size(54, 35);
            addName.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            addName.TabIndex = 123;
            addName.TabStop = false;
            addName.Click += new System.EventHandler(this.addName_Click);
            // 
            // removeName1
            // 
            PictureBox removeName = new PictureBox();
            removeName.Image = global::PersAhwal.Properties.Resources.remove;
            removeName.Location = new System.Drawing.Point(32, 44);
            removeName.Name = "removeName_" + addNameIndex + ".";
            removeName.Size = new System.Drawing.Size(54, 35);
            removeName.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            removeName.TabIndex = 175;
            removeName.TabStop = false;
            removeName.Click += new System.EventHandler(this.removeName_Click);
            //
            // Job
            //
            Label Job = new Label();
            Job.AutoSize = true;
            Job.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            Job.Location = new System.Drawing.Point(1129, 555);
            Job.Name = "label36_" + addNameIndex + ".";
            Job.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            Job.Size = new System.Drawing.Size(40, 27);
            Job.TabIndex = 604;
            Job.Text = "المهنة:";
            // 
            // المهنة
            // 
            TextBox textJob = new TextBox();
            textJob.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            textJob.Location = new System.Drawing.Point(801, 400);
            textJob.Name = "المهنة_" + addNameIndex + ".";
            textJob.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            textJob.Size = new System.Drawing.Size(550, 35);
            textJob.TabIndex = 603;
            textJob.Text = job;
            //autoCompleteTextBox(textJob, DataSource, "jobs", "TableListCombo");
            textJob.TextChanged += new System.EventHandler(this.textJob_TextChanged);
            // 
            // label37
            // 
            Label Age = new Label();
            Age.AutoSize = true;
            Age.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            Age.Location = new System.Drawing.Point(724, 555);
            Age.Name = "label37_" + addNameIndex + ".";
            Age.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            Age.Size = new System.Drawing.Size(75, 27);
            Age.TabIndex = 606;
            Age.Text = "تاريخ الميلاد:";
            // 
            // تاريخ_الميلاد
            //
            TextBox textAge = new TextBox();
            textAge.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            textAge.Location = new System.Drawing.Point(522, 552);
            textAge.Name = "تاريخ_الميلاد_" + addNameIndex + ".";
            textAge.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            textAge.Size = new System.Drawing.Size(100, 35);
            textAge.TabIndex = 844;
            textAge.Text = age;
            textAge.TextChanged += new System.EventHandler(this.textAge_TextChanged);
            autoCompleteBulk(AppName, DataSource, "الاسم", "TableGenNames");
            autoCompleteBulk(textJob, DataSource, "jobs", "TableListCombo");

            Panelapp.Controls.Add(labelName);
            Panelapp.Controls.Add(AppName);
            Panelapp.Controls.Add(labeltitle1);
            Panelapp.Controls.Add(checkSexType);
            //Panelapp.Controls.Add(combTitle);
            Panelapp.Controls.Add(labeldocType);
            Panelapp.Controls.Add(DocType);
            Panelapp.Controls.Add(labeldocNo);
            Panelapp.Controls.Add(DocNo);
            Panelapp.Controls.Add(labeldocIssue);
            Panelapp.Controls.Add(DocIssue);
            Panelapp.Controls.Add(Age);
            Panelapp.Controls.Add(textAge);
            Panelapp.Controls.Add(Job);
            Panelapp.Controls.Add(textJob);
            Panelapp.Controls.Add(addName);
            Panelapp.Controls.Add(removeName);
            addNameIndex++;
            //autoCompleteTextBox(AppName, DataSource, "الاسم", "TableGenNames");
            صفة_مقدم_الطلب_off.SelectedIndex = Appcases(النوع, addNameIndex);
            Console.WriteLine(addNameIndex + صفة_مقدم_الطلب_off.SelectedIndex + addAuthticIndex + صفة_الموكل_off.SelectedIndex);
            //Panelapp.Height = 130 * (addNameIndex);
        }
        private void autoCompleteBulk(TextBox textbox, string source, string col, string table)
        {

            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct " + col + " from " + table + " where " + col + " is not null";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                DataTable Textboxtable = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(Textboxtable);
                AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();
                bool newSrt = true;
                textbox.AutoCompleteCustomSource.Clear();

                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    autoComplete.Add(dataRow[col].ToString());
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }
        private void autoCompleteBulk(ComboBox comboBox, string source, string col, string table)
        {

            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct " + col + " from " + table + " where " + col + " is not null";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                DataTable Textboxtable = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(Textboxtable);
                AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();
                bool newSrt = true;
                comboBox.AutoCompleteCustomSource.Clear();

                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    autoComplete.Add(dataRow[col].ToString());
                }
                comboBox.AutoCompleteMode = AutoCompleteMode.Suggest;
                comboBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                comboBox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }

        public int Appcases(TextBox text, int index)
        {

            if (index == 0 || index == 1)
            {
                if (text.Text == "ذكر")
                {
                    return 0;//المقيم
                }
                else
                {
                    return 1;//المقيمة
                }
            }

            else if (index == 2)
            {
                if (text.Text.Split('_')[0] != "ذكر" && text.Text.Split('_')[0] != "ذكر")
                {
                    return 3;//المقيمتان
                }
                else
                {
                    return 2;//المقيمان
                }
            }

            else if (index == 3)
            {
                if (text.Text.Split('_')[0] != "ذكر" && text.Text.Split('_')[0] != "ذكر" && text.Text.Split('_')[0] != "ذكر")
                {
                    return 4;//المقيمات
                }
            }
            return 5;//المقيمون
        }
        string lastInput = "";
        private void textAge_TextChanged(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;

            if (textBox.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(textBox.Text, 1, 2));
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //textBox.Text = "";
                    textBox.Text = SpecificDigit(textBox.Text, 3, 10);
                    return;
                }
                else checkChanged(textBox, Panelapp);
            }
            if (textBox.Text.Length == 11)
            {
                textBox.Text = lastInput; return;
            }
            if (textBox.Text.Length == 10) return;
            if (textBox.Text.Length == 4) textBox.Text = "-" + textBox.Text;
            else if (textBox.Text.Length == 7) textBox.Text = "-" + textBox.Text;
            lastInput = textBox.Text;


        }
        private void textJob_TextChanged(object sender, EventArgs e)
        {
            checkChanged(المهنة, Panelapp);
        }
        private void AppName_TextChanged(object sender, EventArgs e)
        {
            checkChanged(مقدم_الطلب, Panelapp);
            if (gridFill) return;
            TextBox textBox = (TextBox)sender;
            string TextID = textBox.Name.Split('_')[2].Replace(".", "");
            string[] text = getID(textBox);
            fillFirstInfo(Panelapp, "", text[5], text[1], text[0], text[2], "العربية", text[3], text[4], TextID);
        }

        public void fillFirstInfo(FlowLayoutPanel panel, string name, string sex, string docType, string docNo, string docIssue, string language, string job, string age, string ID)
        {
            foreach (Control control in panel.Controls)
            {
                if (control.Name == "مقدم_الطلب_" + ID + "." && gridFill)
                    control.Text = name;

                if (control.Name == "الموكَّل_" + ID + "." && gridFill)
                    control.Text = name;

                if (control.Name == "النوع_" + ID + ".")
                {
                    control.Text = sex;
                    if (control.Text == "ذكر")
                        ((CheckBox)control).Checked = true;
                    else ((CheckBox)control).Checked = false;
                    if (language == "العربية")
                    {
                        control.Visible = true;
                    }
                    else
                    {
                        control.Visible = false;
                    }
                }

                if (control.Name == "جنس_الموكَّل" + ID + ".")
                {
                    control.Text = sex;
                    if (control.Text == "ذكر")
                        ((CheckBox)control).Checked = true;
                    else ((CheckBox)control).Checked = false;
                }
                if (control.Name == "title_" + ID + ".")
                {
                    control.Text = sex;
                    if (language == "العربية")
                    {
                        control.Visible = false;
                    }
                    else
                    {
                        control.Visible = true;
                    }
                }
                if (control.Name == "نوع_الهوية_" + ID + ".")
                    control.Text = docType;
                if (control.Name == "رقم_الهوية_" + ID + ".")
                    control.Text = docNo;
                if (control.Name == "هوية_الموكل_" + ID + ".")
                    control.Text = docNo;
                if (control.Name == "مكان_الإصدار_" + ID + ".")
                    control.Text = docIssue;
                if (control.Name == "المهنة_" + ID + ".")
                    control.Text = job;
                if (control.Name == "تاريخ_الميلاد_" + ID + ".")
                    control.Text = age;
            }
        }
        public string[] getID(TextBox text)
        {
            string[] returnedText = new string[6] { "P0", "جواز سفر", "", "", "", "ذكر" };
            string query = "SELECT * FROM TableGenNames where الاسم like N'" + text.Text + "%'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            foreach (DataRow row in dtbl.Rows)
            {
                returnedText[0] = row["رقم_الهوية"].ToString();
                returnedText[1] = row["نوع_الهوية"].ToString();
                returnedText[2] = row["مكان_الإصدار"].ToString();
                returnedText[3] = row["المهنة"].ToString();
                returnedText[4] = row["تاريخ_الميلاد"].ToString();
                returnedText[5] = row["النوع"].ToString();
            }
            return returnedText;
        }
        private void AppName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                TextBox textBox = (TextBox)sender;
                string index = textBox.Name.Split('_')[2].Replace(".", "");
                //MessageBox.Show(textBox.Text);
                writeIDChanged(textBox, Panelapp, "مقدم_الطلب", index);
            }
        }
        private void DocIssue_TextChanged(object sender, EventArgs e)
        {
            checkChanged(مكان_الإصدار, Panelapp);
        }

        private void DocType_TextChanged(object sender, EventArgs e)
        {
            checkChanged(نوع_الهوية, Panelapp);
        }
        private void DocNo_TextChanged(object sender, EventArgs e)
        {
            checkChanged(رقم_الهوية, Panelapp);
        }

        private void DocNo_MouseClick(object sender, EventArgs e)
        {
            //TextBox textBox = (TextBox)sender;
            //if (textBox.Text.Length > 3) return;
            //string index = textBox.Name.Split('_')[2].Replace(".", "");
            //writeIDChanged(textBox, Panelapp, "مقدم_الطلب",  index);
        }

        private void DocAuthNo_MouseClick(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            if (textBox.Text.Length > 3) return;
            string index = textBox.Name.Split('_')[2].Replace(".", "");
            writeIDChanged(textBox, PanelAuthPers, "الموكَّل", index);
        }
        private void sexCheckedChanged(object sender, EventArgs e)
        {

            CheckBox checkBox = (CheckBox)sender;
            if (checkBox.CheckState == CheckState.Unchecked) checkBox.Text = "أنثى";
            else checkBox.Text = "ذكر";
            صفة_مقدم_الطلب_off.SelectedIndex = Appcases(النوع, addNameIndex);
            checkChanged(النوع, Panelapp);
        }


        private void checkChanged(TextBox text, FlowLayoutPanel panel)
        {
            int index = 0;
            foreach (Control control in panel.Controls)
            {
                if (control.Visible && control.Name == text.Name + "_" + index + ".")
                {
                    if (index == 0) text.Text = control.Text;
                    else text.Text = text.Text + "_" + control.Text;
                    index++;
                }
            }
        }

        private void checkIDChanged(TextBox text, FlowLayoutPanel panel)
        {
            int index = 0;
            //MessageBox.Show(text.Name+" - "+ text.Text);
            foreach (Control control in panel.Controls)
            {
                if (control.Name.Contains("مقدم_الطلب_"))
                {
                    if (control.Text == text.Text)
                    {
                        //writeIDChanged("رقم_الهوية", panel, "مقدم_الطلب", index);

                        return;
                    }
                    index++;
                }
            }
        }

        private void writeIDChanged(TextBox textto, FlowLayoutPanel panel, string controlType, string index)
        {
            foreach (Control control in panel.Controls)
            {
                if (control.Name == controlType + "_" + index + ".")
                {
                    foreach (Control control2 in panel.Controls)
                    {
                        if (control2.Name == "رقم_الهوية_" + index + ".")
                            getID((TextBox)control2, control.Text.Trim(), "رقم_الهوية", fisrtWitIndex, "P0");
                        if (control2.Name == "تاريخ_الميلاد_" + index + ".")
                            getID((TextBox)control2, control.Text.Trim(), "تاريخ_الميلاد", fisrtWitIndex, "");
                        if (control2.Name == "المهنة_" + index + ".")
                            getID((TextBox)control2, control.Text.Trim(), "المهنة", fisrtWitIndex, "");
                        if (control2.Name == "نوع_الهوية_" + index + ".")
                            getID((ComboBox)control2, control.Text.Trim(), "نوع_الهوية", fisrtWitIndex, "جواز سفر");
                        if (control2.Name == "مكان_الإصدار_" + index + ".")
                            getID((TextBox)control2, control.Text.Trim(), "مكان_الإصدار", fisrtWitIndex, "");
                        if (control2.Name == "النوع_" + index + ".")
                            getID((CheckBox)control2, control.Text.Trim(), "النوع", fisrtWitIndex, "ذكر");
                        if (control2.Name == "هوية_الموكل_" + index + ".")
                            getID((TextBox)control2, control.Text.Trim(), "رقم_الهوية", fisrtWitIndex, "P0");
                    }
                }
            }
        }


        private bool checkGender(FlowLayoutPanel panel, string controlType, string control2type)
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
        
        private void relatedProUpdate()
        {
            if (!checkEdited()) return;
            string[] relatedPro = new string[17];
            relatedPro[0] = "النوع";
            relatedPro[1] = "نوع_الهوية";
            relatedPro[2] = "رقم_الهوية";
            relatedPro[3] = "مكان_الإصدار";
            relatedPro[4] = "التاريخ_الميلادي";
            relatedPro[5] = "التاريخ_الهجري";
            relatedPro[6] = "طريقة_الطلب";
            relatedPro[7] = "اسم_الموظف";
            relatedPro[8] = "اسم_المندوب";
            relatedPro[9] = "تعليق";
            relatedPro[10] = "تاريخ_الميلاد";
            relatedPro[11] = "المهنة";
            relatedPro[12] = "طريقة_الإجراء";
            relatedPro[13] = "رقم_الوكالة";
            relatedPro[14] = "جهة_إصدار_الوكالة";
            relatedPro[15] = "تاريخ_إصدار_الوكالة";
            relatedPro[16] = "مقدم_الطلب";
            string values= "النوع=@النوع";
            for (int x = 1; x < 17; x++) 
            {
                values = values +", " +relatedPro[x]+"=@"+ relatedPro[x];
            }
            string query = "Update TableCollection set " + values + " where رقم_المعاملة =N'" + رقم_المرجع_المرتبط_off.Text + "'";
            Console.WriteLine(query);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {

                    sqlCon.Open();
                }

                catch (Exception ex)
                {
                    return;
                }

            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            for (int i = 0; i < 17; i++) {
                foreach (Control control in Panelapp.Controls)
                {
                    if (control.Name == relatedPro[i])
                    {
                        Console.WriteLine(relatedPro[i] + " - " + control.Text);
                        sqlCmd.Parameters.AddWithValue("@" + relatedPro[i], control.Text);
                    }
                }
                
                foreach (Control control in panelapplicationInfo.Controls)
                {
                    if (control.Name == relatedPro[i])
                    {
                        Console.WriteLine(relatedPro[i] + " - " + control.Text);
                        sqlCmd.Parameters.AddWithValue("@" + relatedPro[i], control.Text);
                    }
                }foreach (Control control in finalPanel.Controls)
                {
                    if (control.Name == relatedPro[i])
                    {
                        Console.WriteLine(relatedPro[i] + " - " + control.Text);
                        sqlCmd.Parameters.AddWithValue("@" + relatedPro[i], control.Text);
                    }
                }
            }
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        public bool checkEdited()
        {
            if (رقم_المرجع_المرتبط_off.Text == "")
                return false;
            string query = "SELECT مقدم_الطلب FROM TableCollection where رقم_المعاملة = N'" + رقم_المرجع_المرتبط_off.Text+"'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            Console.WriteLine(query);
            sqlDa.Fill(dtbl);
            foreach (DataRow row in dtbl.Rows)
            {
                if (row["مقدم_الطلب"].ToString() != "")
                {
                    Console.WriteLine(row["مقدم_الطلب"].ToString());
                    return false;
                }
            }
            if (dtbl.Rows.Count == 0)
            {
                Console.WriteLine("معاملة غير موجودة");
                return false;
            }
            else return true;
        }


        private void fileUpload(string id, string text)
        {
            //MessageBox.Show(id);
            SqlConnection sqlCon = new SqlConnection(DataSource.Replace("AhwalDataBase", "ArchFilesDB"));
            if (sqlCon.State == ConnectionState.Closed)
                try
                {

                    sqlCon.Open();
                }

                catch (Exception ex)
                {
                    return;
                }

            SqlCommand sqlCmd = new SqlCommand("UPDATE TableGeneralArch SET fileUpload=N'" + text + "' WHERE رقم_معاملة_القسم=N'" + id + "'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        public void getID(TextBox textTo, string name, string controlType, int index, string def)
        {
            if (gridFill) return;
            string query = "SELECT " + controlType + " FROM TableGenNames where الاسم like N'" + name + "%'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            index = 0;
            textTo.Text = "";
            foreach (DataRow row in dtbl.Rows)
            {
                if (index == 0)
                    textTo.Text = row[controlType].ToString();
                else if (!textTo.Text.Contains(row[controlType].ToString()))
                    textTo.Text = textTo.Text + "_" + row[controlType].ToString();
                index++;
            }
            int AllIndex = textTo.Text.Split('_').Length;
            textTo.Text = textTo.Text.Split('_')[AllIndex - 1];
            if (index == 0)
                textTo.Text = def;
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

        public void getID(ComboBox textTo, string name, string controlType, int index, string def)
        {
            if (gridFill) return;
            string query = "SELECT " + controlType + " FROM TableGenNames where الاسم like N'" + name + "%'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            index = 0;
            textTo.Text = "";
            foreach (DataRow row in dtbl.Rows)
            {
                if (index == 0)
                    textTo.Text = row[controlType].ToString();
                else if (!textTo.Text.Contains(row[controlType].ToString()))
                    textTo.Text = textTo.Text + "_" + row[controlType].ToString();
                index++;
            }
            int AllIndex = textTo.Text.Split('_').Length;
            textTo.Text = textTo.Text.Split('_')[AllIndex - 1];
            if (index == 0)
                textTo.Text = def;
        }

        public void getID(CheckBox textTo, string name, string controlType, int index, string def)
        {
            if (gridFill) return;
            string query = "SELECT " + controlType + " FROM TableGenNames where الاسم like N'" + name + "%'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            index = 0;
            textTo.Text = "";
            foreach (DataRow row in dtbl.Rows)
            {
                if (index == 0)
                    textTo.Text = row[controlType].ToString();
                else if (!textTo.Text.Contains(row[controlType].ToString()))
                    textTo.Text = textTo.Text + "_" + row[controlType].ToString();
                index++;
            }
            int AllIndex = textTo.Text.Split('_').Length;
            textTo.Text = textTo.Text.Split('_')[AllIndex - 1];
            if (index == 0)
                textTo.Text = def;
        }
        private void addName_Click(object sender, EventArgs e)
        {
            addNames("", "ذكر", "جواز سفر", "P0", "", "العربية", "", "");
            btnPanelapp.Height = Panelapp.Height = 130 * addNameIndex;
            checkChanged(مقدم_الطلب, Panelapp);
            checkChanged(النوع, Panelapp);
            checkChanged(نوع_الهوية, Panelapp);
            checkChanged(رقم_الهوية, Panelapp);
            checkChanged(مكان_الإصدار, Panelapp);
            checkChanged(تاريخ_الميلاد, Panelapp);
            checkChanged(المهنة, Panelapp);
        }

        private void removeName_Click(object sender, EventArgs e)
        {
            PictureBox pictureBox = (PictureBox)sender;
            string rowID = pictureBox.Name.Split('_')[1];
            foreach (Control control in Panelapp.Controls)
            {
                if (control.Visible && control.Name.Contains("_" + rowID) && control.Name.Contains("."))
                {
                    control.Visible = false;
                    control.Name = "unvalid_" + InvalidControl.ToString();
                    InvalidControl++;
                }
            }
            if (addNameIndex > 0)
            {
                addNameIndex--;
                Panelapp.Height = 130 * addNameIndex;
            }
            else
            {
                Panelapp.Height = 130;
                addNames("", "ذكر", "جواز سفر", "P0", "", "العربية", "", "");

            }
            صفة_مقدم_الطلب_off.SelectedIndex = Appcases(النوع, addNameIndex);
            Console.WriteLine(addNameIndex + صفة_مقدم_الطلب_off.SelectedIndex + addAuthticIndex + صفة_الموكل_off.SelectedIndex);
            checkChanged(مقدم_الطلب, Panelapp);
            checkChanged(النوع, Panelapp);
            checkChanged(نوع_الهوية, Panelapp);
            checkChanged(رقم_الهوية, Panelapp);
            checkChanged(مكان_الإصدار, Panelapp);
            checkChanged(تاريخ_الميلاد, Panelapp);
            checkChanged(المهنة, Panelapp);
        }
        private void reSetPanel(FlowLayoutPanel panel)
        {
            foreach (Control control in panel.Controls)
            {
                if (control is TextBox)
                {
                    control.Text = "";
                }
            }
            التاريخ_الميلادي.Text = GreDate;
            التاريخ_الهجري.Text = HijriDate;
            نوع_التوكيل.Text = "إختر نوع التوكيل";
            إجراء_التوكيل.Text = "إختر الإجراء";
            هوية_الأول.Text = "P0";
            هوية_الثاني.Text = "P0";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //
            //Panel app
            //
            //reSetPanel(Panelapp);
            intID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            addAuthticIndex = addNameIndex = 0;

            if (dataGridView1.CurrentRow.Index != -1)
            {

                fillInfo(Panelapp, true);
                

                fillInfo(PanelAuthPers, true);
                fillInfo(panelAuthOptions, false);

                if (مقدم_الطلب.Text == "") ArchData = true;


                for (int app = 0; app < مقدم_الطلب.Text.Split('_').Length; app++)
                {
                    string appJob, appBirth;
                    try
                    {
                        appJob = المهنة.Text.Split('_')[app];
                        appBirth = تاريخ_الميلاد.Text.Split('_')[app];
                    }
                    catch (Exception ex)
                    {
                        appBirth = appJob = "";
                    }

                    if (مقدم_الطلب.Text.Split('_')[app] != "")
                    {
                        addNames(مقدم_الطلب.Text.Split('_')[app], النوع.Text.Split('_')[app], نوع_الهوية.Text.Split('_')[app], رقم_الهوية.Text.Split('_')[app], مكان_الإصدار.Text.Split('_')[app], "العربية", appJob, appBirth);
                        archState = "old";
                    }
                    else
                    {
                        addNames("", "ذكر", "جواز سفر", "P0", "", "العربية", "", appBirth);
                        archState = "new";

                    }
                }
                if (مقدم_الطلب.Text == "" && File.ReadAllText(FilespathOut + "autoDocs.txt") == "Yes")
                    FillDatafromGenArch("data1", intID.ToString(), "TableAuth");
                for (int app = 0; app < الموكَّل.Text.Split('_').Length; app++)
                {
                    string str = "";
                    try
                    {
                        str = مقدم_الطلب.Text.Split('_')[app];
                    }
                    catch (Exception ex) { }
                    if (str != "")
                        addAuthenticPerson(الموكَّل.Text.Split('_')[app], جنس_الموكَّل.Text.Split('_')[app], جنسية_الموكل.Text.Split('_')[app], هوية_الموكل.Text.Split('_')[app]);
                    else
                        addAuthenticPerson("", "ذكر", "سوداني الجنسية", "P0");
                }
                صفة_مقدم_الطلب_off.SelectedIndex = Appcases(النوع, addNameIndex);
                صفة_الموكل_off.SelectedIndex = Appcases(جنس_الموكَّل, addAuthticIndex);
                //MessageBox.Show("boxesPreparations " + addNameIndex + صفة_مقدم_الطلب_off.SelectedIndex + addAuthticIndex + صفة_الموكل_off.SelectedIndex);

                fillInfo(panelapplicationInfo, false);
                //MessageBox.Show(نوع_التوكيل.Text);
                fillInfo(PanelItemsboxes, false);
                //fillTextBoxesInvers();

                fillInfo(panelAuthRights, false);
                fillInfo(panelRemove, false);
                checkAutoUpdate.Checked = false;
                string[] settings = الاعدادات.Text.Split('_');
                try
                {
                    if (settings[0] == "1") توكيل_الغير.Checked = true;
                    else توكيل_الغير.Checked = false;
                    if (settings[1] == "1")
                    {
                        إلغاء_توكيل_سابق.Checked = true;
                        panelAuthOptions.Size = new System.Drawing.Size(1007, 261);
                    }
                    else
                    {
                        إلغاء_توكيل_سابق.Checked = false;
                        panelAuthOptions.Size = new System.Drawing.Size(1007, 313);
                    }
                    if (txtReview.Text == "" && نوع_التوكيل.Text != "توكيل بصيغة غير مدرجة")
                        checkAutoUpdate.Checked = true;
                }
                catch (Exception ex) { }
                fillInfo(finalPanel, false);
                التعليقات_السابقة_Off.Text = تعليق.Text;
                txtReview.Text = txtReview.Text.Replace("  ", " ");
                currentPanelIndex = 1;
                if (طريقة_الطلب.Text == "حضور مباشرة إلى القنصلية")
                {
                    طريقة_الطلب.Checked = true;
                    اسم_المندوب.Text = "";
                }
                else
                {
                    طريقة_الطلب.Checked = false;
                }
                    panelShow(currentPanelIndex);

                //                checkAutoUpdate.Checked = false; 
            }
            checkChanged(مقدم_الطلب, Panelapp);
            checkChanged(النوع, Panelapp);
            checkChanged(نوع_الهوية, Panelapp);
            checkChanged(رقم_الهوية, Panelapp);
            checkChanged(مكان_الإصدار, Panelapp);
            checkChanged(تاريخ_الميلاد, Panelapp);
            checkChanged(المهنة, Panelapp);
            checkChanged(الموكَّل, PanelAuthPers);
            checkChanged(جنسية_الموكل, PanelAuthPers);
            checkChanged(جنس_الموكَّل, PanelAuthPers);
            checkChanged(هوية_الموكل, PanelAuthPers);
            //
            //Panel app
            //
            gridFill = false;
            return;
        }

        private void prepareDocxfile()
        {

            oBMiss = System.Reflection.Missing.Value;
            oBMicroWord = new Word.Application();

            object objCurrentCopy = localCopy.Text;

            oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);
            oBMicroWord.Selection.Find.ClearFormatting();
            oBMicroWord.Selection.Find.Replacement.ClearFormatting();

        }
        private void fillDocFileAppInfo(FlowLayoutPanel panel)
        {
            foreach (Control control in panel.Controls)
            {
                if (control is TextBox || control is ComboBox)
                {
                    try
                    {
                        //if (control.Name == "التوقيع") 
                        //    MessageBox.Show(panel.Name + control.Text);
                        object ParaAuthIDNo = control.Name;
                        Word.Range BookAuthIDNo = oBDoc.Bookmarks.get_Item(ref ParaAuthIDNo).Range;
                        

                        object rangeAuthIDNo = BookAuthIDNo;
                        oBDoc.Bookmarks.Add(control.Name, ref rangeAuthIDNo);

                        //MessageBox.Show(control.Text);
                    }
                    catch (Exception ex)
                    {
                        //    MessageBox.Show(control.Name); 
                    }
                }
            }
            if (addNameIndex != 1)
            {

                //MessageBox.Show(addNameIndex.ToString());
                Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
                for (int x = 0; x < addNameIndex; x++)
                {
                    if (مقدم_الطلب.Text.Split('_')[x] != "")
                    {
                        table.Rows.Add();
                        table.Rows[x + 2].Cells[1].Range.Text = (x + 1).ToString();
                        table.Rows[x + 2].Cells[2].Range.Text = مقدم_الطلب.Text.Split('_')[x];
                        table.Rows[x + 2].Cells[3].Range.Text = رقم_الهوية.Text.Split('_')[x];
                        table.Rows[x + 2].Cells[4].Range.Text = مكان_الإصدار.Text.Split('_')[x];
                    }
                }
            }

            if (ButtonInfoIndex != 0)
            {
                fillTextBoxesDocx(addNameIndex, true);
            }
            else
                fillTextBoxesDocx(addNameIndex, false);
        }

        private void fillTextBoxesDocx(int index, bool libtnAdd1Vis)
        {
            if (index > 1) index = 2;
            try
            {
                Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[index];
           
            //MessageBox.Show(index.ToString());
            if (!libtnAdd1Vis)
            {
                //MessageBox.Show(index.ToString());
                table.Delete();
                return;
            }

            table.Rows[1].Cells[1].Range.Text = "الرقم";
            table.Rows[1].Cells[2].Range.Text = labl1.Text.Replace(":", "");
            table.Rows[1].Cells[3].Range.Text = labl2.Text.Replace(":", "");
            table.Rows[1].Cells[4].Range.Text = labl3.Text.Replace(":", "");
            table.Rows[1].Cells[5].Range.Text = labl4.Text.Replace(":", "");
            try
            {
                table.Rows[1].Cells[6].Range.Text = labl5.Text.Replace(":", "");
            }
            catch (Exception ex) { }
            for (int x = 0; x <= 4; x++)
            {
                int indBox = 1;
                foreach (Control control in PanelButtonInfo.Controls)
                {
                    if (x == 0)
                    {
                        table.Rows.Add();
                        table.Rows[indBox + 1].Cells[1].Range.Text = indBox.ToString();
                        indBox++;
                    }
                    else
                    {
                        if (control is TextBox && control.Name.Contains("textBox" + x + "_"))
                        {

                            table.Rows[indBox + 1].Cells[x + 1].Range.Text = control.Text;
                            indBox++;
                        }
                    }
                    if (indBox > ButtonInfoIndex) break;
                }
            }
            try
            {
                if (labl5.Text == "" || labl5.Text == "غير مدرج") table.Columns[6].Delete();
                if (labl4.Text == "" || labl4.Text == "غير مدرج") table.Columns[5].Delete();
                if (labl3.Text == "" || labl3.Text == "غير مدرج") table.Columns[4].Delete();
                if (labl2.Text == "" || labl2.Text == "غير مدرج") table.Columns[3].Delete();
                if (labl1.Text == "" || labl1.Text == "غير مدرج") table.Columns[2].Delete();

            }
            catch (Exception ex)
            {
            }
            }
            catch (Exception ex1) { }
        }
        private void fillDocFileInfo(Panel panel)
        {
            //MessageBox.Show(panel.Name);
            foreach (Control control in panel.Controls)
            {
                //MessageBox.Show(control.Text);
                if (control is TextBox || control is ComboBox)
                {
                    //if (control.Name == "التوقيع")
                    //MessageBox.Show(panel.Name + " "+ control.Name + " "+ control.Text);
                    try
                    {
                        object ParaAuthIDNo = control.Name;
                        Word.Range BookAuthIDNo = oBDoc.Bookmarks.get_Item(ref ParaAuthIDNo).Range;
                        if (control.Name == "موقع_التوكيل")
                            BookAuthIDNo.Text = control.Text + AuthTitleLast;
                        else BookAuthIDNo.Text = control.Text;
                        object rangeAuthIDNo = BookAuthIDNo;
                        oBDoc.Bookmarks.Add(control.Name, ref rangeAuthIDNo);


                    }
                    catch (Exception ex)
                    {
                        //    MessageBox.Show(control.Name); 
                    }
                }
            }
        }
        private void fillPrintDocx(string deleteDocxFile)
        {
            btnPrint.Enabled = false;
            //MessageBox.Show(localCopy.Text);
            string pdfFile = localCopy.Text.Replace("docx", "pdf");
            try
            {
                oBDoc.SaveAs2(localCopy.Text);
            }
            catch (Exception ex) { }
            if (deleteDocxFile == "pdf")
                oBDoc.ExportAsFixedFormat(pdfFile, Word.WdExportFormat.wdExportFormatPDF);
            try
            {
                oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            if (deleteDocxFile == "pdf")
            {
                System.Diagnostics.Process.Start(pdfFile);
                File.Delete(localCopy.Text);
            }
            else System.Diagnostics.Process.Start(localCopy.Text);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
            }
            catch (Exception ex1) { }
        }
        public void ComboProcedure_Text()
        {

        }
        public void resetBoxes(bool resetBoxes)
        {
            //MessageBox.Show("resetBoxes");
            checkboxdt = new DataTable();
            checkboxdt.Clear();
            Nobox = 0;
            strRights = "";
            ColName = "Col0";

            foreach (Control control in panelAuthOptions.Controls)
            {
                //control.Visible = false;
                if (control is CheckBox)
                {
                    ((CheckBox)control).Visible = false;
                    ((CheckBox)control).CheckState = CheckState.Unchecked;
                    ((CheckBox)control).Tag = "dispoase";

                }

                if (control is PictureBox)
                {
                    ((PictureBox)control).Visible = false;
                }
            }
            txtReview.Text = "";
            if (resetBoxes)
            {
                foreach (Control control in PanelItemsboxes.Controls)
                {
                    control.Visible = false;
                    if (control is ComboBox || control is TextBox)
                        control.Text = "";
                    if (control is ComboBox)
                    {
                        ((ComboBox)control).Items.Clear();
                    }
                    else if (control is CheckBox) ((CheckBox)control).CheckState = CheckState.Unchecked;
                }
            }
        }

        int listchecked = 0;
        public void PopulateCheckBoxes(bool genForm, string col, string table, string dataSource, int caseIndex, bool changeText)
        {

            LastCol = col;
            if (genForm) LastCol = col = "توكيل_بصيغة_غير_مدرجة";
            if (col == "الحقوق" || col == "Col" || col == "" || table == "" || dataSource == "") return;
            string query = "SELECT ID," + col.Replace("-", "_") + " FROM " + table;
            //قائمة_الحقوق.Text = "";
            string textRights = "";
            //MessageBox.Show(query);
            resetBoxes(false);
            using (SqlConnection con = new SqlConnection(dataSource))
            {

                using (SqlDataAdapter sda = new SqlDataAdapter(query, con))
                {
                    Console.WriteLine(query);
                    try
                    {
                        sda.Fill(checkboxdt);
                    }
                    catch (Exception ex) { return; }
                    listchecked = checkboxdt.Rows.Count;
                    Nobox = 0;
                    int rowsIndex = 0;
                    foreach (DataRow row in checkboxdt.Rows)
                    {
                        if (rowsIndex == 0)
                        {
                            rowsIndex++;
                            continue;
                        }
                        Text_statis = row[col.Replace("-", "_")].ToString().Split('_');
                        if (row[col.Replace("-", "_")].ToString() == "") continue;
                        string text = SuffReplacements(Text_statis[0], caseIndex, صفة_الموكل_off.SelectedIndex);
                        if (text.Contains("الحق في توكيل الغير"))
                            authother = text;

                        if (text.Contains("ويعتبر التوكيل الصادر"))
                            removeAuthother = text;
                        if (text.Contains("لمن يشهد والله"))
                            lastSentence = text;
                        if (checkboxdt.Rows[Nobox][col.Replace("-", "_")].ToString() == "" || checkboxdt.Rows[Nobox][col.Replace("-", "_")].ToString() == "null") return;

                        try
                        {
                            statistic[Nobox] = Convert.ToInt32(Text_statis[1]);
                            times[Nobox] = Convert.ToInt32(Text_statis[2]);
                            staticIndex[Nobox] = Convert.ToInt32(Text_statis[3]);

                            if (!textRights.Contains(text))
                                textRights = textRights + text + " ";

                            if (!توكيل_الغير.Checked && authother != "")
                            {
                                //MessageBox.Show(authother);
                                textRights = textRights.Replace(authother, "") + " ";
                            }

                            if (!إلغاء_توكيل_سابق.Checked && removeAuthother != "")
                            {
                                //MessageBox.Show(removeAuthother);
                                textRights = textRights.Replace(removeAuthother, "") + " ";
                            }
                            //MessageBox.Show(قائمة_الحقوق.Text );
                        }
                        catch (Exception ex) { }
                    }
                }
            }
            قائمة_الحقوق.Text = textRights;
            //autoCompleteTextBox(txtAddRight, DataSource, "قائمة_الحقوق_الكاملة", "TableAuthRight");
        }

        public string ibanText(string col, string table, string dataSource)
        {
            if (col == "Col" || col == "" || table == "" || dataSource == "") return "";
            string query = "SELECT ID," + col.Replace(" ", "_") + " FROM " + table;
            //MessageBox.Show("query " + query);


            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
            }
            catch (Exception ex) { return ""; }

            foreach (DataRow row in dtbl.Rows)
            {
                Text_statis = row[col].ToString().Split('_');
                //MessageBox.Show("Text_statis " + Text_statis[0]);
                if (Text_statis[0].Contains("إيداع "))
                {
                    //MessageBox.Show("إيداع " + Text_statis[0]);
                    return Text_statis[0];
                }
            }

            return "";
        }


        private void drawboxes(string txt, int idbox, bool check)
        {
            CheckBox chk = new CheckBox();
            chk.TabIndex = idbox;
            chk.Width = 80;
            chk.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            if (idbox == 0) chk.Width = panelAuthOptions.Width - 100;
            else chk.Width = panelAuthOptions.Width - 130;
            chk.Height = 33;
            chk.CheckState = CheckState.Unchecked;
            chk.Location = new System.Drawing.Point(130, 3 + idbox * 37);
            chk.Name = "checkBox_" + idbox.ToString();

            chk.Text = txt;
            chk.Tag = "valid";
            chk.CheckedChanged += new System.EventHandler(this.chk_CheckedChanged);
            chk.Checked = check;
            panelAuthOptions.Controls.Add(chk);



            PictureBox picboxedit = new PictureBox();
            picboxedit.Image = global::PersAhwal.Properties.Resources.edit;
            picboxedit.Location = new System.Drawing.Point(55, idbox * 37);
            picboxedit.Name = idbox.ToString();
            picboxedit.Size = new System.Drawing.Size(24, 26);
            picboxedit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            picboxedit.TabIndex = 175 + idbox;
            picboxedit.TabStop = false;
            picboxedit.Click += new System.EventHandler(this.pictureBoxedit_Click);
            panelAuthOptions.Controls.Add(picboxedit);

            PictureBox picboxup = new PictureBox();
            picboxup.Image = global::PersAhwal.Properties.Resources.arrowup;
            picboxup.Location = new System.Drawing.Point(86, idbox * 37);
            picboxup.Name = "Up";
            picboxup.Size = new System.Drawing.Size(24, 26);
            picboxup.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            picboxup.TabIndex = 176 + idbox;
            picboxup.TabStop = false;
            picboxup.Click += new System.EventHandler(this.pictureBoxup_Click);
            if (idbox == 0)
            {
                picboxup.Visible = false;
            }
            if (chk.Text.Contains("لمن يشهد والله خير الشاهدين")) picboxup.Visible = false;
            panelAuthOptions.Controls.Add(picboxup);

            PictureBox picboxdown = new PictureBox();
            picboxdown.Image = global::PersAhwal.Properties.Resources.arrowdown;
            picboxdown.Location = new System.Drawing.Point(107, idbox * 37);
            picboxdown.Name = "Down";
            picboxdown.Size = new System.Drawing.Size(24, 26);
            picboxdown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            picboxdown.TabIndex = 177 + idbox;
            picboxdown.TabStop = false;
            picboxdown.Click += new System.EventHandler(this.pictureBoxdown_Click);
            if (chk.Text.Contains("ويعتبر التوكيل الصادر") || chk.Text.Contains("لمن يشهد والله خير الشاهدين"))
                picboxdown.Visible = false;
            panelAuthOptions.Controls.Add(picboxdown);

            //PictureBox picboRemove = new PictureBox();
            //picboRemove.Image = global::PersAhwal.Properties.Resources.remove;
            //picboRemove.Location = new System.Drawing.Point(24, idbox * 37);
            //picboRemove.Name = "Remove";
            //picboRemove.Size = new System.Drawing.Size(24, 26);
            //picboRemove.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            //picboRemove.TabIndex = 177 + idbox;
            //picboRemove.TabStop = false;
            //picboRemove.Click += new System.EventHandler(this.pictureBoxdown_Click);
            //if (chk.Text.Contains("ويعتبر التوكيل الصادر") || chk.Text.Contains("لمن يشهد والله خير الشاهدين"))
            //    picboRemove.Enabled = false;
            //panelAuthOptions.Controls.Add(picboRemove);

        }

        public void pictureBoxedit_Click(object sender, EventArgs e)
        {
            PictureBox picbox = (PictureBox)sender;
            //MessageBox.Show(picbox.Name);
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    //MessageBox.Show(((CheckBox)control).Name);
                    //if (((CheckBox)control).TabIndex == picbox.TabIndex - 177)
                    if (((CheckBox)control).Name.Split('_')[1] == picbox.Name && ((CheckBox)control).Tag.ToString() == "valid")
                    {
                        //txtAddRight.Text = ((CheckBox)control).Text;
                        //btnAddRight.Text = "تعديل";
                        //btnRemove.Enabled = true;
                        LastTabIndex = ((CheckBox)control).TabIndex;
                        controlName = ((CheckBox)control).Name;
                        //MessageBox.Show(LastTabIndex +"_"+ controlName);
                        return;
                    }
                }
            }
        }

        public void PopulateCheckBoxes(string[] textList)
        {
            resetBoxes(false);
            listchecked = textList.Length;
            Nobox = 0;

            foreach (string str in textList)
            {
                string text = "";
                bool trueFalse = false;
                if (str.Length > 5 && !str.Contains("تحديث تلقائي"))
                {
                    try
                    {
                        if (!str.Contains("والله خير الشاهدين"))
                            text = str.Split('_')[1] + "،";
                        else text = str.Split('_')[1];
                        if (str.Split('_')[0] == "1") trueFalse = true;
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(str);
                        text = str + "،";
                    }
                    drawboxes(text, Nobox, trueFalse);
                    Nobox++;
                }
            }

        }

        private void CreatestrAuthRight()
        {
            حقوق_التوكيل.Text = "1";
            int xindex = 0;
            //if(نوع_التوكيل.Text != "توكيل بصيغة غير مدرجة") قائمة_الحقوق.Text = "";
            //الحقوق_الممنوحة.Text = قائمة_الحقوق.Text;
            //MessageBox.Show(قائمة_الحقوق.Text);
            //if (نوع_التوكيل.SelectedIndex == 0)
            //{                
            //    الحقوق_الممنوحة.Text = قائمة_الحقوق.Text;
            //    txtfinal.Text = txtfinal.Text + txtReview.Text + ", " + قائمة_الحقوق.Text;
            //    return;
            //}
            //string checked_unchecked = "1_";
            //foreach (Control control in panelAuthOptions.Controls)
            //{
            //    if (control.Visible && control is CheckBox && !control.Text.Contains("(نص ملغي)") && !control.Text.Contains("تحديث تلقائي"))
            //    {
            //        if (((CheckBox)control).Checked)
            //        {
            //            checked_unchecked = "1_";
            //            if (xindex == 0)
            //            {
            //                قائمة_الحقوق.Text = ((CheckBox)control).Text;
            //            }
            //            else
            //            {
            //                قائمة_الحقوق.Text = قائمة_الحقوق.Text + " " + ((CheckBox)control).Text;
            //            }
            //        }
            //        else checked_unchecked = "0_";
            //        if (xindex == 0)
            //        {
            //            الحقوق_الممنوحة.Text = checked_unchecked + ((CheckBox)control).Text;
            //            //قائمة_الحقوق.Text = ((CheckBox)control).Text;
            //        }
            //        else
            //        {
            //            الحقوق_الممنوحة.Text = الحقوق_الممنوحة.Text + checked_unchecked + ((CheckBox)control).Text;
            //            //قائمة_الحقوق.Text = قائمة_الحقوق.Text + " " + ((CheckBox)control).Text;
            //        }


            //        xindex++;
            //    }
            //}
        }

        private void chk_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox checkBox = (CheckBox)sender;
            if (checkBox.Text.Contains("لاغ") && checkBox.Checked)
            {
                panelRemove.Visible = true;
            }
            else
            {
                panelRemove.Visible = false;
            }
        }
        public void pictureBoxdown_Click(object sender, EventArgs e)
        {
            PictureBox picbox = (PictureBox)sender;

            string st = "", nd = "";
            bool statest = false, statend = false; bool FirstCase = false;
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (!((CheckBox)control).Text.Contains("والله خير الشاهدين"))
                    {
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 177)
                        {
                            st = ((CheckBox)control).Text;
                            if (((CheckBox)control).CheckState == CheckState.Checked) statest = true;
                            else statest = false;
                        }
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 176)
                        {
                            nd = ((CheckBox)control).Text;
                            if (((CheckBox)control).CheckState == CheckState.Checked) statend = true;
                            else statend = false;
                        }
                        FirstCase = true;
                    }
                    else FirstCase = false;
                }
            }
            int x = 0, y = 0;

            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (!((CheckBox)control).Text.Contains("والله خير الشاهدين"))
                    {
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 177)
                        {
                            ((CheckBox)control).Text = nd;
                            if (statend) ((CheckBox)control).CheckState = CheckState.Checked;
                            else ((CheckBox)control).CheckState = CheckState.Unchecked;
                        }
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 176)
                        {
                            ((CheckBox)control).Text = st;
                            if (statest) ((CheckBox)control).CheckState = CheckState.Checked;
                            else ((CheckBox)control).CheckState = CheckState.Unchecked;
                            y = statistic[x];
                            statistic[x] = statistic[x + 1];
                            statistic[x + 1] = y;
                            y = staticIndex[x];
                            staticIndex[x] = staticIndex[x + 1];
                            staticIndex[x + 1] = y;
                        }
                        x++;
                    }
                }

            }
        }

        public void pictureBoxup_Click(object sender, EventArgs e)
        {


            PictureBox picbox = (PictureBox)sender;

            string st = "", nd = "";
            bool statest = false, statend = false;
            bool FirstCase = false;

            foreach (Control control in panelAuthOptions.Controls)
            {

                if (control is CheckBox)
                {

                    if (!((CheckBox)control).Text.Contains("والله خير الشاهدين"))
                    {
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 176)
                        {
                            st = ((CheckBox)control).Text;
                            if (((CheckBox)control).CheckState == CheckState.Checked) statest = true;
                            else statest = false;

                        }
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 177)
                        {
                            nd = ((CheckBox)control).Text;
                            if (((CheckBox)control).CheckState == CheckState.Checked) statend = true;
                            else statend = false;

                        }
                        FirstCase = true;
                    }
                    else FirstCase = false;

                }
            }
            int x = 0;

            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (!((CheckBox)control).Text.Contains("والله خير الشاهدين"))
                    {
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 176)
                        {
                            ((CheckBox)control).Text = nd;
                            if (statend) ((CheckBox)control).CheckState = CheckState.Checked;
                            else ((CheckBox)control).CheckState = CheckState.Unchecked;
                            int y = 0;

                            y = statistic[x];
                            statistic[x] = statistic[x - 1];
                            statistic[x - 1] = y;

                            y = staticIndex[x];
                            staticIndex[x] = staticIndex[x - 1];
                            staticIndex[x - 1] = y;
                        }
                        if (((CheckBox)control).TabIndex == picbox.TabIndex - 177)
                        {
                            ((CheckBox)control).Text = st;
                            if (statest) ((CheckBox)control).CheckState = CheckState.Checked;
                            else ((CheckBox)control).CheckState = CheckState.Unchecked;
                        }
                        x++;
                    }
                }
            }


        }



        public void pictureRemove_Click(object sender, EventArgs e)
        {
            PictureBox picbox = (PictureBox)sender;
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (((CheckBox)control).TabIndex == Convert.ToInt32(picbox.Name))
                    {
                        ((CheckBox)control).Text = ((CheckBox)control).Text + "(نص ملغي)";
                    }
                }
            }
        }

        private void flllPanelItemsboxes(string rowID, string cellValue)
        {
            if (نوع_التوكيل.Text == "إختر نوع التوكيل" || إجراء_التوكيل.Text == "إختر الإجراء") return;
            //            MessageBox.Show("rowID = " + rowID + " - cellValue=" + cellValue);
            string query = "select * from TableAddContext where altColName = N'" + نوع_التوكيل.Text + "' and altSubColName = N'" + إجراء_التوكيل.Text + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            int checkIndex = 0;
            Console.WriteLine(query);
            //MessageBox.Show(query);
            if (dtbl.Rows.Count > 0)

                foreach (DataRow dr in dtbl.Rows)
                {
                    ColName = dr["ColName"].ToString();
                    الحقوق_off.Text = ColRight = dr["ColRight"].ToString().Replace("-", "_");                    
                    StrSpecPur = TextModel = dr["TextModel"].ToString();
                    

                    txtReviewListStar[0] = TextModel;
                    txtRevStarIndex = 1;
                    //selectTable = "TableAuthsub";
                    //    StrSpecPur = getStarText(selectTable, نوع_التوكيل.Text.Replace(" ", "_") + "_" + إجراء_التوكيل.Text.Replace(" ", "_"), starTextSub.Split('_')[0]);
                    ////StrSpecPur = SuffOrigConvertments(getStarText(selectTable, نوع_التوكيل.Text.Replace(" ", "_") + "_" + إجراء_التوكيل.Text.Replace(" ", "_"), starTextSub.Split('_')[0]));
                    
                    timer1.Enabled = true;
                    foreach (Control Lcontrol in PanelItemsboxes.Controls)
                    {
                        if (Lcontrol is CheckBox)
                        {
                            try
                            {
                                itemsicheck1[checkIndex] = dr[Lcontrol.Name.Replace("V", "") + "Option"].ToString();
                                Lcontrol.Text = itemsicheck1[checkIndex].Split('_')[0];
                                checkIndex++;
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                        if (Lcontrol is ComboBox)
                        {
                            try
                            {
                                //MessageBox.Show(dr[Lcontrol.Name.Replace("V", "") + "Option"].ToString());
                                ((ComboBox)Lcontrol).Items.Clear();
                                string[] items = dr[Lcontrol.Name.Replace("V", "") + "Option"].ToString().Split('_');

                                for (int x = 0; x < items.Length; x++)
                                    ((ComboBox)Lcontrol).Items.Add(items[x]);
                            }
                            catch (Exception ex)
                            {
                            }
                        }

                        try
                        {
                            if (Lcontrol.Name.StartsWith("L"))
                            {
                                Lcontrol.Text = dr[Lcontrol.Name.Replace("L", "")].ToString();
                                if (Lcontrol.Text == "إضافة")
                                {
                                    //MessageBox.Show(Lcontrol.Name);
                                    PanelButtonInfo.Visible = true;
                                    if (dr["itext1"].ToString() != "") labl1.Text = dr["itext1"].ToString();
                                    if (dr["itext2"].ToString() != "") labl2.Text = dr["itext2"].ToString();
                                    if (dr["itext3"].ToString() != "") labl3.Text = dr["itext3"].ToString();
                                    if (dr["itext4"].ToString() != "") labl4.Text = dr["itext4"].ToString();
                                    if (dr["itext5"].ToString() != "") labl5.Text = dr["itext5"].ToString();
                                }

                                if (Lcontrol.Text != "")
                                {
                                    Lcontrol.Visible = true;

                                    foreach (Control Vcontrol in PanelItemsboxes.Controls)
                                    {
                                        if (Vcontrol.Name.Trim() == Lcontrol.Name.Replace("L", "V").Trim())
                                        {
                                            Vcontrol.Visible = true;
                                            string size = dr[Lcontrol.Name.Replace("L", "") + "Length"].ToString();
                                            Vcontrol.Width = Convert.ToInt32(size);
                                            if (Convert.ToInt32(size) >= 700)
                                            {
                                                if (Vcontrol is TextBox) ((TextBox)Vcontrol).Multiline = true;
                                                Vcontrol.Height = 150;
                                            }



                                        }



                                        if (Vcontrol.Name.Contains(Lcontrol.Name.Replace("L", "V") + "V") || Vcontrol.Name.Contains(Lcontrol.Name.Replace("L", "V") + "L"))
                                        {
                                            Vcontrol.Visible = true;
                                        }
                                    }
                                }
                            }

                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    return;
                }

        }

        void FillContextView(string text)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableAddContext order by ID desc", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["ColName"].Width = 250;
            sqlCon.Close();
        }
        //private void btnAddLegacey_Click_1(object sender, EventArgs e)
        //{
        //    string str = "";
        //    if (comboPropertyType.Text.Contains("مركبة"))
        //    {

        //        if (LegtextBox1.Text != "") str = "في السيارة من نوع " + LegtextBox1.Text;
        //        if (LegtextBox5.Text != "") str = str + " موديل العام (" + LegtextBox5.Text + ")";
        //        if (LegtextBox2.Text != "") str = str + "باللون " + LegtextBox2.Text;
        //        if (LegtextBox3.Text != "") str = str + " ورقم لوحة (" + LegtextBox3.Text + " )";
        //        if (LegtextBox4.Text != "") str = str + "وشاسيه بالرقم (" + LegtextBox4.Text + ") ";
        //        if (LegaceyIndex == 0 || LegaceyPreStr == " في التركة المذكورة أعلاه")
        //        {
        //            LegaceyIndex = 0;
        //            LegaceyPreStr = str;
        //        }
        //        else LegaceyPreStr = LegaceyPreStr + " و " + str;
        //    }
        //    else if (comboPropertyType.Text.Contains("عقار"))
        //    {

        //        if (LegtextBox1.Text != "") str = "في العقار بالرقم (" + LegtextBox1.Text;
        //        if (LegtextBox2.Text != "") str = str + ") بمربع رقم (" + LegtextBox2.Text + ")";
        //        if (LegtextBox3.Text != "") str = str + ") البالغ مساحتها(" + LegtextBox3.Text + "م.م)";
        //        if (LegtextBox4.Text != "") str = str + " ب" + LegtextBox4.Text + "-" + LegtextBox5.Text + " )";
        //        if (LegaceyIndex == 0 || LegaceyPreStr == " في التركة المذكورة أعلاه")
        //        {
        //            LegaceyIndex = 0;
        //            LegaceyPreStr = str;
        //        }
        //        else LegaceyPreStr = LegaceyPreStr + " و " + str;
        //    }
        //    else if (comboPropertyType.Text.Contains("أخرى"))
        //    {
        //        if (LegaceyIndex == 0 || LegaceyPreStr == " في التركة المذكورة أعلاه")
        //        {
        //            LegaceyIndex = 0;
        //            LegaceyPreStr = " في " + LegtxtBoxGeneral.Text;
        //        }
        //        else LegaceyPreStr = LegaceyPreStr + " وفي " + LegtxtBoxGeneral.Text;
        //    }
        //    else
        //    {
        //        LegaceyPreStr = " في التركة المذكورة أعلاه";
        //    }
        //    LegtextBox1.Text = LegtextBox2.Text = LegtextBox3.Text = LegtextBox4.Text = LegtextBox5.Text = LegtxtBoxGeneral.Text = "";
        //    LegaceyIndex++;
        //    //txtReviewBody();
        //}


        private string createAuthPart1(bool Auth)
        {
            string authDesc = "";
            string authSexTag = "";

            if (addAuthticIndex == 1)
            {
                if (جنس_الموكَّل.Text == "أنثى")
                    authSexTag = "ة";

                if (Auth)
                    authDesc = " السيد" + authSexTag + "/ " + الموكَّل.Text;
                else authDesc = " للسيد" + authSexTag + "/ " + الموكَّل.Text;

                string authDocType = "إقامة رقم ";
                if (!جنسية_الموكل.Text.Contains("سوداني"))
                    authDocType = "هوية وطينة رقم ";

                if (!جنسية_الموكل.Text.Contains("سوداني") && Auth)
                    authDesc = " السيد" + authSexTag + "/ " + الموكَّل.Text + " (" + جنسية_الموكل.Text + ") حامل" + authSexTag + " " + authDocType + " " + هوية_الموكل.Text;
                else if (!جنسية_الموكل.Text.Contains("سوداني") && !Auth)
                    authDesc = " للسيد" + authSexTag + "/ " + الموكَّل.Text + " (" + جنسية_الموكل.Text + ") حامل" + authSexTag + " " + authDocType + " " + هوية_الموكل.Text;
            }
            else if (addAuthticIndex > 1)
            {
                try
                {
                    authSexTag = "";
                    if (جنس_الموكَّل.Text.Split('_')[0] == "أنثى")
                        authSexTag = "ة";
                    authDesc = " كل من السيد" + authSexTag + "/ " + الموكَّل.Text.Split('_')[0];

                    string authDocType = "إقامة رقم ";
                    if (!جنسية_الموكل.Text.Split('_')[0].Contains("سوداني"))
                        authDocType = "هوية وطينة رقم ";

                    if (هوية_الموكل.Text.Split('_')[0].Length > 8)
                        authDesc = " السيد" + authSexTag + "/ " + الموكَّل.Text.Split('_')[0] + " (" + جنسية_الموكل.Text.Split('_')[0] + ") حامل" + authSexTag + " " + authDocType + " " + هوية_الموكل.Text.Split('_')[0];

                    for (int x = 1; x < addAuthticIndex; x++)
                    {
                        authSexTag = "";
                        if (جنس_الموكَّل.Text.Split('_')[x] == "أنثى")
                            authSexTag = "ة";


                        authDocType = "إقامة رقم ";
                        if (!جنسية_الموكل.Text.Split('_')[x].Contains("سوداني"))
                            authDocType = "هوية وطينة رقم";

                        if (!جنسية_الموكل.Text.Split('_')[x].Contains("سوداني"))
                            authDesc = authDesc + " والسيد" + authSexTag + "/ " + الموكَّل.Text.Split('_')[x] + " (" + جنسية_الموكل.Text.Split('_')[x] + ") حامل" + authSexTag + " " + authDocType + " " + هوية_الموكل.Text.Split('_')[x];
                        else authDesc = authDesc + " السيد" + authSexTag + "/ " + الموكَّل.Text.Split('_')[x];
                    }
                }
                catch (Exception ex)
                {

                }
            }
            //MessageBox.Show(authDesc);
            return authDesc;
        }

        private string SuffReplacements(string text, int appCaseIndex, int intAuthcases)
        {
            if (appCaseIndex < 0) appCaseIndex = 0;
            if (intAuthcases < 0) intAuthcases = 0;
            if (text.Contains("auth1"))
                text = text.Replace("auth1", legaceyAuthInfo);
            if (text.Contains("  "))
                text = text.Replace("  ", " ");
            if (text.Contains("t1"))
                text = text.Replace("t1", Vitext1.Text);
            if (text.Contains("t2"))
                text = text.Replace("t2", Vitext2.Text);
            if (text.Contains("t3"))
                text = text.Replace("t3", Vitext3.Text);
            if (text.Contains("t4"))
                text = text.Replace("t4", Vitext4.Text);
            if (text.Contains("t5"))
                text = text.Replace("t5", Vitext5.Text);
            if (text.Contains("c1"))
                text = text.Replace("c1", Vicheck1.Text);
            if (text.Contains("m1"))
                text = text.Replace("m1", Vicombo1.Text);
            if (text.Contains("m2"))
                text = text.Replace("m2", Vicombo2.Text);
            if (text.Contains("a1"))
                text = text.Replace("a1", LibtnAdd1.Text);
            if (text.Contains("n1"))
                text = text.Replace("n1", " " + VitxtDate1.Text + " ");
            if (text.Contains("@*@"))
            {
                spacialCharacter = "@*@";
                text = text.Replace("@*@", "لدى  برقم الايبان (" + Vitext3.Text + ")");
            }
            if (text.Contains("#8"))
                text = text.Replace("#8", removedDocNo.Text);
            if (text.Contains("#6"))
                text = text.Replace("#6", removedDocSource.Text);
            if (text.Contains("#7"))
                text = text.Replace("#7", removedDocDate.Text);
            
            if (text.Contains("AuthName"))
                text = text.Replace("AuthName", الموكَّل.Text);
            if (text.Contains("AuthNatio"))
                text = text.Replace("AuthNatio", جنسية_الموكل.Text);
            if (text.Contains("AuthDoc"))
                text = text.Replace("AuthDoc", هوية_الموكل.Text);

            for (int gridIndex = 0; gridIndex < dataGridView2.Rows.Count - 1; gridIndex++)
            {
                string code = dataGridView2.Rows[gridIndex].Cells["الرموز"].Value.ToString();
                string person = dataGridView2.Rows[gridIndex].Cells["الضمير"].Value.ToString();
                string[] replacemest = new string[6];
                try
                {
                    replacemest[0] = dataGridView2.Rows[gridIndex].Cells["المقابل1"].Value.ToString();
                    replacemest[1] = dataGridView2.Rows[gridIndex].Cells["المقابل2"].Value.ToString();
                    replacemest[2] = dataGridView2.Rows[gridIndex].Cells["المقابل3"].Value.ToString();
                    replacemest[3] = dataGridView2.Rows[gridIndex].Cells["المقابل4"].Value.ToString();
                    replacemest[4] = dataGridView2.Rows[gridIndex].Cells["المقابل5"].Value.ToString();
                    replacemest[5] = dataGridView2.Rows[gridIndex].Cells["المقابل6"].Value.ToString();
                }
                catch (Exception ex) { return text; }
                if (text.Contains(code))
                {
                    if (person == "1")
                        text = text.Replace(code, replacemest[appCaseIndex]);
                    else if (person == "2")
                        text = text.Replace(code, replacemest[intAuthcases]);
                }
            }
            return text;
        }



        private void chooseDocxFile(string appName, string docId, string docType, bool visible)
        {
            string RouteFile;
            string strID = "1";
            string docName = docId.Split('/')[2] + docId.Split('/')[3] + docId.Split('/')[4] + "_";

            if (visible)
            {
                strID = "2";
            }

            if (addNameIndex == 1)
            {
                RouteFile = FilespathIn + "SingleAuth" + strID + ".docx";
            }
            else
            {
                RouteFile = FilespathIn + "MultiAuth" + strID + ".docx";
            }
            if (docType == "شهادة ميلاد")
                RouteFile = FilespathIn + "newAuthbirth" + strID + ".docx";

            localCopy.Text = FilespathOut + docName + DateTime.Now.ToString("ddmmss") + ".docx";
            while (File.Exists(localCopy.Text))
            {
                localCopy.Text = FilespathOut + docName + DateTime.Now.ToString("ddmmss") + ".docx";
            }
            System.IO.File.Copy(RouteFile, localCopy.Text);
            FileInfo fileInfo = new FileInfo(localCopy.Text);
            if (fileInfo.IsReadOnly) fileInfo.IsReadOnly = false;
        }

        private void fillInfo(FlowLayoutPanel panel, bool hide)
        {
            foreach (Control control in panel.Controls)
            {
                if (hide)
                {
                    control.Visible = false;
                    //control.Text = "";
                }
                if (control.Name.Contains("."))
                {
                    control.Name = "unvalid_" + InvalidControl.ToString();
                    InvalidControl++;
                }
                else
                {
                    panelFill(control);
                }
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (currentPanelIndex <= 4)
                currentPanelIndex++;
            else return;
            panelShow(currentPanelIndex);

            //fillChart();
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            if (currentPanelIndex > 0) currentPanelIndex--;
            else return;
            if (currentPanelIndex == 0) FillDataGridView(DataSource);
            panelShow(currentPanelIndex);
            btnPrevious.BringToFront();
            btnNext.BringToFront();
            //if(currentPanelIndex == 0) FillDataGridView(DataSource);
        }
        public void panelShow(int panelIndex)
        {
            switch (panelIndex)
            {
                case 0:
                    //DataGrid
                    PanelDataGrid.Size = new System.Drawing.Size(1318, 600);
                    PanelDataGrid.Location = new System.Drawing.Point(12, 38);
                    PanelDataGrid.BringToFront();
                    btnNext.Visible = ListSearch.Visible = btnListView.Visible = PanelDataGrid.Visible = labDescribed.Visible = true;

                    //false
                    btnSettings.Visible = btnDelete.Visible = btnFile1.Visible = btnFile2.Visible = btnFile3.Visible = Panelapp.Visible = false;
                    finalPanel.Visible = panelAuthRights.Visible = panelAuthRights.Visible = btnPrevious.Visible = panelapplicationInfo.Visible = false;


                    break;
                case 1:
                    //Basic Info
                    صفة_مقدم_الطلب_off.SelectedIndex = 0;
                    صفة_الموكل_off.SelectedIndex = 0;
                    panelapplicationInfo.Size = new System.Drawing.Size(821, 624);
                    panelapplicationInfo.Location = new System.Drawing.Point(294, 38);
                    panelapplicationInfo.BringToFront();
                    btnPrevious.Visible = panelapplicationInfo.Visible = true;
                    //false
                    finalPanel.Visible = panelAuthRights.Visible = ListSearch.Visible = btnListView.Visible = panelAuthRights.Visible = PanelDataGrid.Visible = labDescribed.Visible = false;
                    btnDelete.Visible = btnFile1.Visible = btnFile2.Visible = btnFile3.Visible = Panelapp.Visible = true;

                    break;
                case 2:
                    

            التاريخ_الميلادي.Text = GreDate;
            التاريخ_الهجري.Text = HijriDate;
            
                    updateGenName(مقدم_الطلب.Text, رقم_التوكيل.Text);
                    relatedProUpdate();
                    reversTextRights();
                    if (إلغاء_توكيل_سابق.Checked)
                        panelRemove.Visible = true;

                    التعليقات_السابقة_Off.Text = اسم_الموظف.Text + "- قام ببعض التعديلات" + "*" + Environment.NewLine + DateTime.Now.ToString("g") + Environment.NewLine + "--------------" + Environment.NewLine + التعليقات_السابقة_Off.Text;

                    if (!backgroundWorker1.IsBusy)
                        backgroundWorker1.RunWorkerAsync();
                    //authrights
                    panelAuthOptions.Size = new System.Drawing.Size(1007, 313);
                    if (نوع_التوكيل.Text == "توكيل بصيغة غير مدرجة" && إجراء_التوكيل.Text == "توكيل بصيغة غير مدرجة")
                    {
                        إجراء_التوكيل.BackColor = System.Drawing.Color.MistyRose;
                        MessageBox.Show("يرجى اقتراح اسم للمعاملة");
                        currentPanelIndex--; return;
                    }
                    if (!checkGender(Panelapp, "مقدم_الطلب_", "النوع_"))
                    {
                        currentPanelIndex--; return;
                    }
                    else addNewAppNameInfo();

                    if (!checkEmpty(Panelapp))
                    {
                        currentPanelIndex--; return;
                    }
                    save2DataBase(Panelapp);

                    if (!checkGender(Panelapp, "الموكَّل_", "جنس_الموكَّل_"))
                    {
                        currentPanelIndex--; return;
                    }
                    else addNewAuthNameInfo();

                    if (!checkEmpty(PanelAuthPers))
                    {
                        currentPanelIndex--; return;
                    }
                    save2DataBase(PanelAuthPers);


                    اسم_الموظف.Text = EmpName;

                    string gender = getGender(الشاهد_الأول.Text.Split(' ')[0]);
                    if (gender != "ذكر")
                    {
                        var selectedOption = MessageBox.Show("تم رصد اسم سيدة في حقل الشاهد الأول", "هل تود تغيير إعدادات البرنامج الداخلية والمتابعة للصفحة التالية؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (selectedOption == DialogResult.No)
                        {
                            currentPanelIndex--; return;
                        }
                        else if (selectedOption == DialogResult.Yes)
                        {
                            updateGender("ذكر", getSexIndex);
                        }
                    }
                    gender = getGender(الشاهد_الثاني.Text.Split(' ')[0]);
                    if (gender != "ذكر")
                    {
                        var selectedOption = MessageBox.Show("تم رصد اسم سيدة في حقل الشاهد الأول", "هل تود تغيير إعدادات البرنامج الداخلية والمتابعة للصفحة التالية؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (selectedOption == DialogResult.No)
                        {
                            currentPanelIndex--; return;
                        }
                        else if (selectedOption == DialogResult.Yes)
                        {
                            updateGender("ذكر", getSexIndex);
                        }
                    }

                    if (!checkEmpty(panelapplicationInfo))
                    {
                        currentPanelIndex--; return;
                    }
                    save2DataBase(panelapplicationInfo);


                    //MessageBox.Show("صفة_الموكل_off.SelectedIndex " + صفة_الموكل_off.SelectedIndex.ToString() + " - addAuthticIndex " + addAuthticIndex.ToString());
                    صفة_مقدم_الطلب_off.SelectedIndex = Appcases(النوع, addNameIndex);
                    //MessageBox.Show("2- " + LibtnAdd1.Text);
                    notAllowed = false;
                    boxesPreparations();
                    //MessageBox.Show("8- " + LibtnAdd1.Text);
                    panelAuthRights.Size = new System.Drawing.Size(1315, 622);
                    panelAuthRights.Location = new System.Drawing.Point(10, 36);
                    panelAuthRights.BringToFront();
                    panelAuthRights.Visible = btnNext.Visible = true;
                    finalPanel.Visible = PanelDataGrid.Visible = panelapplicationInfo.Visible = false;
                    if (LibtnAdd1.Visible)
                    {
                        //MessageBox.Show("Visible");
                        LibtnAdd1Vis = true;
                        fillTextBoxesInvers();
                    }
                    if (txtReview.Text == "")
                    {
                        txtReview.Text = SuffReplacements(StrSpecPur, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);
                        checkAutoUpdate.Checked = true;
                    }

                    if (starTextSub == "" && txtReview.Text == "")
                        checkAutoUpdate.Checked = true;
                    
                    timer5.Enabled = true;
                    if (قائمة_الحقوق.Text == "")
                        getDefualtRight(DataSource);
                    //else MessageBox.Show(قائمة_الحقوق.Text);
                    break;
                case 3:
                    checkAutoUpdate.Checked = false;
                    قائمة_الحقوق.Text = removeSpace(قائمة_الحقوق.Text, true);
                    txtReview.Text = removeSpace(txtReview.Text, false);
                    SuffConvertments(txtReview, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);
                    SuffConvertments(قائمة_الحقوق, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);
                    
                    timer5.Enabled = false;
                    txtReview.Text = txtReview.Text.Replace("،،", "،");
                    if (توكيل_الغير.Checked)
                        الاعدادات.Text = "1";
                    else الاعدادات.Text = "0";

                    if (إلغاء_توكيل_سابق.Checked)
                        الاعدادات.Text = الاعدادات.Text + "_1";
                    else الاعدادات.Text = الاعدادات.Text + "_0";


                    //finalPanel
                    //if (backgroundWorker1.IsBusy) { currentPanelIndex--; return; }


                    //الحقوق_الممنوحة.Text = قائمة_الحقوق.Text;
                    flowLayoutPanel2.Size = new System.Drawing.Size(940, 241);
                    CreatestrAuthRight();
                    if (PanelButtonInfo.Visible)
                    {
                        fillTextBoxes(Vitext1, 1);
                        fillTextBoxes(Vitext2, 2);
                        fillTextBoxes(Vitext3, 3);
                        fillTextBoxes(Vitext4, 4);
                        fillTextBoxes(Vitext5, 5);
                        txtReview.Size = new System.Drawing.Size(1007, 57);
                    }
                    else
                        txtReview.Size = new System.Drawing.Size(1007, 150);

                    if (!checkEmpty(PanelItemsboxes))
                    {
                        currentPanelIndex--; return;
                    }
                    save2DataBase(PanelItemsboxes);

                    if (PanelButtonInfo.Visible)
                    {
                        Vitext1.Text = Vitext2.Text = Vitext3.Text = Vitext4.Text = Vitext5.Text = "";
                    }

                    if (!checkEmpty(panelAuthRights))
                    {
                        currentPanelIndex--; return;
                    }
                    save2DataBase(panelAuthRights);


                    if (!checkEmpty(panelAuthOptions))
                    {
                        currentPanelIndex--; return;
                    }
                    save2DataBase(panelAuthOptions);


                    
                    if (!checkEmpty(panelRemove))
                    {
                        currentPanelIndex--; return;
                    }
                    save2DataBase(panelRemove);

                    
                    if (!طريقة_الإجراء.Checked)
                    {
                        التوثيق.Size = new System.Drawing.Size(923, 123);
                        تعليق_جديد_Off.Size = new System.Drawing.Size(923, 53);
                        flowLayoutPanel1.Size = new System.Drawing.Size(923, 149);
                        flowLayoutPanel2.Size = new System.Drawing.Size(940, 175);
                    }
                    else 
                    {
                        التوثيق.Size = new System.Drawing.Size(923, 123);
                        تعليق_جديد_Off.Size = new System.Drawing.Size(923, 53);
                        flowLayoutPanel1.Size = new System.Drawing.Size(923, 219);
                        flowLayoutPanel2.Size = new System.Drawing.Size(940, 175);
                    }
                    
                    finalPanel.Size = new System.Drawing.Size(944, 622);
                    finalPanel.Location = new System.Drawing.Point(192, 38);
                    finalPanel.BringToFront();
                    finalPanel.Visible = true;
                    panelAuthRights.Visible = btnNext.Visible = PanelDataGrid.Visible = panelapplicationInfo.Visible = false;
                    authJob();
                    break;
            }
        }
        private void updateGenName(string name, string idDoc)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource.Replace("AhwalDataBase", "ArchFilesDB"));
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string query = "update TableGeneralArch set الاسم=N'" + name + "' where رقم_معاملة_القسم = '" + idDoc + "'";
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        private void pictureBox13_Click(object sender, EventArgs e)
        {

        }

        private void رقم_الهوية_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnPanelapp_Click(object sender, EventArgs e)
        {
            if (btnPanelapp.Height != 130)
            {
                btnPanelapp.Height = Panelapp.Height = 130;
                btnPanelapp.Text = "عرض";
            }
            else
            {
                btnPanelapp.Height = Panelapp.Height = 130 * addNameIndex;
                btnPanelapp.Text = "إخفاء";
            }

        }

        private void btnPanelAuthPers_Click(object sender, EventArgs e)
        {
            if (btnPanelAuthPers.Height != 90)
            {
                btnPanelAuthPers.Height = PanelAuthPers.Height = 90;
                btnPanelAuthPers.Text = "عرض";
            }
            else
            {
                btnPanelAuthPers.Height = PanelAuthPers.Height = 90 * addAuthticIndex;
                btnPanelAuthPers.Text = "إخفاء";
            }
        }

        private void طريقة_الطلب_CheckedChanged(object sender, EventArgs e)
        {
            if (طريقة_الطلب.Checked)
            {
                اسم_المندوب.Text = "";
                طريقة_الطلب.Text = "حضور مباشرة إلى القنصلية";
                //MessageBox.Show(طريقة_الطلب.Text + " true");
                mandoubLabel.Visible = اسم_المندوب.Visible = false;

                proType1 = false;
                الشاهد_الأول.Text = هوية_الأول.Text = "";

            }
            else if (!طريقة_الطلب.Checked)
            {

                طريقة_الطلب.Text = "عن طريق أحد مندوبي القنصلية";
                اسم_المندوب.Visible = mandoubLabel.Visible = true;
                proType1 = true;
                //MessageBox.Show(طريقة_الطلب.Text + " false");
                اسم_المندوب.Text = "إختر اسم المندوب";

            }

        }



        private void FormAuth_Load(object sender, EventArgs e)
        {
            diplomats(موقع_التوكيل, DataSource);
            if (موقع_التوكيل.Items.Count > Atvc)
                موقع_التوكيل.SelectedIndex = Atvc;
            getTitle(DataSource, موقع_التوكيل.Text);
            fileColComboBox(نوع_التوكيل, DataSource, "altColName");
            fileComboBox(الحقوق_off, DataSource, "ColRight", "TableAddContext", true);
            fileComboBoxMandoub(اسم_المندوب, DataSource, "TableMandoudList");
            autoCompleteBulk(إجراء_التوكيل, DataSource, "إجراء_التوكيل", "TableAuth");
            autoCompleteBulk(الشاهد_الأول, DataSource, "الاسم", "TableGenNames");
            autoCompleteBulk(الشاهد_الثاني, DataSource, "الاسم", "TableGenNames");
        }

        private void diplomats(ComboBox combbox, string source)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct EmployeeName from TableUser where EmployeeName is not null and الدبلوماسيون = N'yes' and Aproved like N'%أكده%' order by EmployeeName asc";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;

                Console.WriteLine(query);
                try
                {
                    cmd.ExecuteNonQuery();
                    DataTable table = new DataTable();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                    dataAdapter.Fill(table);

                    foreach (DataRow dataRow in table.Rows)
                    {
                        combbox.Items.Add(dataRow["EmployeeName"].ToString());
                    }
                }
                catch (Exception ex) { }
                saConn.Close();
            }
        }

        private void fileColComboBox(ComboBox combbox, string source, string comlumnName)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct " + comlumnName + " from TableAddContext where " + comlumnName + " is not null and ColRight <> '' order by " + comlumnName + " asc";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;

                Console.WriteLine(query);
                try
                {
                    cmd.ExecuteNonQuery();
                    DataTable table = new DataTable();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                    dataAdapter.Fill(table);

                    foreach (DataRow dataRow in table.Rows)
                    {
                        combbox.Items.Add(dataRow[comlumnName].ToString());
                    }
                }
                catch (Exception ex) { }
                saConn.Close();
            }
        }


        private void fileComboBoxMandoub(ComboBox combbox, string source, string tableName)
        {
            //combbox.Visible = true;
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

        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName, bool order)
        {
            //MessageBox.Show("source += "+source);
            combbox.Visible = true;
            //MessageBox.Show(source);
            //MessageBox.Show(Server);
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName;
                if (order) query = "select " + comlumnName + " from " + tableName + " order by " + comlumnName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                foreach (DataRow dataRow in table.Rows)
                {
                    if (!String.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                        combbox.Items.Add(dataRow[comlumnName].ToString().Replace("-", "_"));
                }
                saConn.Close();
            }
            //if (combbox.Items.Count > 0) combbox.SelectedIndex = 0;
        }

        private void fileComboBoxMan(ComboBox combbox, string source, string comlumnName, string tableName)
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

        private void txtReview_MouseHover(object sender, EventArgs e)
        {
            //txtReviewBody();
        }

        private void txtReviewBody()
        {
            fillSamplesCodes(DataSource);
            
            string text = StrSpecPur + LegaceyPreStr;
            if (إجراء_التوكيل.Text == "إقرار بالتنازل") text = StrSpecPur;
            text = SuffReplacements(text, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);


            if (نوع_التوكيل.Text.Contains("ورثة"))
            {
                //if (إجراء_التوكيل.Text == "إقرار بالتنازل")
                txtReview.Text = text.Trim() + "،";

            }
            else
            {
                if (إجراء_التوكيل.Text == "إقرار بالتنازل")
                    txtReview.Text = text.Trim() + "،";
                else if (إجراء_التوكيل.Text == "إجراء_التوكيل")

                    txtReview.Text = " ل" + preffix[صفة_الموكل_off.SelectedIndex, 7] + " ع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 2] + " و" + preffix[صفة_الموكل_off.SelectedIndex, 8] + " مقام" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " " + text.Trim();
                else

                    txtReview.Text = " ل" + preffix[صفة_الموكل_off.SelectedIndex, 7] + " ع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 2] + " و" + preffix[صفة_الموكل_off.SelectedIndex, 8] + " مقام" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " " + text.Trim() + "،";
            }

            txtReview.Text = txtReview.Text.Replace("  ", " ");
            if (txtRev.Text != "")
            {
                checkAutoUpdate.Checked = false;
                txtReview.Text = txtRev.Text;
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var settings = new Settings("57", false, "", DataSource, false, FilespathIn, FilespathOut, FilespathOut, FilespathIn, إجراء_التوكيل.SelectedIndex.ToString() + "-" + نوع_التوكيل.SelectedIndex.ToString());
            settings.Show();
        }

        private void نوع_التوكيل_TextChanged(object sender, EventArgs e)
        {
            for (int item = 0; item < نوع_التوكيل.Items.Count; item++)
            {

                if (نوع_التوكيل.Items[item].ToString().Trim() == نوع_التوكيل.Text.Trim())
                {
                    //MessageBox.Show("not found list - " + نوع_التوكيل.Text.Trim());
                    نوع_التوكيل.SelectedIndex = item;
                    return;
                }
            }


        }

        private void نوع_التوكيل_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (checkColumnName(نوع_التوكيل.Text.Replace(" ", "_").Trim()))
            {
                إجراء_التوكيل.Items.Clear();
                fillSubComboBox(إجراء_التوكيل, DataSource, نوع_التوكيل.Text.Replace(" ", "_"), "TableListCombo", false);
                checkBoxesJob();

                if (نوع_التوكيل.SelectedIndex == 0)
                {
                    generalForms(true);
                    return;
                }
                else
                {
                    generalForms(false);
                }
                if (وجهة_التوكيل.Items.Count > 0) وجهة_التوكيل.SelectedIndex = 0;
                if (وجهة_التوكيل.Text == "إقرار بالتنازل")
                    نوع_المعاملة.Text = "إقرار بالتنازل";
                else نوع_المعاملة.Text = "توكيل";

                if (نوع_التوكيل.SelectedIndex == 6)
                {
                    LegaceyPreStr = " في التركة المذكورة أعلاه";
                }
                else
                {
                    LegaceyPreStr = "";
                }

                return;
            }

        }

        private void fillSubComboBox(ComboBox combbox, string source, string comlumnName, string tableName, bool select)
        {
            //MessageBox.Show("source += "+source);
            combbox.Visible = true;
            //MessageBox.Show(source);
            //MessageBox.Show(Server);
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
                        if (!String.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                        {
                            //MessageBox.Show(dataRow[comlumnName].ToString());
                            combbox.Items.Add(dataRow[comlumnName].ToString());
                        }
                    }
                }
                catch (Exception ex) { }

                saConn.Close();
            }
            if (select && combbox.Items.Count > 0) combbox.SelectedIndex = 0;
        }

        private void newFillComboBox1(ComboBox combbox, string source, string id, string Language)
        {
            combbox.Visible = true;
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();
                string query = "select ColName,ColRight,Lang from TableAddContext";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                foreach (DataRow dataRow in table.Rows)
                {
                    if (dataRow["Lang"].ToString() == Language && dataRow["ColRight"].ToString() != "" && !String.IsNullOrEmpty(dataRow["ColName"].ToString()) && dataRow["ColName"].ToString().Contains("-"))
                    {
                        if (dataRow["ColName"].ToString().Split('-')[1].All(char.IsDigit))
                        {

                            try
                            {
                                if (id == dataRow["ColName"].ToString().Split('-')[1])
                                {
                                    //MessageBox.Show(dataRow["ColName"].ToString().Split('-')[0]);
                                    combbox.Items.Add(dataRow["ColName"].ToString().Split('-')[0]);
                                }
                            }
                            catch (Exception exp)
                            {

                            }
                        }
                    }
                }
                saConn.Close();
            }
            //if (combbox.Items.Count > 0) combbox.SelectedIndex = 0;
        }
        private bool checkColumnName(string colNo)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableListCombo", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(dataRow["COLUMN_NAME"].ToString()))
                {
                    if (dataRow["COLUMN_NAME"].ToString() == colNo)
                        return true;
                }
            }
            return false;
        }

        private void إجراء_التوكيل_SelectedIndexChanged(object sender, EventArgs e)
        {
            Pro();
            //if (نوع_التوكيل.Text == "شهادة ميلاد" && Vitext1.Text != "")
            //{
            //    BirthName = Vitext1.Text.Split('_');
            //    BirthPlace = Vitext2.Text.Split('_');
            //    BirthDate = Vitext3.Text.Split('_');
            //    BirthMother = Vitext4.Text.Split('_');
            //    //MessageBox.Show(Vitext1.Text);

            //    ButtonInfoIndex = Vitext1.Text.Split('_').Length;

            //    //MessageBox.Show(Vitext1.Text +" __ "+LibtnAdd1.Text);
            //    Vitext1.Text = Vitext2.Text = Vitext3.Text = Vitext4.Text = "";

            //}
        }

        private void Pro() {
            reversTextReview();
            reversTextRight();

            checkBoxesJob();
            fillInfo(PanelItemsboxes, false);
        }
            private void checkBoxesJob()
        {

            resetBoxes(true);
            
//            MessageBox.Show(قائمة_الحقوق.Text);
            صفة_الموكل_off.SelectedIndex = Appcases(جنس_الموكَّل, addAuthticIndex);
            صفة_مقدم_الطلب_off.SelectedIndex = Appcases(النوع, addNameIndex);

            if (قائمة_الحقوق.Text == "")
            {
                savedRights.Checked = false;
                صفة_الموكل_off.Enabled = true;
                bool genForm = false;
                if (نوع_التوكيل.Text == "توكيل بصيغة غير مدرجة") genForm = true;

                flllPanelItemsboxes("ColName", إجراء_التوكيل.Text + "-" + نوع_التوكيل.SelectedIndex.ToString());
                if (starText != "")
                {
                    starTectRightJob();
                }
                else
                {
                    PopulateCheckBoxes(genForm, ColRight.Replace(" ", "_"), "TableAuthRights", DataSource, صفة_مقدم_الطلب_off.SelectedIndex, true);
                }
                //autoCompleteTextBox(txtAddRight, DataSource, "قائمة_الحقوق_الكاملة", "TableAuthRight");
            }
            else
            {
                قائمة_الحقوق.Text = removeSpace(قائمة_الحقوق.Text, true);
                //MessageBox.Show(الحقوق_الممنوحة.Text);
                //صفة_الموكل_off.Enabled = false;
                savedRights.Checked = true;
                flllPanelItemsboxes("ColName", إجراء_التوكيل.Text + "-" + نوع_التوكيل.SelectedIndex.ToString());
                //PopulateCheckBoxes(الحقوق_الممنوحة.Text.Split('،'));
                //قائمة_الحقوق.Text = الحقوق_الممنوحة.Text;
                //autoCompleteTextBox(txtAddRight, DataSource, "قائمة_الحقوق_الكاملة", "TableAuthRight");
            }

            

            IBAN = ibanText(ColRight.Replace(" ", "_"), "TableAuthRights", DataSource);

            //MessageBox.Show("IBAN " + IBAN);
            if (txtRev.Text != "")
            {
                checkAutoUpdate.Checked = false;
                txtReview.Text = txtRev.Text;
            }
        }

        private void starTectRightJob()
        {
            picStarRight.Visible = true;
            if (txtRightList == null) return;
            if (starText == "0") starText = "1";
            //MessageBox.Show(txtRightList.Length.ToString() +" - "+ starText);
            try
            {
                قائمة_الحقوق.Text = txtRightList[Convert.ToInt32(starText) - 1];
            }
            catch (Exception ex) { return; }
            string[] textRights = قائمة_الحقوق.Text.Split('،');

            foreach (string text in textRights)
            {
                if (text.Contains("الحق في توكيل الغير"))
                    authother = text;

                if (text.Contains("ويعتبر التوكيل الصادر"))
                    removeAuthother = text;
                if (text.Contains("لمن يشهد والله"))
                    lastSentence = text;
            }

            if (!توكيل_الغير.Checked && authother != "")
            {
                قائمة_الحقوق.Text = قائمة_الحقوق.Text.Replace(authother, "") + " ";
            }

            if (!إلغاء_توكيل_سابق.Checked && removeAuthother != "")
            {
                قائمة_الحقوق.Text = قائمة_الحقوق.Text.Replace(removeAuthother, "") + " ";
            }
        }

        private void إجراء_التوكيل_TextChanged(object sender, EventArgs e)
        {
            if (إجراء_التوكيل.Text != "إختر الإجراء")
            {
                for (int item = 0; item < إجراء_التوكيل.Items.Count; item++)
                {
                    if (إجراء_التوكيل.Items[item].ToString() == إجراء_التوكيل.Text)
                        إجراء_التوكيل.SelectedIndex = item;
                }
                //MessageBox.Show(إجراء_التوكيل.SelectedIndex.ToString());
            }
            Pro();
        }

        private void VitxtDate1VD_TextChanged(object sender, EventArgs e)
        {

            //if (VitxtDate1.Text.Length == 10)
            //{
            //    int month = Convert.ToInt32(SpecificDigit(VitxtDate1.Text,4, 5));
            //    if (month > 12)
            //    {
            //        MessageBox.Show("الشهر يحب أن يكون أقل من 12");
            //        //VitxtDate1.Text = "";
            //        VitxtDate1.Text = SpecificDigit(VitxtDate1.Text, 3, 10);
            //        return;
            //    }
            //}

            if (VitxtDate1.Text.Length == 11)
            {
                VitxtDate1.Text = lastInput2; return;
            }
            if (VitxtDate1.Text.Length == 10) return;
            if (VitxtDate1.Text.Length == 4) VitxtDate1.Text = "-" + VitxtDate1.Text;
            else if (VitxtDate1.Text.Length == 7) VitxtDate1.Text = "-" + VitxtDate1.Text;
            lastInput2 = VitxtDate1.Text;
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

        private void VitxtDate2_TextChanged(object sender, EventArgs e)
        {
            if (VitxtDate2.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(VitxtDate2.Text, 1, 2));
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //VitxtDate2.Text = "";
                    VitxtDate2.Text = SpecificDigit(VitxtDate2.Text, 3, 10);
                    return;
                }
            }

            if (VitxtDate2.Text.Length == 11)
            {
                VitxtDate2.Text = lastInput2; return;
            }
            if (VitxtDate2.Text.Length == 10) return;
            if (VitxtDate2.Text.Length == 4) VitxtDate2.Text = "-" + VitxtDate2.Text;
            else if (VitxtDate2.Text.Length == 7) VitxtDate2.Text = "-" + VitxtDate2.Text;
            lastInput2 = VitxtDate2.Text;
        }

        private void VitxtDate3_TextChanged(object sender, EventArgs e)
        {
            if (VitxtDate3.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(VitxtDate3.Text, 1, 2));
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //VitxtDate3.Text = "";
                    VitxtDate3.Text = SpecificDigit(VitxtDate3.Text, 3, 10);
                    return;
                }
            }

            if (VitxtDate3.Text.Length == 11)
            {
                VitxtDate3.Text = lastInput3; return;
            }
            if (VitxtDate3.Text.Length == 10) return;
            if (VitxtDate3.Text.Length == 4) VitxtDate3.Text = "-" + VitxtDate3.Text;
            else if (VitxtDate3.Text.Length == 7) VitxtDate3.Text = "-" + VitxtDate3.Text;
            lastInput3 = VitxtDate3.Text;
        }

        private void VitxtDate4_TextChanged(object sender, EventArgs e)
        {
            if (VitxtDate4.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(VitxtDate4.Text, 1, 2));
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //VitxtDate4.Text = "";
                    VitxtDate4.Text = SpecificDigit(VitxtDate4.Text, 3, 10);
                    return;
                }
            }

            if (VitxtDate4.Text.Length == 11)
            {
                VitxtDate4.Text = lastInput4; return;
            }
            if (VitxtDate4.Text.Length == 10) return;
            if (VitxtDate4.Text.Length == 4) VitxtDate4.Text = "-" + VitxtDate4.Text;
            else if (VitxtDate4.Text.Length == 7) VitxtDate4.Text = "-" + VitxtDate4.Text;
            lastInput4 = VitxtDate4.Text;
        }

        private void VitxtDate5_TextChanged(object sender, EventArgs e)
        {
            if (VitxtDate5.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(VitxtDate5.Text, 1, 2));
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //VitxtDate5.Text = "";
                    VitxtDate5.Text = SpecificDigit(VitxtDate5.Text, 3, 10);
                    return;
                }
            }

            if (VitxtDate5.Text.Length == 11)
            {
                VitxtDate5.Text = lastInput3; return;
            }
            if (VitxtDate5.Text.Length == 10) return;
            if (VitxtDate5.Text.Length == 4) VitxtDate5.Text = "-" + VitxtDate5.Text;
            else if (VitxtDate5.Text.Length == 7) VitxtDate5.Text = "-" + VitxtDate5.Text;
            lastInput3 = VitxtDate5.Text;
        }

        private void checkAutoUpdate_CheckedChanged(object sender, EventArgs e)
        {
            if (checkAutoUpdate.Checked)
            {
                checkAutoUpdate.Text = "تحديث النص";

            }
            else
            {
                boxesPreparations();                
                oldText = txtReview.Text;
                checkAutoUpdate.Text = "إيقاف التحديث";

            }
        }

        private bool specialChar(string text)
        {
            string str = "#*@&%^$";
            Char[] ca = text.ToCharArray();
            foreach (Char c in ca)
            {
                if (str.Contains(c))
                {
                    //MessageBox.Show("char " + c.ToString());
                    return true;
                }
            }
            return false;
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (checkAutoUpdate.Checked && currentPanelIndex > 0)
            {
                txtReviewBody();

                //التواكيل المتعلقة بتعويض حوادث السير
                if (IBAN == "") return;
                if (إجراء_التوكيل.Text == "استلام تأمين")
                {
                    string iban = IBAN;
                    iban = SuffReplacements(iban, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);

                    foreach (Control control in panelAuthOptions.Controls)
                    {
                        if (control is CheckBox)
                        {
                            if (((CheckBox)control).Text.Contains("إيداع") && ((CheckBox)control).Tag.ToString() == "valid")
                            {
                                ((CheckBox)control).Text = iban;
                                return;
                            }
                        }
                    }
                }
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (addNameIndex > 1) btnPanelapp.Enabled = true;
            else btnPanelapp.Enabled = false;

            if (addAuthticIndex > 1) btnPanelAuthPers.Enabled = true;
            else btnPanelAuthPers.Enabled = false;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            int count = getTodaDocxPdf();
            if (count < 4)
            {
                var selectedOption = MessageBox.Show("لديك عدد " + (5 - count).ToString() + " عرض متوفر", "طباعة ملف وورد على اي حال", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (selectedOption == DialogResult.Yes)
                {
                    FinalPrint("docx");
                }
                else if (selectedOption == DialogResult.No)
                {
                    FinalPrint("pdf");
                }


            }
            else
            {
                MessageBox.Show("تم استنفاد طلبات التعديل على المعاملات سيتم الطباعة بصيغة pdf");
                FinalPrint("pdf");
            }

            FinalPrint("docx");
            
        }

        private string getDocxPdf()
        {
            string doc = "";
            string query = "select الإجراء_الأخير from TableAuth where رقم_التوكيل = N'" + رقم_التوكيل.Text + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                doc = dataRow["الإجراء_الأخير"].ToString();
            }
            return doc;
        }

        private int getTodaDocxPdf()
        {
            string query = "select الإجراء_الأخير from TableAuth where التاريخ_الميلادي = N'" + GreDate + "' and الإجراء_الأخير = N'docx'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            return dtbl.Rows.Count;
        }

        private void setDocxPdf(string doc)
        {
            string query = "update TableAuth set الإجراء_الأخير = N'" + doc + "' where رقم_التوكيل = N'" + رقم_التوكيل.Text + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void FinalPrint(string doc) {

            if (getDocxPdf() != "docx")
                setDocxPdf(doc);

            if (!AuthenticatOther()) return;
            if (طريقة_الإجراء.Checked)
                SuffConvertments(التوثيق, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);
            else
            {
                if (نوع_الموقع.Checked)
                    SuffConvertments(التوثيق, 0, صفة_الموكل_off.SelectedIndex);
                else SuffConvertments(التوثيق, 1, صفة_الموكل_off.SelectedIndex);
            }

            if (وجهة_التوكيل.Text == "إختر نوع الإجراء")
                وجهة_التوكيل.SelectedIndex = 0;

            commentInfo();
            التاريخ_الميلادي.Text = GreDate;
            التاريخ_الهجري.Text = HijriDate;

            save2DataBase(panelapplicationInfo);



            footers();
            التاريخ_الميلادي.Text = GreDate;
            التاريخ_الهجري.Text = HijriDate;

            //MessageBox.Show(قائمة_الحقوق.Text);

            fillDocFileInfo(panelapplicationInfo);
            fillDocFileAppInfo(Panelapp);
            fillDocFileInfo(panelAuthRights);
            fillDocFileInfo(panelAuthOptions);
            fillDocFileInfo(finalPanel);
            fillDocFileInfo(flowLayoutPanel2);
            notFiled = false;
            حالة_الارشفة.Text = "غير مؤرشف";


            if (checkAutoUpdate.Checked) txtRev.Text = "";
            else txtRev.Text = txtReview.Text;
            if (!save2DataBase(finalPanel)) return;

            fillPrintDocx(doc);

            addarchives();
            if (!وجهة_التوكيل.Text.Contains("الخرطوم"))
                CreateMessageWord(مقدم_الطلب.Text.Replace("_", " و"), وجهة_التوكيل.Text, رقم_التوكيل.Text, "توكيلا", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 17], التاريخ_الميلادي_off.Text, HijriDate, موقع_التوكيل.Text);
            fileUpload(رقم_التوكيل.Text, "missed");
            relatedProUpdate();
            this.Close();
        }
            private void authJob()
        {
            string auth = " المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15] + " المذكور" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15] + " أعلاه قد حضر" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " ووقع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " بتوقيع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4] + " على هذا التوكيل في حضور الشاهدين المشار إليهما أعلاه وذلك بعد تلاوته علي" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4] + " وبعد أن فهم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " مضمونه ومحتواه";
            if (!طريقة_الطلب.Checked)
                auth = " المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15] + " المذكور" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15] + " أعلاه قد وقع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " بتوقيع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4] + " على هذا التوكيل في حضور الشاهدين المشار إليهما أعلاه وذلك بعد تلاوته علي" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4] + " وبعد أن فهم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " مضمونه ومحتواه";
            if (طريقة_الطلب.Checked)
            {
                التوثيق.Text = "قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن" + auth + "، صدر تحت توقيعي وختم القنصلية العامة";
                if (طريقة_الإجراء.Checked)
                    التوثيق.Text = AuthTitle + " بالقنصلية العامة لجمهورية السودان بجدة، بأن" + auth + "، صدر تحت توقيعي وختم القنصلية العامة";
                else
                {
                    auth = " بأن المواطن" + preffix[onBehalfIndex, 5] + " /" + اسم_الموكل_بالتوقيع.Text + " قد حضر" + preffix[onBehalfIndex, 3] + " ووقع" + preffix[onBehalfIndex, 3] + " بتوقيع" + preffix[onBehalfIndex, 4] + " على هذا التوكيل في حضور الشهود المشار إليهما أعلاه بعد تلاوته علي" + preffix[onBehalfIndex, 4] + " وبعد أن فهم" + preffix[onBehalfIndex, 3] + " مضمونه ومحتواه، وذلك بناءً على حق توكي الغير الممنوح ل"+ preffix[onBehalfIndex, 4] + " والمنصوص عليه بموجب التوكيل الصادر عن " + جهة_إصدار_الوكالة.Text + " بالرقم " + رقم_الوكالة.Text + " بتاريخ " + تاريخ_إصدار_الوكالة.Text;
                    توقيع_مقدم_الطلب.Text = اسم_الموكل_بالتوقيع.Text;
                    التوثيق.Text = auth + "، صدر تحت توقيعي وختم القنصلية العامة";
                }
            }
            else
                التوثيق.Text = auth + "، صدر تحت توقيعي وختم القنصلية العامة";

        }
        private void CreateMessageWord(string ApplicantName, string EmbassySource, string IqrarNo, string MessageType, string ApplicantSex, string GregorianDate, string HijriDate, string ViseConsul)
        {
            string ActiveCopy;
            string ReportName = DateTime.Now.ToString("mmss");
            string routeDoc = FilespathIn + @"\MessageCap.docx";
            loadMessageNo();
            ActiveCopy = FilespathOut + "Message" + ApplicantName + ReportName + ".docx";
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
                string noID = "ق س ج/80/01/" + (MessageDocNo + 1).ToString();
                BookMApplicantName.Text = ApplicantName;
                BookcapitalMessage.Text = EmbassySource;
                BookMassageNo.Text = noID;
                BookMassageIqrarNo.Text = IqrarNo;
                BookApliSex.Text = ApplicantSex;
                BookDateGre.Text = GregorianDate;
                BookGregorDate2.Text = GregorianDate;
                BookHijriDate.Text = HijriDate;
                BookViseConsul1.Text = ViseConsul;
                BookMassageTitle.Text = MessageType;

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

                oBDoc2.Activate();
                oBDoc2.Save();
                //addMessageArch(ActiveCopy, noID);
                oBMicroWord2.Visible = true;
                NewMessageNo();
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
        private void generalForms(bool genType)
        {

            if (genType)
            {
                checkAutoUpdate.Checked = checkAutoUpdate.Visible = PanelItemsboxes.Visible = false;
                //إجراء_التوكيل.Text = "صيغة عامة";
                //txtReview.Size = new System.Drawing.Size(1290, 182);

                txtReviewBody();
                //قائمة_الحقوق.Visible = true;
                //قائمة_الحقوق.Size = new System.Drawing.Size(1290, 1500);
                //panelAuthRights.AutoScroll = true;
            }
            else
            {
                إجراء_التوكيل.Enabled = checkAutoUpdate.Checked = checkAutoUpdate.Visible = PanelItemsboxes.Visible = panelAuthOptions.Visible = true;                //txtAddRight.Visible = btnAddRight.Visible =
                //txtReview.Size = new System.Drawing.Size(1176, 57);
                //قائمة_الحقوق.Visible = false;
                //قائمة_الحقوق.Size = new System.Drawing.Size(226, 46);
                //panelAuthRights.AutoScroll = false;
            }
        }
        private void getMaxRange(string dataSource)
        {
            SqlConnection Con = new SqlConnection(dataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select allowedEdit  from TableSettings where ID=1", Con);

            try
            {
                if (Con.State == ConnectionState.Closed)
                    Con.Open();
                sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
                var reader = sqlCmd1.ExecuteReader();

                if (reader.Read())
                {
                    //allowedEdit.Text = reader["allowedEdit"].ToString();
                }
            }
            catch (Exception ex)
            {
                //allowedEdit.Text = "0";
                Con.Close();
            }
        }
        private void getDefualtRight(string dataSource)
        {
            SqlConnection Con = new SqlConnection(dataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select توكيل_بصيغة_غير_مدرجة_عام from TableAuthRightStarText where ID=1", Con);

            try
            {
                if (Con.State == ConnectionState.Closed)
                    Con.Open();                
                var reader = sqlCmd1.ExecuteReader();

                if (reader.Read())
                {
                    قائمة_الحقوق.Text = reader["توكيل_بصيغة_غير_مدرجة_عام"].ToString();
                    قائمة_الحقوق.Text = SuffReplacements(قائمة_الحقوق.Text, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);
                    starTectRightJob();
                    قائمة_الحقوق.Text = removeSpace(قائمة_الحقوق.Text, true);
                }
                
            }
            catch (Exception ex)
            {
                قائمة_الحقوق.Text = "";
                //MessageBox.Show(intID.ToString());
                Con.Close();
            }
        }

        private void addarchives()
        {
            Console.WriteLine(رقم_التوكيل.Text);

            if (checkArchives(رقم_التوكيل.Text)) return;// else MessageBox.Show("not found");

            colIDs[0] = مقدم_الطلب.Text.Split('_')[0];
            colIDs[1] = archState;
            colIDs[2] = طريقة_الطلب.Text;
            colIDs[3] = التاريخ_الميلادي.Text;
            colIDs[4] = رقم_التوكيل.Text;
            colIDs[5] = intID.ToString();


            colIDs[6] = EmpName;

            colIDs[7] = اسم_المندوب.Text;


            string[] allArchList = getColList("archives");
            string strList = "";
            for (int i = 0; i < allArchList.Length; i++)
            {
                Console.WriteLine(i.ToString() + " - " + allArchList[i]);
                if (i == 0) strList = "@" + allArchList[0];
                else strList = strList + ",@" + allArchList[i];
            }
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();

            SqlCommand sqlCommand = new SqlCommand("insert into archives (" + strList.Replace("@", "") + ") values (" + strList + ");SELECT @@IDENTITY as lastid", sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            for (int i = 0; i < allArchList.Length; i++)
            {

                sqlCommand.Parameters.AddWithValue("@" + allArchList[i], colIDs[i]);
            }
            Console.WriteLine("insert into archives (" + strList.Replace("@", "") + ") values (" + strList + ")");
            //MessageBox.Show("lastid");
            var reader = sqlCommand.ExecuteReader();
            if (reader.Read())
            {
                //MessageBox.Show(reader["lastid"].ToString());
            }
            try
            {


            }
            catch (Exception ex) { MessageBox.Show("insert into archives (" + strList.Replace("@", "") + ") values (" + strList + ")"); }
        }

        private void addNewAppNameInfo()
        {

            string query = "insert into TableGenNames ([الاسم], رقم_الهوية,تاريخ_الميلاد,المهنة,النوع,نوع_الهوية,مكان_الإصدار) values (@col1,@col2,@col3,@col4,@col5,@col6,@col7) ;SELECT @@IDENTITY as lastid";
            for (int x = 0; x < addNameIndex; x++)
            {
                string id = checkExist(مقدم_الطلب.Text.Split('_')[x]);
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
                sqlCommand.Parameters.AddWithValue("@col1", مقدم_الطلب.Text.Split('_')[x]);
                sqlCommand.Parameters.AddWithValue("@col2", رقم_الهوية.Text.Split('_')[x]);
                sqlCommand.Parameters.AddWithValue("@col3", تاريخ_الميلاد.Text.Split('_')[x]);
                sqlCommand.Parameters.AddWithValue("@col4", المهنة.Text.Split('_')[x]);
                sqlCommand.Parameters.AddWithValue("@col5", النوع.Text.Split('_')[x]);
                sqlCommand.Parameters.AddWithValue("@col6", نوع_الهوية.Text.Split('_')[x]);
                sqlCommand.Parameters.AddWithValue("@col7", مكان_الإصدار.Text.Split('_')[x]);

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
        }

        private void addNewAuthNameInfo()
        {

            string query = "insert into TableGenNames ([الاسم], رقم_الهوية,النوع) values (@col1,@col2,@col5) ;SELECT @@IDENTITY as lastid";
            for (int x = 0; x < addAuthticIndex; x++)
            {
                string id = checkExist(الموكَّل.Text.Split('_')[x]);
                if (id != "0")
                {
                    query = "update TableGenNames set [الاسم] =  @col1,[رقم_الهوية] = @col2,النوع = @col5 where ID = " + id;
                    //MessageBox.Show(query);
                }
                SqlConnection sqlConnection = new SqlConnection(DataSource);
                if (sqlConnection.State == ConnectionState.Closed)
                    sqlConnection.Open();

                SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
                sqlCommand.CommandType = CommandType.Text;
                sqlCommand.Parameters.AddWithValue("@col1", الموكَّل.Text.Split('_')[x]);
                sqlCommand.Parameters.AddWithValue("@col2", هوية_الموكل.Text.Split('_')[x]);
                sqlCommand.Parameters.AddWithValue("@col5", جنس_الموكَّل.Text.Split('_')[x]);

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
                    MessageBox.Show(query);
                }
            }
        }
        private bool checkArchives(string name)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT ID FROM archives where docID=N'" + name + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count > 0) { return true; }
            else return false;
        }
        private bool AuthenticatOther()
        {
            bool moveOn = true;
            if (طريقة_الإجراء.Checked) return true;
            
            if (AuthenticName != مقدم_الطلب.Text && AuthenticName != "")
            {
                MessageBox.Show("اسم صاحب الوكالة المشار إليها غير متطابق مع مقدم الطلب بالوكالة الحالية");
                moveOn = false;
            }

            if (archStat != "مؤرشف نهائي")
            {
                MessageBox.Show("المكاتبة المرجعية غير مؤرشفة، ولا يمكن إجراء توكيل بناء عليها");
                moveOn = false;
            }
            
            if (removedStat != "")
            {
                MessageBox.Show("المكاتبة المرجعية ملغية، ولا يمكن إجراء توكيل بناء عليها");
                moveOn = false;
            }

            if (autheticatingOthes == "0")
            {
                MessageBox.Show("لم يتم منح " + اسم_الموكل_بالتوقيع.Text + " حق توكيل غيره" +" بالوكالة المشار إليها");
                moveOn = false;
            }
            if (!moveOn)
            {
                var selectedOption = MessageBox.Show("يجب على مقدم الطلب الحضور والتوقيع بنفسه على المكاتبة", "إصدار المكاتبة بالحضور الشخصي؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (selectedOption == DialogResult.Yes)
                {
                    طريقة_الإجراء.Checked = true;
                }
                else if (selectedOption == DialogResult.No)
                {
                    return false;
                }
            }
            return true;
        }
            private void commentInfo()
        {
            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text == "")
                تعليق.Text = "قام  " + اسم_الموظف.Text + " بإدخال البيانات " + Environment.NewLine + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine;

            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text != "")
                تعليق.Text = "قام  " + اسم_الموظف.Text + " ببعض التعديلات " + Environment.NewLine + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine + التعليقات_السابقة_Off.Text;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text == "")
                تعليق.Text = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + "قام  " + اسم_الموظف.Text + " ببعض التعديلات " + Environment.NewLine + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text != "")
                تعليق.Text = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + "قام  " + اسم_الموظف.Text + " ببعض التعديلات " + Environment.NewLine + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine + "*" + التعليقات_السابقة_Off.Text.Trim();
        }

        private void وجهة_التوكيل_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (وجهة_التوكيل.SelectedIndex == 0)
            {
                if (إجراء_التوكيل.Text.Contains("تنازل"))
                    مدة_الاعتماد.Text = "لا يعتمد هذا الاقرار ما لم يتم توثيقه خلال عام من تاريخ إصدارة من وزارة خارجية جمهورية السودان";
                else
                    مدة_الاعتماد.Text = "لا يعتمد هذا التوكيل ما لم يتم توثيقه خلال عام من تاريخ إصدارة من وزارة خارجية جمهورية السودان";
            }
            else مدة_الاعتماد.Text = "";
        }
        private void footers()
        {
            if (وجهة_التوكيل.SelectedIndex == 0)
            {
                if (إجراء_التوكيل.Text.Contains("إقرار"))
                {
                    نوع_المعاملة.Text = إجراء_التوكيل.Text;
                    مدة_الاعتماد.Text = "لا يعتمد هذا الاقرار ما لم يتم توثيقه خلال عام من تاريخ إصدارة من وزارة خارجية جمهورية السودان";
                }
                else
                {
                    نوع_المعاملة.Text = "توكيل";
                    مدة_الاعتماد.Text = "لا يعتمد هذا التوكيل ما لم يتم توثيقه خلال عام من تاريخ إصدارة من وزارة خارجية جمهورية السودان";
                }
            }
            else مدة_الاعتماد.Text = "";
        }

        private void btnFile1_Click(object sender, EventArgs e)
        {
            btnFile1.Enabled = false;
            FillDatafromGenArch("data1", intID.ToString(), "TableAuth");
            btnFile1.Enabled = true;
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

        


        private void btnFile2_Click(object sender, EventArgs e)
        {
            btnFile2.Enabled = false;
            FillDatafromGenArch("data2", intID.ToString(), "TableAuth");
            btnFile2.Enabled = true;
        }

        private void btnFile3_Click(object sender, EventArgs e)
        {
            btnFile3.Enabled = false;
            FillDatafromGenArch("data3", intID.ToString(), "TableAuth");
            btnFile3.Enabled = true;
        }

        private void ListSearch_TextChanged(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '" + ListSearch.Text + "%'";
            dataGridView1.DataSource = bs;
            ColorFulGrid9();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            int ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            deleteRowsData(رقم_التوكيل.Text, "رقم_التوكيل", "TableAuth", DataSource);
            deleteRowsData(رقم_التوكيل.Text, "docID", "archives", DataSource);
            deleteRowsData(رقم_التوكيل.Text, "رقم_معاملة_القسم", "TableGeneralArch", DataSource.Replace("AhwalDataBase", "ArchFilesDB"));

            FillDataGridView(dataSource);
        }
        private void deleteRowsData(string docID, string docIDName, string v2, string source)
        {
            string query;
            SqlConnection Con = new SqlConnection(source);
            query = "DELETE FROM " + v2 + " where " + docIDName + " = @" + docIDName;
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@" + docIDName, docID);
            sqlCmd.ExecuteNonQuery();
            Con.Close();

        }


        private void timer3_Tick(object sender, EventArgs e)
        {
            if (labDescribed.Text != "الملخص")
                timer3.Enabled = false;
            ColorFulGrid9();
        }

        private void LibtnAdd1_Click(object sender, EventArgs e)
        {
            //BirthName[birthindex] = Vitext1.Text;
            //BirthPlace[birthindex] = Vitext2.Text;
            //BirthDate[birthindex] = Vitext3.Text;
            //BirthMother[birthindex] = Vitext4.Text;

            //Vitext1.Text = Vitext2.Text = Vitext3.Text = Vitext4.Text = "";
            //if (birthindex == 0) specialDataSum = BirthName[birthindex] + "_" + BirthPlace[birthindex] + "_" + BirthDate[birthindex] + "_" + BirthMother[birthindex] + "_" + BirthDecs[birthindex];
            //else specialDataSum = specialDataSum + "*" + BirthName[birthindex] + "_" + BirthPlace[birthindex] + "_" + BirthDate[birthindex] + "_" + BirthMother[birthindex] + "_" + BirthDecs[birthindex];

            //if (birthindex == 0 && Vicombo2.SelectedIndex > 0)
            //{

            //    if (Vicombo2.SelectedIndex == 1)
            //    {

            //        Mentioned = "لابني";

            //    }
            //    else if (Vicombo2.SelectedIndex == 2)
            //    {

            //        Mentioned = "لابنتي";

            //    }
            //}
            //else if (birthindex == 1 && Vicombo2.SelectedIndex > 0)
            //{
            //    if (Vicombo2.SelectedIndex == 1 && Mentioned == "لابني")
            //    {
            //        Mentioned = "لابنيَّ";

            //    }
            //    else if (Vicombo2.SelectedIndex == 2 && Mentioned == "لابنتي")
            //    {
            //        Mentioned = "لابنتيَّ";

            //    }
            //    else
            //    {
            //        Mentioned = "لابنائي";

            //    }

            //}
            //else if (birthindex >= 2 && Vicombo2.SelectedIndex > 0)
            //{
            //    if (Vicombo2.SelectedIndex == 2 && Mentioned == "لابنتيَّ")
            //    {
            //        Mentioned = "لبناتي";
            //    }
            //    else
            //    {
            //        Mentioned = "لأبنائي";
            //    }

            //}

            //BirthDecs[birthindex] = Mentioned;

            //birthindex++;
            //idShow = birthindex;
            addButtonInfo(Vitext1.Text, Vitext2.Text, Vitext3.Text, Vitext4.Text, Vitext5.Text);
            Vitext1.Text = Vitext2.Text = Vitext3.Text = Vitext4.Text = Vitext5.Text = "";

            LibtnAdd1.Text = "اضافة (" + idShow.ToString() + "/" + ButtonInfoIndex.ToString() + ")" + "   ";

            //MessageBox.Show(birthindex.ToString());
        }

        private void txtfinal_MouseClick(object sender, MouseEventArgs e)
        {
            flowLayoutPanel2.Size = new System.Drawing.Size(940, 178);
        }

        
        



        private void updateerrorList(string text, string col)
        {
            string colName = إجراء_التوكيل.Text + "-" + نوع_التوكيل.SelectedIndex.ToString();
            string query = "update TableAddContext set " + col + " = N'" + text + "_" + EmpName + "' where altColName = N'" + نوع_التوكيل.Text + "' and altSubColName = N'" + إجراء_التوكيل.Text + "'";
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;

            sqlCommand.ExecuteNonQuery();
        }

        private void btnAddRight_Click(object sender, EventArgs e)
        {
            //if (txtAddRight.Text == "") return;
            //if (btnAddRight.Text == "إضافة")
            //{
            //    CheckBox chk = new CheckBox();
            //    chk.TabIndex = Nobox;
            //    chk.Width = 80;
            //    chk.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            //    chk.Width = panelAuthOptions.Width - 130;
            //    chk.Height = 33;
            //    chk.Tag = "valid";
            //    chk.CheckState = CheckState.Checked;
            //    chk.Location = new System.Drawing.Point(60, 3 + Nobox * 37);
            //    chk.Name = "checkBox_" + Nobox.ToString();
            //    chk.Text = txtAddRight.Text;
            //    txtAddRight.Clear();
            //    statistic[Nobox] = 1;
            //    times[Nobox] = 1;
            //    panelAuthOptions.Controls.Add(chk);

            //    PictureBox picboxedit = new PictureBox();
            //    picboxedit.Image = global::PersAhwal.Properties.Resources.edit;
            //    picboxedit.Location = new System.Drawing.Point(55, Nobox * 37);
            //    picboxedit.Name = "Edit";
            //    picboxedit.Size = new System.Drawing.Size(24, 26);
            //    picboxedit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            //    picboxedit.TabIndex = 175 + Nobox;
            //    picboxedit.TabStop = false;
            //    picboxedit.Click += new System.EventHandler(this.pictureBoxedit_Click);
            //    panelAuthOptions.Controls.Add(picboxedit);

            //    PictureBox picboxup = new PictureBox();
            //    picboxup.Image = global::PersAhwal.Properties.Resources.arrowup;
            //    picboxup.Location = new System.Drawing.Point(76, Nobox * 37);
            //    picboxup.Name = "Up";
            //    picboxup.Size = new System.Drawing.Size(24, 26);
            //    picboxup.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            //    picboxup.TabIndex = 176 + Nobox;
            //    picboxup.TabStop = false;
            //    picboxup.Visible = false;
            //    picboxup.Click += new System.EventHandler(this.pictureBoxup_Click);
            //    panelAuthOptions.Controls.Add(picboxup);

            //    PictureBox picboxdown = new PictureBox();
            //    picboxdown.Image = global::PersAhwal.Properties.Resources.arrowdown;
            //    picboxdown.Location = new System.Drawing.Point(45, Nobox * 37);
            //    picboxdown.Size = new System.Drawing.Size(24, 26);
            //    picboxdown.Name = "Down";
            //    picboxdown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            //    picboxdown.TabIndex = 177 + Nobox; ;
            //    picboxdown.TabStop = false;
            //    picboxdown.Visible = false;
            //    picboxdown.Click += new System.EventHandler(this.pictureBoxdown_Click);

            //    panelAuthOptions.Controls.Add(picboxdown);
            //    Nobox++;
            //    for (int swap = 0; swap < 2; swap++)

            //    {
            //        SwapText(Nobox - swap);
            //        ShowArrows(Nobox, swap);
            //    }

            //}
            //else if (btnAddRight.Text == "تعديل")
            //{
            //    foreach (Control control in panelAuthOptions.Controls)
            //    {
            //        if (control is CheckBox)
            //        {
            //            //MessageBox.Show(control.Name +"_" +control.Tag);
            //            if (((CheckBox)control).Name == controlName && control.Tag.ToString() == "valid")
            //            {
            //                ((CheckBox)control).Text = txtAddRight.Text;
            //                txtAddRight.Text = "";
            //                btnAddRight.Text = "إضافة";

            //                editsMade[1] = true;
            //                error4.Checked = true;
            //                error4.Enabled = false;

            //            }
            //        }
            //    }

            //foreach (Control control in panelAuthOptions.Controls)
            //{
            //    if (control is CheckBox && !txtAddRight.Text.Contains("نص ملغي") && txtAddRight.Text != "")
            //    {
            //        if (((CheckBox)control).TabIndex == LastTabIndex)
            //        {
            //            control.Text = txtAddRight.Text;
            //            btnAddRight.Text = "إضافة";

            //            editsMade[1] = true;
            //            error4.Checked = true;
            //            error4.Enabled = false;
            //            txtAddRight.Text = "";
            //            return;
            //        }
            //    }
            //}
            //}
            //editsMade[1] = true;
            //error4.Checked = true;
            //error4.Enabled = false;
        }

        private void editIBan(string iban)
        {
            //foreach (Control control in panelAuthOptions.Controls)
            //{
            //    if (control is CheckBox)
            //    {
            //        //MessageBox.Show(control.Name +"_" +control.Tag);
            //        if (((CheckBox)control).Name == "checkBox_"+iban.Split('_')[1] && control.Tag.ToString() == "valid")
            //        {
            //            ((CheckBox)control).Text = txtAddRight.Text;
            //            txtAddRight.Text = "";
            //            btnAddRight.Text = "إضافة";

            //            editsMade[1] = true;
            //            error4.Checked = true;
            //            error4.Enabled = false;

            //        }
            //    }
            //}
        }

        private void ShowArrows(int tabindex, int indexMinus)
        {
            foreach (Control control in panelAuthOptions.Controls)
            {

                if (control is PictureBox)
                {

                    if (((PictureBox)control).Name == "Down" && ((PictureBox)control).TabIndex == 177 + tabindex - 3)
                    {
                        ((PictureBox)control).Visible = true;
                    }
                    if (((PictureBox)control).Name == "Up" && ((PictureBox)control).TabIndex == 176 + tabindex - 2 - indexMinus)
                    {
                        ((PictureBox)control).Visible = true;
                    }
                }
            }
        }

        private void SwapText(int tabindex)
        {
            string st = "", nd = "";
            bool statest = false, statend = false;


            foreach (Control control in panelAuthOptions.Controls)
            {

                if (control is CheckBox)
                {

                    if (((CheckBox)control).TabIndex == tabindex - 1)
                    {
                        st = ((CheckBox)control).Text;
                        if (((CheckBox)control).CheckState == CheckState.Checked) statest = true;
                        else statest = false;

                    }
                    if (((CheckBox)control).TabIndex == tabindex - 2)
                    {
                        nd = ((CheckBox)control).Text;
                        if (((CheckBox)control).CheckState == CheckState.Checked) statend = true;
                        else statend = false;

                    }
                }
            }
            int x = 0;

            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (((CheckBox)control).TabIndex == tabindex - 1)
                    {
                        ((CheckBox)control).Text = nd;
                        if (statend) ((CheckBox)control).CheckState = CheckState.Checked;
                        else ((CheckBox)control).CheckState = CheckState.Unchecked;
                        int y = 0;

                        y = statistic[x];
                        statistic[x] = statistic[x - 1];
                        statistic[x - 1] = y;

                        y = staticIndex[x];
                        staticIndex[x] = staticIndex[x - 1];
                        staticIndex[x - 1] = y;
                    }
                    if (((CheckBox)control).TabIndex == tabindex - 2)
                    {
                        ((CheckBox)control).Text = st;
                        if (statest) ((CheckBox)control).CheckState = CheckState.Checked;
                        else ((CheckBox)control).CheckState = CheckState.Unchecked;
                    }
                    x++;
                }

            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

            string docType = "";
            if (btnPrint.InvokeRequired)
            {
                btnPrint.Invoke(new MethodInvoker(delegate { btnPrint.Enabled = false; }));
            }

            if (نوع_التوكيل.InvokeRequired)
            {
                نوع_التوكيل.Invoke(new MethodInvoker(delegate { docType = نوع_التوكيل.Text; }));
            }
            //if (btnPrint.InvokeRequired)
            //{
            //    btnPrint.Invoke(new MethodInvoker(delegate { btnPrint.Enabled = false; }));
            //}

            //if (نوع_التوكيل.InvokeRequired)
            //{
            //    نوع_التوكيل.Invoke(new MethodInvoker(delegate { docType = نوع_التوكيل.Text; }));
            //}
            chooseDocxFile(مقدم_الطلب.Text.Split('_')[0], رقم_التوكيل.Text, docType, proType1);
            prepareDocxfile();
            if (btnPrint.InvokeRequired)
            {
                btnPrint.Invoke(new MethodInvoker(delegate { btnPrint.Enabled = true; }));
            }

        }


        private void txtReview_TextChanged(object sender, EventArgs e)
        {
            //if (!checkAutoUpdate.Checked)
            //{
            //    if (oldText != txtReview.Text)
            //    {
            //        editsMade[0] = true;
            //        error2.Checked = true;
            //        error2.Enabled = false;
            //    }
            //}
            ////txtReview.Text = removeSpace(txtReview.Text,false);
        }

        private void allowedEdit_SelectedIndexChanged(object sender, EventArgs e)
        {
            //SqlConnection sqlCon = new SqlConnection(DataSource);
            //if (sqlCon.State == ConnectionState.Closed)
            //    sqlCon.Open();
            //SqlCommand sqlCmd = new SqlCommand("update TableSettings set allowedEdit=@allowedEdit", sqlCon);
            //sqlCmd.CommandType = CommandType.Text;
            //sqlCmd.Parameters.AddWithValue("@allowedEdit", Convert.ToInt32(allowedEdit.Text));
            //sqlCmd.ExecuteNonQuery();
            //sqlCon.Close();
        }


        private void التاريخ_الميلادي_TextChanged(object sender, EventArgs e)
        {
            التاريخ_الميلادي_off.Text = التاريخ_الميلادي.Text.Split('-')[1] + " - " + التاريخ_الميلادي.Text.Split('-')[0] + " - " + التاريخ_الميلادي.Text.Split('-')[2];
            //MessageBox.Show(التاريخ_الميلادي_off.Text);
        }

        private void صفة_الموكل_off_SelectedIndexChanged(object sender, EventArgs e)
        {
            //autoCompleteTextBox(txtAddRight, DataSource, "قائمة_الحقوق_الكاملة", "TableAuthRight");            
        }

        private void txtAddRight_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (e.KeyChar == (char)13) btnAddRight.PerformClick();
        }

        private void savedRights_CheckedChanged(object sender, EventArgs e)
        {
            صفة_الموكل_off.Enabled = true;
            if (savedRights.Checked) savedRights.Text = "قائمة الحقوق المحفوظة";
            else savedRights.Text = "قائمة حقوق جديدة";
        }

        private void ListSearch_MouseClick(object sender, MouseEventArgs e)
        {

            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
        }

        private void LibtnAdd1_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("LibtnAdd1 - " + LibtnAdd1.Text);
        }

        private void Vicheck1_CheckedChanged(object sender, EventArgs e)
        {
            if (Vicheck1.Checked)
                Vicheck1.Text = itemsicheck1[0].Split('_')[0];
            else Vicheck1.Text = itemsicheck1[0].Split('_')[1];

        }

        private void Vicheck2_CheckedChanged(object sender, EventArgs e)
        {
            if (Vicheck1.Checked)
                Vicheck1.Text = itemsicheck1[1].Split('_')[0];
            else Vicheck1.Text = itemsicheck1[1].Split('_')[1];
        }

        private void الحقوق_off_TextChanged(object sender, EventArgs e)
        {
            resetBoxes(false);
            bool genForm = false;
            if (نوع_التوكيل.Text == "توكيل بصيغة غير مدرجة") 
                genForm = true;
            PopulateCheckBoxes(genForm, الحقوق_off.Text.Replace(" ", "_").Replace("-", "_"), "TableAuthRights", DataSource, صفة_مقدم_الطلب_off.SelectedIndex, false);
            قائمة_الحقوق.Text = removeSpace(قائمة_الحقوق.Text, true);
            //MessageBox.Show(قائمة_الحقوق.Text);
        }

        private void الحقوق_off_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void صفة_مقدم_الطلب_off_SelectedIndexChanged(object sender, EventArgs e)
        {
            //autoCompleteTextBox(txtAddRight, DataSource, "قائمة_الحقوق_الكاملة", "TableAuthRight");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            resetBoxes(false);
            bool genForm = false;
            if (نوع_التوكيل.Text == "توكيل بصيغة غير مدرجة") genForm = true;
            PopulateCheckBoxes(genForm, الحقوق_off.Text.Replace(" ", "_"), "TableAuthRights", DataSource, صفة_مقدم_الطلب_off.SelectedIndex, false);
            //autoCompleteTextBox(txtAddRight, DataSource, "قائمة_الحقوق_الكاملة", "TableAuthRight");
        }

        private void autoCompleteTextBox(TextBox textbox, string source, string comlumnName, string tableName)
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
                    text = SuffReplacements(text, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);
                    Console.WriteLine("autoCompleteTextBox " + text);
                    autoComplete.Add(text);
                }
                textbox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }

        private void autoCompleteTextBox(ComboBox textbox, string source, string comlumnName, string tableName)
        {

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
                    text = SuffReplacements(text, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);
                    Console.WriteLine("autoCompleteTextBox " + text);
                    autoComplete.Add(text);
                }
                textbox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }

        private void طريقة_الطلب_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtAddRight_TextChanged(object sender, EventArgs e)
        {

        }

        private void موقع_التوكيل_TextUpdate(object sender, EventArgs e)
        {

        }

        private void التوثيق_TextChanged(object sender, EventArgs e)
        {

        }

        public void FillDataGridView(string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableAuth", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            rowCount = dtbl.Rows.Count.ToString();
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            //dataGridView1.Columns[0].Visible = false ;
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[3].Width = 50;
            dataGridView1.Columns[8].Width = 50;
            dataGridView1.Columns[9].Width = 170;
            dataGridView1.Columns[7].Width = dataGridView1.Columns[2].Width = 200;
            AuthNoPart2 = dataGridView1.Rows.Count.ToString();
            sqlCon.Close();
            int bre = 0;
            ColorFulGrid9();

        }
        private void addButtonInfo(string text1, string text2, string text3, string text4, string text5)
        {
            //MessageBox.Show(text1);
            // 
            // textBox1
            // 
            TextBox textBox1 = new TextBox();
            textBox1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            textBox1.Location = new System.Drawing.Point(1007, 44);
            textBox1.Name = "textBox1_" + ButtonInfoIndex.ToString() + ".";
            textBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            textBox1.Size = new System.Drawing.Size(230, 35);
            textBox1.TabIndex = 570;
            textBox1.Text = text1;
            textBox1.Visible = true;
            if (labl1.Text == "غير مدرج")
                textBox1.Enabled = false;
            // 
            // textBox2
            // 
            TextBox textBox2 = new TextBox();
            textBox2.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            textBox2.Location = new System.Drawing.Point(817, 44);
            textBox2.Name = "textBox2_" + ButtonInfoIndex.ToString() + ".";
            textBox2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            textBox2.Size = new System.Drawing.Size(230, 35);
            textBox2.TabIndex = 572;
            textBox2.Visible = true;
            textBox2.Text = text2;
            if (labl2.Text == "غير مدرج")
                textBox2.Enabled = false;
            // 
            // textBox3
            // 
            TextBox textBox3 = new TextBox();
            textBox3.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            textBox3.Location = new System.Drawing.Point(571, 44);
            textBox3.Name = "textBox3_" + ButtonInfoIndex.ToString() + ".";
            textBox3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            textBox3.Size = new System.Drawing.Size(230, 35);
            textBox3.TabIndex = 574;
            textBox3.Visible = true;
            textBox3.Text = text3;
            if (labl3.Text == "غير مدرج")
                textBox3.Enabled = false;
            // 
            // textBox4
            // 
            TextBox textBox4 = new TextBox();
            textBox4.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            textBox4.Location = new System.Drawing.Point(325, 44);
            textBox4.Name = "textBox4_" + ButtonInfoIndex.ToString() + ".";
            textBox4.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            textBox4.Size = new System.Drawing.Size(230, 35);
            textBox4.TabIndex = 576;
            textBox4.Visible = true;
            textBox4.Text = text4;
            if (labl4.Text == "غير مدرج")
                textBox4.Enabled = false;
            // 
            // textBox5
            // 
            TextBox textBox5 = new TextBox();
            textBox5.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            textBox5.Location = new System.Drawing.Point(79, 44);
            textBox5.Name = "textBox5_" + ButtonInfoIndex.ToString() + ".";
            textBox5.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            textBox5.Size = new System.Drawing.Size(230, 35);
            textBox5.TabIndex = 578;
            textBox5.Visible = true;
            textBox5.Text = text5;
            if (labl5.Text == "غير مدرج")
                textBox5.Enabled = false;

            PictureBox addName = new PictureBox();
            addName.Image = global::PersAhwal.Properties.Resources.add;
            addName.Location = new System.Drawing.Point(92, 44);
            addName.Name = "addName_" + ButtonInfoIndex.ToString() + ".";
            addName.Size = new System.Drawing.Size(50, 35);
            addName.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            addName.TabIndex = 123;
            addName.TabStop = false;
            addName.Click += new System.EventHandler(this.addNameBtnName_Click);
            // 
            // removeName1
            // 
            PictureBox removeName = new PictureBox();
            removeName.Image = global::PersAhwal.Properties.Resources.remove;
            removeName.Location = new System.Drawing.Point(32, 44);
            removeName.Name = "removeName_" + ButtonInfoIndex.ToString() + ".";
            removeName.Size = new System.Drawing.Size(50, 35);
            removeName.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            removeName.TabIndex = 175;
            removeName.TabStop = false;
            removeName.Click += new System.EventHandler(this.removeBtnName_Click);

            PanelButtonInfo.Controls.Add(textBox1);
            PanelButtonInfo.Controls.Add(textBox2);
            PanelButtonInfo.Controls.Add(textBox3);
            PanelButtonInfo.Controls.Add(textBox4);
            PanelButtonInfo.Controls.Add(textBox5);
            PanelButtonInfo.Controls.Add(addName);
            PanelButtonInfo.Controls.Add(removeName);

            ButtonInfoIndex++;
            LibtnAdd1.Text = "اضافة (" + ButtonInfoIndex.ToString() + "/" + ButtonInfoIndex.ToString() + ")" + "   ";
        }
        private void addNameBtnName_Click(object sender, EventArgs e)
        {
            addButtonInfo("", "", "", "", "");
        }

        private void Vitext1_TextChanged(object sender, EventArgs e)
        {

        }

        private void طريقة_الإجراء_CheckedChanged(object sender, EventArgs e)
        {
            if (طريقة_الإجراء.Checked)
            {

                طريقة_الإجراء.Text = "حضور بالأصالة";
                label18.Visible = تاريخ_إصدار_الوكالة.Visible = label15.Visible = اسم_الموكل_بالتوقيع.Visible = label16.Visible = رقم_الوكالة.Visible = label17.Visible = جهة_إصدار_الوكالة.Visible = label18.Visible = تاريخ_إصدار_الوكالة.Visible = نوع_الموقع.Visible = false;
                اسم_الموكل_بالتوقيع.Text = رقم_الوكالة.Text = جهة_إصدار_الوكالة.Text = تاريخ_إصدار_الوكالة.Text = "بدون";
                
            }
            else
            {
                اسم_الموكل_بالتوقيع.Text = رقم_الوكالة.Text = جهة_إصدار_الوكالة.Text = تاريخ_إصدار_الوكالة.Text = "";
                تاريخ_إصدار_الوكالة.Visible = label18.Visible = جهة_إصدار_الوكالة.Visible = label17.Visible = رقم_الوكالة.Visible = label16.Visible = اسم_الموكل_بالتوقيع.Visible = label15.Visible = نوع_الموقع.Visible = true;
                طريقة_الإجراء.Text = "حضور بالإنابة";

            }

            if (!طريقة_الإجراء.Checked)
            {
                التوثيق.Size = new System.Drawing.Size(923, 123);
                تعليق_جديد_Off.Size = new System.Drawing.Size(923, 53);
                flowLayoutPanel1.Size = new System.Drawing.Size(923, 149);
                flowLayoutPanel2.Size = new System.Drawing.Size(940, 175);
            }
            else
            {
                التوثيق.Size = new System.Drawing.Size(923, 123);
                تعليق_جديد_Off.Size = new System.Drawing.Size(923, 53);
                flowLayoutPanel1.Size = new System.Drawing.Size(923, 219);
                flowLayoutPanel2.Size = new System.Drawing.Size(940, 175);
            }

        }

        private void طريقة_الإجراء_TextChanged(object sender, EventArgs e)
        {
            if (طريقة_الإجراء.Text == "حضور بالأصالة")
                طريقة_الإجراء.Checked = true;
            else طريقة_الإجراء.Checked = false;
        }

        private void نوع_الموقع_TextChanged(object sender, EventArgs e)
        {
            if (نوع_الموقع.Text == "السيد")
                نوع_الموقع.Checked = true;
            else نوع_الموقع.Checked = false;
        }

        private void نوع_الموقع_CheckedChanged(object sender, EventArgs e)
        {
            if (نوع_الموقع.Checked)
                onBehalfIndex = 0;
            else onBehalfIndex = 1;
            authJob();
        }

        private void timer4_Tick(object sender, EventArgs e)
        {

            if (ButtonInfoIndex != 0)
            {
                PanelButtonInfo.Visible = true;

                txtReview.Size = new System.Drawing.Size(1007, 57);
                checkAutoUpdate.Height = 57;
            }
            else
            {
                PanelButtonInfo.Visible = false;
                txtReview.Size = new System.Drawing.Size(1007, 154);
                checkAutoUpdate.Height = 154;
            }

            if (طريقة_الإجراء.Checked)
            {
                if (!timer4.Enabled) timer4.Enabled = true;
                return;
            }
            else
            {
                timer4.Enabled = false;
            }
            if (طريقة_الطلب.Checked)
            {
                mandoubLabel.Visible = اسم_المندوب.Visible = false; اسم_المندوب.Text = "";
            }
        }

        private void اسم_المندوب_TextUpdate(object sender, EventArgs e)
        {
            if (اسم_المندوب.Text != "" && !اسم_المندوب.Text.Contains("ختر"))
            {
                //MessageBox.Show("update");

                الشاهد_الأول.Text = اسم_المندوب.Text.Split('-')[0].Trim();
                هوية_الأول.Text = getMandoubPass(DataSource, اسم_المندوب.Text.Split('-')[0].Trim());
            }
        }

        private void اسم_المندوب_TextChanged(object sender, EventArgs e)
        {
            
            if (اسم_المندوب.Text != "" && اسم_المندوب.Text != "حضور مباشرة إلى القنصلية" && اسم_المندوب.Text != "إختر اسم المندوب" && طريقة_الطلب.Checked)
            {
                //MessageBox.Show("change");

                الشاهد_الأول.Enabled = هوية_الأول.Enabled = false;
                الشاهد_الأول.Text = اسم_المندوب.Text.Split('-')[0].Trim();
                هوية_الأول.Text = getMandoubPass(DataSource, اسم_المندوب.Text.Split('-')[0].Trim());
            }
            else
                الشاهد_الأول.Enabled = هوية_الأول.Enabled = true;
        }

        private void FormAuth_FormClosed(object sender, FormClosedEventArgs e)
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

        private void مقدم_الطلب_TextChanged(object sender, EventArgs e)
        {

        }
        int fisrtWitIndex = 0;
        private void الشاهد_الأول_TextChanged(object sender, EventArgs e)
        {
            string[] text = getID(الشاهد_الأول);
            هوية_الأول.Text = text[0];
        }
        int secondWitIndex = 0;
        private void الشاهد_الثاني_TextChanged(object sender, EventArgs e)
        {
            string[] text = getID(الشاهد_الثاني);
            هوية_الثاني.Text = text[0];
        }

        private void هوية_الأول_KeyDown(object sender, KeyEventArgs e)
        {
            //if (fisrtWitIndex < autoFillIndex - 1) fisrtWitIndex++;
            //else return;
            //getID(هوية_الأول, الشاهد_الأول.Text, "رقم_الهوية", fisrtWitIndex);
        }

        private void هوية_الأول_KeyUp(object sender, KeyEventArgs e)
        {
            //if (fisrtWitIndex > 0) fisrtWitIndex--;
            //else return;
            //getID(هوية_الأول, الشاهد_الأول.Text, "رقم_الهوية", fisrtWitIndex);
        }

        private void هوية_الأول_MouseClick(object sender, MouseEventArgs e)
        {
            //getID(هوية_الأول, الشاهد_الأول.Text, "رقم_الهوية", fisrtWitIndex,"P0");
        }

        private void هوية_الثاني_MouseClick(object sender, MouseEventArgs e)
        {
            // getID(هوية_الثاني, الشاهد_الثاني.Text, "رقم_الهوية", secondWitIndex, "P0");
        }

        private void removeBtnName_Click(object sender, EventArgs e)
        {
            PictureBox pictureBox = (PictureBox)sender;
            string rowID = pictureBox.Name.Split('_')[1];
            //MessageBox.Show(rowID);
            foreach (Control control in PanelButtonInfo.Controls)
            {
                if (control.Visible && control.Name.Contains("_" + rowID) && control.Name.Contains("."))
                {
                    control.Visible = false;
                    control.Name = "unvalid_" + InvalidControl.ToString();
                    InvalidControl++;
                }
            }
        }
        private void fillTextBoxes(TextBox textbox, int index)
        {
            int id = 0;
            foreach (Control control in PanelButtonInfo.Controls)
            {
                if (control.Visible && control.Name.Contains("textBox" + index + "_"))
                {
                    if (id == 0)
                    {
                        textbox.Text = control.Text;
                    }
                    else
                    {
                        textbox.Text = textbox.Text + "_" + control.Text;
                    }
                    id++;
                }
            }
        }

        private void اسم_المندوب_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void توكيل_الغير_CheckedChanged(object sender, EventArgs e)
        {
            authOthers();
        }

        private void authOthers()
        {
            try
            {
                if (!قائمة_الحقوق.Text.Contains(authother) && توكيل_الغير.Checked)
                {
                    if (!إلغاء_توكيل_سابق.Checked)
                        قائمة_الحقوق.Text = قائمة_الحقوق.Text.Replace(lastSentence, authother + lastSentence);
                    else قائمة_الحقوق.Text = قائمة_الحقوق.Text.Replace(removeAuthother, authother + removeAuthother);
                }
                else if (!توكيل_الغير.Checked)
                    قائمة_الحقوق.Text = قائمة_الحقوق.Text.Replace(authother, "");
            }
            catch (Exception ex) { }
        }

        private void إلغاء_توكيل_سابق_CheckedChanged(object sender, EventArgs e)
        {
            removedPreAuth();
        }


        private void removedPreAuth()
        {
            if (قائمة_الحقوق.Text == "" && إلغاء_توكيل_سابق.Checked)
                قائمة_الحقوق.Text = "ويعتبر التوكيل الصادر من #6 بتاريخ #7 بالرقم #8 لاغٍ،";

            if (!قائمة_الحقوق.Text.Contains("ويعتبر التوكيل الصادر") && إلغاء_توكيل_سابق.Checked)
            {
                panelAuthOptions.Size = new System.Drawing.Size(1007, 261);
                try
                {
                    قائمة_الحقوق.Text = قائمة_الحقوق.Text.Replace(lastSentence, removeAuthother + lastSentence);
                }catch (Exception ex) {

                    string removeOrigin = "ويعتبر التوكيل الصادر من #6 بتاريخ #7 بالرقم #8 لاغٍ،";
                    string lastRemoveAuth = removeAuthother;
                    getDocInfoRemove(DataSource);
                    removeAuthother = removeOrigin.Replace("#6", removedDocSource.Text);
                    removeAuthother = removeAuthother.Replace("#7", removedDocDate.Text);
                    removeAuthother = removeAuthother.Replace("#8", removedDocNo.Text);
                    قائمة_الحقوق.Text = removeAuthother;
                }
                panelRemove.Visible = buttonRemove.Visible = true;
            }
            else if (!إلغاء_توكيل_سابق.Checked)
            {
                panelAuthOptions.Size = new System.Drawing.Size(1007, 313);
                try
                {
                    قائمة_الحقوق.Text = قائمة_الحقوق.Text.Replace(removeAuthother, "");
                }
                catch (Exception ex) { }
                panelRemove.Visible = buttonRemove.Visible = false;
            }
        }

        private void removedDocNo_TextChanged(object sender, EventArgs e)
        {
            string removeOrigin = "ويعتبر التوكيل الصادر من #6 بتاريخ #7 بالرقم #8 لاغٍ،";
            string lastRemoveAuth = removeAuthother;
            getDocInfoRemove(DataSource);
            removeAuthother = removeOrigin.Replace("#6", removedDocSource.Text);
            removeAuthother = removeAuthother.Replace("#7", removedDocDate.Text);
            removeAuthother = removeAuthother.Replace("#8", removedDocNo.Text);
            قائمة_الحقوق.Text = قائمة_الحقوق.Text.Replace(lastRemoveAuth, removeAuthother);
        }

        private void removedDocSource_TextChanged(object sender, EventArgs e)
        {
            string removeOrigin = "ويعتبر التوكيل الصادر من #6 بتاريخ #7 بالرقم #8 لاغٍ،";
            string lastRemoveAuth = removeAuthother;
            removeAuthother = removeOrigin.Replace("#6", removedDocSource.Text);
            removeAuthother = removeAuthother.Replace("#7", removedDocDate.Text);
            removeAuthother = removeAuthother.Replace("#8", removedDocNo.Text);
            قائمة_الحقوق.Text = قائمة_الحقوق.Text.Replace(lastRemoveAuth, removeAuthother);
        }

        private void removedDocDate_TextChanged(object sender, EventArgs e)
        {
            string removeOrigin = "ويعتبر التوكيل الصادر من #6 بتاريخ #7 بالرقم #8 لاغٍ،";
            string lastRemoveAuth = removeAuthother;
            removeAuthother = removeOrigin.Replace("#6", removedDocSource.Text);
            removeAuthother = removeAuthother.Replace("#7", removedDocDate.Text);
            removeAuthother = removeAuthother.Replace("#8", removedDocNo.Text);
            قائمة_الحقوق.Text = قائمة_الحقوق.Text.Replace(lastRemoveAuth, removeAuthother);
        }

        private void labDescribed_Click(object sender, EventArgs e)
        {

        }

        private void panelRemove_VisibleChanged(object sender, EventArgs e)
        {
            if (panelRemove.Visible)
            {
                panelAuthOptions.Size = new System.Drawing.Size(1007, 261);
                buttonRemove.Visible = true;
            }
            else
            {
                panelAuthOptions.Size = new System.Drawing.Size(1007, 313);
                buttonRemove.Visible = false;
            }
        }

        private void timer5_Tick(object sender, EventArgs e)
        {
            //MessageBox.Show("صفة_مقدم_الطلب " +  Appcases(النوع, addNameIndex).ToString());
            صفة_مقدم_الطلب_off.SelectedIndex = Appcases(النوع, addNameIndex);
            //MessageBox.Show("صفة_الموكل  " + Appcases(جنس_الموكَّل, addAuthticIndex));
            صفة_الموكل_off.SelectedIndex = Appcases(جنس_الموكَّل, addAuthticIndex);



        }

        private void SuffConvertments(TextBox text, int person1, int person2)
        {
            string[] words = text.Text.Split(' ');

            foreach (string word in words)
            {
                if (word == "" || word == " ") continue;
                for (int gridIndex = 0; gridIndex < dataGridView2.Rows.Count - 1; gridIndex++)
                {
                    string code = dataGridView2.Rows[gridIndex].Cells["الرموز"].Value.ToString();
                    string person = dataGridView2.Rows[gridIndex].Cells["الضمير"].Value.ToString();

                    string replacemest1 = dataGridView2.Rows[gridIndex].Cells["المقابل" + (person1 + 1).ToString()].Value.ToString();
                    string replacemest2 = dataGridView2.Rows[gridIndex].Cells["المقابل" + (person2 + 1).ToString()].Value.ToString();

                    string[] replacemests = new string[6];
                    replacemests[0] = dataGridView2.Rows[gridIndex].Cells["المقابل1"].Value.ToString();
                    replacemests[1] = dataGridView2.Rows[gridIndex].Cells["المقابل2"].Value.ToString();
                    replacemests[2] = dataGridView2.Rows[gridIndex].Cells["المقابل3"].Value.ToString();
                    replacemests[3] = dataGridView2.Rows[gridIndex].Cells["المقابل4"].Value.ToString();
                    replacemests[4] = dataGridView2.Rows[gridIndex].Cells["المقابل5"].Value.ToString();
                    replacemests[5] = dataGridView2.Rows[gridIndex].Cells["المقابل6"].Value.ToString();

                    for (int cellIndex = 0; cellIndex < 6; cellIndex++)
                    {
                        if (word == replacemests[cellIndex] || word == replacemests[cellIndex] + "،")
                        {
                            Console.WriteLine(word);
                            if (person == "1")
                            {
                                if (word != replacemests[person1] && word != replacemests[person1]+"،")
                                {
                                    var selectedOption = MessageBox.Show("هل تود إجراء التصحيح التلقائي (" + replacemests[person1] + ")", "تم رصد خطاء في الصياغة (" + word + ")", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (selectedOption == DialogResult.Yes)
                                    {
                                        text.Text = text.Text.Replace(word, replacemests[person1]);
                                        break;
                                    }
                                }
                            }
                            if (person == "2")
                            {
                                if (word != replacemests[person2] && word != replacemests[person2] + "،")
                                {
                                    var selectedOption = MessageBox.Show("هل تود إجراء التصحيح التلقائي (" + replacemests[person2] + ")", "تم رصد خطاء في الصياغة (" + word + ")", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (selectedOption == DialogResult.Yes)
                                    {
                                        text.Text = text.Text.Replace(word, replacemests[person2]);
                                        break;
                                    }
                                }
                            }
                            if (person == "3")
                            {
                                if (word != replacemests[person1] && word != replacemests[person1] + "،")
                                {
                                    var selectedOption = MessageBox.Show("هل تود إجراء التصحيح التلقائي (" + replacemests[person1] + ")", "تم رصد خطاء في الصياغة (" + word + ")", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (selectedOption == DialogResult.Yes)
                                    {
                                        text.Text = text.Text.Replace(word, replacemests[person1]);
                                        break;
                                    }
                                }
                            }
                        }
                    }

                }
            }

        }

        private string SuffConvertments(string text, int person1, int person2)
        {
            string[] words = text.Split(' ');

            foreach (string word in words)
            {
                if (word == "" || word == " ") continue;
                for (int gridIndex = 0; gridIndex < dataGridView2.Rows.Count - 1; gridIndex++)
                {
                    string code = dataGridView2.Rows[gridIndex].Cells["الرموز"].Value.ToString();
                    string person = dataGridView2.Rows[gridIndex].Cells["الضمير"].Value.ToString();

                    string replacemest1 = dataGridView2.Rows[gridIndex].Cells["المقابل" + (person1 + 1).ToString()].Value.ToString();
                    string replacemest2 = dataGridView2.Rows[gridIndex].Cells["المقابل" + (person2 + 1).ToString()].Value.ToString();

                    string[] replacemests = new string[6];
                    replacemests[0] = dataGridView2.Rows[gridIndex].Cells["المقابل1"].Value.ToString();
                    replacemests[1] = dataGridView2.Rows[gridIndex].Cells["المقابل2"].Value.ToString();
                    replacemests[2] = dataGridView2.Rows[gridIndex].Cells["المقابل3"].Value.ToString();
                    replacemests[3] = dataGridView2.Rows[gridIndex].Cells["المقابل4"].Value.ToString();
                    replacemests[4] = dataGridView2.Rows[gridIndex].Cells["المقابل5"].Value.ToString();
                    replacemests[5] = dataGridView2.Rows[gridIndex].Cells["المقابل6"].Value.ToString();

                    for (int cellIndex = 0; cellIndex < 6; cellIndex++)
                    {
                        if (word == replacemests[cellIndex] || word == replacemests[cellIndex] + "،")
                        {
                            Console.WriteLine(word);
                            if (person == "1")
                            {
                                if (word != replacemests[person1])
                                {
                                    var selectedOption = MessageBox.Show("هل تود إجراء التصحيح التلقائي (" + replacemests[person2] + ")", "تم رصد خطاء في الصياغة (" + word + ")", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (selectedOption == DialogResult.Yes)
                                    {
                                        text = text.Replace(word, replacemests[person1]);
                                        break;
                                    }
                                    //    MessageBox.Show(word); 
                                }
                            }
                            if (person == "2")
                            {
                                if (word != replacemests[person2])
                                {
                                    var selectedOption = MessageBox.Show("هل تود إجراء التصحيح التلقائي (" + replacemests[person2] + ")", "تم رصد خطاء في الصياغة (" + word + ")", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (selectedOption == DialogResult.Yes)
                                    {
                                        text = text.Replace(word, replacemests[person2]);
                                        break;
                                    }
                                    //    MessageBox.Show(word); 
                                }
                            }
                            if (person == "3")
                            {
                                if (word != replacemests[person1])
                                {
                                    var selectedOption = MessageBox.Show("هل تود إجراء التصحيح التلقائي (" + replacemests[person1] + ")", "تم رصد خطاء في الصياغة (" + word + ")", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (selectedOption == DialogResult.Yes)
                                    {
                                        text = text.Replace(word, replacemests[person1]);
                                        break;
                                    }
                                    //    MessageBox.Show(word); 
                                }
                            }
                            //text = text.Replace(replacemest[cellIndex], code);
                            //break;
                        }
                    }

                }
            }
            return text;
        }
        //private string SuffOrigConvertments(string text)
        //{
        //    try
        //    {
        //        string[] words = text.Split(' ');

        //        foreach (string word in words)
        //        {
        //            if (word == "" || word == " ") continue;
        //            for (int gridIndex = 0; gridIndex < dataGridView2.Rows.Count - 1; gridIndex++)
        //            {
        //                string code = dataGridView2.Rows[gridIndex].Cells["الرموز"].Value.ToString();
        //                string[] replacemests = new string[6];
        //                replacemests[0] = dataGridView2.Rows[gridIndex].Cells["المقابل1"].Value.ToString();
        //                replacemests[1] = dataGridView2.Rows[gridIndex].Cells["المقابل2"].Value.ToString();
        //                replacemests[2] = dataGridView2.Rows[gridIndex].Cells["المقابل3"].Value.ToString();
        //                replacemests[3] = dataGridView2.Rows[gridIndex].Cells["المقابل4"].Value.ToString();
        //                replacemests[4] = dataGridView2.Rows[gridIndex].Cells["المقابل5"].Value.ToString();
        //                replacemests[5] = dataGridView2.Rows[gridIndex].Cells["المقابل6"].Value.ToString();

        //                for (int cellIndex = 0; cellIndex < 6; cellIndex++)
        //                {
        //                    if (word == replacemests[cellIndex])
        //                    {
        //                        text = text.Replace(word, code);
        //                        break;
        //                    }
        //                    else if (word == replacemests[cellIndex] + "،")
        //                    {
        //                        text = text.Replace(word, code + "،");
        //                        break;
        //                    }
        //                }

        //            }
        //        }
        //    }
        //    catch (Exception ex) { }
        //    return text;
        //}

        private void اسم_المندوب_VisibleChanged(object sender, EventArgs e)
        {

        }

        private void reversTextReview()
        {
            string col = نوع_التوكيل.Text.Replace(" ", "_") + "_" + إجراء_التوكيل.Text.Replace(" ", "_");
            string query = "select " + col + " from TableAuthsub order by ID desc";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            //try
            //{
                sqlDa.Fill(dtbl);
                sqlCon.Close();
                txtReviewListIndex = 0;
                txtReviewStarIndex = 1;
                txtReviewList = new string[dtbl.Rows.Count];
                txtReviewListStar = new string[dtbl.Rows.Count+1];
                if (dtbl.Rows.Count == 0)
                    fileUpdate.Visible = false;
                foreach (DataRow dataRow in dtbl.Rows)
                {
                    if (dataRow[col].ToString() == "" || dataRow[col].ToString().Contains("removed")) continue;
                    {
                        txtReviewList[txtReviewListIndex] = dataRow[col].ToString().Replace("Star", "");
                        Console.WriteLine(txtReviewList[txtReviewListIndex]);
                        txtReviewListIndex++;
                    }
                    if (dataRow[col].ToString() == "" || !dataRow[col].ToString().Contains("Star") || dataRow[col].ToString().Contains("removed")) continue;
                    {
                        txtReviewListStar[txtReviewStarIndex] = dataRow[col].ToString().Replace("Star","");
                        //MessageBox.Show("txtReviewStarIndex " + txtReviewListStar[txtReviewStarIndex]);
                        txtReviewStarIndex++;
                    }
                }
            //}
            //catch (Exception ex) {
            //    Console.WriteLine(query); 
            //    Console.WriteLine(DataSource);
            //    MessageBox.Show(query);
            //    }

        }
        private void reversTextRight()
        {
            string col = إجراء_التوكيل.Text.Replace(" ", "_") + "_" + نوع_التوكيل.Text.Replace(" ", "_");
            string query = "select " + col + " from TableAuthRightStarText order by ID desc";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
                sqlCon.Close();
                txtRightListIndex = 0;
                txtRightList = new string[dtbl.Rows.Count];
                foreach (DataRow dataRow in dtbl.Rows)
                {
                    if (dataRow[col].ToString() == "" || dataRow[col].ToString().Contains("removed")) continue;
                    txtRightList[txtRightListIndex] = dataRow[إجراء_التوكيل.Text.Replace(" ", "_")].ToString();
                    Console.WriteLine(txtRightList[txtRightListIndex]);
                    txtRightListIndex++;
                }
            }
            catch (Exception ex) { }

        }


        private void reversTextReviewold()
        {
            string query = "select * from TableAuth where إجراء_التوكيل = N'" + إجراء_التوكيل.Text + "' order by ID desc";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            int index = 0;
            txtReviewList = new string[dtbl.Rows.Count];
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["txtReview"].ToString() == "") continue;

                txtReviewList[index] = dataRow["txtReview"].ToString();

                Console.WriteLine(txtReviewList[index]);
                if (dataRow["itext1"].ToString() != "" && txtReviewList[index].Contains(dataRow["itext1"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["itext1"].ToString(), "t1");
                if (dataRow["itext2"].ToString() != "" && txtReviewList[index].Contains(dataRow["itext2"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["itext2"].ToString(), "t2");
                if (dataRow["itext3"].ToString() != "" && txtReviewList[index].Contains(dataRow["itext3"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["itext3"].ToString(), "t3");
                if (dataRow["itext4"].ToString() != "" && txtReviewList[index].Contains(dataRow["itext4"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["itext4"].ToString(), "t4");
                if (dataRow["itext5"].ToString() != "" && txtReviewList[index].Contains(dataRow["itext5"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["itext5"].ToString(), "t5");

                if (dataRow["icheck1"].ToString() != "" && txtReviewList[index].Contains(dataRow["icheck1"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["icheck1"].ToString(), "c1");

                if (dataRow["icombo1"].ToString() != "" && txtReviewList[index].Contains(dataRow["icombo1"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["icombo1"].ToString(), "m1");
                if (dataRow["icombo2"].ToString() != "" && txtReviewList[index].Contains(dataRow["icombo2"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["icombo2"].ToString(), "m2");

                if (dataRow["ibtnAdd1"].ToString() != "" && txtReviewList[index].Contains(dataRow["ibtnAdd1"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["ibtnAdd1"].ToString(), "a1");
                if (dataRow["itxtDate1"].ToString() != "" && txtReviewList[index].Contains(dataRow["itxtDate1"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["itxtDate1"].ToString(), "n1");
                Console.WriteLine(txtReviewList[index]);
                index++;
            }
        }
        private void reversTextRights()
        {
        }


        private string SuffReversReplacements(string text, int appCaseIndex, int intAuthcases)
        {
            //MessageBox.Show(text);
            try
            {
                if (text.Contains("auth1"))
                    text = text.Replace("auth1", legaceyAuthInfo);
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains("  "))
                    text = text.Replace("  ", " ");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vitext1.Text))
                    text = text.Replace(Vitext1.Text, "t1");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vitext2.Text))
                    text = text.Replace(Vitext2.Text, "t2");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vitext3.Text))
                    text = text.Replace(Vitext3.Text, "t3");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vitext4.Text))
                    text = text.Replace(Vitext4.Text, "t4");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vitext5.Text))
                    text = text.Replace(Vitext5.Text, "t5");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vicheck1.Text))
                    text = text.Replace(Vicheck1.Text, "c1");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vicombo1.Text))
                    text = text.Replace(Vicombo1.Text, "m1");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vicombo2.Text))
                    text = text.Replace(Vicombo2.Text, "m2");
            }
            catch (Exception ex1) { }

            try
            {
                if (text.Contains(LibtnAdd1.Text))
                    text = text.Replace(LibtnAdd1.Text, "a1");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(VitxtDate1.Text))
                    text = text.Replace(VitxtDate1.Text, "n1");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(VitxtDate1.Text))
                    text = text.Replace(VitxtDate1.Text, "n1");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains("لدى  برقم الايبان (" + Vitext3.Text + ")"))
                    text = text.Replace("لدى  برقم الايبان (" + Vitext3.Text + ")", "لدى  برقم الايبان (@*@)");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(removedDocNo.Text))
                    text = text.Replace(removedDocNo.Text, "#8");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(removedDocSource.Text))
                    text = text.Replace(removedDocSource.Text, "#6");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(removedDocDate.Text))
                    text = text.Replace(removedDocDate.Text, "#7");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(الموكَّل.Text))
                    text = text.Replace(الموكَّل.Text, "AuthName");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(جنسية_الموكل.Text))
                    text = text.Replace(جنسية_الموكل.Text, "AuthNatio");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(هوية_الموكل.Text))
                    text = text.Replace(هوية_الموكل.Text, "AuthDoc");
            }
            catch (Exception ex1) { }
            //MessageBox.Show(text);
            text = SuffConvertments(text, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);
            //MessageBox.Show(text);
            return text;
        }
        private void fileUpdate_Click(object sender, EventArgs e)
        {
            if (txtReviewList == null)
                return;

            checkAutoUpdate.Checked = false;
            if (txtReviewLast == "")
                txtReviewLast = txtReview.Text;
            int len = 0;
            try
            {
                len = txtReviewList.Length;                
            }
            catch (Exception ex1) { return; }
            if (txtRevIndex == len) txtRevIndex = 0;
            Console.WriteLine(txtReviewList[txtRevIndex]);
            //    StrSpecPur = SuffOrigConvertments(txtReviewList[txtRevIndex]);
            StrSpecPur = txtReviewList[txtRevIndex];
                //MessageBox.Show(txtReview.Text);
                txtReview.Text = SuffReplacements(StrSpecPur, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex) ;
                txtRevIndex++;
                txtReview.Text = removeSpace(txtReview.Text, false);
            
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            if (txtRightList == null)
                return; 
            قائمة_الحقوقLast = قائمة_الحقوق.Text;
            if (txtRigIndex == txtRightList.Length) txtRigIndex = 0;
//            قائمة_الحقوق.Text = SuffOrigConvertments(txtRightList[txtRigIndex]);
            قائمة_الحقوق.Text = txtRightList[txtRigIndex];
            //MessageBox.Show(قائمة_الحقوق.Text);
            قائمة_الحقوق.Text = SuffReplacements(قائمة_الحقوق.Text, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);
            highlightSentences();
            authOthers();
            removedPreAuth();
            txtRigIndex++;
            قائمة_الحقوق.Text = removeSpace(قائمة_الحقوق.Text, true);
        }


        private string removeSpace(string text, bool addLast)
        {
            string authother = "";
            string removeAuthother = "";
            string lastSentence = "";
            string[] sentences = text.Split('،');
            foreach (string sentence in sentences)
            {
                if (sentence.Contains("الحق في توكيل الغير"))
                    authother = sentence;
                if (sentence.Contains("ويعتبر التوكيل الصادر"))
                    removeAuthother = sentence;
                if (sentence.Contains("لمن يشهد والله"))
                    lastSentence = sentence;
            }
            if (addLast)
            {
                if (!text.Contains("لمن يشهد والله"))
                    text = text + "، وأذنت لمن يشهد والله خير الشاهدين";
                else
                    text = text.Replace(lastSentence, "، وأذنت لمن يشهد والله خير الشاهدين");
            }
            try
            {
                text = text.Replace(authother, "");
                text = text.Replace(removeAuthother, "");
            }
            catch (Exception ex) { }

            for (; text.Contains("،،");)
            {
                text = text.Replace("،،", "، ");
            }
            for (; text.Contains("..");)
            {
                text = text.Replace("..", ".");
            }
            
            for (; text.Contains("  ");)
            {
                text = text.Replace("  ", " ");
            }
            text = text.Replace("، ،", "، ");
            text = text.Replace("،", "، ");
            text = text.Replace("1_", "");
            text = text.Replace("0_", "");
            text = text.Replace("،،", "،");
            text = text.Replace("..", ".");
            for (; text.Contains("  ");)
            {
                text = text.Replace("  ", " ");
            }
            text = text.Trim();


            return text;
        }
        private void highlightSentences()
        {
            authother = removeAuthother = lastSentence = "";
            string[] sentences = قائمة_الحقوق.Text.Split('،');
            foreach (string sentence in sentences)
            {
                if (sentence.Contains("الحق في توكيل الغير"))
                    authother = sentence;
                if (sentence.Contains("ويعتبر التوكيل الصادر"))
                    removeAuthother = sentence;
                if (sentence.Contains("لمن يشهد والله"))
                    lastSentence = sentence;
            }

            if (lastSentence == "")
            {
                lastSentence = " ،وأذ#$# لمن يشهد والله خير الشاهدين";
                قائمة_الحقوق.Text = قائمة_الحقوق.Text + lastSentence.Replace("#$#", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 13]);
            }
        }

        private void قائمة_الحقوق_Click(object sender, EventArgs e)
        {

        }

        private void picStarSub_VisibleChanged(object sender, EventArgs e)
        {
        }

        private void picStarRight_VisibleChanged(object sender, EventArgs e)
        {
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            checkAutoUpdate.Checked = false; 
            txtReview.Text = txtReviewLast;
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            قائمة_الحقوق.Text = قائمة_الحقوقLast;
        }

        private void قائمة_الحقوق_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(قائمة_الحقوق.Text);
        }

        private void panelAuthRights_Paint(object sender, PaintEventArgs e)
        {

        }

        private void picStarSubAdd_Click(object sender, EventArgs e)
        {
            selectTable = "TableAuthsub";
            updateText(txtReview);
            reversTextReview();
        }
        private void updateText(TextBox text)
        {
            string starButton = "";
            if (!checkColExistance(selectTable, نوع_التوكيل.Text.Replace(" ", "_") + "_" + إجراء_التوكيل.Text.Replace(" ", "_")))
                CreateColumn(نوع_التوكيل.Text.Replace(" ", "_") + "_" + إجراء_التوكيل.Text.Replace(" ", "_"), selectTable, "max");
            else
            {
                starButton = getStarID(selectTable, نوع_التوكيل.Text.Replace(" ", "_") + "_" + إجراء_التوكيل.Text.Replace(" ", "_"), text.Text);
                //MessageBox.Show("نص موجود بالرقم " + starButton);
            }

            string textRevers = SuffReversReplacements(text.Text, 0, 0) + "Star";
            string query = "UPDATE " + selectTable + " SET " + نوع_التوكيل.Text.Replace(" ", "_") + "_" + إجراء_التوكيل.Text.Replace(" ", "_") + "=N'" + textRevers + "' where ID = " + starButton;

            if (!checkStarTextExist(dataSource, نوع_التوكيل.Text.Replace(" ", "_") + "_" + إجراء_التوكيل.Text.Replace(" ", "_"), SuffReversReplacements(text.Text, 0, 0), "TableAuthsub"))
                query = "insert into " + selectTable + " (" + نوع_التوكيل.Text.Replace(" ", "_") + "_" + إجراء_التوكيل.Text.Replace(" ", "_") + ") values (N'" + textRevers + "')";

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            Console.WriteLine("updateText " + query);
            sqlCmd.ExecuteNonQuery();                        
        }

        private bool checkStarTextExist(string dataSource, string col, string text, string genTable)
        {
            string query = "select * from " + genTable + " where " + col + "=N'" + text + "' or " + col + " = N'" + text + "'removed'"+ "' or " + col + " = N'" + text + "'Star'";
            Console.WriteLine(query);
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
            }
            catch (Exception ex) { }
            if (dtbl.Rows.Count > 0) return true;
            else return false;
            sqlCon.Close();
        }

        private void CreateColumn(string Columnname, string tableName, string size)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("alter table " + tableName + " add " + Columnname + " nvarchar(" + size + ")", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            try
            {
                sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex) { return; }
            sqlCon.Close();
        }
        private bool checkColExistance(string table, string colName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('" + table + "')", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            foreach (DataRow row in dtbl.Rows)
            {

                if (row["name"].ToString() == colName)
                {
                    return true;
                }
            }
            //MessageBox.Show(table+" - "+ colName);
            return false;

        }

        private string getStarID(string table, string colName, string text)
        {
            string ID = "";
            string query = "select ID from " + table + " where " + colName + " = N'" + text + "'";
            Console.WriteLine("getStarID " + query);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return ""; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            foreach (DataRow row in dtbl.Rows)
            {
                ID = row["ID"].ToString();
            }

            return ID;

        }



        private void picStarRightAdd_Click(object sender, EventArgs e)
        {
            selectTable = "TableAuthRightStarText";
            updateText(قائمة_الحقوق);            
        }

        private void picStarSub_Click(object sender, EventArgs e)
        {
            checkAutoUpdate.Checked = false;
            if (txtReviewLast == "")
                txtReviewLast = txtReview.Text;
            if (txtRevStarIndex == txtReviewStarIndex) 
                txtRevStarIndex = 0;
            //Console.WriteLine(txtReviewListStar[txtRevStarIndex]);
            //MessageBox.Show("1");
            //StrSpecPur = SuffOrigConvertments(txtReviewListStar[txtRevStarIndex]);
            StrSpecPur = txtReviewListStar[txtRevStarIndex];
            Console.WriteLine(StrSpecPur);
            MessageBox.Show(StrSpecPur);
            txtReview.Text = SuffReplacements(StrSpecPur, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);
            Console.WriteLine(txtReview.Text);
            MessageBox.Show(txtReview.Text);
            txtRevStarIndex++;
            txtReview.Text = removeSpace(txtReview.Text, false);
            Console.WriteLine(txtReview.Text);
            //MessageBox.Show("4");
            checkAutoUpdate.Checked = true;
            label12.Text = "الاختيار من قائمة المفضلة (" + txtRevStarIndex.ToString() + "/" + txtReviewStarIndex.ToString() + ")";
        }
        private string getStarText(string table, string colName, string ID)
        {
            string text = "";
            string query = "select " + colName + " from " + table + " where ID = '" + ID + "'";
            Console.WriteLine(query);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return ""; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            foreach (DataRow row in dtbl.Rows)
            {
                text = row[colName].ToString();
            }
            return text;
        }
        private void picStarRight_Click(object sender, EventArgs e)
        {
            selectTable = "TableAuthRightStarText";            
            الحقوق_off.Text = إجراء_التوكيل.Text.Replace(" ", "_") + "_" + نوع_التوكيل.Text.Replace(" ", "_");
            PopulateCheckBoxes(false, الحقوق_off.Text.Replace(" ", "_").Replace("-", "_"), "TableAuthRights", DataSource, صفة_مقدم_الطلب_off.SelectedIndex, false);
            قائمة_الحقوق.Text = removeSpace(قائمة_الحقوق.Text, true);

            Console.WriteLine(" 2 " + الحقوق_off.Text);
            //checkAutoUpdate.Checked = false;
            //if (starTextIndex == starText.Split('_').Length)
            //    starTextIndex = 0;

            //if (قائمة_الحقوقLast == "")
            //    قائمة_الحقوقLast = قائمة_الحقوق.Text;

            //قائمة_الحقوق.Text = SuffOrigConvertments(getStarText(selectTable, نوع_التوكيل.Text.Replace(" ", "_") + "_" + إجراء_التوكيل.Text.Replace(" ", "_"), starText.Split('_')[starTextIndex]));
            //قائمة_الحقوق.Text = SuffReplacements(قائمة_الحقوق.Text, صفة_مقدم_الطلب_off.SelectedIndex, صفة_الموكل_off.SelectedIndex);
            //starTextIndex++;
            //قائمة_الحقوق.Text = removeSpace(قائمة_الحقوق.Text, true);
        }

        private void picStarRightAdd_MouseEnter(object sender, EventArgs e)
        {
            label11.Text = "إضافة إلى المفضلة";
        }

        private void pictureBox3_MouseEnter(object sender, EventArgs e)
        {
            label11.Text = "الاختيار من قائمة المحفوظات العامة";
        }

        private void pictureBox7_MouseEnter(object sender, EventArgs e)
        {
            label11.Text = "الرجوع إلى النصل قبل التعديل";
        }

        private void picStarRight_MouseEnter(object sender, EventArgs e)
        {
            label11.Text = "الاختيار من قائمة المفضلة";
        }

        private void اسم_الموكل_بالتوقيع_TextChanged(object sender, EventArgs e)
        {
            authJob();
        }

        private void رقم_الوكالة_TextChanged(object sender, EventArgs e)
        {
            getDocInfoOnBehalf(DataSource);
            authJob();
        }


        private void getDocInfoOnBehalf(string source)
        {
            string query = "select * from TableAuth where رقم_التوكيل = N'" + رقم_الوكالة.Text + "'";
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
                autheticatingOthes = dataRow["الاعدادات"].ToString().Split('_')[0];
                archStat = dataRow["حالة_الارشفة"].ToString();
                removedStat = dataRow["المكاتبات_الملغية"].ToString();
                AuthenticName = dataRow["مقدم_الطلب"].ToString();
                تاريخ_إصدار_الوكالة.Text = dataRow["التاريخ_الميلادي"].ToString();
                اسم_الموكل_بالتوقيع.Text = dataRow["الموكَّل"].ToString();
                جهة_إصدار_الوكالة.Text = "القنصلية العامة لجمهورية السودان بجدة";
                جهة_إصدار_الوكالة.Enabled = اسم_الموكل_بالتوقيع.Enabled = تاريخ_إصدار_الوكالة.Enabled = false;
            }
            if (dtbl.Rows.Count == 0)
                جهة_إصدار_الوكالة.Enabled = اسم_الموكل_بالتوقيع.Enabled = تاريخ_إصدار_الوكالة.Enabled = true;
        }

        private void getDocInfoRemove(string source)
        {
            string query = "select * from TableAuth where رقم_التوكيل = N'" + removedDocNo.Text + "'";
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
                removedDocDate.Text = dataRow["التاريخ_الميلادي"].ToString();
                removedPersonName = dataRow["مقدم_الطلب"].ToString();
                removedDocSource.Text = "القنصلية العامة لجمهورية السودان بجدة";
                removedDocSource.Enabled = removedDocDate.Enabled = false;
            }
            if (dtbl.Rows.Count == 0)
            {
                removedDocSource.Enabled = removedDocDate.Enabled = true;
                removedDocSource.Text= removedDocDate.Text = "";
                removedPersonName = "";
            }
        }

        private void جهة_إصدار_الوكالة_TextChanged(object sender, EventArgs e)
        {
            authJob();
        }

        private void تاريخ_إصدار_الوكالة_TextChanged(object sender, EventArgs e)
        {
            authJob();
        }

        private void flowLayoutPanel3_MouseLeave(object sender, EventArgs e)
        {
            label11.Text = "";
        }

        private void رقم_المرجع_المرتبط_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            FinalPrint("pdf");
        }

        private void picStarSub_MouseEnter(object sender, EventArgs e)
        {
            label12.Text = "الاختيار من قائمة المفضلة (" + txtRevStarIndex.ToString() +"/"+ txtReviewStarIndex.ToString() + ")";
        }

        private void fileUpdate_MouseEnter(object sender, EventArgs e)
        {
            label12.Text = "الاختيار من القائمة العامة(" + txtRevIndex.ToString() + "/" + txtReviewListIndex.ToString() + ")";
        }

        private void picStarSubAdd_MouseEnter(object sender, EventArgs e)
        {
            label12.Text = "إضافة إلى المفضلة";
        }

        private void pictureBox5_MouseEnter(object sender, EventArgs e)
        {
            label12.Text = "الرجوع إلى النصل قبل التعديل";
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            //fillSamplesCodes(DataSource);
            if (Vitext1.InvokeRequired)
            {
                Vitext1.Invoke(new MethodInvoker(delegate {
                    autoCompleteTextBox(Vitext1, DataSource, "itext1", "TableAuth");
                }));
            }
            if (Vitext2.InvokeRequired)
            {
                Vitext2.Invoke(new MethodInvoker(delegate {
                    autoCompleteTextBox(Vitext2, DataSource, "itext2", "TableAuth");
                }));
            }
            if (Vitext3.InvokeRequired)
            {
                Vitext3.Invoke(new MethodInvoker(delegate {
                    autoCompleteTextBox(Vitext3, DataSource, "itext3", "TableAuth");
                }));
            }
            if (Vitext4.InvokeRequired)
            {
                Vitext4.Invoke(new MethodInvoker(delegate {
                    autoCompleteTextBox(Vitext4, DataSource, "itext4", "TableAuth");
                }));
            }
            if (Vitext5.InvokeRequired)
            {
                Vitext5.Invoke(new MethodInvoker(delegate {
                    autoCompleteTextBox(Vitext5, DataSource, "itext5", "TableAuth");
                }));
            }

            if (الشاهد_الأول.InvokeRequired)
            {
                الشاهد_الأول.Invoke(new MethodInvoker(delegate {
                    autoCompleteTextBox(الشاهد_الأول, DataSource, "الاسم", "TableGenNames");
                }));
            }

            if (الشاهد_الثاني.InvokeRequired)
            {
                الشاهد_الثاني.Invoke(new MethodInvoker(delegate {
                    autoCompleteTextBox(الشاهد_الثاني, DataSource, "الاسم", "TableGenNames");
                }));
            }

        }

        private void fillTextBoxesInvers()
        {
            if (!Vitext1.Text.Contains('_'))
            {
                PanelButtonInfo.Visible = false;
                return;
            }
            for (int x = 0; x < Vitext1.Text.Split('_').Length; x++)
            {
                addButtonInfo(Vitext1.Text.Split('_')[x], Vitext2.Text.Split('_')[x], Vitext3.Text.Split('_')[x], Vitext4.Text.Split('_')[x], Vitext5.Text.Split('_')[x]);
            }
            Vitext1.Text = Vitext2.Text = Vitext3.Text = Vitext4.Text = Vitext5.Text = "";
        }

        private void Suffex_preffixList()
        {

            //return;
            preffix[0, 0] = "ي"; //$$$
            preffix[1, 0] = "ي";
            preffix[2, 0] = "ا";
            preffix[3, 0] = "ا";
            preffix[4, 0] = "ا";
            preffix[5, 0] = "ا";


            preffix[0, 1] = "ت";//&&&
            preffix[1, 1] = "ت";
            preffix[2, 1] = "نا";
            preffix[3, 1] = "نا";
            preffix[4, 1] = "نا";
            preffix[5, 1] = "نا";


            preffix[0, 2] = "ني";//^^^
            preffix[1, 2] = "ني";
            preffix[2, 2] = "نا";
            preffix[3, 2] = "نا";
            preffix[4, 2] = "نا";
            preffix[5, 2] = "نا";


            preffix[0, 3] = "";//***
            preffix[1, 3] = "ت";
            preffix[2, 3] = "ا";
            preffix[3, 3] = "تا";
            preffix[4, 3] = "ن";
            preffix[5, 3] = "وا";


            preffix[0, 4] = "ه";//###
            preffix[1, 4] = "ها";
            preffix[2, 4] = "هما";
            preffix[3, 4] = "هما";
            preffix[4, 4] = "هن";
            preffix[5, 4] = "هم";


            preffix[0, 5] = "";
            preffix[1, 5] = "ة";
            preffix[2, 5] = "ان";
            preffix[3, 5] = "تان";
            preffix[4, 5] = "ات";
            preffix[5, 5] = "ون";


            preffix[0, 6] = "";//#5
            preffix[1, 6] = "ة";
            preffix[2, 6] = "ين";
            preffix[3, 6] = "تين";
            preffix[4, 6] = "ات";
            preffix[5, 6] = "ين";


            preffix[0, 7] = "ينوب";
            preffix[1, 7] = "تنوب";
            preffix[2, 7] = "ينوبا";
            preffix[3, 7] = "تنوبا";
            preffix[4, 7] = "ينبن";
            preffix[5, 7] = "ينوبوا";


            preffix[0, 8] = "يقوم";
            preffix[1, 8] = "تقوم";
            preffix[2, 8] = "يقوما";
            preffix[3, 8] = "تقوما";
            preffix[4, 8] = "يقمن";
            preffix[5, 8] = "يقوموا";


            preffix[0, 9] = "نصيبي";//#9
            preffix[1, 9] = "نصيبي";
            preffix[2, 9] = "نصيبينا";
            preffix[3, 9] = "نصيبينا";
            preffix[4, 9] = "أنصبتنا";
            preffix[5, 9] = "أنصبتنا";


            preffix[0, 10] = "ت";//#*#
            preffix[1, 10] = "";


            preffix[0, 11] = "تنازلت تنازلاً نهائياً";//&&&
            preffix[1, 11] = "تنازلت تنازلاً نهائياً";
            preffix[2, 11] = "تنازلنا تنازلاً نهائياً";
            preffix[3, 11] = "تنازلنا تنازلاً نهائياً";
            preffix[4, 11] = "تنازلنا تنازلاً نهائياً";
            preffix[5, 11] = "تنازلنا تنازلاً نهائياً";


            preffix[0, 12] = "ي";//"%&%
            preffix[1, 12] = "ي";
            preffix[2, 12] = "نا";
            preffix[3, 12] = "نا";
            preffix[4, 12] = "نا";
            preffix[5, 12] = "نا";


            preffix[0, 13] = "نت";//#$#
            preffix[1, 13] = "نت";
            preffix[2, 13] = "نا";
            preffix[3, 13] = "نا";
            preffix[4, 13] = "نا";
            preffix[5, 13] = "نا";

            preffix[0, 14] = "أ";//&^&
            preffix[1, 14] = "إ";
            preffix[2, 14] = "ن";
            preffix[3, 14] = "ن";
            preffix[4, 14] = "ن";
            preffix[5, 14] = "ن";

            preffix[0, 15] = "";
            preffix[1, 15] = "ة";
            preffix[2, 15] = "ين";
            preffix[3, 15] = "تين";
            preffix[4, 15] = "ات";
            preffix[5, 15] = "ين";


            preffix[0, 16] = "اسمي";//$$&
            preffix[1, 16] = "اسمي";
            preffix[2, 16] = "اسمينا";
            preffix[3, 16] = "اسمينا";
            preffix[4, 16] = "اسمائنا";
            preffix[5, 16] = "اسمائنا";

            preffix[0, 17] = "للسيد";//
            preffix[1, 17] = "للسيدة";
            preffix[2, 17] = "لكل من ";
            preffix[3, 17] = "لكل من ";
            preffix[4, 17] = "لكل من ";
            preffix[5, 17] = "لكل من ";

        }

    }
}
