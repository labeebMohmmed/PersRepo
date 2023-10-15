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
using Microsoft.Office.Core;
using static Azure.Core.HttpHeader;
using Aspose.Words.Settings;
using DocumentFormat.OpenXml.Presentation;
using Control = System.Windows.Forms.Control;
using Microsoft.Reporting.WinForms;
using System.Windows.Documents;
using System.Reflection;
using System.Text.RegularExpressions;
using SautinSoft.Document;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using SixLabors.ImageSharp.Drawing;
using DocumentFormat.OpenXml.Drawing;
using System.Data.SqlTypes;

namespace PersAhwal
{
    public partial class FormCollection : Form
    {
        string DataSource = "";
        bool addInfo = false;
        string autheticatingOthes = "";
        string archStat = "";
        string removedStat = "";
        string AuthenticName = "";
        string FilespathIn = "";
        string FilespathOut = "";
        string updateAll = "";
        string[] allList;
        bool AddEdit = true;
        string EmpName = "";
        int genIDNo = 0;
        int AtVCIndex = 0;
        string GregorianDate = "";
        string HijriDate = "";
        bool newData = false;
        string[] colIDs = new string[100];
        string[] boldTexts = new string[100];
        string[] forbidDs = new string[100];
        static string[,] preffix = new string[10, 20];
        string Jobposition = "";
        int currentPanelIndex = 0;
        int intID = 0;
        int InvalidControl = 0;
        int addNameIndex = 0;
        bool ArchData = false;
        string archState = "new";
        string[] foundList;
        string[] checkOptions = new string[5];
        int checkIndex = 0;
        int txtReviewListIndex = 0;
        int txtReviewListIndexStar = 0;
        string StrSpecPur = "";
        string textModel = "";
        string ColName = "";
        string ColRight = "Col0";
        string startText = "";
        string oldText = "";
        Word.Document oBDoc;
        object oBMiss;
        Word.Application oBMicroWord;
        bool notFiled = true;
        string proType1 = "1";
        int ButtonInfoIndex = 0;
        bool LibtnAdd1Vis = false;
        int MessageDocNo = 0;
        int onBehalfIndex = 0;
        string AuthTitle = "نائب قنصل";
        string AuthTitleLast = "نائب قنصل";
        bool goBack = false;
        string[] txtReviewList;
        string[] txtReviewListStar;
        string[] txtPurposeList;
        int txtRigIndex = 0;
        int txtPurIndex = 0;
        string originTextReview = "";
        string originTextPurpose = "";
        string starText = "";
        string ProTitle = "";
        bool gridFill = false;
        string selectTable = "TableCollectStarText";
        int starTextIndex = 0;
        string ageDetected = "";
        int txtPurposeListIndex = 0;
        string txtReviewLast = "";
        int starRightIndexStar = 0;
        int AllowedTimes = 5;

        string appName = "";
        string sex = "";
        string docType = "";
        string docNo = "";
        string docIsuue = "";
        string appJob = "";
        string appBirth = "";
        string appExp = "";
        string wordCheck = "";
        string checkEnd = "";
        string year = "";

        public FormCollection(int allowedTimes, int Atvc, int currentRow, int DocumentType, string empName, string dataSource, string filepathOut, string jobposition, string gregorianDate, string hijriDate)
        {
            InitializeComponent();
            definColumn(dataSource);
            DataSource = dataSource;
            //FilespathIn = filepathIn;
            FilespathOut = filepathOut + @"\";
            year = gregorianDate.Split('-')[2];
            //MessageBox.Show(FilespathOut);
            fillYears(yearSel);
            fillSamplesCodes(dataSource);
            AtVCIndex = Atvc;
            AllowedTimes = allowedTimes;
            EmpName = empName;
            Jobposition = jobposition;
            التاريخ_الميلادي.Text = GregorianDate = gregorianDate;
            التاريخ_الهجري.Text = HijriDate = hijriDate;
            Console.WriteLine("1");
            genPreperations();
            Console.WriteLine("2");
            FillDataGridView(DataSource, year);
            Console.WriteLine("3");
            getMaxRange(DataSource);
            backgroundWorker2.RunWorkerAsync();
            try
            {
                string[] info = missionBasicInfo().Split('*');
                txtArabName.Text = info[0];
                txtEngName.Text = info[1];
                txtMissionAddress.Text = info[2];
                //txtMissionCode.Text = info[3];
            }
            catch (Exception ex)
            {

            }

        }

        private void fillYears(ComboBox combo)
        {
            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            combo.Items.Clear();
            combo.Items.Add("جميع الأعوام");
            string query = "select distinct DATENAME(YEAR, التاريخ) as years from TableGeneralArch order by DATENAME(YEAR, التاريخ) desc";
            SqlConnection Con = new SqlConnection(DataSource.Replace("AhwalDataBase", "ArchFilesDB"));
            if (Con.State == ConnectionState.Closed)
                try
                {
                    Con.Open();
                    SqlDataAdapter sqlDa = new SqlDataAdapter(query, Con);
                    sqlDa.SelectCommand.CommandType = CommandType.Text;
                    DataTable dtbl2 = new DataTable();
                    sqlDa.Fill(dtbl2);
                    sqlCon.Close();
                    foreach (DataRow dataRow in dtbl2.Rows)
                    {
                        if (dataRow["years"].ToString().Length == 4)
                            combo.Items.Add(dataRow["years"].ToString());
                    }
                }
                catch (Exception ex) { }
            
        }

        private string missionBasicInfo()
        {
            string infoDet = "";
            string query = "select بيانات_البعثة from TableSettings";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return ""; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
            }
            catch (Exception ex)
            {
                return "";
            }



            sqlCon.Close();

            foreach (DataRow dataRow in dtbl.Rows)
            {
                try
                {
                    infoDet = dataRow["بيانات_البعثة"].ToString();
                }
                catch (Exception ex)
                {

                }
            }
            return infoDet;
        }
        private void getMaxRange(string dataSource)
        {
            SqlConnection Con = new SqlConnection(dataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select allowedEditCollec  from TableSettings where ID=1", Con);

            try
            {
                if (Con.State == ConnectionState.Closed)
                    Con.Open();
                sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
                var reader = sqlCmd1.ExecuteReader();

                if (reader.Read())
                {
                    //allowedEdit.Text = reader["allowedEditCollec"].ToString();
                }
            }
            catch (Exception ex)
            {
                Con.Close();
            }
        }


        public void FillDataGridView(string dataSource, string year)
        {
            string query = "select * from TableCollection where DATEPART(year,تاريخ_الارشفة1) =" + year +" order by ID";
            if (year == "جميع الأعوام")
                query = "select * from TableCollection order by ID";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            dataGridView1.Columns[0].Visible = false;
            //dataGridView1.Columns["نوع_المعاملة"].Visible = false ;
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[2].Width = 350;
            dataGridView1.Columns[3].Width = 40;
            dataGridView1.Columns[5].Width = dataGridView1.Columns[6].Width = 200;
            sqlCon.Close();
            int bre = 0;
            ColorFulGrid9();

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


            }
            labDescribed.Text = "عدد (" + i.ToString() + ") معاملة .. عدد (" + inComb.ToString() + ") غير مكتمل.. والمؤرشف منها عدد (" + arch.ToString() + ")...";

        }

        public void genPreperations()
        {
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);

            allList = getColList("TableCollection");
            label36.Text = "الموظف:" + EmpName;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            dataGridView1.BackgroundColor = Color.White;
            PanelDataGrid.Size = new System.Drawing.Size(1318, 622);
            PanelDataGrid.Location = new System.Drawing.Point(12, 38);
            //
            PanelDataGrid.BringToFront();
            //
            Suffex_preffixList();
            if (Jobposition.Contains("قنصل"))
            {
                btnDelete.Visible = true;
                نوع_المعاملة.Enabled = نوع_الإجراء.Enabled = طريقة_الطلب.Enabled = اسم_المندوب.Enabled = true;
            }
            else {
                نوع_المعاملة.Enabled = نوع_الإجراء.Enabled = طريقة_الطلب.Enabled = اسم_المندوب.Enabled = false;
            }

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
                    panelAuthRights.Visible = panelAuthRights.Visible = btnPrevious.Visible = panelapplicationInfo.Visible = false;
                    break;
                case 1:
                    //Basic Info                   
                    panelapplicationInfo.Size = new System.Drawing.Size(893, 649); 
                    panelapplicationInfo.Location = new System.Drawing.Point(225, 40); 
                    panelapplicationInfo.BringToFront();
                    btnPrevious.Visible = panelapplicationInfo.Visible = true;
                    //false
                    panelAuthRights.Visible = ListSearch.Visible = btnListView.Visible = panelAuthRights.Visible = PanelDataGrid.Visible = labDescribed.Visible = false;
                    btnDelete.Visible = btnFile1.Visible = btnFile2.Visible = btnFile3.Visible = Panelapp.Visible = true;
                    break;
                case 2:
                    if (نوع_الإجراء.Text.Contains("عامة")||نوع_الإجراء.Text.Contains("توكيل"))
                    {
                        نوع_الإجراء.BackColor = System.Drawing.Color.MistyRose;
                        MessageBox.Show("يرجى اقتراح أو ختيار من القائمة اسم للمعاملة");
                        نوع_الإجراء.Enabled = true;
                        currentPanelIndex--; return;
                    }
                    نوع_الإجراء.Enabled = false;
                    التاريخ_الميلادي.Text = GregorianDate;
                    التاريخ_الهجري.Text = HijriDate;

                    updateGenName(مقدم_الطلب.Text, رقم_المعاملة.Text);
                    relatedProUpdate();
                    صفة_مقدم_الطلب_off.SelectedIndex = Appcases(النوع, addNameIndex);
                    اسم_الموظف.Text = EmpName;
                    if (!checkEmpty(panelapplicationInfo))
                    {
                        currentPanelIndex--; return;
                    }
                    save2DataBase(panelapplicationInfo, "البيانات العامة");

                    if (!checkEmpty(Panelapp) || !checkDate(Panelapp))
                    {
                        currentPanelIndex--; return;
                    }
                    save2DataBase(Panelapp, "مقدم الطلب");


                    if (!checkGender(Panelapp, "مقدم_الطلب_", "النوع_"))
                    {
                        currentPanelIndex--; return;
                    }
                    else addNewAppNameInfo(مقدم_الطلب);

                    if (!طريقة_الطلب.Checked) proType1 = "2";
                    if (!backgroundWorker1.IsBusy) backgroundWorker1.RunWorkerAsync();

                    if (اللغة.Checked)
                        boxesPreparationsEnglish(addNameIndex, نوع_المعاملة.SelectedIndex);
                    else 
                        boxesPreparationsArabic(addNameIndex, نوع_المعاملة.SelectedIndex);


                    //txtReview.Text = writeStrSpecPur();
                    panelAuthRights.Size = new System.Drawing.Size(1315, 622);
                    panelAuthRights.Location = new System.Drawing.Point(10, 36);
                    panelAuthRights.BringToFront();
                    panelAuthRights.Visible = btnNext.Visible = true;
                    PanelDataGrid.Visible = panelapplicationInfo.Visible = false;
                    timer1.Enabled = true;
                    if (LibtnAdd1.Visible && (Vitext1.Text.Contains("_") || Vitext2.Text.Contains("_") || Vitext3.Text.Contains("_") || Vitext4.Text.Contains("_") || Vitext5.Text.Contains("_")))
                    {
                        LibtnAdd1Vis = true;
                        //MessageBox.Show("addvis");
                        fillTextBoxesInvers();
                    }
                    if (ProTitle != "")
                        عنوان_المكاتبة.Text = ProTitle;

                    if (نوع_المعاملة.Text == "إقرار" || نوع_المعاملة.Text == "إقرار مشفوع باليمين")
                    {

                        authJob();
                    }
                    if (txtReview.Text == "")
                    {
                        checkAutoUpdate.Checked = true;
                    }
                    autoCompleteBulk(عنوان_المكاتبة, DataSource, "عنوان_المكاتبة", "TableCollection");
                    طريقة_الطلب.Checked = الشاهد_الأول.Enabled = هوية_الأول.Enabled = true;
                    
                    if (طريقة_الطلب.Text != "حضور مباشرة إلى القنصلية")
                        طريقة_الطلب.Checked = الشاهد_الأول.Enabled = هوية_الأول.Enabled = false;

                    if (PanelButtonInfo.Visible)
                    {
                        PaneltxtReview.Height = 231;
                        PaneltxtReview.AutoScroll = true;
                    }
                    else
                    {
                        PaneltxtReview.Height = 410;
                        PaneltxtReview.AutoScroll = false;
                    }

                    break;
                case 3:

                    if (!اللغة.Checked)
                    {
                        if (checkRelevencey())
                        {
                            MessageBox.Show("المحتوى غير لمطابق لنوع المكاتبة، يرجى إضافة لفظ " + wordCheck);
                            var selectedOption = MessageBox.Show("هل تود تصحيح المحتوى؟(", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (selectedOption == DialogResult.Yes)
                            {
                                currentPanelIndex--;
                                return;
                            }
                        }

                        if (checkEnding())
                        {
                            MessageBox.Show("المحتوى غير لمطابق لنوع المكاتبة، يرجى إضافة لفظ " + checkEnd);
                            var selectedOption = MessageBox.Show("هل تود تصحيح المحتوى؟(", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (selectedOption == DialogResult.Yes)
                            {
                                currentPanelIndex--;
                                return;
                            }
                        }
                    }

                    if (!save2DataBase(PanelItemsboxes, "بيانات الموضوع"))
                    {
                        MessageBox.Show("بيانات الموضوع");
                        currentPanelIndex--; return;
                    }
                    if (!save2DataBase(PaneltxtReview, "موضوع الطلب"))
                    {
                        MessageBox.Show("موضوع الطلب");
                        currentPanelIndex--; return;
                    }

                    if (panelRemove.Visible)
                        if (!checkEmpty(panelRemove) || !save2DataBase(panelRemove, "بيانات الغاء الطلب"))
                        {
                            MessageBox.Show("بيانات الغاء الطلب");
                            currentPanelIndex--; return;
                        }

                    Console.WriteLine(txtReview.Text);
                    //MessageBox.Show(txtReview.Text);
                    txtReview.Text = SuffConvertments(txtReview.Text, صفة_مقدم_الطلب_off.SelectedIndex, 0, true);
                    if (نوع_المعاملة.Text == "مذكرة لسفارة عربية")
                        غرض_المعاملة.Text = "القنصلية العامة ل" + Vitext1.Text + " – جـدة";
                    else if (نوع_المعاملة.Text == "مذكرة لسفارة أجنبية")
                        غرض_المعاملة.Text = "The Consulate General of " + Vitext1.Text + " in Jeddah";

                    غرض_المعاملة.Text = SuffConvertments(غرض_المعاملة.Text, صفة_مقدم_الطلب_off.SelectedIndex, 0, true);

                    if (اللغة.Checked && (نوع_المعاملة.Text == "إقرار" ||نوع_المعاملة.Text == "إقرار مشفوع باليمين"))
                        غرض_المعاملة.Text = "";

                    if (addInfo && ButtonInfoIndex == 0)
                    {
                        currentPanelIndex--;
                        MessageBox.Show("يرجى إضافة البيانات أولا");
                        return;
                    }
                    
                    
                    if (!checkEmpty(PanelItemsboxes))
                    {
                        currentPanelIndex--; return;
                    }
                    timer1.Enabled = false;
                    addNonEmptyFields(PanelItemsboxes);
                    txtReview.Text = removeSpace(txtReview.Text, false);
                    if (Vitext1.Text == "" && Vitext2.Text == "" && Vitext3.Text == "" && Vitext4.Text == "" && Vitext5.Text == "" && ButtonInfoIndex > 0)
                    {
                        fillTextBoxes(Vitext1, 1);

                        fillTextBoxes(Vitext2, 2);
                        fillTextBoxes(Vitext3, 3);
                        fillTextBoxes(Vitext4, 4);
                        fillTextBoxes(Vitext5, 5);
                    }
                    if (عنوان_المكاتبة.Text == "عنوان المكاتبة" && !نوع_المعاملة.Text.Contains("مذكر"))
                    {
                        MessageBox.Show("يرجى إختيار عنوان المكاتبة");
                        currentPanelIndex--; return;
                    }
                    

                    else if (PanelButtonInfo.Visible && ButtonInfoIndex > 0)
                    {

                        Vitext1.Text = Vitext2.Text = Vitext3.Text = Vitext4.Text = Vitext5.Text = "";
                    }
                    //MessageBox.Show(Vitext1.Text);
                    

                    if (goBack)
                    {
                        MessageBox.Show("تعذر الوصول إلى نموذج أولى.. يرجى التواصل مع مدير النظام لإعداده");
                        currentPanelIndex--; return;
                    }
                    finalPanel.Size = new System.Drawing.Size(944, 624);
                    finalPanel.Location = new System.Drawing.Point(192, 38);
                    finalPanel.BringToFront();
                    finalPanel.Visible = true;
                    panelAuthRights.Visible = btnNext.Visible = PanelDataGrid.Visible = panelapplicationInfo.Visible = false;
                    removeSpace(txtReview);


                    string codedText = TextReviewCoding(txtReview.Text);
                    int TotalRows = checkTotalRows(DataSource, "TableCollectStarText");
                    int TotalcolRows = checkTotalcolRows(DataSource, "TableCollectStarText", نوع_المعاملة.Text.Replace(" ", "_") + "_" + نوع_الإجراء.Text.Replace(" ", "_"));
                    Console.WriteLine("checkTotalRows = " + TotalRows);
                    Console.WriteLine("checkTotalcolRows = " + TotalcolRows);




                    if (!checkStarTextExist(DataSource, نوع_المعاملة.Text.Replace(" ", "_") + "_" + نوع_الإجراء.Text.Replace(" ", "_"), codedText, "TableCollectStarText"))
                    {
                        if (TotalcolRows < TotalRows)
                            updateNewText(DataSource, نوع_المعاملة.Text.Replace(" ", "_") + "_" + نوع_الإجراء.Text.Replace(" ", "_"), codedText, "TableCollectStarText", (TotalcolRows + 1).ToString());
                        else
                            TotalcolRows = insertNewText(DataSource, نوع_المعاملة.Text.Replace(" ", "_") + "_" + نوع_الإجراء.Text.Replace(" ", "_"), codedText, "TableCollectStarText");

                    }

                    relatedProUpdate();
                    if (عنوان_المكاتبة.Text == "إفادة لمن يهمه الأمر")
                    {
                        var selectedOption = MessageBox.Show("هل تود تعديل العنوان إلى إفادة (" ,"", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (selectedOption == DialogResult.Yes)
                        {
                            عنوان_المكاتبة.Text = "إفادة";                            
                        }
                    }
                    else if (عنوان_المكاتبة.Text == "TO WHOM IT MAY CONCERN")
                    {
                        var selectedOption = MessageBox.Show("هل تود تعديل العنوان إلى Certificate (", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (selectedOption == DialogResult.Yes)
                        {
                            عنوان_المكاتبة.Text = "Certificate";                            
                        }
                    }
                    if (وجهة_المعاملة.Text == "إختر وجهة المعاملة")
                    {
                        var selectedOption1 = MessageBox.Show("هل الإجراء محلي خاص بالمملكة؟", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (selectedOption1 == DialogResult.Yes)
                        {
                            وجهة_المعاملة.Text = "جدة";
                        }
                    }

                    if (Jobposition.Contains("قنصل"))
                    {
                        btnPrintDocx.Select();
                    }
                    else btnPrintPdf.Select();

                    break;
            }
        }

        private bool checkRelevencey() 
        {
            
            switch (نوع_المعاملة.Text) {
                case "إقرار":
                    wordCheck = SuffConvertments("أقر", صفة_مقدم_الطلب_off.SelectedIndex, 0, true);
                    checkEnd = SuffConvertments("وهذا إقرار مني بذلك، والله على ما اقول شهيد.", صفة_مقدم_الطلب_off.SelectedIndex, 0, true);
                    break;
                case "إقرار مشفوع باليمين":
                    wordCheck = SuffConvertments("أقسم بالله العظيم وأقر", صفة_مقدم_الطلب_off.SelectedIndex, 0, true);
                    checkEnd = SuffConvertments("وهذا إقرار مني بذلك، والله على ما اقول شهيد.", صفة_مقدم_الطلب_off.SelectedIndex, 0, true);
                    break;                
            }
            //MessageBox.Show(wordCheck);
            if (!txtReview.Text.Contains(wordCheck))
                return true;
            else
                return false;
        }
        private bool checkEnding() 
        {
            
            switch (نوع_المعاملة.Text) {
                case "إقرار":
                    checkEnd = SuffConvertments("وهذا إقرار مني بذلك، والله على ما أقول شهيد.", صفة_مقدم_الطلب_off.SelectedIndex, 0, true);
                    break;
                case "إقرار مشفوع باليمين":
                    checkEnd = SuffConvertments("وهذا إقرار مني بذلك، والله على ما أقول شهيد", صفة_مقدم_الطلب_off.SelectedIndex, 0, true);
                    break;                
            }
            if (!txtReview.Text.Contains(checkEnd))
                return true;
            else
                return false;
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
        private string SuffConvertments(string text, int person1, int person2, bool ask)
        {
            string[] words = text.Split(' ');
            checkAutoUpdate.Checked = false;
            foreach (string word in words)
            {
                if (word == "" || word == " ") continue;
                Console.WriteLine(word);
                for (int gridIndex = 0; gridIndex < dataGridView2.Rows.Count - 1; gridIndex++)
                {
                    string code = dataGridView2.Rows[gridIndex].Cells["الرموز"].Value.ToString();
                    string person = dataGridView2.Rows[gridIndex].Cells["الضمير"].Value.ToString();
                    Console.WriteLine(person);
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
                        if (word == replacemests[cellIndex] || word == replacemests[cellIndex] + "،" || word.Contains(code))
                        {
                            Console.WriteLine(word);
                            if (person == "1")
                            {
                                if (word != replacemests[person1] && word != replacemests[person1] + "،")
                                {

                                    //if (ask)
                                    //{
                                        var selectedOption = MessageBox.Show("هل تود إجراء التصحيح التلقائي (" + replacemests[person1] + ")", "تم رصد خطاء في الصياغة (" + word + ")", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                        if (selectedOption == DialogResult.Yes)
                                        {
                                            text = text.Replace(word, replacemests[person1]);
                                        
                                            break;
                                        }
                                }
                                return text;
                            }
                            else if (person == "2")
                            {
                                if (word != replacemests[person2] && word != replacemests[person2] + "،")
                                {
                                    var selectedOption = MessageBox.Show("هل تود إجراء التصحيح التلقائي (" + replacemests[person2] + ")", "تم رصد خطاء في الصياغة (" + word + ")", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                    
                                        if (selectedOption == DialogResult.Yes)
                                        {
                                            text = text.Replace(word, replacemests[person2]);
                                            break;
                                        }
                                    
                                }
                            }
                            else if (person == "3")
                            {
                                if (word != replacemests[person1] && word != replacemests[person1] + "،")
                                {
                                    if (ask)
                                    {
                                        var selectedOption = MessageBox.Show("هل تود إجراء التصحيح التلقائي (" + replacemests[person1] + ")", "تم رصد خطاء في الصياغة (" + word + ")", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                        if (selectedOption == DialogResult.Yes)
                                        {
                                            text = text.Replace(word, replacemests[person1]);
                                            break;
                                        }
                                    }
                                    else {
                                        text = text.Replace(word, replacemests[person1]);
                                        break;
                                    }
                                }
                            }

                            else if (person == "5")
                            {
                                //
                                if (word.Contains(code))
                                {
                                    text = text.Replace(code, replacemests[person1]);
                                    break;
                                }
                            }
                            else if (person == "*")
                            {
                                //
                                if (word.Contains(code))
                                {
                                    text = text.Replace(word, replacemests[person1]);
                                    break;
                                }
                            }

                            else if (person == "#")
                            {
                                if (word != replacemests[person1] && word != replacemests[person1] + "،")
                                {
                                    text = text.Replace(word, replacemests[person1]);
                                    break;
                                }
                            }
                        }
                    }

                }
            }
            return text;
        }

        private void updateProType(string table, ComboBox altColName, ComboBox altSubColName, string proName, string proID)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string query = "update "+ table+" set "+ altColName.Name + " =N'" + altColName.Text + "', "+ altSubColName.Name + " = N'"+ altSubColName.Text+"' where "+ proName + " = N'" + proID + "'";
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
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

            relatedPro[9] = "تاريخ_الميلاد";
            relatedPro[10] = "المهنة";
            relatedPro[11] = "طريقة_الإجراء";
            relatedPro[12] = "رقم_الوكالة";
            relatedPro[13] = "جهة_إصدار_الوكالة";
            relatedPro[14] = "تاريخ_إصدار_الوكالة";
            relatedPro[15] = "مقدم_الطلب";
            string values = "النوع=@النوع";
            for (int x = 1; x < 16; x++)
            {
                values = values + ", " + relatedPro[x] + "=@" + relatedPro[x];
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
            for (int i = 0; i < 16; i++)
            {
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
                }
                foreach (Control control in finalPanel.Controls)
                {
                    if (control.Name == relatedPro[i])
                    {
                        Console.WriteLine(relatedPro[i] + " - " + control.Text);
                        sqlCmd.Parameters.AddWithValue("@" + relatedPro[i], control.Text);
                    }
                }
                foreach (Control control in panelAuthen.Controls)
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
            if (رقم_المرجع_المرتبط_off.Text == "") return false;
            string query = "SELECT مقدم_الطلب FROM TableCollection where رقم_المعاملة =N'" + رقم_المرجع_المرتبط_off.Text + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
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


        private int insertNewText(string dataSource, string col, string text, string genTable)
        {
            string query = "INSERT INTO " + genTable + " (" + col + ")  values (N'" + text + "') ;SELECT @@IDENTITY as lastid";
            Console.WriteLine(query);
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return 0; }
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            try
            {
                var reader = sqlCmd.ExecuteReader();
            
            Console.WriteLine(query);
            //MessageBox.Show(query);
            if (reader.Read())
            {
                return Convert.ToInt32(reader["lastid"].ToString());
            }
            }
            catch (Exception ex) { return 0; }
            return 0;
        }

        private void updateNewText(string dataSource, string col, string text, string genTable, string ID)
        {
            string query = "update " + genTable + " set " + col + "=N'" + text + "' where ID=" + ID;
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            //MessageBox.Show("update " + text);
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            try
            {
                sqlCmd.ExecuteNonQuery();
            } catch (Exception ex) { return; }
            sqlCon.Close();
        }



        private bool checkStarTextExist(string dataSource, string col, string text, string genTable)
        {
            string query = "select * from " + genTable + " where " + col + "=N'" + text + "'";
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
            catch (Exception ex) {
                return false;
            }
            if (dtbl.Rows.Count > 0) return true;
            else return false;
            sqlCon.Close();
        }
        private int checkTotalRows(string dataSource, string genTable)
        {
            string query = "select * from " + genTable;
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
            catch (Exception ex) { return 0; }

            return dtbl.Rows.Count;
            sqlCon.Close();
        }

        private int checkTotalcolRows(string dataSource, string genTable, string col)
        {
            string query = "select * from " + genTable + " where " + col + " is not null";
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
            catch (Exception ex) { return 0; }

            return dtbl.Rows.Count;
            sqlCon.Close();
        }

        public void boxesPreparationsEnglish(int index, int proTypeIndex)
        {
            //addNameIndex

            switch (نوع_المعاملة.Text)
            {
                case "إقرار":
                    صفة_مقدم_الطلب_off.SelectedIndex = Appcases(النوع, index);
                    //إقرار 
                    if (index == 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "I. the undersigned,";
                        نص_مقدم_الطلب1_off.Text = "";// + مقدم_الطلب.Text + "holding a " + نوع_الهوية.Text + " No. " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P")+ " issued on " + مكان_الإصدار.Text + " solemnly declare and affirm that, ";
                    }
                    else if (index > 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "We, the undersigned:";
                        نص_مقدم_الطلب1_off.Text = "";// + مقدم_الطلب.Text + "holding a " + نوع_الهوية.Text + " No. " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " issued on " + مكان_الإصدار.Text + " solemnly declare and affirm that, ";
                        //MessageBox.Show(نص_مقدم_الطلب0_off.Text);
                        return;
                    }

                    موقع_المعاملة_off.Text = " " + مقدم_الطلب.Text.Trim();
                    break;
                case "إفادة لمن يهمه الأمر":
                    // افادة وشهادة لمن يهمه الامر
                    if (index == 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "";
                        نص_مقدم_الطلب1_off.Text = "";// + مقدم_الطلب.Text + "holding a " + نوع_الهوية.Text + " No. " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " issued on " + مكان_الإصدار.Text;
                    }
                    else if (index > 1)
                    {
                        نص_مقدم_الطلب0_off.Text = " Sudanese citizen mentioned below:";
                        نص_مقدم_الطلب1_off.Text = "";
                    }
                    if(غرض_المعاملة.Text == "")
                        غرض_المعاملة.Text = "This certificate has been issued upon " + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 18] + " request,,,";
                    break;
                case "مذكرة لسفارة أجنبية":
                    نص_مقدم_الطلب1_off.Text = "The Consulate General of the Republic of Sudan avails itself this opportunity to renew to the esteemed Consulate the assurances of its highest consideration.";
                    غرض_المعاملة.Text = "The Consulate General of " + Vitext1.Text + " in Jeddah";
                    break;

            }
        }
        public void boxesPreparationsArabic(int index, int proTypeIndex)
        {
            if (صفة_مقدم_الطلب_off.SelectedIndex < 0) صفة_مقدم_الطلب_off.SelectedIndex = 0;
            for (int x = 0; x < 100; x++)
                boldTexts[x] = "";
            boldTexts[0] = موقع_المعاملة.Text.Trim();
            boldTexts[1] = موقع_المعاملة_off.Text.Trim();
            for (int x = 10; x < مقدم_الطلب.Text.Split('_').Length; x++)
                boldTexts[x] = مقدم_الطلب.Text.Split('_')[x - 10];

            for (int x = 30; x < مقدم_الطلب.Text.Split('_').Length; x++)
                boldTexts[x] = مقدم_الطلب.Text.Split('_')[x - 30];
            //addNameIndex


            switch (نوع_المعاملة.Text)
            {
                case "إقرار":
                    صفة_مقدم_الطلب_off.SelectedIndex = Appcases(النوع, index);
                    //إقرار 
                    if (index == 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "أنا المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5];
                        نص_مقدم_الطلب1_off.Text = "/ " + مقدم_الطلب.Text + "، المقيم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " بالمملكة العربية السعودية حامل" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " إصدار " + مكان_الإصدار.Text + "، وبكامل قوا" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " العقلية وبطوع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " واختيار" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " وحالت" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " المعتبرة شرعا وقانونا ";
                        boldTexts[0] = مقدم_الطلب.Text;
                    }
                    else if (index > 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "نحن المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " الموقع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " أعلاه";
                        نص_مقدم_الطلب1_off.Text = "والمقيم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " بالمملكة العربية السعودية، وبكامل قوا" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " العقلية وبطوع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " واختيار" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " وحالت" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " المعتبرة شرعا وقانونا ";
                    }

                    موقع_المعاملة_off.Text = موقع_المعاملة.Text.Trim();

                    التوقيع_off.Text = مقدم_الطلب.Text;

                    if (نوع_الإجراء.Text == "إقرار بالإتفاق")
                    {
                        التوقيع_off.Text = مقدم_الطلب.Text.Split('_')[0] + Environment.NewLine + "توقيع "
                            + مقدم_الطلب.Text.Split('_')[1] + "/ ـ..................................";
                    }

                    break;
                case "إقرار مشفوع باليمين":
                    صفة_مقدم_الطلب_off.SelectedIndex = Appcases(النوع, index);
                    //إقرار مشفوع باليمين 
                    if (index == 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "أنا المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5];
                        نص_مقدم_الطلب1_off.Text = "/ " + مقدم_الطلب.Text + "، المقيم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " بالمملكة العربية السعودية حامل" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " إصدار " + مكان_الإصدار.Text + "، وبكامل قوا" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " العقلية وبطوع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " واختيار" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " وحالت" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " المعتبرة شرعا وقانونا ";
                    }
                    else if (index > 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "نحن المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " الموقع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " أعلاه";
                        نص_مقدم_الطلب1_off.Text = "والمقيم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " بالمملكة العربية السعودية، وبكامل قوا" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " العقلية وبطوع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " واختيار" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " وحالت" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12] + " المعتبرة شرعا وقانونا ";
                    }

                    موقع_المعاملة_off.Text = موقع_المعاملة.Text.Trim();
                    //                    MessageBox.Show(موقع_المعاملة_off.Text);
                    التوقيع_off.Text = مقدم_الطلب.Text;
                    if (نوع_الإجراء.Text == "إقرار بالإتفاق")
                    {
                        التوقيع_off.Text = مقدم_الطلب.Text.Split('_')[0] + Environment.NewLine + "توقيع "
                            + مقدم_الطلب.Text.Split('_')[1] + "/ ـ..................................";
                    }
                    //MessageBox.Show("التوقيع_off " + التوقيع_off.Text);
                    break;
                case "إفادة لمن يهمه الأمر":
                    // افادة وشهادة لمن يهمه الامر
                    if (index == 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "";// " المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " السواني" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " السيد" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5];
                        نص_مقدم_الطلب1_off.Text = "";// " + مقدم_الطلب.Text + "، المقيم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " بالمملكة العربية السعودية حامل" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " إصدار " + مكان_الإصدار.Text + "،";
                    }
                    else if (index > 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "";// المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " السوداني" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " المذكور" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " أدناه:";
                        نص_مقدم_الطلب1_off.Text = "";
                    }
                    غرض_المعاملة.Text = "حررت هذه الإفادة بناء على طلب المذكور" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " أعلاه لاستخدامها على الوجه المشروع";

                    break;
                case "شهادة لمن يهمه الأمر":
                    // افادة وشهادة لمن يهمه الامر
                    if (index == 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "";// "المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + "السواني" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + "السيد" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5];
                        نص_مقدم_الطلب1_off.Text = "";// "/ " + مقدم_الطلب.Text + "، المقيم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " بالمملكة العربية السعودية حامل" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " إصدار " + مكان_الإصدار.Text + "،";
                    }
                    else if (index > 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "";// "المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " السوداني" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " المذكور" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " أدناه:";
                        نص_مقدم_الطلب1_off.Text = "";
                    }
                    غرض_المعاملة.Text = "حررت هذه الشهادة بناء على طلب المذكور" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 5] + " أعلاه لاستخدامها على الوجه المشروع";

                    break;
                case "مذكرة لسفارة عربية":
                    نص_مقدم_الطلب1_off.Text = "تنتهز القنصلية العامة لجمهورية السودان بجدة هذه السانحة لتُعرب للقنصلية العامة ل" + Vitext1.Text + " بجدة عن فائق شكرها وتقديرها واحترامها.";
                    نص_مقدم_الطلب1_off.Text = SuffConvertments(نص_مقدم_الطلب1_off.Text, صفة_مقدم_الطلب_off.SelectedIndex, 0, false);
                    غرض_المعاملة.Text = "القنصلية العامة ل" + Vitext1.Text + " – جـدة";
                    break;

            }
            string auth = "";
            //string witnesses = "";
            //if (الشاهد_الأول.Text != "" && الشاهد_الأول.Text != "") witnesses = " في حضور الشهود المذكورين أعلاه ";
            //auth = " المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15] + " المذكور" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15] + " أعلاه قد حضر" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " ووقع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " بتوقيع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4] + " على هذا الإقرار " + witnesses + " وذلك بعد تلاوته علي" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4] + " وبعد أن فهم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " مضمونه ومحتواه";
            //if (!طريقة_الطلب.Checked)
            //    auth = " المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15] + " المذكور" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15] + " أعلاه قد وقع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " بتوقيع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4] + " على هذا التوكيل في حضور الشاهدين المذكورين أعلاه وذلك بعد تلاوته علي" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4] + " وبعد أن فهم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " مضمونه ومحتواه";
            //if (!اسم_المندوب.Visible)
            //{
            //    if (طريقة_الإجراء.Checked)
            //        التوثيق_off.Text = AuthTitle + " بالقنصلية العامة لجمهورية السودان بجدة، بأن" + auth + "، صدر تحت توقيعي وختم القنصلية العامة";
            //    else
            //    {
            //        auth = " بأن المواطن" + preffix[onBehalfIndex, 5] + " /" + اسم_الموكل_بالتوقيع.Text + " قد حضر" + preffix[onBehalfIndex, 3] + " ووقع" + preffix[onBehalfIndex, 3] + " بتوقيع" + preffix[onBehalfIndex, 4] + " على هذا الإقرار في حضور الشهود المذكورين أعلاه بعد تلاوته علي" + preffix[onBehalfIndex, 4] + " وبعد أن فهم" + preffix[onBehalfIndex, 3] + " مضمونه ومحتواه، وذلك بناءً على الحق الممنوح لها بموجب التوكيل الصادر عن " + جهة_إصدار_الوكالة.Text + " بالرقم " + رقم_الوكالة.Text + " بتاريخ " + تاريخ_إصدار_الوكالة.Text;
            //        //التوقيع_off.Text = اسم_الموكل_بالتوقيع.Text;
            //        التوثيق_off.Text = auth + "، صدر تحت توقيعي وختم القنصلية العامة";
            //    }
            //}
            //else التوثيق_off.Text = auth + "، صدر تحت توقيعي وختم القنصلية العامة";

            if (نوع_المعاملة.Text == "إقرار" || نوع_المعاملة.Text == "إقرار مشفوع باليمين")
                authJob();
            //MessageBox.Show("التوقيع_off " + التوقيع_off.Text);
        }


        public int Appcases(TextBox text, int index)
        {
            if (index == 1)
            {
                if (النوع.Text == "ذكر")
                    return 0;//المقيم
                else
                    return 1;//المقيمة
            }

            else if (index == 2)
            {
                if (text.Text.Split('_')[0] != "ذكر" && text.Text.Split('_')[0] != "ذكر")
                    return 3;//المقيمتان
                else
                    return 2;//المقيمان
            }

            else if (index == 3)
            {
                if (text.Text.Split('_')[0] != "ذكر" && text.Text.Split('_')[0] != "ذكر" && text.Text.Split('_')[0] != "ذكر")
                    return 4;//المقيمات
            }

            return 5;//المقيمون
        }
        private void addNonEmptyFields(FlowLayoutPanel panel)
        {
            foreach (Control Econtrol in panel.Controls)
            {
                if ((Econtrol is TextBox || Econtrol is ComboBox || Econtrol is CheckBox) && Econtrol.Visible && !checkColumnName(Econtrol.Name, DataSource))
                {
                    CreateColumn(Econtrol.Name, DataSource);
                }
            }
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
                    if (control.Text == "" || control.Text == "P0" || control.Text.Contains("إختر"))
                        if (control.Name != "الشاهد_الأول" && control.Name != "الشاهد_الثاني" && control.Name != "هوية_الأول" && control.Name != "هوية_الثاني")
                        {
                            control.BackColor = System.Drawing.Color.MistyRose;
                            if (panel.Name == "Panelapp") { panel.Height = 130 * addNameIndex; }
                            MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل " + control.Name);
                            return false;
                        }
                        else if ((control.Name == "هوية_الأول" && الشاهد_الأول.Text != "") || (control.Name == "هوية_الثاني" && الشاهد_الثاني.Text != ""))
                        {
                            control.BackColor = System.Drawing.Color.MistyRose;
                            if (panel.Name == "Panelapp") { panel.Height = 130 * addNameIndex; }
                            MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل " + control.Name);
                            return false;
                        }                        
                }
            }
            return true;
        }
        
        private bool checkDate(FlowLayoutPanel panel)
        {
            foreach (Control control in panel.Controls)
            {
                if (!control.Visible) continue;
                if ((control.Name == "Vitext2" || control.Name == "Vitext3" || control.Name == "Vitext4" || control.Name == "Vitext5" || control.Name == "Vitext1") && ButtonInfoIndex != 0)
                    continue;
                if (control is TextBox || control is ComboBox)
                {
                    if (control.Name.Contains("انتهاء_الصلاحية") || control.Name.Contains("تاريخ_الميلاد") )
                        {
                            Console.WriteLine(control.Name +" - "+ control.Text);   
                            if (control.Text.Length != 10)
                            {
                                MessageBox.Show("لا يمكن المتابعة يرجى كتابة تاريخ " + control.Name.Replace("_", " ") + " بشكل صحيح");
                                control.BackColor = System.Drawing.Color.MistyRose;
                                return false;
                            }
                            else {
                                int month = Convert.ToInt32(SpecificDigit(control.Text, 4, 5));
                                if (month > 12)
                                {
                                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                                    //textBox.Text = "";
                                    control.Text = SpecificDigit(control.Text, 7, 10);
                                    control.BackColor = System.Drawing.Color.MistyRose;
                                    return false;
                                }

                            }                            
                            if (panel.Name == "Panelapp") { panel.Height = 130 * addNameIndex; }                            
                        }
                }
            }
            return true;
        }

        private bool save2DataBase(FlowLayoutPanel panel, string comment)
        {
            string query = checkList(panel, allList, "TableCollection");
            //MessageBox.Show(query);
            if (query == "UPDATE TableCollection SET  where ID = @id") return true;
            Console.WriteLine(panel.Name + " - " + query);
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
                    Console.WriteLine(foundList[i] +" - "+ commentInfo(comment));
                    //MessageBox.Show(foundList[i] +" - "+ commentInfo());
                    sqlCommand.Parameters.AddWithValue("@" + foundList[i], commentInfo(comment));
                }
                else
                    foreach (Control control in panel.Controls)
                    {
                        string name = control.Name;
                        if (name == foundList[i])
                        {
                            sqlCommand.Parameters.AddWithValue("@" + foundList[i], control.Text);
                            Console.WriteLine(i.ToString() + " " + foundList[i] + " - " + control.Text);
                            break;
                        }
                    }
            }
            sqlCommand.ExecuteNonQuery();
            return true;
        }
        //private bool save2DataBase(FlowLayoutPanel panel)
        //{
        //    string query = checkList(panel, allList, "TableCollection");
        //    //MessageBox.Show(query);
        //    if (query == "UPDATE TableCollection SET  where ID = @id") return true;
        //    Console.WriteLine(panel.Name + " - " + query);
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
        //            sqlCommand.Parameters.AddWithValue("@" + foundList[i], commentInfo());
        //        }
        //        else
        //            foreach (Control control in panel.Controls)
        //            {
        //                string name = control.Name;
        //                //if (panel.Name == "PanelItemsboxes")
        //                //    name = name.Replace("V", "");
        //                if (name == foundList[i])
        //                {
        //                    //if (control.Name == "اسم_المندوب" && control.Visible && !control.Text.Contains("-"))
        //                    //{
        //                    //    control.BackColor = System.Drawing.Color.MistyRose;
        //                    //    MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل اسم المندوب ومنطقة التغطية مفصولين بالعلامة -");
        //                    //    return false;
        //                    //}
        //                    //if (control.Visible && ((control is TextBox && control.Text == "") || (control is ComboBox && control.Text.Contains("إختر"))))
        //                    //    foreach (Control Econtrol in panel.Controls)
        //                    //    {
        //                    //        if ((Econtrol is TextBox || Econtrol is ComboBox) && control.Text == "")
        //                    //            if (panel.Name != "PanelItemsboxes" || (Econtrol.Name != control.Name && Econtrol.Name.Contains(control.Name)) || Econtrol.Name == "اسم_المندوب")
        //                    //            {
        //                    //                //MessageBox.Show(Econtrol.Name + " - " + control.Name);
        //                    //                if (control.Name == "اسم_المندوب" && control.Visible)
        //                    //                {
        //                    //                    //
        //                    //                    control.BackColor = System.Drawing.Color.MistyRose;
        //                    //                    MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل اسم_المندوب ");
        //                    //                    return false;
        //                    //                }
        //                    //                else if (control.Name != "اسم_المندوب" && control.Name != "الشاهد_الأول" && control.Name != "الشاهد_الثاني")
        //                    //                {
        //                    //                    Econtrol.BackColor = System.Drawing.Color.MistyRose;
        //                    //                    if (panel.Name == "Panelapp") { panel.Height = 130 * addNameIndex; }
        //                    //                    MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل " + control.Name.Replace("_", " "));
        //                    //                    return false;
        //                    //                }
        //                    //            }
        //                    //            else if (panel.Name == "PanelItemsboxes")
        //                    //            {
        //                    //                if (control.Visible)
        //                    //                {
        //                    //                    control.BackColor = System.Drawing.Color.MistyRose;
        //                    //                    MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل غير المكتمل");
        //                    //                    return false;
        //                    //                }
        //                    //            }
        //                    //    }

        //                    //if (panel.Name == "panelapplicationInfo") MessageBox.Show(control.Text);
        //                    sqlCommand.Parameters.AddWithValue("@" + foundList[i], control.Text);
        //                    break;
        //                }
        //            }
        //    }
        //    sqlCommand.ExecuteNonQuery();
        //    return true;
        //}
        private string commentInfo(string commentLoc)
        {
            string comment = "";
            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text == "")
                comment = "قام  " + EmpName + " بإدخال البيانات " + commentLoc + Environment.NewLine + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine;

            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text != "")
                comment = "قام  " + EmpName + " ببعض التعديلات " + commentLoc + Environment.NewLine + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine + التعليقات_السابقة_Off.Text;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text == "")
                comment = تعليق_جديد_Off.Text.Trim() + commentLoc + Environment.NewLine + "قام  " + EmpName + " ببعض التعديلات " + Environment.NewLine + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text != "")
                comment = تعليق_جديد_Off.Text.Trim() + commentLoc + Environment.NewLine + "قام  " + EmpName + " ببعض التعديلات " + Environment.NewLine + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine + "*" + التعليقات_السابقة_Off.Text.Trim();

            return comment;
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
                    {
                        //MessageBox.Show(List[col]);
                        if (control.Name == List[col])
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
                //if (panel.Name == "PanelItemsboxes")
                //    name = name.Replace("V", "");
                if (control is TextBox || control is ComboBox || control is CheckBox)
                    for (int col = 0; col < List.Length; col++)
                        if (name == List[col])
                        {
                            foundList[found] = name;
                            //if (panel.Name == "panelapplicationInfo") MessageBox.Show(foundList[found]);
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
            //MessageBox.Show(updateValues);
            return updateAll = "UPDATE " + table + " SET " + updateValues + " where ID = @id";
        }
        private void Suffex_preffixList()
        {


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
            preffix[5, 6] = "رين";


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


            preffix[0, 9] = "نصيبي";//#6
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

            preffix[0, 17] = "للسيد";//$$&
            preffix[1, 17] = "للسيدة";
            preffix[2, 17] = "لكل من ";
            preffix[3, 17] = "لكل من ";
            preffix[4, 17] = "لكل من ";
            preffix[5, 17] = "لكل من ";

            preffix[0, 18] = "his";//$$&
            preffix[1, 18] = "her";
            preffix[2, 18] = "their";
            preffix[3, 18] = "their";
            preffix[4, 18] = "their";
            preffix[5, 18] = "their";

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
            string updateValues = "";
            foreach (DataRow row in dtbl.Rows)
            {
                if (row["name"].ToString() != "ID")
                {

                    allList[i] = row["name"].ToString();
                    //MessageBox.Show(allList[i]);
                    if (i == 0)
                    {
                        updateValues = row["name"].ToString() + "=@" + row["name"].ToString();
                    }
                    else
                    {
                        updateValues = updateValues + "," + row["name"].ToString() + "=@" + row["name"].ToString();
                    }
                    i++;
                }
            }
            updateAll = "UPDATE " + table + " SET " + updateValues + " where ID = @id";
            //MessageBox.Show(updateAll);
            return allList;

        }


        private void definColumn(string dataSource)
        {
            DataSource = dataSource;
            for (int index = 0; index < 100; index++)
                forbidDs[index] = "";

            forbidDs[0] = "تعليق";            
            forbidDs[2] = "sms";
            foreach (System.Windows.Forms.Control control in panelapplicationInfo.Controls)
            {
                if ((control is TextBox || control is ComboBox || control is CheckBox) && !control.Name.Contains("ff"))
                {
                    if (!checkColumnName(control.Name, DataSource))
                    {
                        CreateColumn(control.Name, DataSource);
                    }
                }
            }
            for (int index = 0; forbidDs[index] != ""; index++)
            {
                if (!checkColumnName(forbidDs[index].Replace(" ", "_"), DataSource))
                {
                    CreateColumn(forbidDs[index].Replace(" ", "_"), DataSource);
                }
            }
        }
        private bool checkColumnName(string colNo, string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableCollection", sqlCon);
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
                }
                catch (Exception ex) { }
            return false;
        }
        private void CreateColumn(string Columnname, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try { sqlCon.Open();
                    SqlCommand sqlCmd = new SqlCommand("alter table "+ table+" add " + Columnname.Replace(" ", "_") + " nvarchar(max)", sqlCon);
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.ExecuteNonQuery();
                    sqlCon.Close();
                }
                catch (Exception ex) { }
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {


            intID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            addNameIndex = 0;
            checkAutoUpdate.Checked = false;
            if (dataGridView1.CurrentRow.Index != -1)
            {
                gridFill = true;
                fillInfo(Panelapp, true);
                //MessageBox.Show(مقدم_الطلب.Text);
                string appJob, appBirth;
                try
                {
                    appJob = المهنة.Text.Split('_')[0];
                    appBirth = تاريخ_الميلاد.Text.Split('_')[0];
                }
                catch (Exception ex)
                {
                    appBirth = appJob = "";
                }

                //fillFirstInfo(مقدم_الطلب.Text.Split('_')[0], النوع.Text.Split('_')[0], نوع_الهوية.Text.Split('_')[0], رقم_الهوية.Text.Split('_')[0], مكان_الإصدار.Text.Split('_')[0], "العربية", المهنة.Text.Split('_')[0], تاريخ_الميلاد.Text.Split('_')[0], "0");
                //MessageBox.Show(مقدم_الطلب.Text);


                if (مقدم_الطلب.Text == "")
                {
                    addNames("", "ذكر", "جواز سفر", "P0", "", "العربية", appJob, appBirth, "");
                    ArchData = true;
                }
                for (int app = 0; app < مقدم_الطلب.Text.Split('_').Length; app++)
                {
                    //MessageBox.Show(مقدم_الطلب.Text.Split('_')[app]);

                    try
                    {
                        appJob = المهنة.Text.Split('_')[app];
                        appBirth = تاريخ_الميلاد.Text.Split('_')[app];
                    }
                    catch (Exception ex)
                    {
                        appBirth = appJob = "";
                    }
                    string[] sex = new string[مقدم_الطلب.Text.Split('_').Length];
                    for (int i = 0; i < sex.Length; i++) sex[i] = "ذكر";
                    for (int i = 0; i < النوع.Text.Split('_').Length; i++) sex[i] = النوع.Text.Split('_')[i];
                    if (مقدم_الطلب.Text.Split('_')[app] != "")
                    {
                        addNames(مقدم_الطلب.Text.Split('_')[app], sex[app], نوع_الهوية.Text.Split('_')[app], رقم_الهوية.Text.Split('_')[app], مكان_الإصدار.Text.Split('_')[app], "العربية", appJob, appBirth, انتهاء_الصلاحية.Text.Split('_')[app]);
                        archState = "old";
                    }
                    else
                    {
                        archState = "new";
                        //MessageBox.Show(File.ReadAllText(FilespathOut + "autoDocs.txt"));
                        if (File.ReadAllText(FilespathOut + @"\autoDocs.txt") == "Yes")
                            FillDatafromGenArch("data1", intID.ToString(), "TableCollection");
                    }
                }
                ageDetected = تاريخ_الميلاد.Text;
                fillInfo(PaneltxtReview, false);
                fillInfo(panelapplicationInfo, false);
                fillInfo(PanelItemsboxes, false);
                fillInfo(finalPanel, false);
                currentPanelIndex = 1;
                panelShow(currentPanelIndex);
            }

            checkChanged(مقدم_الطلب, Panelapp);
            checkChanged(النوع, Panelapp);
            checkChanged(نوع_الهوية, Panelapp);
            checkChanged(رقم_الهوية, Panelapp);
            checkChanged(مكان_الإصدار, Panelapp);
            checkChanged(تاريخ_الميلاد, Panelapp);
            checkChanged(انتهاء_الصلاحية, Panelapp);
            checkChanged(المهنة, Panelapp);
            gridFill = false;
            infoCapture();
        }
        private void AppName_MouseClick(object sender, MouseEventArgs e)
        {            
            //appName = مقدم_الطلب.Text;
            //sex = النوع.Text;
            //docType = نوع_الهوية.Text;
            //docNo = رقم_الهوية.Text;
            //docIsuue = مكان_الإصدار.Text;
            //appJob = المهنة.Text;
            //appBirth = تاريخ_الميلاد.Text;
        }
        
        private void infoCapture()
        {            
            appName = مقدم_الطلب.Text;
            sex = النوع.Text;
            docType = نوع_الهوية.Text;
            docNo = رقم_الهوية.Text;
            docIsuue = مكان_الإصدار.Text;
            appJob = المهنة.Text;
            appBirth = تاريخ_الميلاد.Text;
            appExp = انتهاء_الصلاحية.Text;
        }
        public void fillFirstInfo(string name, string sex, string docType, string docNo, string docIssue, string language, string job, string age, string ID, string exp)
        {
            foreach (Control control in Panelapp.Controls)
            {
                if (control.Name == "مقدم_الطلب_" + ID + "." && gridFill)
                    control.Text = name;
                if (control.Name == "النوع_" + ID + ".")
                {
                    control.Text = sex;
                    if (control.Text == "ذكر")
                        ((CheckBox)control).Checked = true;
                    else ((CheckBox)control).Checked = false;                    
                }
                
                if (control.Name == "نوع_الهوية_" + ID + ".")
                    control.Text = docType;
                if (control.Name == "رقم_الهوية_" + ID + ".")
                    control.Text = docNo;
                if (control.Name == "مكان_الإصدار_" + ID + ".")
                    control.Text = docIssue;
                if (control.Name == "المهنة_" + ID + ".")
                    control.Text = job;
                if (control.Name == "تاريخ_الميلاد_" + ID + ".")
                    control.Text = age;
                if (control.Name == "انتهاء_الصلاحية_" + ID + ".")
                    control.Text = exp;
            }
        }
        private bool FillDatafromGenArch(string doc, string id, string table)
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
            if (dtbl.Rows.Count > 0)
                return true;
            else return false;
        }

        private bool FillDatafromGenArchrelated(string id, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource.Replace("AhwalDataBase", "ArchFilesDB"));
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableGeneralArch where  رقم_المرجع_المرتبط_off='" + id + "' and docTable='" + table + "'", sqlCon);
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
            if (dtbl.Rows.Count > 0)
                return true;
            else return false;
        }
        public void addNames(string name, string sex, string docType, string docNo, string docIssue, string language, string job, string age, string Expire_date)
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
            AppName.MouseClick += new System.Windows.Forms.MouseEventHandler(this.AppName_MouseClick);
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
            checkSexType.TextChanged += new System.EventHandler(this.checkSexType_TextChanged); 
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
            //combTitle.Name = "title_" + addNameIndex + ".";
            //combTitle.Size = new System.Drawing.Size(40, 35);
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
            "بطاقة قومية",
            "Passport",
            "Saudi Resident Identity"
            });
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
            DocNo.Size = new System.Drawing.Size(200, 35);
            DocNo.TabIndex = 120;
            DocNo.Tag = "pass";
            DocNo.Text = docNo;
            DocNo.TextChanged += new System.EventHandler(this.DocNo_TextChanged);

            // 
            // label8
            // 
            Label labelExpDate = new Label();
            labelExpDate.AutoSize = true;
            labelExpDate.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labelExpDate.Location = new System.Drawing.Point(371, 41);
            labelExpDate.Name = "label7_" + addNameIndex + ".";
            labelExpDate.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labelExpDate.Size = new System.Drawing.Size(87, 27);
            labelExpDate.TabIndex = 121;
            labelExpDate.Text = "تاريخ انتهاء الصلاحية:";
            // 
            // DocIssue1
            // 
            TextBox DocExpireDate = new TextBox();
            DocExpireDate.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            DocExpireDate.Location = new System.Drawing.Point(152, 44);
            DocExpireDate.Name = "انتهاء_الصلاحية_" + addNameIndex + ".";
            DocExpireDate.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            DocExpireDate.Size = new System.Drawing.Size(100, 35);
            DocExpireDate.TabIndex = 118;
            DocExpireDate.Text = Expire_date;
            DocExpireDate.TextChanged += new System.EventHandler(this.ExpireDate_TextChanged);
            // 
            // label7
            // 
            Label labeldocIssue = new Label();
            labeldocIssue.AutoSize = true;
            labeldocIssue.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labeldocIssue.Location = new System.Drawing.Point(371, 41);
            labeldocIssue.Name = "label8_" + addNameIndex + ".";
            labeldocIssue.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labeldocIssue.Size = new System.Drawing.Size(87, 27);
            labeldocIssue.TabIndex = 121;
            labeldocIssue.Text = "مكان الإصدار:";
            // 
            // DocIssue1
            // 
            TextBox DocIssue = new TextBox();
            DocIssue.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            DocIssue.Location = new System.Drawing.Point(75, 44);
            DocIssue.Name = "مكان_الإصدار_" + addNameIndex + ".";
            DocIssue.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            DocIssue.Size = new System.Drawing.Size(130, 35);
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
            textJob.Size = new System.Drawing.Size(320, 35);
            textJob.TabIndex = 603;
            textJob.Text = job;
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

            //if (control.Name == "مقدم_الطلب_0.")
            autoCompleteBulk(AppName, DataSource, "الاسم", "TableGenNames");
            autoCompleteBulk(textJob, DataSource, "jobs", "TableListCombo");
            //if (control.Name == "المهنة_0.")
            autoCompleteBulk(DocIssue, DataSource, "مكان_الإصدار", "TableCollection");

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
            Panelapp.Controls.Add(labelExpDate);
            Panelapp.Controls.Add(DocExpireDate);
            Panelapp.Controls.Add(Job);
            Panelapp.Controls.Add(textJob);
            Panelapp.Controls.Add(addName);
            Panelapp.Controls.Add(removeName);
            addNameIndex++;

            //Panelapp.Height = 130 * (addNameIndex);
        }

        public string[] getID(TextBox text, int textID)
        {
            string[] returnedText = new string[7] { "","","","","","",""};

            try
            {
                returnedText[0] = docNo.Split('_')[textID];
                returnedText[1] = docType.Split('_')[textID];
                if (returnedText[1] == "")
                    returnedText[1] = "جواز سفر";
                returnedText[2] = docIsuue.Split('_')[textID];
                returnedText[3] = appJob.Split('_')[textID];
                returnedText[4] = appBirth.Split('_')[textID];
                returnedText[5] = sex.Split('_')[textID];
                if (returnedText[5] == "")
                    returnedText[5] = "ذكر";
                if (returnedText[6] == "")
                    returnedText[6] = appExp;
            }
            catch(Exception ex) {
                returnedText = new string[7] { "P0", "جواز سفر", "", "", "", "ذكر","" };
            }
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
                returnedText[6] = row["انتهاء_الصلاحية"].ToString();
            }
            return returnedText;
        }
        
        public string[] getID(TextBox text)
        {
            string[] returnedText = new string[7] {"P0", "جواز سفر", "","","","ذكر", ""};


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
                returnedText[6] = row["انتهاء_الصلاحية"].ToString();                
            }
            return returnedText;
        }


        private void addButtonInfo(string text1, string text2, string text3, string text4, string text5)
        {

            // 
            // textBox1
            // 
            TextBox textBox1 = new TextBox();
            textBox1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            textBox1.Location = new System.Drawing.Point(1063, 44);
            textBox1.Name = "textBox1_" + ButtonInfoIndex.ToString() + ".";
            textBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            textBox1.Size = new System.Drawing.Size(230, 35);
            textBox1.TabIndex = 570;
            textBox1.Text = text1;
            textBox1.Visible = true;
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
        }
        private void addNameBtnName_Click(object sender, EventArgs e)
        {
            addButtonInfo("", "", "", "", "");
        }
        private void addName_Click(object sender, EventArgs e)
        {
            addNames("", "ذكر", "جواز سفر", "P0", "", "العربية", "", "", "");
            btnPanelapp.Height = Panelapp.Height = 130 * addNameIndex;
            checkChanged(مقدم_الطلب, Panelapp);
            checkChanged(النوع, Panelapp);
            //MessageBox.Show(النوع.Text);
            checkChanged(نوع_الهوية, Panelapp);
            checkChanged(رقم_الهوية, Panelapp);
            checkChanged(مكان_الإصدار, Panelapp);
            checkChanged(تاريخ_الميلاد, Panelapp);
            checkChanged(المهنة, Panelapp);
            checkChanged(انتهاء_الصلاحية, Panelapp);
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



        private void fillTextBoxesInvers()
        {
            for (int x = 0; x < Vitext1.Text.Split('_').Length; x++)
            {
                addButtonInfo(Vitext1.Text.Split('_')[x], Vitext2.Text.Split('_')[x], Vitext3.Text.Split('_')[x], Vitext4.Text.Split('_')[x], Vitext5.Text.Split('_')[x]);
            }
            Vitext1.Text = Vitext2.Text = Vitext3.Text = Vitext4.Text = Vitext5.Text = "";
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
                addNames("", "ذكر", "جواز سفر", "P0", "", "العربية", "", "", "");

            }
            checkChanged(مقدم_الطلب, Panelapp);
            checkChanged(النوع, Panelapp);
            checkChanged(نوع_الهوية, Panelapp);
            checkChanged(رقم_الهوية, Panelapp);
            checkChanged(مكان_الإصدار, Panelapp);
            checkChanged(تاريخ_الميلاد, Panelapp);
            checkChanged(المهنة, Panelapp);
            checkChanged(انتهاء_الصلاحية, Panelapp);
        }
        string lastInput = "";
        private void textAge_TextChanged(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;

            checkChanged(تاريخ_الميلاد, Panelapp);
            Console.WriteLine(تاريخ_الميلاد.Text);            
            if (textBox.Text.Length == 11)
            {
                textBox.Text = lastInput; return;
            }
            if (textBox.Text.Length == 10) return;
            if (textBox.Text.Length == 4) textBox.Text = "-" + textBox.Text;
            else if (textBox.Text.Length == 7) textBox.Text = "-" + textBox.Text;
            lastInput = textBox.Text;


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

        private void textJob_TextChanged(object sender, EventArgs e)
        {
            checkChanged(المهنة, Panelapp);
        }
        private void AppName_TextChanged(object sender, EventArgs e)
        {
            checkChanged(مقدم_الطلب, Panelapp);
            if (gridFill) return;
            int index = 0;
            string[] textInfo = new string[7] { "","","","","","",""};
            //text[6] = "";
            TextBox textBox = (TextBox)sender;
            
            foreach (Control control in Panelapp.Controls) {
                if (control.Name == textBox.Name.Replace("مقدم_الطلب", "تاريخ_الميلاد"))
                {
                    string textNo = control.Name.Replace(".", "");
                    Console.WriteLine(textNo);
                    Console.WriteLine(textNo.Split('_')[2]);
                    index = Convert.ToInt32(textNo.Split('_')[2]);
                    index++;
                }
            }

            string TextID = textBox.Name.Split('_')[2].Replace(".", "");
            int id = Convert.ToInt32(TextID);
            if (textBox.Text == "")
            {
                try
                {
                    textInfo[0] = docNo.Split('_')[id];
                    textInfo[1] = docType.Split('_')[id];
                    if (textInfo[1] == "")
                        textInfo[1] = "جواز سفر";
                    textInfo[2] = docIsuue.Split('_')[id];
                    textInfo[3] = appJob.Split('_')[id];
                    textInfo[4] = appBirth.Split('_')[id];
                    textInfo[5] = sex.Split('_')[id];
                    textInfo[6] = appExp.Split('_')[id];
                    if (textInfo[5] == "")
                        textInfo[5] = "ذكر";
                }
                catch (Exception ex) { }
                //fillFirstInfo("", text[5], text[1], text[0], text[2], اللغة.Text, text[3], text[4], TextID);

            }
            else
            {
                textInfo = getID(textBox, Convert.ToInt32(TextID));

                try
                {
                    if (textInfo[4] == "")
                        textInfo[4] = ageDetected.Split('_')[index];
                }
                catch (Exception ex) { }

                try
                {
                    if (textInfo[5] == "")
                        textInfo[5] = "ذكر";
                }
                catch (Exception ex) { }
                
                try
                {
                    if (textInfo[6] == "")
                        textInfo[6] = appExp.Split('_')[index];
                }
                catch (Exception ex) { }
                
               
            }

            //MessageBox.Show(textInfo[6] );
            try
            {
                fillFirstInfo("", textInfo[5], textInfo[1], textInfo[0], textInfo[2], اللغة.Text, textInfo[3], textInfo[4], TextID, textInfo[6]);
            }catch (Exception ex) {
                fillFirstInfo("", textInfo[5], textInfo[1], textInfo[0], textInfo[2], اللغة.Text, textInfo[3], textInfo[4], TextID, "");
            }

        }

        

        private void ExpireDate_TextChanged(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            checkChanged(انتهاء_الصلاحية, Panelapp);
            Console.WriteLine(انتهاء_الصلاحية.Text);            
            if (textBox.Text.Length == 11)
            {
                textBox.Text = lastInput; return;
            }
            if (textBox.Text.Length == 10) return;
            if (textBox.Text.Length == 4) textBox.Text = "-" + textBox.Text;
            else if (textBox.Text.Length == 7) textBox.Text = "-" + textBox.Text;
            lastInput = textBox.Text;


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
        private void sexCheckedChanged(object sender, EventArgs e)
        {

            CheckBox checkBox = (CheckBox)sender;
            if (checkBox.CheckState == CheckState.Unchecked) checkBox.Text = "أنثى";
            else checkBox.Text = "ذكر";
            checkChanged(النوع, Panelapp);
        }
        
        private void checkSexType_TextChanged(object sender, EventArgs e)
        {

            CheckBox checkBox = (CheckBox)sender;
            if (checkBox.Text == "")
            {
                checkBox.Text = "ذكر";
                checkBox.Checked = true;
            }
            checkChanged(النوع, Panelapp);
        }

        private void fillInfo(FlowLayoutPanel panel, bool hide)
        {
            foreach (Control control in panel.Controls)
            {
                if (hide && !control.Name.Contains("_0."))
                {
                    control.Visible = false;
                    //control.Text = "";
                }
                if (control.Name.Contains(".") && !control.Name.Contains("_0."))
                {
                    control.Name = "unvalid_" + InvalidControl.ToString();
                    InvalidControl++;
                }
                else
                {
                    //MessageBox.Show(control.Name); 
                    panelFill(control);
                }
            }
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

        private void checkChangedAge(TextBox text, FlowLayoutPanel panel)
        {
            int index = 0;
            foreach (Control control in panel.Controls)
            {

                //if (control.Visible && control.Name == "تاريخ_الميلاد_" + index + ".")
                //{
                //    if (index == 0) text.Text = control.Text;
                //    else text.Text = text.Text + "_" + control.Text;
                //    index++;
                //}

                if (control.Visible && control.Name == control.Name + "_" + index + ".")
                {
                    //MessageBox.Show(control.Name +" - "+ control.Text);
                    if (index == 0) text.Text = control.Text;
                    else text.Text = text.Text + "_" + control.Text;
                    //MessageBox.Show(control.Name + " - " + control.Text);
                    index++;
                }

            }
        }
        public void panelFill(Control control)
        {
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
                        Console.WriteLine(control.Text);

                    }

                }
            }
        }

        private void btnFile1_Click(object sender, EventArgs e)
        {
            btnFile1.Enabled = false;
            bool found = FillDatafromGenArch("data1", intID.ToString(), "TableCollection");
            if (!found)
                found = FillDatafromGenArchrelated(intID.ToString(), "TableCollection");
            if (!found)
                FillDatafromGenArchrelated(intID.ToString(), "TableAuth");

            btnFile1.Enabled = true;
        }

        private void btnFile2_Click(object sender, EventArgs e)
        {
            btnFile2.Enabled = false;
            FillDatafromGenArch("data2", intID.ToString(), "TableCollection");
            btnFile2.Enabled = true;
        }

        private void btnFile3_Click(object sender, EventArgs e)
        {
            btnFile3.Enabled = false;
            FillDatafromGenArch("data3", intID.ToString(), "TableCollection");
            btnFile3.Enabled = true;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string docType = "";
            if (btnPrintDocx.InvokeRequired)
            {
                btnPrintDocx.Invoke(new MethodInvoker(delegate { btnPrintDocx.Enabled = false; }));
            }
            if (btnPrintPdf.InvokeRequired)
            {
                btnPrintPdf.Invoke(new MethodInvoker(delegate { btnPrintPdf.Enabled = false; }));
            }

            if (نوع_المعاملة.InvokeRequired)
            {
                نوع_المعاملة.Invoke(new MethodInvoker(delegate { docType = نوع_المعاملة.Text; }));
            }
            if (docType == "إقرار" && اللغة.Checked)
            {
                docType = "Affidavit";
            }
            else if (docType == "إفادة لمن يهمه الأمر" && اللغة.Checked)
            {
                docType = "TO WHOM IT MAY CONCERN";
            }
            chooseDocxFile(مقدم_الطلب.Text.Split('_')[0], رقم_المعاملة.Text, docType);
            prepareDocxfile();
            if (btnPrintDocx.InvokeRequired)
            {
                btnPrintDocx.Invoke(new MethodInvoker(delegate { btnPrintDocx.Enabled = true; btnPrintDocx.Text = "طباعة المعاملة (docx)"; }));
            }

            if (btnPrintPdf.InvokeRequired)
            {
                btnPrintPdf.Invoke(new MethodInvoker(delegate { btnPrintPdf.Enabled = true; btnPrintPdf.Text = "طباعة المعاملة (pdf)"; }));
            }
        }

        private void prepareDocxfile()
        {

            oBMiss = System.Reflection.Missing.Value;
            oBMicroWord = new Word.Application();
            //MessageBox.Show(localCopy.Text);
            object objCurrentCopy = localCopy.Text;

            try
            {
                oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);
            }
            catch (Exception ex) { return; }
            oBMicroWord.Selection.Find.ClearFormatting();
            oBMicroWord.Selection.Find.Replacement.ClearFormatting();

        }
        private void chooseDocxFile(string appName, string docId, string docType)
        {
            string proType = "";

            if (addNameIndex > 1) proType = " متعدد";
            //string RouteFile = FilespathIn + docType + proType + proType1 + ".docx";
            string RouteFile = docType + proType + proType1;
            if (appName != "")
                localCopy.Text = FilespathOut + appName + DateTime.Now.ToString("ddmmss") + ".docx";
            else 
                localCopy.Text = FilespathOut + docId.Replace("/", "_") + DateTime.Now.ToString("ddmmss") + ".docx";
            //MessageBox.Show("1 " + localCopy.Text);
            OpenModelFile(RouteFile, false, localCopy.Text);
            //while (File.Exists(localCopy.Text))
            //{
            //    if (appName != "")
            //        localCopy.Text = FilespathOut + appName + DateTime.Now.ToString("ddmmss") + ".docx";
            //    else localCopy.Text = FilespathOut + docId.Replace("/", "_") + DateTime.Now.ToString("ddmmss") + ".docx";

            //    MessageBox.Show("2 " + localCopy.Text); 
            //    localCopy.Text = OpenModelFile(RouteFile, false, localCopy.Text);
            //}
            

            //try
            //{
            //    //System.IO.File.Copy(RouteFile, localCopy.Text);
            //    //MessageBox.Show(localCopy.Text);
                
            //}
            //catch (Exception ex)
            //{
            //    goBack = true;

            //    return;
            //}
            goBack = false;
            //FileInfo fileInfo = new FileInfo(localCopy.Text);
            //if (fileInfo.IsReadOnly) fileInfo.IsReadOnly = false;
            Console.WriteLine(localCopy.Text);
            //MessageBox.Show("3 " + localCopy.Text);
        }

        private string OpenModelFile(string documen, bool printOut, string FileName)
        {
            string query = "SELECT ID, المستند,Data1, Extension1 from TableModelFiles where المستند=N'" + documen.Split('.')[0] + "'";

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow reader in dtbl.Rows)
            {
                var name = reader["المستند"].ToString();
                if (name == "")
                {
                    return "";
                }
                try
                {
                    var Data = (byte[])reader["Data1"];
                    string ext = ".docx";
                    //FileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                    File.WriteAllBytes(FileName, Data);
                    if (printOut)
                        System.Diagnostics.Process.Start(FileName);
                }
                catch (Exception ex) { return ""; }
            }
            sqlCon.Close();
            return FileName;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (currentPanelIndex <= 4)
                currentPanelIndex++;
            else return;
            panelShow(currentPanelIndex);
        }

        private void addNewAppNameInfo(TextBox textName)
        {
            //MessageBox.Show(addNameIndex.ToString());
            string query = "insert into TableGenNames ([الاسم], رقم_الهوية,تاريخ_الميلاد,المهنة,النوع,نوع_الهوية,مكان_الإصدار,انتهاء_الصلاحية) values (@col1,@col2,@col3,@col4,@col5,@col6,@col7,@col8) ;SELECT @@IDENTITY as lastid";
            for (int x = 0; x < addNameIndex; x++)
            {
                string id = checkExist(textName.Text.Split('_')[x]);
                if (id != "0")
                {
                    query = "update TableGenNames set [الاسم] =  @col1,[رقم_الهوية] = @col2,[تاريخ_الميلاد] = @col3,[المهنة] = @col4,النوع = @col5,نوع_الهوية = @col6,مكان_الإصدار = @col7 ,انتهاء_الصلاحية = @col8 where ID = " + id;
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
                try
                {
                    sqlCommand.Parameters.AddWithValue("@col5", النوع.Text.Split('_')[x]);
                }
                catch (Exception ex) {
                    sqlCommand.Parameters.AddWithValue("@col5", النوع.Text);
                }
                sqlCommand.Parameters.AddWithValue("@col6", نوع_الهوية.Text.Split('_')[x]);
                sqlCommand.Parameters.AddWithValue("@col7", مكان_الإصدار.Text.Split('_')[x]);
                sqlCommand.Parameters.AddWithValue("@col8", انتهاء_الصلاحية.Text.Split('_')[x]);

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
        private bool checkGender(FlowLayoutPanel panel, string controlType, string control2type)
        {
            int index = 0;
            foreach (Control control in panel.Controls)
            {
                if (control.Name == controlType + index + ".")
                {
                    string gender = getGender(control.Text.Split(' ')[0]);
                    if (gender == "") return true;
                    foreach (Control control2 in panel.Controls)
                    {
                        if (control2.Name == control2type + index + ".")
                        {
                            if (gender != control2.Text)
                            {
                                var selectedOption = MessageBox.Show("(ذكر/أنثى)", "الجنس غير مطابق ل" + control.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

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
            string sex = "";
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
                    //MessageBox.Show("UPDATE TableGenGender SET النوع=N'" + newGender + "' WHERE ID=" + id);
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
        private void btnPrevious_Click(object sender, EventArgs e)
        {
            //if (currentPanelIndex > 0) currentPanelIndex--;
            //else return;
            //if (currentPanelIndex == 0) FillDataGridView(DataSource);
            //panelShow(currentPanelIndex);
        }

        private void اللغة_TextChanged(object sender, EventArgs e)
        {
            if (اللغة.Text == "العربية")
            {
                اللغة.Checked = false;

            }
            else
            {
                اللغة.Checked = true;

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
                        if (dataRow[comlumnName].ToString() != "")
                            combbox.Items.Add(dataRow[comlumnName].ToString());
                    }
                }
                saConn.Close();
            }
        }
        private void FormCollection_Load(object sender, EventArgs e)
        {
            fileColComboBox(نوع_المعاملة, DataSource, "altColName");

            //وجهة_المعاملة.SelectedIndex = 0;
            //fileComboBoxAttend(DocType, DataSource, "DocType", "TableListCombo");
            //autoCompleteTextBox(DocSource, DataSource, "SDNIssueSource", "TableListCombo");
            diplomats(موقع_المعاملة, DataSource, اللغة.Text);
            //MessageBox.Show(AtVCIndex.ToString());
            موقع_المعاملة.SelectedIndex = AtVCIndex;
            getTitle(DataSource, موقع_المعاملة.Text);
            fileComboBoxMandoub(اسم_المندوب, DataSource, "TableMandoudList");
            // autoCompleteTextBox(Vitext5, DataSource, "Vitext5", "TableCollection");
            //addNames("", "ذكر", "جواز سفر", "P0", "", "العربية", "", "");

            autoCompleteBulk(الشاهد_الأول, DataSource, "الاسم", "TableGenNames");
            autoCompleteBulk(الشاهد_الثاني, DataSource, "الاسم", "TableGenNames");
        }

        private void fileColComboBox(ComboBox combbox, string source, string comlumnName)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct " + comlumnName + " from TableAddContext where " + comlumnName + " is not null and ColRight = '' order by " + comlumnName + " asc";
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

        private void diplomats(ComboBox combbox, string source, string lang)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct EmployeeName,EngEmployeeName from TableUser where EmployeeName is not null and الدبلوماسيون = N'yes' and Aproved like N'%أكده%' order by EmployeeName asc";
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
                        if (lang == "العربية")
                            combbox.Items.Add(dataRow["EmployeeName"].ToString());
                        else combbox.Items.Add(dataRow["EngEmployeeName"].ToString());
                    }
                }
                catch (Exception ex) { }
                saConn.Close();
            }
        }

        private void fileComboBoxMandoub(ComboBox combbox, string source, string tableName)
        {
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
        private void fileComboBoxAttend(ComboBox combbox, string source, string comlumnName, string tableName)
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
                    if (dataRow[comlumnName].ToString() != "")
                        combbox.Items.Add(dataRow[comlumnName].ToString());

                }
                saConn.Close();
            }
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
                textbox.AutoCompleteCustomSource.Clear();

                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    if (!string.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                    {
                        for (int x = 0; x < Textboxtable.Rows.Count; x++)
                            if (dataRow[comlumnName].ToString().Equals(Textboxtable.Rows[x]))
                                newSrt = false;

                        if (newSrt)
                            autoComplete.Add(dataRow[comlumnName].ToString());
                    }
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        } private void autoCompleteTextBox(ComboBox textbox, string source, string comlumnName, string tableName)
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
                textbox.AutoCompleteCustomSource.Clear();

                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    if (!string.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                    {
                        for (int x = 0; x < Textboxtable.Rows.Count; x++)
                            if (dataRow[comlumnName].ToString().Equals(Textboxtable.Rows[x]))
                                newSrt = false;

                        if (newSrt)
                            autoComplete.Add(dataRow[comlumnName].ToString());
                    }
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
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
        private void autoCompleteBulk(ComboBox textbox, string source, string col, string table)
        {

            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct " + col + " from " + table + " where نوع_المعاملة=N'"+ نوع_المعاملة.Text+ "' and نوع_الإجراء = N'"+ نوع_الإجراء.Text+"'";
                Console.WriteLine(query);
                //MessageBox.Show(query);
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
                    textbox.Items.Add(dataRow[col].ToString());
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }


        private void fileComboBox2(ComboBox combbox, string source, string comlumnName, string tableName)
        {
            combbox.Visible = true;

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
                    if (!String.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                        combbox.Items.Add(dataRow[comlumnName].ToString());
                }
                saConn.Close();
            }
            if (combbox.Items.Count > 0) combbox.SelectedIndex = 0;
        }

        private void نوع_المعاملة_SelectedIndexChanged(object sender, EventArgs e)
        {
            نوع_الإجراء.Items.Clear();
            fileComboBoxSub(نوع_الإجراء, DataSource, "altColName", "altSubColName");
            autoCompleteTextBox(نوع_الإجراء, DataSource, "نوع_الإجراء", "TableCollection");
            عنوان_المكاتبة.Items.Clear();
                عنوان_المكاتبة.Items.Add(نوع_المعاملة.Text);
            عنوان_المكاتبة.Items.Add(نوع_الإجراء.Text);

                if (نوع_المعاملة.Text == "إفادة لمن يهمه الأمر")
                {
                    تفيد_تشهد_off.Text = "فيد";
                    عنوان_المكاتبة.Items.Add("إفادة");
                if (عنوان_المكاتبة.Text == "")
                    عنوان_المكاتبة.Text = "إفادة";
            }
                else if (نوع_المعاملة.Text == "شهادة لمن يهمه الأمر")
                {
                    تفيد_تشهد_off.Text = "شهد";
                    عنوان_المكاتبة.Items.Add("شهادة");

                if (عنوان_المكاتبة.Text == "")
                    عنوان_المكاتبة.Text = نوع_المعاملة.Text;
            }
            if (نوع_المعاملة.Text == "مذكرة لسفارة عربية")
            {
                autoCompleteTextBox(Vitext1, DataSource, "ArabCountries", "TableListCombo");
                autoCompleteTextBox(Vitext1, DataSource, "ForiegnCountries", "TableListCombo");
                autoCompleteTextBox(Vitext2, DataSource, "Vitext1", "TableCollection");                
            }
        }

        private void fileComboBoxSub(ComboBox combbox, string source, string comlumnName, string SubComlumnName)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct " + SubComlumnName + " from TableAddContext where " + SubComlumnName + " is not null and " + comlumnName + " =N'" + نوع_المعاملة.Text + "' order by " + SubComlumnName + " asc";
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
                        combbox.Items.Add(dataRow[SubComlumnName].ToString());
                    }
                }
                catch (Exception ex) { }
                saConn.Close();
            }
        }

        private void newFillComboBox1(ComboBox combbox, string source, string colName)
        {
            combbox.Visible = true;
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();
                string query = "select " + colName + " from TableListCombo where " + colName + " is not null";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                foreach (DataRow dataRow in table.Rows)
                {
                    combbox.Items.Add(dataRow[colName].ToString());
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

        private void نوع_المعاملة_TextChanged(object sender, EventArgs e)
        {
            if (نوع_المعاملة.Text != "إختر نوع المعاملة")
            {
                for (int item = 0; item < نوع_المعاملة.Items.Count; item++)
                {
                    if (نوع_المعاملة.Items[item].ToString() == نوع_المعاملة.Text)
                        نوع_المعاملة.SelectedIndex = item;
                }
            }
        }

        private void اللغة_CheckedChanged(object sender, EventArgs e)
        {
            if (!اللغة.Checked)
            {
                اللغة.Text = "العربية";
                diplomats(موقع_المعاملة, DataSource, اللغة.Text);
                موقع_المعاملة.Width = 150;
                غرض_المعاملة.RightToLeft = PanelItemsboxes.RightToLeft = System.Windows.Forms.RightToLeft.No;
                txtReview.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
                نوع_الإجراء.Width = 329;
                موقع_المعاملة.Width = 184;
            }
            else
            {
                اللغة.Text = "الانجليزية";
                diplomats(موقع_المعاملة, DataSource, اللغة.Text);
                نوع_الإجراء.Width = 300;
                PanelItemsboxes.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
                غرض_المعاملة.RightToLeft = txtReview.RightToLeft = System.Windows.Forms.RightToLeft.No;
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("en-US");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
                موقع_المعاملة.Width = 300;
            }
            getTitle(DataSource, موقع_المعاملة.Text);
            //if (نوع_المعاملة.Items.Count > 0) نوع_المعاملة.SelectedIndex = 0;
            if (موقع_المعاملة.Items.Count > AtVCIndex) موقع_المعاملة.SelectedIndex = AtVCIndex;

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (checkAutoUpdate.Checked)
                txtReview.Text = writeStrSpecPur();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {

            if (وجهة_المعاملة.Text == "" || وجهة_المعاملة.Text.Contains("إختر"))
            {
                MessageBox.Show("يرجى إختيار وجهة المعاملة");
                return;
            }
            
            
            if (Jobposition.Contains("قنصل")) {
                btnPrintDocx.Select();
                FinalPrint("docxO");
                return;
            }
            else btnPrintPdf.Select();

            int count = getTodaDocxPdf();

            if (count < AllowedTimes-1)
            {
                var selectedOption = MessageBox.Show("لديك عدد " + (AllowedTimes - count).ToString() + " عرض متوفر", "طباعة ملف وورد على اي حال", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

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
           
        }

        private string getDocxPdf()
        {
            string doc = "";
            string query = "select الإجراء_الأخير from TableCollection where رقم_المعاملة = N'" + رقم_المعاملة.Text + "'and اسم_الموظف = N'" + اسم_الموظف.Text + "'";
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
            string query = "select الإجراء_الأخير from TableCollection where التاريخ_الميلادي = N'" + GregorianDate + "' and الإجراء_الأخير = N'docx' and اسم_الموظف = N'" + اسم_الموظف.Text + "'";
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
            string query = "update TableCollection set الإجراء_الأخير = N'" + doc + "' where رقم_المعاملة = N'" + رقم_المعاملة.Text + "'and اسم_الموظف = N'" + اسم_الموظف.Text + "'";
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
                MessageBox.Show("لم يتم منح " + اسم_الموكل_بالتوقيع.Text + " حق توكيل غيره" + " بالوكالة المشار إليها");
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
        private void getTitle(string source, string empName)
        {
            string query = "select AuthenticType,AuthenticTypeEng from TableUser where EmployeeName = N'" + empName + "' or EngEmployeeName =N'" + empName +"'";
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
                if (اللغة.Text == "العربية")
                {
                    AuthTitle = dataRow["AuthenticType"].ToString();
                    AuthTitleLast = Environment.NewLine + AuthTitle;
                }
                else
                {
                    AuthTitle = dataRow["AuthenticTypeEng"].ToString();
                    AuthTitleLast = Environment.NewLine + AuthTitle;
                }
            }
            //MessageBox.Show(AuthTitleLast);

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

        private void CreateMessageWord(string ApplicantName, string EmbassySource, string IqrarNo, string MessageType, string ApplicantSex, string GregorianDate, string HijriDate, string ViseConsul)
        {
            string ActiveCopy;
            string ReportName = DateTime.Now.ToString("mmss");
            string routeDoc = FilespathIn + @"\MessageCap.docx";
            loadMessageNo();
            ActiveCopy = FilespathOut + @"\Message" + ApplicantName.Replace("/", "_") + ReportName + ".docx";
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
        private void addarchives()
        {
            Console.WriteLine(رقم_المعاملة.Text);

            if (checkArchives(رقم_المعاملة.Text)) return;// else MessageBox.Show("not found");

            colIDs[3] = مقدم_الطلب.Text.Split('_')[0];
            colIDs[7] = archState;
            colIDs[5] = طريقة_الطلب.Text;
            colIDs[2] = التاريخ_الميلادي.Text;
            colIDs[0] = رقم_المعاملة.Text;
            colIDs[1] = intID.ToString();


            colIDs[4] = EmpName;

            colIDs[6] = اسم_المندوب.Text;


            string[] allArchList = getColList("archives");
            string strList = "";
            for (int i = 0; i < 8; i++)
            {

                if (i == 0) strList = "@" + allArchList[0];
                if (!string.IsNullOrEmpty(allArchList[i]) && allArchList[i] != "")
                {
                    Console.WriteLine(i.ToString() + " - " + allArchList[i]);
                    //MessageBox.Show(i.ToString() + " - " + allArchList[i]);                    
                    if (i != 0)
                        strList = strList + ",@" + allArchList[i];
                }
            }
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();

            SqlCommand sqlCommand = new SqlCommand("insert into archives (" + strList.Replace("@", "") + ") values (" + strList + ");SELECT @@IDENTITY as lastid", sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            for (int i = 0; i < 8; i++)
            {
                Console.WriteLine(allArchList[i] + " - " + colIDs[i]);
                //MessageBox.Show(allArchList[i]+" - "+ colIDs[i]);

                sqlCommand.Parameters.AddWithValue("@" + allArchList[i], colIDs[i]);
            }
            Console.WriteLine("insert into archives (" + strList.Replace("@", "") + ") values (" + strList + ")");
            //MessageBox.Show("lastid");

            try
            {
                var reader = sqlCommand.ExecuteReader();
                if (reader.Read())
                {
                    //MessageBox.Show(reader["lastid"].ToString());
                }

            }
            catch (Exception ex) { MessageBox.Show("insert into archives (" + strList.Replace("@", "") + ") values (" + strList + ")"); }
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
        private void chooseBtnTable()
        {
            switch (نوع_المعاملة.Text)
            {
                case "إقرار":
                    if(!اللغة.Checked)
                    fillTextBoxesDocx(1, addInfo);
                    else fillTextBoxesDocx(ButtonInfoIndex, addInfo);
                    break;
                case "إقرار مشفوع باليمين":
                    fillTextBoxesDocx(1, addInfo);
                    break;
                case "إفادة لمن يهمه الأمر":
                    if (!اللغة.Checked)
                        fillTextBoxesDocx(addNameIndex, addInfo);
                    else fillTextBoxesDocx(addNameIndex, addInfo);
                    break;
                
                case "شهادة لمن يهمه الأمر":
                    fillTextBoxesDocx(addNameIndex, addInfo);
                    break;
                case "مذكرة لسفارة عربية":
                    //fillTextBoxesDocx(addNameIndex, addInfo);
                    break;
            }
        }
        private void fillTextBoxesDocx(int index, bool libtnAdd1Vis)
        {
            if (index > 1) index = 2;
            //MessageBox.Show(index.ToString());
            try
            {
                Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[index];
                if (!libtnAdd1Vis) return;
                //if (نوع_المعاملة.SelectedIndex > 1 && !libtnAdd1Vis)
                //{ table.Delete(); return; }
                if (اللغة.Text == "العربية")
                {
                    table.Rows[1].Cells[1].Range.Text = "الرقم";
                    table.Rows[1].Cells[2].Range.Text = labl1.Text.Replace(":", "");
                    table.Rows[1].Cells[3].Range.Text = labl2.Text.Replace(":", "");
                    table.Rows[1].Cells[4].Range.Text = labl3.Text.Replace(":", "");
                    table.Rows[1].Cells[5].Range.Text = labl4.Text.Replace(":", "");
                    table.Rows[1].Cells[6].Range.Text = labl5.Text.Replace(":", "");
                    for (int x = 0; x <= 5; x++)
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
                }
                else {
                    table.Rows[1].Cells[6].Range.Text = "الرقم";
                    table.Rows[1].Cells[5].Range.Text = labl1.Text.Replace(":", "");
                    table.Rows[1].Cells[4].Range.Text = labl2.Text.Replace(":", "");
                    table.Rows[1].Cells[3].Range.Text = labl3.Text.Replace(":", "");
                    table.Rows[1].Cells[2].Range.Text = labl4.Text.Replace(":", "");
                    table.Rows[1].Cells[1].Range.Text = labl5.Text.Replace(":","");
                    for (int x = 0; x <= 5; x++)
                    {
                        int indBox = 1;
                        foreach (Control control in PanelButtonInfo.Controls)
                        {
                            if (x == 0)
                            {
                                table.Rows.Add();
                                table.Rows[indBox + 1].Cells[6].Range.Text = indBox.ToString();
                                indBox++;
                            }
                            else
                            {
                                if (control is TextBox && control.Name.Contains("textBox" + x + "_"))
                                {

                                    table.Rows[indBox + 1].Cells[6-x].Range.Text = control.Text;
                                    indBox++;
                                }

                            }
                            if (indBox > ButtonInfoIndex) break;
                        }
                    }
                }

                try
                {
                       
                    if (labl1.Text == "") table.Columns[2].Delete();
                    if (labl2.Text == "") table.Columns[3].Delete();
                    if (labl3.Text == "") table.Columns[4].Delete();
                    if (labl4.Text == "") table.Columns[5].Delete();
                    if (labl5.Text == "")
                    {

                        
                        table.Columns[6].Delete();
                        //MessageBox.Show(labl5.Text);
                    }
                }
                catch (Exception ex)
                {
                }
            }
            catch (Exception ex) { }
        }
        private void fillDocFileAppInfo(FlowLayoutPanel panel)
        {
            //MessageBox.Show(panel.Name);
            foreach (Control control in panel.Controls)
            {
                //MessageBox.Show(panel.Name + " - " + control.Name + " - " + control.Text);
                if (control is TextBox || control is ComboBox)
                {
                    try
                    {
                        object ParaAuthIDNo = control.Name;
                        Word.Range BookAuthIDNo = oBDoc.Bookmarks.get_Item(ref ParaAuthIDNo).Range;
                        if (control.Name == "موقع_المعاملة")
                            BookAuthIDNo.Text = control.Text + AuthTitleLast;
                        else BookAuthIDNo.Text = control.Text;
                        if ((control.Name == "التاريخ_الميلادي" || control.Name == "التاريخ_الهجري") && اللغة.Checked)
                            BookAuthIDNo.Text = control.Text.Split('-')[1] + "-" + control.Text.Split('-')[0] + "-" + control.Text.Split('-')[2];

                        object rangeAuthIDNo = BookAuthIDNo;
                        oBDoc.Bookmarks.Add(control.Name, ref rangeAuthIDNo);

                        //MessageBox.Show(panel.Name+ " - "+control.Name+ " - "+control.Text);
                    Console.WriteLine(panel.Name+ " - "+control.Name+ " - "+control.Text);
                    }
                    catch (Exception ex)
                    {
                        //    MessageBox.Show(control.Name); 
                    }
                }
            }
            //MessageBox.Show(مقدم_الطلب.Text);
            if (notFiled)
            {
                Console.WriteLine("notFiled");
                appNameInfo(نوع_المعاملة.Text);
                notFiled = false;
            }
        }

        private void appNameInfo(string appindex)
        {

            switch (appindex)
            {
                case "إقرار":
                    //MessageBox.Show("اقرار");
                    if (addNameIndex == 1)
                    {
                        //MessageBox.Show("فردي");
                        if (!اللغة.Checked)
                        {
                            try
                            {
                                Microsoft.Office.Interop.Word.Table table1 = oBDoc.Tables[1];
                                Microsoft.Office.Interop.Word.Table table2 = oBDoc.Tables[2];
                                if (ButtonInfoIndex == 0)
                                {
                                    table1.Delete();
                                    //MessageBox.Show("1");
                                }

                                if (الشاهد_الأول.Text == "")
                                {
                                    table2.Delete();
                                    //MessageBox.Show("2");
                                }
                            }
                            catch (Exception ex) { return; }
                        }
                        else if (اللغة.Checked)
                        {
                            try
                            {
                                Microsoft.Office.Interop.Word.Table table1 = oBDoc.Tables[1];
                                if (ButtonInfoIndex == 0)
                                {
                                    table1.Delete();
                                    //MessageBox.Show("1");
                                }
                            }
                            catch (Exception ex) { return; }
                        }
                    }
                    else
                    {
                       // MessageBox.Show("متعدد");
                        if (!اللغة.Checked)
                        {
                            //MessageBox.Show("عربي");
                            try
                            {
                                Microsoft.Office.Interop.Word.Table table1 = oBDoc.Tables[1];
                                Microsoft.Office.Interop.Word.Table table2 = oBDoc.Tables[2];
                                Microsoft.Office.Interop.Word.Table table3 = oBDoc.Tables[3];

                                if (!LibtnAdd1Vis)
                                    table2.Delete();

                                if (الشاهد_الأول.Text == "")
                                    table3.Delete();
                                for (int x = 0; x < addNameIndex; x++)
                                {
                                    if (مقدم_الطلب.Text.Split('_')[x] != "")
                                    {
                                        table1.Rows.Add();
                                        table1.Rows[x + 2].Cells[1].Range.Text = (x + 1).ToString();
                                        table1.Rows[x + 2].Cells[2].Range.Text = مقدم_الطلب.Text.Split('_')[x];
                                        table1.Rows[x + 2].Cells[3].Range.Text = رقم_الهوية.Text.Split('_')[x];
                                        table1.Rows[x + 2].Cells[4].Range.Text = مكان_الإصدار.Text.Split('_')[x];
                                    }
                                }
                            }
                            catch (Exception ex) { }
                        } 
                        
                        else if (اللغة.Checked)
                        {
                           
                            try
                            {
                                Microsoft.Office.Interop.Word.Table table1 = oBDoc.Tables[1];
                                Microsoft.Office.Interop.Word.Table table2 = oBDoc.Tables[2];

                                if (!LibtnAdd1Vis)
                                {
                                    table2.Delete();
                                    //MessageBox.Show("dele 2");
                                }

                                //MessageBox.Show(مقدم_الطلب.Text +" - "+ addNameIndex.ToString());
                                for (int x = 0; x < addNameIndex; x++)
                                {
                                    //MessageBox.Show(مقدم_الطلب.Text.Split('_')[x]);
                                    if (مقدم_الطلب.Text.Split('_')[x] != "")
                                    {
                                        table1.Rows.Add();
                                        table1.Rows[x + 2].Cells[5].Range.Text = (x + 1).ToString();
                                        table1.Rows[x + 2].Cells[4].Range.Text = مقدم_الطلب.Text.Split('_')[x];
                                        table1.Rows[x + 2].Cells[3].Range.Text = رقم_الهوية.Text.Split('_')[x];
                                        table1.Rows[x + 2].Cells[2].Range.Text = انتهاء_الصلاحية.Text.Split('_')[x];
                                    }
                                }
                            }
                            catch (Exception ex) { }
                        }
                    }
                    break;
                case "إقرار مشفوع باليمين":
                    
                    if (addNameIndex == 1)
                    {
                        //MessageBox.Show(addNameIndex.ToString());
                        try
                        {
                            Microsoft.Office.Interop.Word.Table table1 = oBDoc.Tables[1];
                            Microsoft.Office.Interop.Word.Table table2 = oBDoc.Tables[2];
                            if (!LibtnAdd1Vis)
                                table1.Delete();

                            if (الشاهد_الأول.Text == "")
                                table2.Delete();
                        }
                        catch (Exception ex) { return; }
                    }
                    else
                    {
                        //MessageBox.Show(addNameIndex.ToString());
                        //MessageBox.Show(الشاهد_الأول.Text);
                        //if(PanelButtonInfo.Visible)
                        //    MessageBox.Show("PanelButtonInfo.Visible"); 
                        //else MessageBox.Show("PanelButtonInfo. not Visible");
                        

                        try
                        {
                            Microsoft.Office.Interop.Word.Table table1 = oBDoc.Tables[1];
                            Microsoft.Office.Interop.Word.Table table2 = oBDoc.Tables[2];
                            Microsoft.Office.Interop.Word.Table table3 = oBDoc.Tables[3];
                            if (الشاهد_الأول.Text == "")
                            {
                                table3.Delete();
                                //MessageBox.Show("الغاء الشهود");
                            }

                            if (!PanelButtonInfo.Visible)
                            {
                                //MessageBox.Show("الغاء الادخال المتعدد");
                                table2.Delete();
                            }

                           

                            for (int x = 0; x < addNameIndex; x++)
                            {
                                if (مقدم_الطلب.Text.Split('_')[x] != "")
                                {
                                    table1.Rows.Add();
                                    //MessageBox.Show(x.ToString() +" - "+مقدم_الطلب.Text.Split('_')[x]);
                                    table1.Rows[x + 2].Cells[1].Range.Text = (x + 1).ToString();
                                    table1.Rows[x + 2].Cells[2].Range.Text = مقدم_الطلب.Text.Split('_')[x];
                                    table1.Rows[x + 2].Cells[3].Range.Text = رقم_الهوية.Text.Split('_')[x];
                                    table1.Rows[x + 2].Cells[4].Range.Text = مكان_الإصدار.Text.Split('_')[x];
                                }
                            }
                        }
                        catch (Exception ex) { }
                    }
                    break;
                case "إفادة لمن يهمه الأمر":

                    try
                    {
                        if (ButtonInfoIndex == 0)
                        {
                            Microsoft.Office.Interop.Word.Table table1 = oBDoc.Tables[1];
                            table1.Delete();
                        }
                    }
                    catch (Exception ex) { }
                    break;
                case "شهادة لمن يهمه الأمر":
                    try
                    {
                        if (ButtonInfoIndex == 0)
                    {
                        Microsoft.Office.Interop.Word.Table table1 = oBDoc.Tables[1];
                        table1.Delete();
                    }
                    }
                    catch (Exception ex) { }                    
                    break;
                
                case "مذكرة لسفارة عربية":
                    try
                    {
                        Microsoft.Office.Interop.Word.Table table1 = oBDoc.Tables[1];                        
                        for (int x = 0; x < addNameIndex; x++)
                        {
                            if (مقدم_الطلب.Text.Split('_')[x] != "")
                            {
                                table1.Rows.Add();
                                table1.Rows[x + 2].Cells[1].Range.Text = (x + 1).ToString();
                                table1.Rows[x + 2].Cells[2].Range.Text = مقدم_الطلب.Text.Split('_')[x];
                                table1.Rows[x + 2].Cells[3].Range.Text = رقم_الهوية.Text.Split('_')[x];
                                table1.Rows[x + 2].Cells[4].Range.Text = مكان_الإصدار.Text.Split('_')[x];
                                table1.Rows[x + 2].Cells[5].Range.Text = انتهاء_الصلاحية.Text.Split('_')[x];
                            }
                        }
                    }
                    catch (Exception ex) { }
                    break;
                case "مذكرة لسفارة أجنبية":
                    try
                    {
                        Microsoft.Office.Interop.Word.Table table1 = oBDoc.Tables[1];                        
                        for (int x = 0; x < addNameIndex; x++)
                        {
                            if (مقدم_الطلب.Text.Split('_')[x] != "")
                            {
                                table1.Rows.Add();
                                table1.Rows[x + 2].Cells[5].Range.Text = (x + 1).ToString();
                                table1.Rows[x + 2].Cells[4].Range.Text = مقدم_الطلب.Text.Split('_')[x];
                                table1.Rows[x + 2].Cells[3].Range.Text = رقم_الهوية.Text.Split('_')[x];
                                table1.Rows[x + 2].Cells[2].Range.Text = مكان_الإصدار.Text.Split('_')[x];
                                table1.Rows[x + 2].Cells[1].Range.Text = انتهاء_الصلاحية.Text.Split('_')[x];
                            }
                        }
                    }
                    catch (Exception ex) { }
                    break;
            }
        }
        private void fillDocFileInfo(Panel panel)
        {
            foreach (Control control in panel.Controls)
            {
                if (control is TextBox || control is ComboBox)
                {
                    try
                    {
                        object ParaAuthIDNo = control.Name;
                        Word.Range BookAuthIDNo = oBDoc.Bookmarks.get_Item(ref ParaAuthIDNo).Range;
                        BookAuthIDNo.Text = control.Text;
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
        }
        public static void FindAndReplace(string loadPath, string text, bool remove)
        {
            DocumentCore dc = DocumentCore.Load(loadPath);

            // Find "Bean" and Replace everywhere on "Joker"
            Regex regex = new Regex(@text, RegexOptions.IgnoreCase);

            // Start:

            // Please note, Reverse() makes sure that action Replace() doesn't affect to Find().
            foreach (ContentRange item in dc.Content.Find(regex).Reverse())
            {
                if (remove)
                    item.Replace("", new CharacterFormat() { FontName = "Traditional Arabic", Size = 19.0, Bold = true });
                else item.Replace(text, new CharacterFormat() { FontName = "Traditional Arabic", Size = 19.0, Bold = true });
            }

            // End:

            // The code above finds and replaces the content in the whole document.
            // Let us say, you want to replace a text inside shape blocks only:

            // 1. Comment the code above from the line "Start" to the "End".
            // 2. Uncomment this code:
            //foreach (Shape shp in dc.GetChildElements(true, ElementType.Shape).Reverse())
            //{
            //    foreach (ContentRange item in shp.Content.Find(regex).Reverse())
            //    {
            //        item.Replace("Joker", new CharacterFormat() { BackgroundColor = Color.Yellow, FontName = "Arial", Size = 16.0 });
            //    }
            //}

            // Save the document as DOCX format.
            //string savePath = Path.ChangeExtension(loadPath, ".replaced.docx");
            dc.Save(loadPath, SaveOptions.DocxDefault);

            // Open the original and result documents for demonstration purposes.
            //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(loadPath) { UseShellExecute = true });
            //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(loadPath) { UseShellExecute = true });
        }
        private void fillPrintDocx(string deleteDocxFile)
        {
            btnPrintDocx.Enabled = btnPrintPdf.Enabled = false;
            //MessageBox.Show(localCopy.Text);
            string pdfFile = localCopy.Text.Replace("docx", "pdf");

            oBDoc.SaveAs2(localCopy.Text);
            if (deleteDocxFile == "pdf")
                oBDoc.ExportAsFixedFormat(pdfFile, Word.WdExportFormat.wdExportFormatPDF);
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
        private int getEdited(string date)
        {
            int count = -1;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select COUNT(edited) as edit from TableAuth where التاريخ_الميلادي =N'" + date + "' and edited = 'YES'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow reader in dtbl.Rows)
            {
                count = Convert.ToInt32(reader["edit"].ToString());
            }

            return count;
        }
        private void وجهة_المعاملة_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void وجهة_المعاملة_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(وجهة_المعاملة.Text);
            //if (وجهة_المعاملة.Items.Count > 0 && وجهة_المعاملة.Text == "")
            //    وجهة_المعاملة.SelectedIndex = 0;
            ////MessageBox.Show(وجهة_المعاملة.Text);
        }

        private string writeStrSpecPur()
        {
            //MessageBox.Show(StrSpecPur);
            return SuffReplacements(StrSpecPur, صفة_مقدم_الطلب_off.SelectedIndex);
        }

        private string SuffReplacements(string text, int index)
        {
            string str = "";
            if (النوع.Text != "ذكر") str = "ة";
            if (text == "") return "";

            if (text.Contains("  "))
                text = text.Replace("  ", " ");
            if (text.Contains("tN"))
                text = text.Replace("tN", مقدم_الطلب.Text);
            if (text.Contains("tP"))
                text = text.Replace("tP", رقم_الهوية.Text);
            if (text.Contains("tS"))
                text = text.Replace("tS", مكان_الإصدار.Text);
            if (text.Contains("tX"))
                text = text.Replace("tX", str);
            //if (text.Contains("tT"))
            //    text = text.Replace("tT", title.Text);
            //    text = text.Replace("tT", title.Text);إفادة
            if (text.Contains("tB"))
                text = text.Replace("tB", تاريخ_الميلاد.Text);
            if (text.Contains("tD"))
                text = text.Replace("tD", نوع_الهوية.Text);
            if (text.Contains("fD"))
                text = text.Replace("fD", انتهاء_الصلاحية.Text);
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
            if (text.Contains("t6"))
                text = text.Replace("t6", Vitext6.Text);
            if (text.Contains("t7"))
                text = text.Replace("t7", Vitext7.Text);
            if (text.Contains("t8"))
                text = text.Replace("t8", Vitext8.Text);
            if (text.Contains("t9"))
                text = text.Replace("t9", Vitext9.Text);
            if (text.Contains("t0"))
                text = text.Replace("t0", Vitext0.Text);

            if (text.Contains("c1"))
                text = text.Replace("c1", Vicheck1.Text);
            if (text.Contains("c2"))
                text = text.Replace("c2", Vicheck2.Text);
            if (text.Contains("c3"))
                text = text.Replace("c3", Vicheck3.Text);
            if (text.Contains("c4"))
                text = text.Replace("c4", Vicheck4.Text);
            if (text.Contains("c5"))
                text = text.Replace("c5", Vicheck5.Text);

            if (text.Contains("m1"))
                text = text.Replace("m1", Vicombo1.Text);
            if (text.Contains("m2"))
                text = text.Replace("m2", Vicombo2.Text);
            if (text.Contains("m3"))
                text = text.Replace("m3", Vicombo3.Text);
            if (text.Contains("m4"))
                text = text.Replace("m4", Vicombo4.Text);
            if (text.Contains("m5"))
                text = text.Replace("m5", Vicombo5.Text);
            
            if (text.Contains("a1"))
                text = text.Replace("a1", LibtnAdd1.Text);

            if (text.Contains("n1"))
                text = text.Replace("n1", " " + VitxtDate1.Text + " ");
            if (text.Contains("n2"))
                text = text.Replace("n2", " " + VitxtDate2.Text + " ");
            if (text.Contains("n3"))
                text = text.Replace("n3", " " + VitxtDate3.Text + " ");
            if (text.Contains("n4"))
                text = text.Replace("n4", " " + VitxtDate4.Text + " ");
            if (text.Contains("n5"))
                text = text.Replace("n5", " " + VitxtDate5.Text + " ");
            if (text.Contains("#*#"))
                text = text.Replace("#*#", preffix[0, 10]);

            if (text.Contains("#1"))
                text = text.Replace("#1", preffix[0, 11]);
            if (text.Contains("#2"))
                text = text.Replace("#2", preffix[0, 12]);

            //if (text.Contains("$$$"))
            //    text = text.Replace("$$$", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 0]);
            //if (text.Contains("&&&"))
            //    text = text.Replace("&&&", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 1]);
            //if (text.Contains("^^^"))
            //    text = text.Replace("^^^", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 2]);
            //if (text.Contains("###"))
            //    text = text.Replace("###", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4]);
            //if (text.Contains("***"))
            //    text = text.Replace("***", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3]);
            //if (text.Contains("%&%"))
            //    text = text.Replace("%&%", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 12]);
            //if (text.Contains("#$#"))
            //    text = text.Replace("#$#", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 13]);
            //if (text.Contains("&^&"))
            //    text = text.Replace("&^&", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 14]);
            //if (text.Contains("&^^"))
            //    text = text.Replace("&^^", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15]);
            //if (text.Contains("*%*"))
            //    text = text.Replace("*%*", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 16]);            
            //if (text.Contains("&&*"))
            //    text = text.Replace("&&*", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 17]);

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
                        text = text.Replace(code, replacemest[index]);
                }
            }
            return text;
        }


        private void نوع_الإجراء_SelectedIndexChanged(object sender, EventArgs e)
        {
            resetBoxes();
            reversTextReview();
            reversTextPurpose(); 
            flllPanelItemsboxes("ColName", نوع_الإجراء.Text + "-" + نوع_المعاملة.SelectedIndex.ToString());
            fillInfo(PanelItemsboxes, false);
            updateProType("TableCollection", نوع_المعاملة, نوع_الإجراء, "رقم_المعاملة", رقم_المعاملة.Text);

        }
        private void reversTextReview()
        {
            string column = نوع_المعاملة.Text.Replace(" ", "_") + "_" + نوع_الإجراء.Text.Replace(" ", "_");
            string query = "select " + column + " from TableCollectStarText where "+column+" is not null order by ID desc";
            Console.WriteLine(query);
            //MessageBox.Show(query);
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
                txtReviewListIndex = 0;
                txtReviewListIndexStar = 0;
                txtReviewList = new string[dtbl.Rows.Count];
                txtReviewListStar = new string[dtbl.Rows.Count + 1];
                txtReviewListStar[0] = "";
                //MessageBox.Show("count " + dtbl.Rows.Count.ToString());
                foreach (DataRow dataRow in dtbl.Rows)
                {
                    if (dataRow[column].ToString() != "" && !dataRow[column].ToString().Contains("removed"))
                    {
                        txtReviewList[txtReviewListIndex] = dataRow[column].ToString();
                        Console.WriteLine("txtReviewList" + txtReviewList[txtReviewListIndex]);
                        //MessageBox.Show("txtReviewList" + txtReviewList[txtReviewListIndex]);
                        txtReviewListIndex++;
                    }
                    
                    if (dataRow[column].ToString() != "" && !dataRow[column].ToString().Contains("removed") && dataRow[column].ToString().Contains("Star"))
                    {
                        txtReviewListStar[txtReviewListIndexStar+1] = dataRow[column].ToString();
                        Console.WriteLine("txtReviewList" + txtReviewListStar[txtReviewListIndexStar+1]);
                        txtReviewListIndexStar++;
                    }
                    
                }
            }
            catch (Exception ex) {
                txtReviewListStar = new string[1];
                txtReviewListStar[0] = "";
            }
        }
        
        private void reversTextPurpose()
        {
            string query = "select distinct غرض_المعاملة from TableCollection where غرض_المعاملة is not null and نوع_المعاملة = N'" + نوع_المعاملة.Text + "' and نوع_الإجراء=N'" + نوع_الإجراء.Text + "'";
            Console.WriteLine(query);
            //MessageBox.Show(query);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();

            sqlDa.Fill(dtbl);
            sqlCon.Close();
            txtPurposeListIndex = 0;
            txtPurposeList = new string[dtbl.Rows.Count];
            
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["غرض_المعاملة"].ToString() == "") continue;
                txtPurposeList[txtPurposeListIndex] = dataRow["غرض_المعاملة"].ToString();
                //MessageBox.Show(txtPurposeList[txtPurposeListIndex]);
                txtPurposeListIndex++;
            }
        }

       
            private void reversTextReviewold()
        {
            string query = "select * from TableCollection where نوع_الإجراء = N'" + نوع_الإجراء.Text + "' order by ID desc";
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
                if (dataRow["Vitext1"].ToString() != "" && txtReviewList[index].Contains(dataRow["Vitext1"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["Vitext1"].ToString(), "t1");
                if (dataRow["Vitext2"].ToString() != "" && txtReviewList[index].Contains(dataRow["Vitext2"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["Vitext2"].ToString(), "t2");
                if (dataRow["Vitext3"].ToString() != "" && txtReviewList[index].Contains(dataRow["Vitext3"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["Vitext3"].ToString(), "t3");
                if (dataRow["Vitext4"].ToString() != "" && txtReviewList[index].Contains(dataRow["Vitext4"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["Vitext4"].ToString(), "t4");
                if (dataRow["Vitext5"].ToString() != "" && txtReviewList[index].Contains(dataRow["Vitext5"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["Vitext5"].ToString(), "t5");

                if (dataRow["Vicheck1"].ToString() != "" && txtReviewList[index].Contains(dataRow["Vicheck1"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["Vicheck1"].ToString(), "c1");

                if (dataRow["Vicombo1"].ToString() != "" && txtReviewList[index].Contains(dataRow["Vicombo1"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["Vicombo1"].ToString(), "m1");
                if (dataRow["Vicombo1"].ToString() != "" && txtReviewList[index].Contains(dataRow["Vicombo1"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["Vicombo1"].ToString(), "m2");

                if (dataRow["LibtnAdd1"].ToString() != "" && txtReviewList[index].Contains(dataRow["LibtnAdd1"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["LibtnAdd1"].ToString(), "a1");
                if (dataRow["VitxtDate1"].ToString() != "" && txtReviewList[index].Contains(dataRow["VitxtDate1"].ToString()))
                    txtReviewList[index] = txtReviewList[index].Replace(dataRow["VitxtDate1"].ToString(), "n1");
                if (txtReviewList[index].Contains(مقدم_الطلب.Text))
                    txtReviewList[index] = txtReviewList[index].Replace(مقدم_الطلب.Text, "tN");
                if (txtReviewList[index].Contains(رقم_الهوية.Text))
                    txtReviewList[index] = txtReviewList[index].Replace(رقم_الهوية.Text, "tP");
                if (txtReviewList[index].Contains(مكان_الإصدار.Text))
                    txtReviewList[index] = txtReviewList[index].Replace(مكان_الإصدار.Text, "tS");
                //if (txtReviewList[index].Contains("tX"))
                //    txtReviewList[index] = txtReviewList[index].Replace("tX", str);

                if (txtReviewList[index].Contains(تاريخ_الميلاد.Text))
                    txtReviewList[index] = txtReviewList[index].Replace(تاريخ_الميلاد.Text, "tB");
                if (txtReviewList[index].Contains(نوع_الهوية.Text))
                    txtReviewList[index] = txtReviewList[index].Replace(نوع_الهوية.Text, "tD");
                if (txtReviewList[index].Contains(انتهاء_الصلاحية.Text))
                    txtReviewList[index] = txtReviewList[index].Replace(نوع_الهوية.Text, "fD");
                try
                {
                    if (txtReviewList[index].Contains(title.Text))
                        txtReviewList[index] = txtReviewList[index].Replace(title.Text, "tT");
                }
                catch (Exception ex) { }

                Console.WriteLine(txtReviewList[index]);
                index++;
            }



        }

        private string TextReviewCoding(string txtReviewList)
        {
            
            Console.WriteLine(txtReviewList);
            //MessageBox.Show(txtReviewList);
            //Console.WriteLine(Vitext1.Text);
            //MessageBox.Show(Vitext1.Text);
            //Console.WriteLine(Vitext2.Text);
            //MessageBox.Show(Vitext2.Text);
            
            if (Vitext1.Text != "" && txtReviewList.Contains(Vitext1.Text))
                txtReviewList = txtReviewList.Replace(Vitext1.Text, "t1");

            if (Vitext2.Text != "" && txtReviewList.Contains(Vitext2.Text))
                txtReviewList = txtReviewList.Replace(Vitext2.Text, "t2");

            if (Vitext3.Text != "" && txtReviewList.Contains(Vitext3.Text))
                txtReviewList = txtReviewList.Replace(Vitext3.Text, "t3");

            if (Vitext4.Text != "" && txtReviewList.Contains(Vitext4.Text))
                txtReviewList = txtReviewList.Replace(Vitext4.Text, "t4");

            if (Vitext5.Text != "" && txtReviewList.Contains(Vitext5.Text))
                txtReviewList = txtReviewList.Replace(Vitext5.Text, "t5");

            if (Vicheck1.Text != "" && txtReviewList.Contains(Vicheck1.Text))
                txtReviewList = txtReviewList.Replace(Vicheck1.Text, "c1");

            if (Vicombo1.Text != "" && txtReviewList.Contains(Vicombo1.Text))
                txtReviewList = txtReviewList.Replace(Vicombo1.Text, "m1");
            if (Vicombo2.Text != "" && txtReviewList.Contains(Vicombo2.Text))
                txtReviewList = txtReviewList.Replace(Vicombo2.Text, "m2");

            if (LibtnAdd1.Text != "" && txtReviewList.Contains(LibtnAdd1.Text))
                txtReviewList = txtReviewList.Replace(LibtnAdd1.Text, "a1");
            if (VitxtDate1.Text != "" && txtReviewList.Contains(VitxtDate1.Text))
                txtReviewList = txtReviewList.Replace(VitxtDate1.Text, "n1");
            if (txtReviewList.Contains(مقدم_الطلب.Text))
                txtReviewList = txtReviewList.Replace(مقدم_الطلب.Text, "tN");
            if (txtReviewList.Contains(رقم_الهوية.Text))
                txtReviewList = txtReviewList.Replace(رقم_الهوية.Text, "tP");
            if (txtReviewList.Contains(مكان_الإصدار.Text))
                txtReviewList = txtReviewList.Replace(مكان_الإصدار.Text, "tS");
            try
            {
                if (txtReviewList.Contains(تاريخ_الميلاد.Text))
                txtReviewList = txtReviewList.Replace(تاريخ_الميلاد.Text, "tB");
            if (txtReviewList.Contains(نوع_الهوية.Text))
                txtReviewList = txtReviewList.Replace(نوع_الهوية.Text, "tD");
            
                if (txtReviewList.Contains(title.Text))
                    txtReviewList = txtReviewList.Replace(title.Text, "tT");
                if (txtReviewList.Contains(انتهاء_الصلاحية.Text))
                    txtReviewList = txtReviewList.Replace(title.Text, "fD");
            }
            catch (Exception ex) { }
            txtReviewList = SuffOrigConvertments(txtReviewList);
            Console.WriteLine(txtReviewList);

            return txtReviewList;
        }
        public void resetBoxes()
        {
            //txtReview.Text = "";
            //checkAutoUpdate.Checked = true;
            foreach (Control control in PanelItemsboxes.Controls)
            {
                control.Visible = false;
                control.Text = "";
                control.Height = 35;
                if (control is ComboBox) { ((ComboBox)control).Items.Clear(); }
                else if (control is CheckBox) ((CheckBox)control).CheckState = CheckState.Unchecked;
            }
        }

        private void flllPanelItemsboxes(string rowID, string cellValue)
        {
            //MessageBox.Show("rowID = " + rowID + " - cellValue=" + cellValue);
            string query = "select * from TableAddContext where altColName = N'" + نوع_المعاملة.Text + "' and altSubColName = N'" + نوع_الإجراء.Text + "' and ColRight = ''";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            //MessageBox.Show(query);
            Console.WriteLine(query + " - " + dtbl.Rows.Count.ToString());
            string lang = "العربية";
            if (dtbl.Rows.Count > 0)
                foreach (DataRow dr in dtbl.Rows)
                {
                    ColName = dr["ColName"].ToString();
                    lang = dr["Lang"].ToString();
                    ColRight = dr["ColRight"].ToString();
                    if(dr["TextModel"].ToString().Contains(":"))
                        textModel = dr["TextModel"].ToString().Split(':')[0] + ":"+Environment.NewLine;
                    else 
                        textModel = dr["TextModel"].ToString();
                    Console.WriteLine(textModel);
                    ProTitle = dr["titleDefault"].ToString();
                    //MessageBox.Show(dr["titleDefault"].ToString());
                    try
                    {
                        txtReviewListStar[0] = textModel;
                    }
                    catch (Exception ex) { }
                    if (lang == "الانجليزية")
                        اللغة.Checked = true;
                    else اللغة.Checked = false;
                    
                    btnPrevious.Visible = true;
                    if (dr["TextModel"].ToString().Contains(":"))
                        StrSpecPur = dr["TextModel"].ToString().Split(':')[0] + ":" + Environment.NewLine;
                    else 
                        StrSpecPur = dr["TextModel"].ToString();

                    foreach (Control Lcontrol in PanelItemsboxes.Controls)
                        try
                        {
                            if (Lcontrol is Button && dr["ibtnAdd1"].ToString() != "")
                            {
                                addInfo = true;
                                PanelButtonInfo.Visible = true;
                                labl1.Text = dr["itext1"].ToString();
                                labl2.Text = dr["itext2"].ToString();
                                labl3.Text = dr["itext3"].ToString();
                                labl4.Text = dr["itext4"].ToString();
                                labl5.Text = dr["itext5"].ToString();
                            }
                            if (Lcontrol.Name.StartsWith("L"))
                            {
                                Lcontrol.Text = dr[Lcontrol.Name.Replace("L", "")].ToString();

                                if (Lcontrol.Text != "")
                                {
                                    Lcontrol.Visible = true;


                                    foreach (Control Vcontrol in PanelItemsboxes.Controls)
                                    {
                                        if (Vcontrol.Name.Trim() == Lcontrol.Name.Replace("L", "V").Trim())
                                        {
                                            if (Vcontrol.Name.Contains("combo"))
                                            {
                                                ((ComboBox)Vcontrol).Items.Clear();
                                                string[] items = dr[Lcontrol.Name.Replace("L", "") + "Option"].ToString().Split('_');

                                                for (int x = 0; x < items.Length; x++)
                                                    ((ComboBox)Vcontrol).Items.Add(items[x]);
                                                if (((ComboBox)Vcontrol).Items.Count > 0) ((ComboBox)Vcontrol).SelectedIndex = 0;
                                            }
                                            if (Vcontrol.Name.Contains("check"))
                                            {
                                                Vcontrol.Text = dr[Lcontrol.Name.Replace("L", "") + "Option"].ToString().Split('_')[1];
                                                checkOptions[checkIndex] = dr[Lcontrol.Name.Replace("L", "") + "Option"].ToString();
                                                checkIndex++;
                                            }

                                            if (Vcontrol.Name.Contains("ibtnAdd"))
                                            {
                                                Vcontrol.Text = dr[Lcontrol.Name].ToString();
                                            }
                                            Vcontrol.Visible = true;
                                            //string size = dr[Lcontrol.Name.Replace("L", "") + "Length"].ToString();
                                            //Vcontrol.Width = Convert.ToInt32(size);

                                            Vcontrol.Width = (Vcontrol.Text.Length * 8) + 1;
                                            if (Vcontrol.Width < 100)
                                                Vcontrol.Width = 100;


                                            //if (Convert.ToInt32(size) >= 700)
                                            //{
                                            //    if (Vcontrol is TextBox) ((TextBox)Vcontrol).Multiline = true;
                                            //    Vcontrol.Height = 150;
                                            //}

                                            //if (Vcontrol.Name.Contains("combo"))
                                            //{
                                            //    MessageBox.Show(Vcontrol.Name.Replace("", "") + "Option");
                                            //    ((ComboBox)Vcontrol).Items.Clear();

                                            //    string[] items = dr[Lcontrol.Name.Replace("", "") + "Option"].ToString().Split('_');

                                            //    for (int x = 0; x < items.Length; x++)
                                            //        ((ComboBox)Vcontrol).Items.Add(items[x]);
                                            //}
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
                            Console.WriteLine(Lcontrol.Name.Replace("L", ""));
                        }
                    return;
                }

        }

        private string checkStarTextExist(string dataSource, string col, string star)
        {
            string starText = "";
            string query = "select " + col + " from TableCollectStarText where ID=N'" + star + "'";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            Console.WriteLine(query);
            //MessageBox.Show(col);
            try
            {
                sqlDa.Fill(dtbl);
            }
            catch (Exception ex) { return ""; }
            sqlCon.Close();
            foreach (DataRow dr in dtbl.Rows)
            {
                if(!dr[col].ToString().Contains("remove") && dr[col].ToString().Contains("approved"))
                    starText = dr[col].ToString();
            }
            return starText;
        }

        private void checkAutoUpdate_CheckedChanged(object sender, EventArgs e)
        {
            if (checkAutoUpdate.Checked)
            {
                checkAutoUpdate.Text = "تحديث تلقائي";
            }
            else
            {
                if (اللغة.Checked)
                    boxesPreparationsEnglish(addNameIndex, نوع_المعاملة.SelectedIndex);
                else 
                    boxesPreparationsArabic(addNameIndex, نوع_المعاملة.SelectedIndex);

                oldText = txtReview.Text;
                checkAutoUpdate.Text = "إيقاف التحديث";

            }
        }

        private void التاريخ_الميلادي_TextChanged(object sender, EventArgs e)
        {
            التاريخ_الميلادي_off.Text = التاريخ_الميلادي.Text.Split('-')[1] + " - " + التاريخ_الميلادي.Text.Split('-')[0] + " - " + التاريخ_الميلادي.Text.Split('-')[2];
        }

        private void Vicheck1_CheckedChanged(object sender, EventArgs e)
        {
            if (Vicheck1.Checked) Vicheck1.Text = checkOptions[0].Split('_')[0];
            else Vicheck1.Text = checkOptions[0].Split('_')[1];
        }

        private void Vicheck2_CheckedChanged(object sender, EventArgs e)
        {
            if (Vicheck2.Checked) Vicheck2.Text = checkOptions[1].Split('_')[0];
            else Vicheck2.Text = checkOptions[1].Split('_')[1];
        }

        private void Vicheck3_CheckedChanged(object sender, EventArgs e)
        {
            if (Vicheck3.Checked) Vicheck3.Text = checkOptions[2].Split('_')[0];
            else Vicheck3.Text = checkOptions[2].Split('_')[1];
        }

        private void Vicheck4_CheckedChanged(object sender, EventArgs e)
        {
            if (Vicheck4.Checked) Vicheck4.Text = checkOptions[3].Split('_')[0];
            else Vicheck4.Text = checkOptions[3].Split('_')[1];
        }

        private void Vicheck5_CheckedChanged(object sender, EventArgs e)
        {
            if (Vicheck5.Checked) Vicheck5.Text = checkOptions[4].Split('_')[0];
            else Vicheck5.Text = checkOptions[4].Split('_')[1];
        }

        private void طريقة_الطلب_CheckedChanged(object sender, EventArgs e)
        {
            if (طريقة_الطلب.Checked)
            {
                طريقة_الطلب.Text = "حضور مباشرة إلى القنصلية";
                mandoubLabel.Visible = اسم_المندوب.Visible = false;
                اسم_المندوب.Text = "";
            }
            else
            {
                طريقة_الطلب.Text = "عن طريق أحد مندوبي القنصلية";
                اسم_المندوب.Visible = mandoubLabel.Visible = true;

                اسم_المندوب.Text = "إختر اسم المندوب";

            }


        }

        private void btnPanelapp_Click(object sender, EventArgs e)
        {
            if (btnPanelapp.Height != 130)
            {
                btnPanelapp.Height = Panelapp.Height = btnSave.Height = 130;
                btnPanelapp.Text = "عرض";
            }
            else
            {
                btnPanelapp.Height = Panelapp.Height = btnSave.Height = 130 * addNameIndex;
                btnPanelapp.Text = "تصغير";
            }
        }

        private void موقع_المعاملة_SelectedIndexChanged(object sender, EventArgs e)
        {
            //التوقيع_off.Text = موقع_المعاملة.Text;
        }

        private void موقع_المعاملة_TextChanged(object sender, EventArgs e)
        {
             موقع_المعاملة_off.Text = موقع_المعاملة.Text;
            getTitle(DataSource, موقع_المعاملة.Text);

        }

        private void طريقة_الطلب_TextChanged(object sender, EventArgs e)
        {
            if (طريقة_الطلب.Text == "حضور مباشرة إلى القنصلية")
                طريقة_الطلب.Checked = الشاهد_الأول.Enabled = هوية_الأول.Enabled = true;
            else طريقة_الطلب.Checked = الشاهد_الأول.Enabled = هوية_الأول.Enabled = false;

        }

        private void LibtnAdd1_Click(object sender, EventArgs e)
        {
            addButtonInfo(Vitext1.Text, Vitext2.Text, Vitext3.Text, Vitext4.Text, Vitext5.Text);
            Vitext1.Text = Vitext2.Text = Vitext3.Text = Vitext4.Text = Vitext5.Text = "";
        }

        private void طريقة_الإجراء_CheckedChanged(object sender, EventArgs e)
        {
            if (طريقة_الإجراء.Checked)
            {                
                panelAuthen.Visible = false;
                flowLayoutPanel1.Size = new System.Drawing.Size(940, 110);
                طريقة_الإجراء.Text = "حضور بالأصالة";
                label18.Visible = تاريخ_إصدار_الوكالة.Visible = label15.Visible = اسم_الموكل_بالتوقيع.Visible = label16.Visible = رقم_الوكالة.Visible = label17.Visible = جهة_إصدار_الوكالة.Visible = label18.Visible = تاريخ_إصدار_الوكالة.Visible = false;
                اسم_الموكل_بالتوقيع.Text = رقم_الوكالة.Text = جهة_إصدار_الوكالة.Text = تاريخ_إصدار_الوكالة.Text = "بدون";
                
            }
            else
            {
                panelAuthen.Visible = true;
                flowLayoutPanel1.Size = new System.Drawing.Size(940, 188);
                اسم_الموكل_بالتوقيع.Text = رقم_الوكالة.Text = جهة_إصدار_الوكالة.Text = تاريخ_إصدار_الوكالة.Text = "";
                تاريخ_إصدار_الوكالة.Visible = label18.Visible = جهة_إصدار_الوكالة.Visible = label17.Visible = رقم_الوكالة.Visible = label16.Visible = اسم_الموكل_بالتوقيع.Visible = label15.Visible = نوع_الموقع.Visible = true;
                طريقة_الإجراء.Text = "حضور بالإنابة";
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (طريقة_الإجراء.Checked) return;
            else طريقة_الطلب.Checked = true;

        }

        private void VitxtDate2_TextChanged(object sender, EventArgs e)
        {

        }
        string lastInput2 = "";
        private void VitxtDate1_TextChanged(object sender, EventArgs e)
        {
            if (VitxtDate1.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(VitxtDate1.Text, 4, 5));
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //VitxtDate1.Text = "";
                    VitxtDate1.Text = SpecificDigit(VitxtDate1.Text, 3, 10);
                    return;
                }
            }

            if (VitxtDate1.Text.Length == 11)
            {
                VitxtDate1.Text = lastInput2; return;
            }
            if (VitxtDate1.Text.Length == 10) return;
            if (VitxtDate1.Text.Length == 4) VitxtDate1.Text = "-" + VitxtDate1.Text;
            else if (VitxtDate1.Text.Length == 7) VitxtDate1.Text = "-" + VitxtDate1.Text;
            lastInput2 = VitxtDate1.Text;
        }
        string lastInput3 = "";
        private void تاريخ_إصدار_الوكالة_TextChanged(object sender, EventArgs e)
        {
            if (تاريخ_إصدار_الوكالة.Text.Length == 10)
            {
                try
                {
                    int month = Convert.ToInt32(SpecificDigit(تاريخ_إصدار_الوكالة.Text, 4, 5));
                    if (month > 12)
                    {
                        MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                        //تاريخ_إصدار_الوكالة.Text = "";
                        تاريخ_إصدار_الوكالة.Text = SpecificDigit(تاريخ_إصدار_الوكالة.Text, 3, 10);
                        return;
                    }
                }
                catch (Exception ex) { return; }
            }

            if (تاريخ_إصدار_الوكالة.Text.Length == 11)
            {
                تاريخ_إصدار_الوكالة.Text = lastInput3; return;
            }
            if (تاريخ_إصدار_الوكالة.Text.Length == 10) return;
            if (تاريخ_إصدار_الوكالة.Text.Length == 4) تاريخ_إصدار_الوكالة.Text = "-" + تاريخ_إصدار_الوكالة.Text;
            else if (تاريخ_إصدار_الوكالة.Text.Length == 7) تاريخ_إصدار_الوكالة.Text = "-" + تاريخ_إصدار_الوكالة.Text;
            lastInput3 = تاريخ_إصدار_الوكالة.Text;
            if (نوع_المعاملة.Text == "إقرار" || نوع_المعاملة.Text == "إقرار مشفوع باليمين")
                authJob();
        }

        private void طريقة_الإجراء_TextChanged(object sender, EventArgs e)
        {
            if (طريقة_الإجراء.Text == "حضور بالأصالة")
                طريقة_الإجراء.Checked = true;
            else طريقة_الإجراء.Checked = false;
        }

        private void نوع_الموقع_CheckedChanged_1(object sender, EventArgs e)
        {
            if (نوع_الموقع.Text == "السيد")
                نوع_الموقع.Checked = true;
            else نوع_الموقع.Checked = false;
        }

        private void نوع_الموقع_TextChanged(object sender, EventArgs e)
        {
            if (نوع_الموقع.Text == "السيد")
                نوع_الموقع.Checked = true;
            else نوع_الموقع.Checked = false;
        }

        private void FormCollection_FormClosed(object sender, FormClosedEventArgs e)
        {
            string primeryLink = @"D:\PrimariFiles\";
            if (!Directory.Exists(@"D:\"))
            {
                string appFileName = Environment.GetCommandLineArgs()[0];
                string directory = System.IO.Path.GetDirectoryName(appFileName);
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

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            autoCompleteTextBox(Vitext1, DataSource, "Vitext1", "TableCollection");
            autoCompleteTextBox(Vitext2, DataSource, "Vitext2", "TableCollection");
            autoCompleteTextBox(Vitext3, DataSource, "Vitext3", "TableCollection");
            autoCompleteTextBox(Vitext4, DataSource, "Vitext4", "TableCollection");
            autoCompleteTextBox(Vitext5, DataSource, "Vitext5", "TableCollection");
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            ColorFulGrid9();
        }

        private void fileUpdate_Click(object sender, EventArgs e)
        {            
            if (txtReviewLast == "")
                txtReviewLast = txtReview.Text;

            if (txtRigIndex >= txtReviewListIndex)
                txtRigIndex = 0;
            
            if (txtReviewList[txtRigIndex] != "")
            {
                string text = txtReviewList[txtRigIndex];
                txtReview.Text = SuffReplacements(text, صفة_مقدم_الطلب_off.SelectedIndex);
                txtReview.Text = removeSpace(txtReview.Text, false);
            }
            txtRigIndex++; 
            button3.Text = "الاختيار من القائمة العامة (" + txtReviewListIndex.ToString() + "/" + txtRigIndex.ToString() + ")";
            
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
                if (sentence.Contains("ويعتبر االإقرار الصادر"))
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
        private void removeSpace(TextBox text)
        {
            text.Text = text.Text.Replace("،", "، ");
            for (; text.Text.Contains("  ");)
            {
                text.Text = text.Text.Replace("  ", " ");
            }
            text.Text = text.Text.Trim();
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

        private string SuffOrigConvertments(string text)
        {
            if (text == null || text == "") return "";
            //MessageBox.Show(text);
            foreach (Control control in PanelItemsboxes.Controls) 
            {
                int textIn = 0;
                
                if (control.Text != "" && control.Name == "Vitext"+ textIn.ToString() && text.Contains(control.Text)) 
                {
                    text = text.Replace(control.Text, "t"+ textIn.ToString());
                    textIn++;
                }
                textIn = 0;
                if (control.Text != "" && control.Name == "Vicheck" + textIn.ToString() && text.Contains(control.Text)) 
                {
                    text = text.Replace(control.Text, "c"+ textIn.ToString());
                    textIn++;
                }
                textIn = 0;
                if (control.Text != "" && control.Name == "Vicombo" + textIn.ToString() && text.Contains(control.Text)) 
                {
                    text = text.Replace(control.Text, "m"+ textIn.ToString());
                    textIn++;
                }
                textIn = 0;
                if (control.Text != "" && control.Name == "LibtnAdd" + textIn.ToString() && text.Contains(control.Text)) 
                {
                    text = text.Replace(control.Text, "a"+ textIn.ToString());
                    textIn++;
                }
                textIn = 0;
                if (control.Text != "" && control.Name == "VitxtDate" + textIn.ToString() && text.Contains(control.Text)) 
                {
                    text = text.Replace(control.Text, "n"+ textIn.ToString());
                    textIn++;
                }
            }
               
            if (text.Contains(مقدم_الطلب.Text))
                text = text.Replace(مقدم_الطلب.Text, "tN");
            if (text.Contains(رقم_الهوية.Text))
                text = text.Replace(رقم_الهوية.Text, "tP");
            if (text.Contains(مكان_الإصدار.Text))
                text = text.Replace(مكان_الإصدار.Text, "tS");
            if (text.Contains(تاريخ_الميلاد.Text))
                text = text.Replace(تاريخ_الميلاد.Text, "tB");
            if (text.Contains(نوع_الهوية.Text))
                text = text.Replace(نوع_الهوية.Text, "tD");
            if (text.Contains(انتهاء_الصلاحية.Text))
                text = text.Replace(انتهاء_الصلاحية.Text, "fD");

            try
            {
                try
                {
                    string[] words = text.Split(' ');


                    foreach (string word in words)
                    {
                        if (word == "" || word == " ") continue;
                        for (int gridIndex = 0; gridIndex < dataGridView2.Rows.Count - 1; gridIndex++)
                        {
                            string code = dataGridView2.Rows[gridIndex].Cells["الرموز"].Value.ToString();
                            string[] replacemests = new string[6];
                            replacemests[0] = dataGridView2.Rows[gridIndex].Cells["المقابل1"].Value.ToString();
                            replacemests[1] = dataGridView2.Rows[gridIndex].Cells["المقابل2"].Value.ToString();
                            replacemests[2] = dataGridView2.Rows[gridIndex].Cells["المقابل3"].Value.ToString();
                            replacemests[3] = dataGridView2.Rows[gridIndex].Cells["المقابل4"].Value.ToString();
                            replacemests[4] = dataGridView2.Rows[gridIndex].Cells["المقابل5"].Value.ToString();
                            replacemests[5] = dataGridView2.Rows[gridIndex].Cells["المقابل6"].Value.ToString();

                            for (int cellIndex = 0; cellIndex < 6; cellIndex++)
                            {
                                if (word == replacemests[cellIndex])
                                {
                                    text = text.Replace(word, replacemests[0]);
                                    break;
                                }
                                else if (word == replacemests[cellIndex] + "،")
                                {
                                    text = text.Replace(word, replacemests[0] + "،");
                                    break;
                                }
                            }

                        }
                    }
                }
                catch (Exception ex) { }
            }
            catch (Exception ex) { }
            //MessageBox.Show(text);
            return text;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            txtReview.Text = originTextReview;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (currentPanelIndex > 0) currentPanelIndex--;
            else return;
            if (currentPanelIndex == 0) FillDataGridView(DataSource, year);
            panelShow(currentPanelIndex);
        }

        private void ListSearch_TextChanged(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[آلية_البحث.Text.Replace(" ","_")].HeaderText.ToString() + " LIKE '" + ListSearch.Text + "%'";
            dataGridView1.DataSource = bs;
            ColorFulGrid9();
        }

        private void picStar_VisibleChanged(object sender, EventArgs e)
        {

        }

        private void PanelButtonInfo_VisibleChanged(object sender, EventArgs e)
        {
            if (PanelButtonInfo.Visible)
            {
                PaneltxtReview.Height = 231;
                PaneltxtReview.AutoScroll = true;
            }
            else
            {
                PaneltxtReview.Height = 410;
                PaneltxtReview.AutoScroll = false;
            }
        }

        private void picStar_Click(object sender, EventArgs e)
        {            
            if (txtReviewLast == "")
                txtReviewLast = txtReview.Text;
            if (starRightIndexStar >= txtReviewListIndexStar)
                starRightIndexStar = 0;
            if (txtReviewListStar[starRightIndexStar] != "")
            {
                string text = txtReviewListStar[starRightIndexStar];
                txtReview.Text = SuffReplacements(text, صفة_مقدم_الطلب_off.SelectedIndex);
                txtReview.Text = removeSpace(txtReview.Text, false);
            }
            button3.Text = "الاختيار من قائمة المفضلة (" + txtReviewListIndexStar.ToString() + "/" + starRightIndexStar.ToString() + ")";            
            starRightIndexStar++;
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
        private void الشاهد_الأول_TextChanged(object sender, EventArgs e)
        {
            string[] text = getID(الشاهد_الأول);
            هوية_الأول.Text = text[0];
        }

        private void الشاهد_الثاني_TextChanged(object sender, EventArgs e)
        {
            string[] text = getID(الشاهد_الثاني);
            هوية_الثاني.Text = text[0];
        }

        private void picStarRightAdd_Click(object sender, EventArgs e)
        {
            txtReview.Text = SuffConvertments(txtReview.Text, صفة_مقدم_الطلب_off.SelectedIndex, 0, false);
            if (txtReviewLast == "")
                txtReviewLast = txtReview.Text;

            //string ID = updateText(txtReview);
            //if (starText.Contains(ID) && ID != "") return;
            //if (starText == "")
            //    starText = ID;
            //else starText = ID + "_" + starText;
            //string query = "UPDATE TableAddContext SET starText = N'" + starText + "' where altColName = N'" + نوع_المعاملة.Text + "' and altSubColName = N'" + نوع_الإجراء.Text.ToString() + "'";
            //Console.WriteLine(query);
            //SqlConnection sqlCon = new SqlConnection(DataSource);
            //if (sqlCon.State == ConnectionState.Closed)
            //    sqlCon.Open();
            //SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            //sqlCmd.CommandType = CommandType.Text;
            //sqlCmd.ExecuteNonQuery();
        }
        private string updateText(TextBox text)
        {
            string starButton = "";
            if (!checkColExistance(selectTable, نوع_المعاملة.Text.Replace(" ", "_") + "_" + نوع_الإجراء.Text.Replace(" ", "_")))
                CreateColumn(نوع_المعاملة.Text.Replace(" ", "_") + "_" + نوع_الإجراء.Text.Replace(" ", "_"), selectTable, "max");
            else
            {
                starButton = getStarID(selectTable, نوع_المعاملة.Text.Replace(" ", "_") + "_" + نوع_الإجراء.Text.Replace(" ", "_"), text.Text);
                //MessageBox.Show("نص موجود بالرقم " + starButton);
            }

            string textRevers = SuffReversReplacements(text.Text, 0, 0);
            string query = "UPDATE " + selectTable + " SET " + نوع_المعاملة.Text.Replace(" ", "_") + "_" + نوع_الإجراء.Text.Replace(" ", "_") + "=N'" + textRevers + "' where ID = " + starButton;
            if (starButton == "")
                query = "insert into " + selectTable + " (" + نوع_المعاملة.Text.Replace(" ", "_") + "_" + نوع_الإجراء.Text.Replace(" ", "_") + ") values (N'" + textRevers + "');SELECT @@IDENTITY as lastid";

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            Console.WriteLine("updateText " + query);
            //MessageBox.Show("starButton " + starButton);

            if (starButton != "")
                sqlCmd.ExecuteNonQuery();
            else
            {
                var reader = sqlCmd.ExecuteReader();
                if (reader.Read())
                {
                    return reader["lastid"].ToString();
                }
                sqlCon.Close();
            }
            return starButton;
        }
        private string SuffReversReplacements(string text, int appCaseIndex, int intAuthcases)
        {
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
                if (text.Contains(Vitext6.Text))
                    text = text.Replace(Vitext6.Text, "t6");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vitext7.Text))
                    text = text.Replace(Vitext7.Text, "t7");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vitext8.Text))
                    text = text.Replace(Vitext8.Text, "t8");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vitext9.Text))
                    text = text.Replace(Vitext9.Text, "t9");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vitext0.Text))
                    text = text.Replace(Vitext0.Text, "t0");
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
                if (text.Contains(Vicheck2.Text))
                    text = text.Replace(Vicheck2.Text, "c2");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vicheck3.Text))
                    text = text.Replace(Vicheck3.Text, "c3");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vicheck4.Text))
                    text = text.Replace(Vicheck4.Text, "c4");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vicheck5.Text))
                    text = text.Replace(Vicheck5.Text, "c5");
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
                if (text.Contains(Vicombo3.Text))
                    text = text.Replace(Vicombo3.Text, "m3");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vicombo4.Text))
                    text = text.Replace(Vicombo4.Text, "m4");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(Vicombo5.Text))
                    text = text.Replace(Vicombo5.Text, "m5");
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
                if (text.Contains(VitxtDate2.Text))
                    text = text.Replace(VitxtDate2.Text, "n2");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(VitxtDate3.Text))
                    text = text.Replace(VitxtDate3.Text, "n3");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(VitxtDate4.Text))
                    text = text.Replace(VitxtDate4.Text, "n4");
            }
            catch (Exception ex1) { }
            try
            {
                if (text.Contains(VitxtDate5.Text))
                    text = text.Replace(VitxtDate5.Text, "n5");
            }
            catch (Exception ex1) { }
            text = SuffOrigConvertments(text);
            return text;
        }

        private string getStarID(string table, string colName, string text)
        {
            string ID = "";
            string query = "select ID from " + table + " where " + colName + " = N'" + text + "'";
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
                ID = row["ID"].ToString();
            }

            return ID;

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

        private void Vitext1_TextChanged(object sender, EventArgs e)
        {
            Vitext1.Width = (Vitext1.Text.Length * 8) + 1;
            if (Vitext1.Width < 100)
                Vitext1.Width = 100;

            if (Vitext1.Width > 500)
            {
                //Vitext1.Multiline = true;
                Vitext1.Height = (Vitext1.Width / 500 + 1) * 35;
            }

        }

        private void Vitext2_TextChanged(object sender, EventArgs e)
        {
            Vitext2.Width = (Vitext2.Text.Length * 8) + 1;
            if (Vitext2.Width < 100)
                Vitext2.Width = 100;
            if (Vitext2.Width > 500)
            {
                //Vitext2.Multiline = true;
                Vitext2.Height = (Vitext2.Width / 500 + 1) * 35;
            }

        }

        private void Vitext3_TextChanged(object sender, EventArgs e)
        {
            Vitext3.Width = (Vitext2.Text.Length * 8) + 1;
            if (Vitext2.Width < 100)
                Vitext2.Width = 100;
            if (Vitext2.Width > 500)
            {
                Vitext2.Multiline = true;
                Vitext2.Height = (Vitext2.Width / 500 + 1) * 35;
            }
        }

        private void Vitext4_TextChanged(object sender, EventArgs e)
        {
            Vitext4.Width = (Vitext4.Text.Length * 8) + 1;
            if (Vitext4.Width < 100)
                Vitext4.Width = 100;
            if (Vitext4.Width > 500)
            {
                Vitext4.Multiline = true;
                Vitext4.Height = (Vitext4.Width / 500 + 1) * 35;
            }
        }

        private void Vitext5_TextChanged(object sender, EventArgs e)
        {
            Vitext5.Width = (Vitext5.Text.Length * 8) + 1;
            if (Vitext5.Width < 100)
                Vitext5.Width = 100;
            if (Vitext5.Width > 500)
            {
                Vitext5.Multiline = true;
                Vitext5.Height = (Vitext5.Width / 500 + 1) * 35;
            }
        }

        private void Vitext6_TextChanged(object sender, EventArgs e)
        {
            Vitext6.Width = (Vitext6.Text.Length * 8) + 1;
            if (Vitext6.Width < 100)
                Vitext6.Width = 100;
            if (Vitext6.Width > 500)
            {
                Vitext6.Multiline = true;
                Vitext6.Height = (Vitext6.Width / 500 + 1) * 35;
            }
        }

        private void Vitext7_TextChanged(object sender, EventArgs e)
        {
            Vitext7.Width = (Vitext7.Text.Length * 8) + 1;
            if (Vitext7.Width < 100)
                Vitext7.Width = 100;
            if (Vitext7.Width > 500)
            {
                Vitext7.Multiline = true;
                Vitext7.Height = (Vitext7.Width / 500 + 1) * 35;
            }
        }

        private void Vitext8_TextChanged(object sender, EventArgs e)
        {
            Vitext8.Width = (Vitext8.Text.Length * 8) + 1;
            if (Vitext8.Width < 100)
                Vitext8.Width = 100;
            if (Vitext8.Width > 500)
            {
                Vitext8.Multiline = true;
                Vitext8.Height = (Vitext8.Width / 500 + 1) * 35;
            }
        }

        private void Vitext9_TextChanged(object sender, EventArgs e)
        {
            Vitext9.Width = (Vitext9.Text.Length * 8) + 1;
            if (Vitext9.Width < 100)
                Vitext9.Width = 100;
            if (Vitext9.Width > 500)
            {
                Vitext9.Multiline = true;
                Vitext9.Height = (Vitext9.Width / 500 + 1) * 35;
            }
        }

        private void Vitext0_TextChanged(object sender, EventArgs e)
        {
            Vitext0.Width = (Vitext0.Text.Length * 8) + 1;
            if (Vitext0.Width < 100)
                Vitext0.Width = 100;
            if (Vitext0.Width > 500)
            {
                Vitext0.Multiline = true;
                Vitext0.Height = (Vitext0.Width / 500 + 1) * 35;
            }

        }

        private void رقم_الوكالة_TextChanged(object sender, EventArgs e)
        {
            getDocInfoOnBehalf(DataSource);
            authJob();
        }

        private void getDocInfoOnBehalf(string source)
        {
            string query = "select * from TableAuth where رقم_المعاملة = N'" + رقم_الوكالة.Text + "'";
            SqlConnection sqlCon = new SqlConnection(source);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {

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
            catch (Exception ex) { }
        }

        private void اسم_الموكل_بالتوقيع_TextChanged(object sender, EventArgs e)
        {
            authJob();
        }
        private void authJob()
        {
            string witnesses = "في حضور الشاهدين المشار إليهما أعلاه ";
            if (الشاهد_الأول.Text == "" || الشاهد_الثاني.Text == "")
                witnesses = "";
            string auth = " المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15] + " المذكور" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15] + " أعلاه قد حضر" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " ووقع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " بتوقيع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4] + " على هذا الإقرار "+ witnesses+"وذلك بعد تلاوته علي" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4] + " وبعد أن فهم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " مضمونه ومحتواه";
            if (!طريقة_الطلب.Checked)
                auth = " المواطن" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15] + " المذكور" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 15] + " أعلاه قد وقع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " بتوقيع" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4] + " على هذا الإقرار "+ witnesses+"وذلك بعد تلاوته علي" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 4] + " وبعد أن فهم" + preffix[صفة_مقدم_الطلب_off.SelectedIndex, 3] + " مضمونه ومحتواه";
            
            if (طريقة_الطلب.Checked)
            {
                غرض_المعاملة.Text = "قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن" + auth + "، صدر تحت توقيعي وختم القنصلية العامة";
                if (طريقة_الإجراء.Checked)
                    غرض_المعاملة.Text = AuthTitle + " بالقنصلية العامة لجمهورية السودان بجدة، بأن" + auth + "، صدر تحت توقيعي وختم القنصلية العامة";
                else
                {
                    auth = " بأن المواطن" + preffix[onBehalfIndex, 5] + " /" + اسم_الموكل_بالتوقيع.Text + " قد حضر" + preffix[onBehalfIndex, 3] + " ووقع" + preffix[onBehalfIndex, 3] + " بتوقيع" + preffix[onBehalfIndex, 4] + " على هذا الإقرار "+ witnesses+"بعد تلاوته علي" + preffix[onBehalfIndex, 4] + " وبعد أن فهم" + preffix[onBehalfIndex, 3] + " مضمونه ومحتواه، وذلك بناءً على حق توكي الغير الممنوح ل" + preffix[onBehalfIndex, 4] + " والمنصوص عليه بموجب التوكيل الصادر عن " + جهة_إصدار_الوكالة.Text + " بالرقم " + رقم_الوكالة.Text + " بتاريخ " + تاريخ_إصدار_الوكالة.Text;
                    التوقيع_off.Text = اسم_الموكل_بالتوقيع.Text;
                    غرض_المعاملة.Text = auth + "، صدر تحت توقيعي وختم القنصلية العامة";
                }
            }
            else
                غرض_المعاملة.Text = auth + "، صدر تحت توقيعي وختم القنصلية العامة";

        }

        private void جهة_إصدار_الوكالة_TextChanged(object sender, EventArgs e)
        {
            authJob();
        }

        private void رقم_المرجع_المرتبط_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(رقم_المرجع_المرتبط.Text);
        }

        private void اسم_الموظف_TextChanged(object sender, EventArgs e)
        {
            اسم_الموظف.Text = label36.Text.Replace("الموظف:","");
        }

        private void نوع_الإجراء_TextUpdate(object sender, EventArgs e)
        {
            reversTextReview();
        }

        private void نوع_الإجراء_TextChanged(object sender, EventArgs e)
        {
            reversTextReview();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtPurposeList.Length == 0 || txtPurposeList[txtPurIndex] == "")
                    return;

                if (originTextPurpose == "")
                    originTextPurpose = غرض_المعاملة.Text;

                if (txtPurIndex == txtPurposeList.Length)
                    txtPurIndex = 0;

                Console.WriteLine(غرض_المعاملة.Text);
                غرض_المعاملة.Text = SuffOrigConvertments(txtPurposeList[txtPurIndex]);
                Console.WriteLine(غرض_المعاملة.Text);
                غرض_المعاملة.Text = SuffReplacements(غرض_المعاملة.Text, صفة_مقدم_الطلب_off.SelectedIndex);
                Console.WriteLine(غرض_المعاملة.Text);
                Console.WriteLine("txtRigIndex = " + txtPurIndex.ToString());
                غرض_المعاملة.Text = removeSpace(غرض_المعاملة.Text, false);
                txtPurIndex++;
            }
            catch (Exception ex) { }
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            غرض_المعاملة.Text = originTextPurpose;
        }

        private void اسم_المندوب_TextChanged(object sender, EventArgs e)
        {
            if (اسم_المندوب.Text != "" && اسم_المندوب.Text != "إختر اسم المندوب" && اسم_المندوب.Text != "حضور مباشرة إلى القنصلية")
            {
                //MessageBox.Show("change");
                الشاهد_الأول.Enabled = هوية_الأول.Enabled = false;
                الشاهد_الأول.Text = اسم_المندوب.Text.Split('-')[0].Trim();
                هوية_الأول.Text = getMandoubPass(DataSource, اسم_المندوب.Text.Split('-')[0].Trim());
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

        private void txtReview_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(txtReview.Text);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            deleteFromCollection();
            deleteFromArch();
            FillDataGridView(DataSource, year);
            panelShow(0);
        }
        
        private void deleteFromCollection()
        {
            string query = "DELETE FROM TableCollection where ID = " + intID;
            SqlConnection Con = new SqlConnection(DataSource);            
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
        }
        
        private void deleteFromArch()
        {
            string query = "DELETE FROM TableGeneralArch where رقم_المرجع = '" + intID+ "' and docTable = 'TableCollection'";
            SqlConnection Con = new SqlConnection(DataSource.Replace("AhwalDataBase", "ArchFilesDB"));            
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FinalPrint("pdf");
        }

        private void FinalPrint(string doc)
        {
            if (getDocxPdf() != "docx")
                setDocxPdf(doc);

            حالة_الارشفة.Text = "غير مؤرشف";
            if (!save2DataBase(finalPanel, "البيانات النهائية"))
            {
                return;
            }

            if (!AuthenticatOther()) return;
            if (طريقة_الإجراء.Checked)
                التوثيق.Text = SuffReplacements(التوثيق.Text, صفة_مقدم_الطلب_off.SelectedIndex);
            else
            {
                if (نوع_الموقع.Checked)
                    التوثيق.Text = SuffReplacements(التوثيق.Text, 0);
                else التوثيق.Text = SuffReplacements(التوثيق.Text, 1);
            }
            if (!وجهة_المعاملة.Text.Contains("الخرطوم"))
            {
                direction_off.Text = "";
            }
            else 
                direction_off.Text = "- لا يعتمد هذا الإقرار بجمهورية السودان ما لم يتم توثيقه من وزارة خارجية جمهورية السودان";
            التاريخ_الميلادي.Text = GregorianDate;
            التاريخ_الهجري.Text = HijriDate;
            
            save2DataBase(panelapplicationInfo, "");
            //MessageBox.Show(نص_مقدم_الطلب0_off.Text);
            fillDocFileAppInfo(panelapplicationInfo);

            fillDocFileAppInfo(Panelapp);
            fillDocFileAppInfo(PaneltxtReview);
            fillDocFileAppInfo(finalPanel);
            fillDocFileAppInfo(PanelButtonInfo);
            chooseBtnTable();
            fillPrintDocx(doc);
            if (!وجهة_المعاملة.Text.Contains("الخرطوم") && !وجهة_المعاملة.Text.Contains("جدة"))
            {
                CreateMessageWord(مقدم_الطلب.Text, وجهة_المعاملة.Text, رقم_المعاملة.Text.Replace("_", " و"), "إقرار", preffix[صفة_مقدم_الطلب_off.SelectedIndex, 17], التاريخ_الميلادي_off.Text, HijriDate, موقع_المعاملة.Text);
            }
            string col = نوع_المعاملة .Text.Replace(" ", "_") + "_" + نوع_الإجراء.Text.Replace(" ", "_");
            calcAuthSub(DataSource, col, "TableCollection", "TableCollectStarText");
            addarchives();
            //fileUpload(رقم_المعاملة.Text, "missed");
            this.Close();
        }


        private void calcAuthSub(string dataSource, string column, string table, string genTable)
        {
            string txtReviewOrg = "";
            if (!checkColExist(genTable, column))
            {
                CreateColumn(column, genTable);
            }
            try
            {
                txtReviewOrg = SuffOrigConvertments(txtReview.Text);
            }
            catch (Exception ex1) { return; }
            if(!checkTextExist(genTable, column, txtReviewOrg))
            insertNewText(dataSource, column, txtReviewOrg, genTable);
        }
        private bool checkColExist(string table, string colName)
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
            return false;

        }
        private bool checkTextExist(string table, string colName, string text)
        {
            string query = "SELECT " + colName + " FROM " + table + " where " + colName + "=N'" + text.Replace("'","~") + "'";
            Console.WriteLine("checkTextExist " + query);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
            }catch (Exception ex) { return false; }
            sqlCon.Close();

            if (dtbl.Rows.Count > 0)
                return true;
            else 
            return false;

        }

        private string[] getMainCol(string dataSource, string colMain, string table)
        {
            string query = "select distinct " + colMain + " from " + table;
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            string[] columns = new string[dtbl.Rows.Count];
            int count = 0;
            foreach (DataRow row in dtbl.Rows)
            {
                columns[count] = row[colMain].ToString();
                count++;
            }
            sqlCon.Close();
            return columns;
        }
        private void مكان_الإصدار_TextChanged(object sender, EventArgs e)
        {

        }

        private void تاريخ_الميلاد_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnPanelapp_SizeChanged(object sender, EventArgs e)
        {
        }

        private void انتهاء_الصلاحية_TextChanged(object sender, EventArgs e)
        {

        }

        private void fileUpdate_MouseEnter(object sender, EventArgs e)
        {
            button3.Text = "الاختيار من القائمة العامة (" + txtReviewListIndex.ToString() + "/" + txtRigIndex.ToString() + ")";
        }

        private void picStar_MouseEnter(object sender, EventArgs e)
        {
            button3.Text = "الاختيار من قائمة المفضلة (" + txtReviewListIndexStar.ToString() + "/" + starRightIndexStar.ToString() + ")";
        }

        private void label6_MouseEnter(object sender, EventArgs e)
        {
            
        }

        private void pictureBox5_MouseEnter(object sender, EventArgs e)
        {
            label6.Text = "قائمة الأغراض (" + txtPurposeListIndex.ToString() + "/" + txtPurIndex.ToString() + ")";
        }

        private void اسم_المندوب_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (اسم_المندوب.SelectedIndex > 1 || اسم_المندوب.Text == "حضور مباشرة إلى القنصلية")
            {
                الشاهد_الأول.Enabled = هوية_الأول.Enabled = false;
            }
            else
                الشاهد_الأول.Enabled = هوية_الأول.Enabled = true;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (!checkEmpty(Panelapp) || !checkDate(Panelapp))
            {
                currentPanelIndex--; return;
            }
            save2DataBase(Panelapp, "مقدم الطلب");
        }

        private void PaneltxtReview_MouseEnter(object sender, EventArgs e)
        {
            //MessageBox.Show(PaneltxtReview.Height.ToString());
        }

        private void txtReview_MouseEnter(object sender, EventArgs e)
        {
            //MessageBox.Show(txtReview.Height.ToString());
        }

        private void yearSel_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillDataGridView(DataSource, yearSel.Text);
            
        }

        private void panelAuthen_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
