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

namespace PersAhwal
{
    public partial class FormCollection : Form
    {
        string DataSource = "";
        
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
        int appCaseIndex = 0;
        string StrSpecPur = "";
        string ColName = "";
        string ColRight = "Col0";
        string startText = "";
        string oldText = "";
        Word.Document oBDoc;
        object oBMiss;
        Word.Application oBMicroWord;
        bool notFiled = true;
        public FormCollection(int Atvc, int currentRow, int DocumentType, string empName, string dataSource, string filepathIn, string filepathOut, string jobposition, string gregorianDate, string hijriDate)
        {
            InitializeComponent();
            definColumn(dataSource);
            DataSource = dataSource;
            FilespathIn = filepathIn;
            FilespathOut = filepathOut;

            AtVCIndex = Atvc;
            
            EmpName = empName;
            Jobposition = jobposition;
            التاريخ_الميلادي.Text = GregorianDate = gregorianDate;
            التاريخ_الهجري.Text = HijriDate = hijriDate;
            genPreperations();
            FillDataGridView(DataSource);
            getMaxRange(DataSource);
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
                    allowedEdit.Text = reader["allowedEditCollec"].ToString();
                }
            }
            catch (Exception ex)
            {
                allowedEdit.Text = "0";
                Con.Close();
            }
        }
        public void FillDataGridView(string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableCollection", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            dataGridView1.Columns[0].Visible = false ;
            //dataGridView1.Columns["نوع_المعاملة"].Visible = false ;
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[2].Width = 200;
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
                allowedEdit.Enabled = true;
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

                    panelapplicationInfo.Size = new System.Drawing.Size(821, 649);
                    panelapplicationInfo.Location = new System.Drawing.Point(294, 38);
                    panelapplicationInfo.BringToFront();
                    btnPrevious.Visible = panelapplicationInfo.Visible = true;
                    //false
                    panelAuthRights.Visible = ListSearch.Visible = btnListView.Visible = panelAuthRights.Visible = PanelDataGrid.Visible = labDescribed.Visible = false;
                    btnDelete.Visible = btnFile1.Visible = btnFile2.Visible = btnFile3.Visible = Panelapp.Visible = true;
                    
                    break;
                case 2:
                    
                    if (!backgroundWorker1.IsBusy) backgroundWorker1.RunWorkerAsync();
                    //authrights
                    if (!save2DataBase(Panelapp) || !save2DataBase(panelapplicationInfo) ) 
                    {
                        currentPanelIndex--; return;
                    }
                    if(اللغة.Checked)
                    boxesPreparationsEnglish(addNameIndex, نوع_المعاملة.SelectedIndex);
                    else boxesPreparationsArabic(addNameIndex, نوع_المعاملة.SelectedIndex);
                    
                    
                    txtReview.Text = writeStrSpecPur();
                    panelAuthRights.Size = new System.Drawing.Size(1315, 622);
                    panelAuthRights.Location = new System.Drawing.Point(10, 36);
                    panelAuthRights.BringToFront();
                    panelAuthRights.Visible = btnNext.Visible = true;
                    PanelDataGrid.Visible = panelapplicationInfo.Visible = false;
                    timer1.Enabled = true;
                    break;
                case 3:
                    timer1.Enabled = false;
                    //finalPanel
                    //flowLayoutPanel2.Size = new System.Drawing.Size(940, 242);
                    addNonEmptyFields(PanelItemsboxes);
                    if (!save2DataBase(PanelItemsboxes))
                    {
                        currentPanelIndex--; return;
                    }
                    if (!save2DataBase(PaneltxtReview))
                    {
                        currentPanelIndex--; return;
                    }
                    
                    if (panelRemove.Visible)
                        if (!save2DataBase(panelRemove))
                        {
                            currentPanelIndex--; return;
                        }
                    finalPanel.Size = new System.Drawing.Size(944, 616);
                    finalPanel.Location = new System.Drawing.Point(192, 38);
                    finalPanel.BringToFront();
                    finalPanel.Visible = true;
                    panelAuthRights.Visible = btnNext.Visible = PanelDataGrid.Visible = panelapplicationInfo.Visible = false;

                    break;
            }
        }

        public void boxesPreparationsEnglish(int index, int proTypeIndex)
        {
            //addNameIndex
            switch (proTypeIndex) {
                case 0:
                    appCaseIndex = Appcases(النوع, index);
                    //إقرار 
                    if (index == 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "I " ;
                        نص_مقدم_الطلب1_off.Text = title.Text + ". " + مقدم_الطلب.Text + "holding a " + نوع_الهوية.Text + " No. " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P")+ " issued on " + مكان_الإصدار.Text + " solemnly declare and affirm that, ";
                    }
                    else if (index > 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "I ";
                        نص_مقدم_الطلب1_off.Text = title.Text + ". " + مقدم_الطلب.Text + "holding a " + نوع_الهوية.Text + " No. " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " issued on " + مكان_الإصدار.Text + " solemnly declare and affirm that, ";
                        return;
                    }

                    موقع_المعاملة_off.Text = "/" + موقع_المعاملة.Text.Trim();
                    break;
                case 2:
                    // افادة وشهادة لمن يهمه الامر
                    if (index == 1)
                    {
                        نص_مقدم_الطلب0_off.Text = " Sudanese citizen ";
                        نص_مقدم_الطلب1_off.Text =  title.Text + ". " + مقدم_الطلب.Text + "holding a " + نوع_الهوية.Text + " No. " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " issued on " + مكان_الإصدار.Text;
                    }
                    else if (index > 1)
                    {
                        نص_مقدم_الطلب0_off.Text = " Sudanese citizen mentioned below:";
                        نص_مقدم_الطلب1_off.Text = "";
                    }
                    التوثيق_off.Text = "This certificate has been issued upon " + preffix[appCaseIndex, 17] + " request,,,";
                    break;
                
            }
        } 
        public void boxesPreparationsArabic(int index, int proTypeIndex)
        {
            //addNameIndex
            switch (proTypeIndex) {
                case 0:
                    appCaseIndex = Appcases(النوع, index);
                    //إقرار 
                    if (index == 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "أنا المواطن" + preffix[appCaseIndex, 5];
                        نص_مقدم_الطلب1_off.Text = "/ " + مقدم_الطلب.Text + "، المقيم" + preffix[appCaseIndex, 5] + " بالمملكة العربية السعودية حامل" + preffix[appCaseIndex, 5] + " " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " إصدار " + مكان_الإصدار.Text + "، وبكامل قوا" + preffix[appCaseIndex, 12] + " العقلية وبطوع" + preffix[appCaseIndex, 12] + " واختيار" + preffix[appCaseIndex, 12] + " وحالت" + preffix[appCaseIndex, 12] + " المعتبرة شرعا وقانونا ";
                    }
                    else if (index > 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "نحن المواطن" + preffix[appCaseIndex, 5] + " الموقع" + preffix[appCaseIndex, 5] + " أدناه:";
                        نص_مقدم_الطلب1_off.Text = "المقيم" + preffix[appCaseIndex, 5] + " بالمملكة العربية السعودية، وبكامل قوا" + preffix[appCaseIndex, 12] + " العقلية وبطوع" + preffix[appCaseIndex, 12] + " واختيار" + preffix[appCaseIndex, 12] + " وحالت" + preffix[appCaseIndex, 12] + " المعتبرة شرعا وقانونا ";
                    }

                    موقع_المعاملة_off.Text = "/" + موقع_المعاملة.Text.Trim();
//                    MessageBox.Show(موقع_المعاملة_off.Text);
                    التوقيع_off.Text = مقدم_الطلب.Text;
  //                  MessageBox.Show(التوقيع_off.Text);
                    
                    string auth = " المواطن" + preffix[appCaseIndex, 5] + " المذكور" + preffix[appCaseIndex, 5] + " أعلاه قد حضر" + preffix[appCaseIndex, 3] + " ووقع" + preffix[appCaseIndex, 3] + " بتوقيع" + preffix[appCaseIndex, 4] + " على هذا الإقرار، وذلك بعد تلاوته علي" + preffix[appCaseIndex, 4] + " وبعد أن فهم" + preffix[appCaseIndex, 3] + " مضمونه ومحتواه";
                    if (!طريقة_الطلب.Checked)
                        auth = " المواطن" + preffix[appCaseIndex, 5] + " المذكور" + preffix[appCaseIndex, 5] + " أعلاه قد حضر" + preffix[appCaseIndex, 3] + " ووقع" + preffix[appCaseIndex, 3] + " بتوقيع" + preffix[appCaseIndex, 4] + " على هذا الإقرار" + " بعد تلاوته علي" + preffix[appCaseIndex, 4] + " وبعد أن فهم" + preffix[appCaseIndex, 3] + " مضمونه ومحتواه" + " وذلك أمام مندوب جالية منطقة " + اسم_المندوب.Text.Split('_')[1] + " السيد/ " + اسم_المندوب.Text.Split('_')[0] + " بموجب التفويض الممنوح له من القنصلية العامة ";
                    التوثيق_off.Text = "نائب قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن" + auth + "، صدر تحت توقيعي وختم القنصلية العامة";
                    
                    break;
                case 2:
                    // افادة وشهادة لمن يهمه الامر
                    if (index == 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "";// " المواطن" + preffix[appCaseIndex, 5] + " السواني" + preffix[appCaseIndex, 5] + " السيد" + preffix[appCaseIndex, 5];
                        نص_مقدم_الطلب1_off.Text = "";// "/ " + مقدم_الطلب.Text + "، المقيم" + preffix[appCaseIndex, 5] + " بالمملكة العربية السعودية حامل" + preffix[appCaseIndex, 5] + " " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " إصدار " + مكان_الإصدار.Text + "،";
                    }
                    else if (index > 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "";// المواطن" + preffix[appCaseIndex, 5] + " السوداني" + preffix[appCaseIndex, 5] + " المذكور" + preffix[appCaseIndex, 5] + " أدناه:";
                        نص_مقدم_الطلب1_off.Text = "";
                    }
                    التوثيق_off.Text = "حررت هذه الإفادة بناء على طلب المذكور" + preffix[appCaseIndex, 5] + " أعلاه لاستخدامها على الوجه المشروع";
                    break;
                case 3:
                    // افادة وشهادة لمن يهمه الامر
                    if (index == 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "المواطن" + preffix[appCaseIndex, 5] + "السواني" + preffix[appCaseIndex, 5] + "السيد" + preffix[appCaseIndex, 5];
                        نص_مقدم_الطلب1_off.Text = "";// "/ " + مقدم_الطلب.Text + "، المقيم" + preffix[appCaseIndex, 5] + " بالمملكة العربية السعودية حامل" + preffix[appCaseIndex, 5] + " " + نوع_الهوية.Text + " رقم " + رقم_الهوية.Text.Replace("p", "P") + " إصدار " + مكان_الإصدار.Text + "،";
                    }
                    else if (index > 1)
                    {
                        نص_مقدم_الطلب0_off.Text = "المواطن" + preffix[appCaseIndex, 5] + " السوداني" + preffix[appCaseIndex, 5] + " المذكور" + preffix[appCaseIndex, 5] + " أدناه:";
                        نص_مقدم_الطلب1_off.Text = "";
                    }
                    التوثيق_off.Text = "حررت هذه الشهادة بناء على طلب المذكور" + preffix[appCaseIndex, 5] + " أعلاه لاستخدامها على الوجه المشروع";
                    break;
            }
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
        private void addNonEmptyFields(FlowLayoutPanel panel) {
            foreach (Control Econtrol in panel.Controls)
            {
                if ((Econtrol is TextBox || Econtrol is ComboBox|| Econtrol is CheckBox) && Econtrol.Visible && !checkColumnName(Econtrol.Name, DataSource))
                {
                    CreateColumn(Econtrol.Name, DataSource);
                }
            }
        }
        private bool save2DataBase(FlowLayoutPanel panel)
        {
            string query = checkList(panel, allList, "TableCollection");
            //MessageBox.Show(query);
            if (query == "UPDATE TableCollection SET  where ID = @id") return true;
            Console.WriteLine(panel.Name +" - " +query);
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
                    sqlCommand.Parameters.AddWithValue("@" + foundList[i], commentInfo());
                }
                else
                    foreach (Control control in panel.Controls)
                    {
                        string name = control.Name;
                        //if (panel.Name == "PanelItemsboxes")
                        //    name = name.Replace("V", "");
                        if (name == foundList[i])
                        {
                            if (control.Name == "اسم_المندوب" && control.Visible && !control.Text.Contains("-"))
                            {
                                control.BackColor = System.Drawing.Color.MistyRose;
                                MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل اسم المندوب ومنطقة التغطية مفصولين بالعلامة -");
                                return false;
                            }
                            if (control.Visible && ((control is TextBox && control.Text == "") || (control is ComboBox && control.Text.Contains("إختر"))))
                                foreach (Control Econtrol in panel.Controls)
                                {
                                    if ((Econtrol is TextBox  || Econtrol is ComboBox) && control.Text == "")
                                        if (panel.Name != "PanelItemsboxes" || (Econtrol.Name != control.Name && Econtrol.Name.Contains(control.Name)) || Econtrol.Name == "اسم_المندوب")
                                        {
                                            //MessageBox.Show(Econtrol.Name + " - " + control.Name);
                                            if (control.Name == "اسم_المندوب" && control.Visible)
                                            {
                                                //
                                                control.BackColor = System.Drawing.Color.MistyRose;
                                                MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل اسم_المندوب ");
                                                return false;
                                            }
                                            else if (control.Name != "اسم_المندوب"  && control.Name != "الشاهد_الأول" && control.Name != "الشاهد_الثاني")
                                            {
                                                Econtrol.BackColor = System.Drawing.Color.MistyRose;
                                                if (panel.Name == "Panelapp") { panel.Height = 130 * addNameIndex; }
                                                MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل " + control.Name.Replace("_", " "));
                                                return false;
                                            }
                                        }
                                        else if (panel.Name == "PanelItemsboxes")
                                        {
                                            if (control.Visible)
                                            {
                                                control.BackColor = System.Drawing.Color.MistyRose;
                                                MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل غير المكتمل");
                                                return false;
                                            }
                                        }
                                }

                            //if (panel.Name == "PanelItemsboxes") MessageBox.Show(control.Text);
                            sqlCommand.Parameters.AddWithValue("@" + foundList[i], control.Text);
                            break;
                        }
                    }
            }
            sqlCommand.ExecuteNonQuery();
            return true;
        }
        private string commentInfo()
        {
            string comment = "";
            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text == "")
                comment = "";

            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text != "")
                comment = التعليقات_السابقة_Off.Text;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text == "")
                comment = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + التاريخ_الميلادي.Text + Environment.NewLine + "--------------" + Environment.NewLine;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text != "")
                comment = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + التاريخ_الميلادي.Text + Environment.NewLine + "--------------" + Environment.NewLine + "*" + التعليقات_السابقة_Off.Text.Trim();

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
                if (control is TextBox || control is ComboBox)
                    for (int col = 0; col < List.Length; col++)
                        if (name == List[col])
                        {
                            foundList[found] = name;
                            //if (panel.Name == "PanelItemsboxes") MessageBox.Show(foundList[found]);
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


            preffix[0, 9] = "نصيبي";//#5
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
            
            preffix[0, 15] = "ي";//&^^
            preffix[1, 15] = "ت";
            preffix[2, 15] = "ي";
            preffix[3, 15] = "ت";
            preffix[4, 15] = "ي";
            preffix[5, 15] = "ي";

            preffix[0, 16] = "";//*%*
            preffix[1, 16] = "";
            preffix[2, 16] = "ا";
            preffix[3, 16] = "ا";
            preffix[4, 16] = "ن";
            preffix[5, 16] = "وا";

            preffix[0, 17] = "his";//&&*
            preffix[1, 17] = "her";
            preffix[2, 17] = "their";
            preffix[3, 17] = "their";
            preffix[4, 17] = "their";
            preffix[5, 17] = "their";

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
            forbidDs[1] = "حالة_الارشفة";
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
            return false;
        }
        private void CreateColumn(string Columnname, string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("alter table TableCollection add " + Columnname.Replace(" ", "_") + " nvarchar(150)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void allowedEdit_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("update TableSettings set allowedEdit=@allowedEdit", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@allowedEdit", Convert.ToInt32(allowedEdit.Text));
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            

            intID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            addNameIndex = 0;

            if (dataGridView1.CurrentRow.Index != -1)
            {
                fillInfo(Panelapp, true);
                fillInfo(panelapplicationInfo, false);
                fillInfo(PanelItemsboxes, false);
                
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
                        addName(مقدم_الطلب.Text.Split('_')[app], النوع.Text.Split('_')[app], نوع_الهوية.Text.Split('_')[app], رقم_الهوية.Text.Split('_')[app], مكان_الإصدار.Text.Split('_')[app], "العربية", appJob, appBirth);
                        archState = "old";
                    }
                    else
                    {
                        addName("", "ذكر", "جواز سفر", "P0", "", "العربية", "", appBirth);
                        archState = "new";
                        FillDatafromGenArch("data1", intID.ToString(), "TableCollection");
                    }
                }

                currentPanelIndex = 1;
                panelShow(currentPanelIndex);
            }
            checkChanged(مقدم_الطلب, Panelapp);
            checkChanged(النوع, Panelapp);
            checkChanged(نوع_الهوية, Panelapp);
            checkChanged(رقم_الهوية, Panelapp);
            checkChanged(مكان_الإصدار, Panelapp);
            checkChanged(تاريخ_الميلاد, Panelapp);
            checkChanged(المهنة, Panelapp);

        }
        void FillDatafromGenArch(string doc, string id, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
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
        public void addName(string name, string sex, string docType, string docNo, string docIssue, string language, string job, string age)
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
            ComboBox combTitle = new ComboBox();
            combTitle.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            combTitle.FormattingEnabled = true;
            combTitle.Items.AddRange(new object[] {
            "Mr",
            "Mrs",
            "Miss",
            "Madam"});
            combTitle.Location = new System.Drawing.Point(291, 3);
            combTitle.Name = "title_" + addNameIndex + ".";
            combTitle.Size = new System.Drawing.Size(15, 35);
            combTitle.TabIndex = 189;
            combTitle.Visible = false;
            combTitle.Text = sex;
            if (language == "العربية")
            {
                checkSexType.Visible = true;
                combTitle.Visible = false;
            }
            else
            {
                checkSexType.Visible = false;
                combTitle.Visible = true;
            }

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
            DocNo.Size = new System.Drawing.Size(120, 35);
            DocNo.TabIndex = 120;
            DocNo.Tag = "pass";
            DocNo.Text = docNo;
            DocNo.TextChanged += new System.EventHandler(this.DocNo_TextChanged);
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
            textJob.Size = new System.Drawing.Size(570, 35);
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

            Panelapp.Controls.Add(labelName);
            Panelapp.Controls.Add(AppName);
            Panelapp.Controls.Add(labeltitle1);
            Panelapp.Controls.Add(checkSexType);
            Panelapp.Controls.Add(combTitle);
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

            //Panelapp.Height = 130 * (addNameIndex);
        }
        private void addName_Click(object sender, EventArgs e)
        {
            addName("", "ذكر", "جواز سفر", "P0", "", "العربية", "", "");
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
                addName("", "ذكر", "جواز سفر", "P0", "", "العربية", "", "");

            }
            checkChanged(مقدم_الطلب, Panelapp);
            checkChanged(النوع, Panelapp);
            checkChanged(نوع_الهوية, Panelapp);
            checkChanged(رقم_الهوية, Panelapp);
            checkChanged(مكان_الإصدار, Panelapp);
            checkChanged(تاريخ_الميلاد, Panelapp);
            checkChanged(المهنة, Panelapp);
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
                        
                    }

                }
            }
        }

        private void btnFile1_Click(object sender, EventArgs e)
        {
            btnFile1.Enabled = false;
            FillDatafromGenArch("data1", intID.ToString(), "TableCollection");
            btnFile1.Enabled = true;
        }

        private void btnFile2_Click(object sender, EventArgs e)
        {
            btnFile2.Enabled = false;
            FillDatafromGenArch("data2", intID.ToString(), "TablTableCollectioneAuth");
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
            if (btnPrint.InvokeRequired)
            {
                btnPrint.Invoke(new MethodInvoker(delegate { btnPrint.Enabled = false; }));
            }

            if (نوع_المعاملة.InvokeRequired)
            {
                نوع_المعاملة.Invoke(new MethodInvoker(delegate { docType = نوع_المعاملة.Text; }));
            }
            chooseDocxFile(مقدم_الطلب.Text.Split('_')[0], رقم_المعاملة.Text, docType);
            prepareDocxfile();
            if (btnPrint.InvokeRequired)
            {
                btnPrint.Invoke(new MethodInvoker(delegate { btnPrint.Enabled = true; btnPrint.Text = "طباعة المعاملة"; }));
            }
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
        private void chooseDocxFile(string appName, string docId, string docType)
        {
            string proType = "";
            if(addNameIndex > 1) proType = " متعدد";
            string RouteFile = RouteFile = FilespathIn + docType + proType +".docx";
            
            if (appName != "")
                localCopy.Text = FilespathOut + appName + DateTime.Now.ToString("ddmmss") + ".docx";
            else localCopy.Text = FilespathOut + docId.Replace("/", "_") + DateTime.Now.ToString("ddmmss") + ".docx";
            while (File.Exists(localCopy.Text))
            {
                if (appName != "")
                    localCopy.Text = FilespathOut + appName + DateTime.Now.ToString("ddmmss") + ".docx";
                else localCopy.Text = FilespathOut + docId.Replace("/", "_") + DateTime.Now.ToString("ddmmss") + ".docx";
            }
            //
            System.IO.File.Copy(RouteFile, localCopy.Text);
            FileInfo fileInfo = new FileInfo(localCopy.Text);
            if (fileInfo.IsReadOnly) fileInfo.IsReadOnly = false;
            //MessageBox.Show(localCopy.Text);
        }
        private void btnNext_Click(object sender, EventArgs e)
        {
            if (currentPanelIndex <= 4)
                currentPanelIndex++;
            else return;
            panelShow(currentPanelIndex);
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            if (currentPanelIndex > 0) currentPanelIndex--;
            else return;
            if (currentPanelIndex == 0) FillDataGridView(DataSource);
            panelShow(currentPanelIndex);
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
            fileComboBox2(نوع_المعاملة, DataSource, "ArabicGenIgrar", "TableListCombo");
            fileComboBox(اسم_المندوب, DataSource, "MandoubNames", "TableListCombo");
            fileComboBox(وجهة_المعاملة, DataSource, "ArabCountries", "TableListCombo");
            if(وجهة_المعاملة.Items.Count > 0 ) وجهة_المعاملة.SelectedIndex = 0; 
            //fileComboBoxAttend(DocType, DataSource, "DocType", "TableListCombo");
            //autoCompleteTextBox(DocSource, DataSource, "SDNIssueSource", "TableListCombo");
            fileComboBox(موقع_المعاملة, DataSource, "ArabicAttendVC", "TableListCombo");
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
            if (checkColumnName(نوع_المعاملة.Text.Replace(" ", "_")))
            {
                إجراء_التوكيل.Items.Clear();
                if (وجهة_المعاملة.Items.Count > 0)
                    وجهة_المعاملة.SelectedIndex = 0;
                newFillComboBox1(نوع_الإجراء, DataSource, نوع_المعاملة.Text.Replace(" ","_"));
                عنوان_المكاتبة.Items.Clear();

                عنوان_المكاتبة.Items.Add(نوع_المعاملة.Text);
                if (نوع_المعاملة.SelectedIndex == 2)
                {
                    تفيد_تشهد_off.Text = "فيد";                    
                    عنوان_المكاتبة.Items.Add("إفادة");
                }
                else if (نوع_المعاملة.SelectedIndex == 3)
                {
                    تفيد_تشهد_off.Text = "شهد";                    
                    عنوان_المكاتبة.Items.Add("شهادة");
                }
                else تفيد_تشهد_off.Text = "";
                عنوان_المكاتبة.SelectedIndex = 0;

            }
        }
        private void newFillComboBox1(ComboBox combbox, string source, string colName)
        {
            combbox.Visible = true;
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();
                string query = "select "+ colName+" from TableListCombo where "+ colName+" is not null";
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
                //MessageBox.Show(نوع_المعاملة.SelectedIndex.ToString());
            }
        }

        private void اللغة_CheckedChanged(object sender, EventArgs e)
        {
            if (!اللغة.Checked)
            {
                اللغة.Text = "العربية";
                fileComboBox2(نوع_المعاملة, DataSource, "ArabicGenIgrar", "TableListCombo");
                fileComboBox(موقع_المعاملة, DataSource, "ArabicAttendVC", "TableListCombo");
                موقع_المعاملة.Width = 150; 
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            }
            else
            {
                اللغة.Text = "الامجليزية";
                fileComboBox(نوع_المعاملة, DataSource, "EnglishGenIgrar", "TableListCombo");
                fileComboBoxAttend(موقع_المعاملة, DataSource, "EnglishAttendVC", "TableListCombo");
                موقع_المعاملة.Width = 250;
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("en-US");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            }
            if (نوع_المعاملة.Items.Count > 0) نوع_المعاملة.SelectedIndex = 0;
            if (موقع_المعاملة.Items.Count > AtVCIndex) موقع_المعاملة.SelectedIndex = AtVCIndex;
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (checkAutoUpdate.Checked)
                txtReview.Text = writeStrSpecPur() + ".";
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            int count = getEdited(GregorianDate);

            fillDocFileAppInfo(panelapplicationInfo);
            fillDocFileAppInfo(Panelapp);
            fillDocFileAppInfo(PaneltxtReview);
            fillDocFileAppInfo(finalPanel);

            fillPrintDocx(edited.Text);
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
                        BookAuthIDNo.Text = control.Text;
                        object rangeAuthIDNo = BookAuthIDNo;
                        oBDoc.Bookmarks.Add(control.Name, ref rangeAuthIDNo);

                        //MessageBox.Show(panel.Name+ " - "+control.Name+ " - "+control.Text);
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
                appNameInfo(نوع_المعاملة.SelectedIndex);
                notFiled = false;
            }
        }

        private void appNameInfo(int appindex) 
        {
            switch (appindex) {
                case 0:
                    if (addNameIndex == 1)
                    {
                        Microsoft.Office.Interop.Word.Table table1 = oBDoc.Tables[1];
                        if (الشاهد_الأول.Text == "")
                            table1.Delete();
                    }
                    else
                    {
                        Microsoft.Office.Interop.Word.Table table1 = oBDoc.Tables[1];
                        Microsoft.Office.Interop.Word.Table table2 = oBDoc.Tables[2];
                        if (الشاهد_الأول.Text == "")
                            table2.Delete();
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
                    break;
                case 2:
                    if (addNameIndex != 1)
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
                            }
                        }
                    }
                    break; 
                case 3:
                    if (addNameIndex != 1)
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
                            }
                        }
                    }
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
        private void fillPrintDocx(string deleteDocxFile)
        {
            btnPrint.Enabled = false;
            //MessageBox.Show(localCopy.Text);
            string pdfFile = localCopy.Text.Replace("docx", "pdf");
            oBDoc.SaveAs2(localCopy.Text);
            if (deleteDocxFile == "no")
                oBDoc.ExportAsFixedFormat(pdfFile, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            if (deleteDocxFile == "no")
            {
                System.Diagnostics.Process.Start(pdfFile);
                File.Delete(localCopy.Text);
            }
            else System.Diagnostics.Process.Start(localCopy.Text);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
            this.Close();
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
            if (وجهة_المعاملة.Text.Length == 0) وجهة_المعاملة.SelectedIndex = 0;
        }

        private string writeStrSpecPur() {
            //MessageBox.Show(StrSpecPur);
            string specialValue = SuffPrefReplacements(StrSpecPur);
            for(int x=0;x<25;x++)
                specialValue = SuffPrefReplacements(specialValue);
            return specialValue;    
        }

        private string SuffPrefReplacements(string text)
        {            
            string str = "";
            if (النوع.Text != "ذكر") str = "ة";

            if (text.Contains("  "))
                return text.Replace("  "," ");
            if (text.Contains("tN"))
                return text.Replace("tN", مقدم_الطلب.Text);
            if (text.Contains("tP"))
                return text.Replace("tP", رقم_الهوية.Text);
            if (text.Contains("tS"))
                return text.Replace("tS", مكان_الإصدار.Text);
            if (text.Contains("tX"))
                return text.Replace("tX", str);
            if (text.Contains("tT"))
                return text.Replace("tT", title.Text);
            if (text.Contains("tD"))
                return text.Replace("tD", نوع_الهوية.Text);
            if (text.Contains("t1"))
                return text.Replace("t1", Vitext1.Text);
            if (text.Contains("t2"))
                return text.Replace("t2", Vitext2.Text);
            if (text.Contains("t3"))
                return text.Replace("t3", Vitext3.Text);
            if (text.Contains("t4"))
                return text.Replace("t4", Vitext4.Text);
            if (text.Contains("t5"))
                return text.Replace("t5", Vitext5.Text);
            if (text.Contains("c1"))
                return text.Replace("c1", Vicheck1.Text);
            if (text.Contains("c2"))
                return text.Replace("c2", Vicheck2.Text);
            if (text.Contains("c3"))
                return text.Replace("c3", Vicheck3.Text);
            if (text.Contains("c4"))
                return text.Replace("c4", Vicheck4.Text);
            if (text.Contains("c5"))
                return text.Replace("c5", Vicheck5.Text);

            if (text.Contains("m1"))
                return text.Replace("m1", Vicombo1.Text);
            if (text.Contains("m2"))
                return text.Replace("m2", Vicombo2.Text);

            if (text.Contains("a1"))
                return text.Replace("a1", LibtnAdd1.Text);

            if (text.Contains("n1"))
                return text.Replace("n1", " " + VitxtDate1.Text + " ");
            if (text.Contains("#*#"))
                return text.Replace("#*#", preffix[0, 10]);

            if (text.Contains("#1"))
                return text.Replace("#1", preffix[0, 11]);
            if (text.Contains("#2"))
                return text.Replace("#2", preffix[0, 12]);

            if (text.Contains("$$$"))
                return text.Replace("$$$", preffix[appCaseIndex, 0]);
            if (text.Contains("&&&"))
                return text.Replace("&&&", preffix[appCaseIndex, 1]);
            if (text.Contains("^^^"))
                return text.Replace("^^^", preffix[appCaseIndex, 2]);
            if (text.Contains("###"))
                return text.Replace("###", preffix[appCaseIndex, 4]);
            if (text.Contains("***"))
                return text.Replace("***", preffix[appCaseIndex, 3]);
            if (text.Contains("%&%"))
                return text.Replace("%&%", preffix[appCaseIndex, 12]);
            if (text.Contains("#$#"))
                return text.Replace("#$#", preffix[appCaseIndex, 13]);
            if (text.Contains("&^&"))
                return text.Replace("&^&", preffix[appCaseIndex, 14]);
            if (text.Contains("&^^"))
                return text.Replace("&^^", preffix[appCaseIndex, 15]);
            if (text.Contains("*%*"))
                return text.Replace("*%*", preffix[appCaseIndex, 16]);            
            if (text.Contains("&&*"))
                return text.Replace("&&*", preffix[appCaseIndex, 17]);
            else return text;
        }

       
        private void نوع_الإجراء_SelectedIndexChanged(object sender, EventArgs e)
        {
            resetBoxes();
            flllPanelItemsboxes("ColName", نوع_الإجراء.Text + "-" + نوع_المعاملة.SelectedIndex.ToString());
            fillInfo(PanelItemsboxes, false);
        }
        public void resetBoxes()
        {            
            txtReview.Text = "";
            checkAutoUpdate.Checked = true;
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
            string query = "select * from TableAddContext where " + rowID + "=N'" + cellValue + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            //MessageBox.Show(query);
            Console.WriteLine(query + " - " +dtbl.Rows.Count.ToString());
            if (dtbl.Rows.Count > 0)

                foreach (DataRow dr in dtbl.Rows)
                //if (cellValue == dataGridView1.Rows[index].Cells[rowID].Value.ToString())
                {
                    ColName = dr["ColName"].ToString();
                    ColRight = dr["ColRight"].ToString();
                    StrSpecPur = dr["TextModel"].ToString();
                    
                    foreach (Control Lcontrol in PanelItemsboxes.Controls)
                        try
                        {
                            if (Lcontrol.Name.StartsWith("L"))
                            {
                                Lcontrol.Text = dr[Lcontrol.Name.Replace("L", "")].ToString();
                                if (Lcontrol.Text != "")
                                {
                                    Lcontrol.Visible = true;
                                    foreach (Control Vcontrol in PanelItemsboxes.Controls)
                                    {
                                        //MessageBox.Show(Vcontrol.Name);
                                        if (Vcontrol.Name.Trim() == Lcontrol.Name.Replace("L", "V").Trim())
                                        {
                                            if (Vcontrol.Name.Contains("combo"))
                                            {
                                                //MessageBox.Show(dr[Lcontrol.Name.Replace("L", "") + "Option"].ToString());
                                                //MessageBox.Show(Lcontrol.Name.Replace("L", "") + "Option");
                                                ((ComboBox)Vcontrol).Items.Clear();

                                                string[] items = dr[Lcontrol.Name.Replace("L", "") + "Option"].ToString().Split('_');

                                                for (int x = 0; x < items.Length; x++)
                                                    ((ComboBox)Vcontrol).Items.Add(items[x]);
                                                if (((ComboBox)Vcontrol).Items.Count > 0) ((ComboBox)Vcontrol).SelectedIndex = 0;
                                            }
                                            if (Vcontrol.Name.Contains("check"))
                                            {
                                                //MessageBox.Show(dr[Lcontrol.Name.Replace("L", "") + "Option"].ToString());
                                                //MessageBox.Show(Lcontrol.Name.Replace("L", "") + "Option");
                                                Vcontrol.Text = dr[Lcontrol.Name.Replace("L", "") + "Option"].ToString().Split('_')[1];
                                                checkOptions[checkIndex] = dr[Lcontrol.Name.Replace("L", "") + "Option"].ToString();
                                                checkIndex++;
                                            }
                                            Vcontrol.Visible = true;
                                            string size = dr[Lcontrol.Name.Replace("L", "") + "Length"].ToString();
                                            Vcontrol.Width = Convert.ToInt32(size);
                                            if (Convert.ToInt32(size) >= 700)
                                            {
                                                if (Vcontrol is TextBox) ((TextBox)Vcontrol).Multiline = true;
                                                Vcontrol.Height = 150;
                                            }
                                            
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
                else boxesPreparationsArabic(addNameIndex, نوع_المعاملة.SelectedIndex);

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
    }    
}
