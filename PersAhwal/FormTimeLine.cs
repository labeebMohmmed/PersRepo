using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WIA;
using Xceed.Words.NET;
using System.Diagnostics;
using Xceed.Document.NET;
using System.Data.SqlClient;
using System.Net;
using System.Globalization;
using System.Threading;
using System;
using Aspose.Words.Settings;
using System.Data.SqlTypes;
using DocumentFormat.OpenXml.Office2016.Excel;

namespace PersAhwal
{
    public partial class FormTimeLine : Form
    {
        DeviceInfo AvailableScanner = null;

        string[] PathImage = new string[100];
        int imagecount = 0;
        bool ChangeEmp = false;
        bool passed = false;
        TextBox text;
        string comment = "";
        string PrimariFiles = "";
        string rowCount = "";
        string DataSource = "";
        string JobPosition;
        string PhoneNo = "";
        string TaskDesc = "";
        int firstTime = 0;
        bool dontCheck = false;
        bool dontCheckG2 = false;
        string Employee = "";
        string FilesPathIn, FilesPathOut;
        int taskIDNo = 0;
        int AttendVC = 0;
        string GregorianDate = "";
        int[] Months = new int[12];
        string lastUpdate = "";
        string[] colIDs = new string[100];
        string[] forbidDs = new string[100];
        string[] allList = new string[100];
        string updateAll = "";
        string insertAll = "";
        int FormType = 18;
        DataTable genReportTable;
        bool grdiFill = false;
        int indexGrid2 = 0;
        public FormTimeLine(int attendVC, string dataSource, string jobPosition, string filesPathIn, string filesPathOut, string employee, string gregorianDate)
        {
            InitializeComponent();
            AttendVC = attendVC;
            تاريخ_التكليف.Text = تاريخ_التحديث.Text = GregorianDate = gregorianDate;
            genReportTable = new DataTable();
            //MessageBox.Show(primariFiles);
            DataSource = dataSource;
            definColumn(DataSource);
            allList = getColList("TableTasks");
            JobPosition = jobPosition;
            PrimariFiles = FilesPathIn = filesPathIn;
            FilesPathOut = filesPathOut;
            الموظف.Text = اسم_الموظف.Text = Employee = employee;
            نوع_الحالة.SelectedIndex = 0;
            DocIDGenerator();
            تحديث_الحالة.SelectedIndex = 0;
            getFellowUpState();
            FillDataGridView(dataGridView1,"");
            updateDocIDs(891);
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
            string insertItems = "";
            string insertValues = "";
            string updateValues = "";
            foreach (DataRow row in dtbl.Rows)
            {

                if (row["name"].ToString() != "ID" && row["name"].ToString() != "حالة_الارشفة" && row["name"].ToString() != "sms")
                {
                    allList[i] = row["name"].ToString();
                    //MessageBox.Show(row["name"].ToString());
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
            insertAll = "INSERT INTO " + table + "(" + insertItems + ") values (" + insertValues + ");SELECT @@IDENTITY as lastid";

            return allList;

        }

        private int selectTopTask()
        {


            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) رقم_المهمة from TableTasks order by ID desc", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string rowCnt = "0";
            foreach (DataRow row in dtbl.Rows)
            {
                rowCnt = row["رقم_المهمة"].ToString();
            }

            return Convert.ToInt32(rowCnt);

        }


        private int checkISUniqueNess(string docName, int lstid)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableTasks where رقم_المهمة='" + docName + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            int x = 0;

            if (dtbl.Rows.Count != 1)
            {
                //MessageBox.Show("id no " + docName + " is repeated " + dtbl.Rows.Count.ToString() + " times ");
                foreach (DataRow row in dtbl.Rows)
                {
                    if (x != 0)
                    {
                        int id = Convert.ToInt32(row["ID"].ToString());
                        lstid++;
                        Console.WriteLine("id no " + id.ToString() + " will be chaned to " + lstid.ToString());
                        UpdateDocID(id, lstid.ToString());
                    }
                    x++;
                }
            }
            return lstid;
        } 
        private bool checkISUnique(string docName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableTasks where رقم_المهمة='" + docName + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            int x = 0;

            if (dtbl.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }
         private int updateDocIDs(int lstid)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select ID,رقم_المهمة from TableTasks", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            
            int id = 0;

            foreach (DataRow row in dtbl.Rows)
            {
                lstid = checkISUniqueNess(row["رقم_المهمة"].ToString(), lstid );                
                //if(row["رقم_المهمة"].ToString() == "") 
                //    UpdateDocID(Convert.ToInt32(row["ID"].ToString()),"1");
            }
            return id;
        }

        private string Days(string startingDate, string endDate)
        {
            Console.WriteLine("startingDate = " + startingDate);
            Console.WriteLine("endDate = " + endDate);
            int dayS,monthS,yearS,dayE,monthE,yearE;
            int days = 0;
            if (startingDate.Contains("-")  )
            {
                if (startingDate.Split('-').Length != 3) return "";
                string[] datetime = startingDate.Split('-');
                dayS = Convert.ToInt32(datetime[0]);
                monthS = Convert.ToInt32(datetime[1]);
                yearS = Convert.ToInt32(datetime[2]);
            }
            else return "";

            if (endDate.Contains("-"))
            {
                if (endDate.Split('-').Length != 3) return "";
                string[] datetime = endDate.Split('-');
                dayE = Convert.ToInt32(datetime[0]);
                monthE = Convert.ToInt32(datetime[1]);
                yearE = Convert.ToInt32(datetime[2]);
            }
            else return "";


            

            for (int y = yearS; y <= yearE; y++)
            {
                for (int m = monthS; m <= monthE; m++)
                {
                    for (int d = dayS; d <= daysOfMonth(m, y); d++)
                    {
                        if ((m == monthS && d > dayE) || (m == monthE && d > dayE)) break;
                        days++;
                    }
                }
            }


            return days.ToString();
        }

        private void dateReport(string reportName,string date, DataTable dataTable)
        {
            string route = FilesPathIn + "DailyReportCopy.docx";
            string[] reportItems = new string[10];
            string empEntry = "موظف الادخال";
            int colNo = 6;
            //MessageBox.Show(date);
            //MessageBox.Show(route);
            string ActiveCopy = FilesPathOut + reportName;
            System.IO.File.Copy(route, ActiveCopy);

            
            //foreach (DataRow row in dtbl.Rows)
            //{
            //    rowCnt = row["ID"].ToString();
            //}

            using (DocX document = DocX.Load(ActiveCopy))
            {
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
                
                if (الملف.Text == "0")
                {
                    string strHeader1 = "رقم التقرير اليومي: " + DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day + Environment.NewLine;
                    document.InsertParagraph(strHeader1)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(18d)
                    .Alignment = Alignment.left;
                    
                    string strHeader2 = "تقرير الاتصالات الهاتفية وقضايا المواطنيين ليوم: " + GregorianDate + " م";
                    document.InsertParagraph(strHeader2)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(18d)
                    .Alignment = Alignment.center;
                }else {
                    string strHeader = "الرقم: " + DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day +"                                  "+ GregorianDate + " م" + Environment.NewLine + "تقرير قائمة متابعة الملفات " ;
                    document.InsertParagraph(strHeader)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(18d)
                    .Alignment = Alignment.center;
                    empEntry = "الموظف المكلف";
                    colNo = 7;
                }
                
                var t = document.AddTable(1 + dataTable.Rows.Count, colNo);
                t.Design = TableDesign.TableGrid;
                t.Alignment = Alignment.center;

                t.SetColumnWidth(0, 55);
                t.SetColumnWidth(1, 150);
                t.SetColumnWidth(2, 300);
                t.SetColumnWidth(3, 90);
                t.SetColumnWidth(4, 150);
                t.SetColumnWidth(5, 40);
                if (الملف.Text != "0") 
                    t.SetColumnWidth(6, 70);

                if (الملف.Text == "0")
                {
                    reportItems[5] = "الرقم";
                    reportItems[4] = "اسم المواطن";
                    reportItems[3] = "رقم الهاتف";
                    reportItems[2] = "الموضوع";
                    reportItems[1] = empEntry;
                    reportItems[0] = "مصدر الموضوع";
                }
                else
                {
                    reportItems[6] = "الرقم";
                    reportItems[5] = "اسم المواطن";
                    reportItems[4] = "رقم الهاتف";
                    reportItems[3] = "الموضوع";
                    reportItems[2] = empEntry;
                    reportItems[1] = "مصدر الموضوع";
                    reportItems[0] = "تاريخ التكليف";
                }
                int r = 0; 
                for (; r < colNo; r++)
                {
                    t.Rows[0].Cells[r].Paragraphs[0].Append(reportItems[r]).FontSize(12d).Bold().Alignment = Alignment.center;
                }
                r = 0;


                foreach (DataRow row in dataTable.Rows)
                {
                    if (الملف.Text == "0")
                    {
                        t.Rows[r + 1].Cells[0].Paragraphs[0].Append(row["مصدر_الموضوع"].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[r + 1].Cells[1].Paragraphs[0].Append(row["اسم_الموظف"].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[r + 1].Cells[2].Paragraphs[0].Append(row["نوع_الاستفسار"].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[r + 1].Cells[3].Paragraphs[0].Append(row["رقم_الهاتف"].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[r + 1].Cells[4].Paragraphs[0].Append(row["اسم_المتصل"].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[r + 1].Cells[5].Paragraphs[0].Append((r + 1).ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        

                        
                    }
                    else {
                        t.Rows[r + 1].Cells[0].Paragraphs[0].Append(row["تاريخ_التكليف"].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[r + 1].Cells[1].Paragraphs[0].Append(row["مصدر_الموضوع"].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[r + 1].Cells[2].Paragraphs[0].Append(row["اسم_الموظف"].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[r + 1].Cells[3].Paragraphs[0].Append(row["نوع_الاستفسار"].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[r + 1].Cells[4].Paragraphs[0].Append(row["رقم_الهاتف"].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[r + 1].Cells[5].Paragraphs[0].Append(row["اسم_المتصل"].ToString()).FontSize(12d).Bold().Alignment = Alignment.center;
                        t.Rows[r + 1].Cells[6].Paragraphs[0].Append((r + 1).ToString()).FontSize(12d).Bold().Alignment = Alignment.center;

                    }
                    r++;
                }

                
                var p = document.InsertParagraph(Environment.NewLine);
                p.InsertTableAfterSelf(t);
                string strAttvCo = Environment.NewLine + Environment.NewLine+ "\t\t\t\t\t\t\t\t\t\t" + المشرف_المناوب.Text + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\tع/ القنصل العام بالإنابة";
                var AttvCo = document.InsertParagraph(strAttvCo)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(20d)
                    .Bold()
                    .Alignment = Alignment.center;

                document.Save();
                Process.Start("WINWORD.EXE", ActiveCopy);
            }

        }
        private int loadIDNo()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from TableTasks order by ID desc", sqlCon);
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
        private void DocIDGenerator()
        {
            string year = DateTime.Now.Year.ToString().Replace("20", "");
            string query = "select max(cast (right(رقم_المهمة,LEN(رقم_المهمة) - 15) as int)) as newDocID from TableTasks where رقم_المهمة like N'ق س ج/80/" + year + "/18%'";
            rowCount = getUniqueID(query);            
            رقم_المهمة.Text = "ق س ج/80/" + year + "/18/" + rowCount;       
        }

        private string getUniqueID(string query)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string maxID = "1";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                try
                {
                    maxID = (Convert.ToInt32(dataRow["newDocID"].ToString()) + 1).ToString();
                }
                catch (Exception ex)
                {
                    return maxID;
                }
            }
            return maxID;
        }

        //private string loadRerNo()
        //{
        //    SqlConnection sqlCon = new SqlConnection(DataSource);
        //    if (sqlCon.State == ConnectionState.Closed)
        //        sqlCon.Open();
        //    SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT رقم_المهمة from TableTasks where رقم_المهمة like N'ق س ج/80/" + DateTime.Now.Year.ToString().Replace("20", "") + "/%'", sqlCon);
        //    sqlDa.SelectCommand.CommandType = CommandType.Text;
        //    DataTable dtbl = new DataTable();
        //    sqlDa.Fill(dtbl);
        //    sqlCon.Close();
        //    string rowCnt = "";
        //    int maxID = 0;
        //    foreach (DataRow row in dtbl.Rows)
        //    {
        //        if(Convert.ToInt32(row["رقم_المهمة"].ToString().Split('/')[4]) > maxID)
        //            maxID = Convert.ToInt32(row["رقم_المهمة"].ToString().Split('/')[4]);
        //    }
        //    return (maxID + 1).ToString();
        //}


        private string getPhone(string text)
        {
            string phoneNo = "";
            using (SqlConnection saConn = new SqlConnection(DataSource))
            {
                saConn.Open();
                string query = "select PhoneNo from TableUser where EmployeeName = @EmployeeName";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.AddWithValue("@EmployeeName", text);
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                foreach (DataRow dataRow in table.Rows)
                {
                    if (!String.IsNullOrEmpty(dataRow["PhoneNo"].ToString()))
                        phoneNo = dataRow["PhoneNo"].ToString();
                    if (phoneNo.Length == 9) phoneNo = "966" + phoneNo;
                }
                saConn.Close();
            }
            return phoneNo;
        }

        private void getNames()
        {

            using (SqlConnection saConn = new SqlConnection(DataSource))
            {
                saConn.Open();

                string query = "select EmployeeName,Purpose from TableUser";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                foreach (DataRow dataRow in table.Rows)
                {
                    if (!String.IsNullOrEmpty(dataRow["EmployeeName"].ToString()) && dataRow["Purpose"].ToString().Contains("شؤون رعايا"))
                        المشرف_المناوب.Items.Add (dataRow["EmployeeName"].ToString());
                }
                saConn.Close();
            }
        }

        private void btnAuth_Click(object sender, EventArgs e)
        {
            panel1.BringToFront();
            panel1.Size = new System.Drawing.Size(748, 630);//729, 615//
            panel1.Location = new System.Drawing.Point(10, 12);//10, 3
            loadPic.Enabled = btnSave.Visible = getScan.Enabled = false;
            panel1.Visible = true;
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
                    
                    //759, 173
                    // location
                    //panel1.Visible = false;
                    imagecount++;
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
            rescan.Visible = reLoadPic.Visible = true;
            loadPic.BackColor = getScan.BackColor = System.Drawing.Color.LightGreen;
            loadPic.Text = getScan.Text = "اضافة مستند آخر (" + imagecount.ToString() + ")";
            getScan.Location = new System.Drawing.Point(1070, 579);
            loadPic.Location = new System.Drawing.Point(1070, 625);
            getScan.Size = loadPic.Size = new System.Drawing.Size(168, 42);
            loadPic.Enabled = btnSave.Visible = getScan.Enabled = true;
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


        private void SendSms(string phone, string message)
        {

            string apiText = "https://www.hisms.ws/api.php?send_sms&username=966543321629&password=CZssA58@9QdF&numbers=***&sender=CON-SUDAN&message=&&&";
            apiText = apiText.Replace("***", phone);
            apiText = apiText.Replace("&&&", message);
            try
            {
                if (phone.Length != 12)
                {
                    MessageBox.Show("تعذر الارسال نسبة لعدم رقم هاتف صالح");
                    return;
                }
                WebClient client = new WebClient();
                Stream stream = client.OpenRead(apiText);
                StreamReader streamsread = new StreamReader(stream);
                string result = streamsread.ReadToEnd();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Console.WriteLine("تم ارسال رسالة");
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
            catch (COMException e)
            {
                MessageBox.Show(e.Message);
            }
        }
        private void FormTimeLine_Load(object sender, EventArgs e)
        {
            loadScanner();
            autoCompleteTextBox(رقم_الهاتف, DataSource, "رقم_الهاتف", "TableTasks");
            autoCompleteTextBox(نوع_الاستفسار, DataSource, "نوع_الاستفسار", "TableTasks");
            
            fillDestiCombo(الملف, DataSource, "الملف", "TableTasks");
            if(الملف.Items.Count > 0) الملف.SelectedIndex = الملف.Items.Count - 1;
            if (JobPosition.Contains("قنصل"))
            {
                اسم_المشرف_المناوب.Visible = false;
                btnChangeEmp.Visible = true;
                getNames();
            }
            else
            {
                fileComboTask(DataSource, اسم_الموظف.Text);
                اسم_المشرف_المناوب.Visible = true;
                btnChangeEmp.Visible = false;

                fileVCComboBox(المشرف_المناوب, DataSource, "ArabicAttendVC", "TableListCombo");
                
                المشرف_المناوب.Text = اسم_الموظف.Text;                                
                المشرف_المناوب.SelectedIndex = AttendVC;
                
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

        private void fileVCComboBox(ComboBox combbox, string source, string comlumnName, string tableName)
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
        
        private void fillDestiCombo(ComboBox combbox, string source, string comlumnName, string tableName)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select DISTINCT " + comlumnName + " from " + tableName + " where "+ comlumnName+" is not null";
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

        private string loadDocxFile()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return dlg.FileName;
            }
            return "";
        }
        private void loadPic_Click(object sender, EventArgs e)
        {
            panel1.BringToFront(); 
            panel1.Size = new System.Drawing.Size(748, 630);//729, 615//
            panel1.Location = new System.Drawing.Point(10, 12);//10, 3                                                    
            panel1.Visible = true;
            
            string fileName = loadDocxFile();
            if (fileName != "")
            {
                pictureBox1.ImageLocation = PathImage[imagecount] = fileName;
                imagecount++;
                getScan.BackColor = System.Drawing.Color.LightGreen;
                getScan.Text = "اضافة مستند آخر (" + (imagecount + 1).ToString() + ")";

            }
            getScan.Location = new System.Drawing.Point(1070, 579);
            loadPic.Location = new System.Drawing.Point(1070, 625);
            getScan.Size = loadPic.Size = new System.Drawing.Size(168, 42);
            rescan.Visible = reLoadPic.Visible = true;
        }

        private void smsReq()
        {
            if (JobPosition.Contains("قنصل"))
            {
                var selectedOption = MessageBox.Show("", "تنبيه برسالة", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                {
                    PhoneNo = getPhone(اسم_الموظف.Text);
                    if (PhoneNo != "")
                        SendSms(PhoneNo, sms.Text);
                }
            }
        }
        private bool reReq()
        {
            bool ok = true;
            if (رقم_المهمة.Text == "")
            {
                DocIDGenerator();
                ok = true;
            }

            if (رقم_الهاتف.Text.Length != 12) { MessageBox.Show("رقم الهاتف يجب أن يتكون من 12 رقم"); ok = false; }
            if (نوع_الاستفسار.Text.Length == 0 || نوع_الاستفسار.Text == "أضف نوع") { MessageBox.Show("يرجى اختيار أو إضافة نوع الاستفسار"); ok = false; }
            if (وصف_المهمة.Text.Length == 0 || وصف_المهمة.Text == "أضف موضوع ") { MessageBox.Show("يرجى اختيار أو إضافة الموضوع "); ok = false; }
            if (اسم_المتصل.Text.Length == 0 ) { MessageBox.Show("يرجى إضافة اسم المتصل"); ok = false; }
            if (مدينة_الاقامة.Text.Length == 0 ) { MessageBox.Show("يرجى اختيار أو إضافة مدينة الإقامة"); ok = false; }

            if (!JobPosition.Contains("قنصل") && المشرف_المناوب.Text == "إختر من القائمة")
            {
                MessageBox.Show("يرجى إختيار المشرف");
                ok = false;
            }
            return ok;  
        }
        private void reLoadPic_Click(object sender, EventArgs e)
        {
            string fileName = loadDocxFile();
            if (fileName != "")
            {
                pictureBox1.ImageLocation = PathImage[imagecount - 1] = fileName;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!reReq()) return;
            string docid = save2DataBase(taskIDNo).ToString();            
            addArch(PathImage, docid);
            riskyCase();
            smsReq();
            this.Close();
        }

        private void FillDataGridView(DataGridView dataGridView, string addedQuery)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableTasks " + addedQuery, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            //sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_المهمة", txtTaskID.Text.Trim());
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView.DataSource = dtbl;
            dataGridView.Sort(dataGridView.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            dataGridView.Columns[0].Visible = false;
            dataGridView.Columns[1].Width = 150;
            dataGridView.Columns[2].Width = 200;
            dataGridView.Columns[3].Width = 50;
            dataGridView.Columns[4].Width = 450;
            
            sqlCon.Close();
            
            Console.WriteLine("select * from TableTasks " + addedQuery);
        }

        private void insertDoc(string id, string date, string employee, string dataSource, string extn1, string DocName1, string messNo, string docType, byte[] buffer1)
        {
            string query = "INSERT INTO TableGeneralArch (Data1,Extension1,نوع_المستند,رقم_معاملة_القسم,المستند,الموظف,التاريخ,رقم_المرجع,docTable) values (@Data1,@Extension1,@نوع_المستند,@رقم_معاملة_القسم,@المستند,@الموظف,@التاريخ,@رقم_المرجع,@docTable)";
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
            sqlCmd.Parameters.Add("@docTable", SqlDbType.NVarChar).Value = "TableTasks";            
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }



        private void addArch(string[] location, string docid)
        {
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
                        insertDoc(docid, تاريخ_التحديث.Text, الموظف.Text, DataSource, extn1, DocName1, رقم_المهمة.Text, "data", buffer1);
                    }
                }
            }
        }

                private void OpenFile(int id)
        {
            //MessageBox.Show(id.ToString());
            SqlConnection Con = new SqlConnection(DataSource);

            string query = "select ارشفة_المستندات,Data1,Extension1 from TableTaskDocs where ID=@id";

            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@Id", SqlDbType.Int).Value = id;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                var name = reader["ارشفة_المستندات"].ToString();
                var Data = (byte[])reader["Data1"];
                var ext = reader["Extension1"].ToString();
                string NewFileName = PrimariFiles + DateTime.Now.ToString("mmss")+ name ;
                while (Directory.Exists(NewFileName))
                    NewFileName = PrimariFiles + DateTime.Now.ToString("mmss")+ name ;
                File.WriteAllBytes(NewFileName, Data);
                System.Diagnostics.Process.Start(NewFileName);

            }
            Con.Close();

        }
        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName)
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

        private void fileComboBox1(ComboBox combbox, string source, string comlumnName, string tableName)
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
            //if (combbox.Items.Count > 0) combbox.SelectedIndex = 0;
        }

        private void fileComboTask( string source, string text)
        {            

            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select وصف_المهمة from TableTasks where اسم_الموظف=@اسم_الموظف";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.AddWithValue("@اسم_الموظف", text);

                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                وصف_المهمة.Items.Clear();
                وصف_المهمة.Text = "وصف الإجراء";
                foreach (DataRow dataRow in table.Rows)
                {
                    if (!String.IsNullOrEmpty(dataRow["وصف_المهمة"].ToString()) && dataRow["وصف_المهمة"].ToString() != "وصف الإجراء")
                    {
                       // MessageBox.Show(dataRow["وصف_المهمة"].ToString());
                        وصف_المهمة.Items.Add(dataRow["وصف_المهمة"].ToString());
                    }
                }
                saConn.Close();
            }
            //if (txtTask.Items.Count > 0) txtTask.SelectedIndex = 0;
        }


        private void labEmpName_Click(object sender, EventArgs e)
        {

        }

        private void labEmpName_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void txtTaskID_TextChanged(object sender, EventArgs e)
        {
            
            
        }

        private void txtTask_SelectedIndexChanged(object sender, EventArgs e)
        {
            //رقم_المهمة,وصف_المهمة,تعليق,تحديث_الحالة,الملف,تاريخ_التكليف,المشرف,الاسم

            if (dontCheck) return;
            //MessageBox.Show(وصف_المهمة.Text);
            getData(وصف_المهمة.Text.Trim(), "وصف_المهمة");

        }

        private void getData(string text, string colName)
        {
            //MessageBox.Show(dataGridView1.Rows.Count.ToString());
            if (dataGridView1.Rows.Count > 1)
            {
                
                for (int x = 0; x < dataGridView1.Rows.Count - 1; x++)
                {
                    string refText = dataGridView1.Rows[x].Cells[colName].Value.ToString();
                    
                    if (refText == text)
                    {
                        //MessageBox.Show(refText);
                        taskIDNo = Convert.ToInt32(dataGridView1.Rows[x].Cells[0].Value.ToString());
                        
                        foreach (Control control in MainPanel.Controls)
                        {
                            if ((control is TextBox || control is ComboBox || control is CheckBox) && !control.Name.Contains("Off"))
                            {
                                control.Text = dataGridView1.Rows[x].Cells[control.Name].Value.ToString();
                            }
                        }
                        if (dataGridView1.Rows[x].Cells["مستندات"].Value.ToString() == "yes")
                        {
                            FillDatafromGenArch(taskIDNo.ToString(), "TableTasks", false, "");
                        }
                        if (JobPosition.Contains("قنصل"))
                        {
                            if (تحديث_الحالة.SelectedIndex == 0)
                            {
                                sms.Text = "عزيزي " + اسم_الموظف.Text + " تم تكليفكم لإداء التكليف بالرقم " + رقم_المهمة.Text + " نثق في أداءكم المتميز دوما";
                            }
                            else
                            {
                                sms.Text = "عزيزي " + اسم_الموظف.Text + " يرجى التكرم بمتابعة العمل على التكليف بالرقم (" + رقم_المهمة.Text + ") " + تحديث_الحالة.Text + " نثق في أداءكم المتميز دوما";
                            }
                        }
                    }
                    dontCheck = true;
                }
            }

            //for (int x = 0; x < dataGridView1.Rows.Count - 1; x++)
            //{
            //    string refText = dataGridView1.Rows[x].Cells[cell].Value.ToString();
            //    if (refText == text)
            //    {
            //        btnSave.Text = "تعديل";
            //        btnSave.Enabled = true;

            //        taskIDNo = Convert.ToInt32(dataGridView1.Rows[x].Cells[0].Value.ToString());
            //        .Text = dataGridView1.Rows[x].Cells[8].Value.ToString();
            //        comment = dataGridView1.Rows[x].Cells[3].Value.ToString();
            //        string[] comments = dataGridView1.Rows[x].Cells[3].Value.ToString().Split('*');
            //        for (int comID = 0; comID < comments.Length; comID++)
            //        {

            //            المستندات_المتعلقة.Items.Add(comments[comID]);
            //        }
            //        اسم_المتصل.Text = dataGridView1.Rows[x].Cells[1].Value.ToString();
            //        تاريخ_التكليف.Text = dataGridView1.Rows[x].Cells[6].Value.ToString();

            //        if (Convert.ToInt32(dataGridView1.Rows[x].Cells[4].Value.ToString()) == 0)
            //            تحديث_الحالة.SelectedIndex = 1;
            //        else if (Convert.ToInt32(dataGridView1.Rows[x].Cells[4].Value.ToString()) == 1)
            //            تحديث_الحالة.SelectedIndex = 2;
            //        else
            //            تحديث_الحالة.SelectedIndex = Convert.ToInt32(dataGridView1.Rows[x].Cells[4].Value.ToString());
            //        lastUpdate = dataGridView1.Rows[x].Cells[9].Value.ToString();
            //        if (string.IsNullOrEmpty(lastUpdate))
            //            missionDate.Text = "اخر تحديث منذ " + Days(تاريخ_التكليف.Text, GregorianDate) + " يوم";
            //        else
            //            missionDate.Text = "اخر تحديث منذ " + Days(تاريخ_التكليف.Text, lastUpdate) + " يوم";
            //        رقم_الهاتف.Text = dataGridView1.Rows[x].Cells[11].Value.ToString();
            //        المشرف.Text = dataGridView1.Rows[x].Cells[10].Value.ToString();
            //        //نوع_الاستفسار.Text = dataGridView1.Rows[x].Cells[12].Value.ToString();
            //        النوع.Text = dataGridView1.Rows[x].Cells[13].Value.ToString();
            //        //FillTaskDocs(رقم_المهمة.Text);
            //        dataGridView1.Visible = false;
            //    }
            //}
            //if (JobPosition.Contains("قنصل"))
            //{
            //    if (تحديث_الحالة.SelectedIndex == 0)
            //    {
            //        sms.Text = "عزيزي " + المشرف.Text + " تم تكليفكم لإداء التكليف بالرقم " + رقم_المهمة.Text + " نثق في أداءكم المتميز دوما";
            //    }
            //    else
            //    {
            //        sms.Text = "عزيزي " + المشرف.Text + " يرجى التكرم بمتابعة العمل على التكليف بالرقم (" + رقم_المهمة.Text + ") " + تحديث_الحالة.Text + " نثق في أداءكم المتميز دوما";
            //    }
            //}
        }

        private string confirmPhone(string text)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableTasks where رقم_الهاتف='" + text + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            int x = 0;
            string name = "";
            foreach (DataRow row in dtbl.Rows)
            {
                name = row["اسم_المتصل"].ToString();                    
            }
            return name;
        }

        private void ReporInfo(string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableTaskDocs() values ()", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            
            
            sqlCmd.ExecuteNonQuery();

            sqlCon.Close();
        }

        private void ReportDocx(string dataSource, string filePath)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableTaskDocs(Data1,Extension1,ارشفة_المستندات,التاريخ_الميلادي,الزمن,رقم_المهمة) values (@Data1,@Extension1,@ارشفة_المستندات,@التاريخ_الميلادي,@الزمن,@رقم_المهمة)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_المهمة", رقم_المهمة.Text);
            sqlCmd.Parameters.AddWithValue("@الزمن", DateTime.Now.ToString("mm:hh"));
            sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", GregorianDate);

            if (filePath != "")
            {
                using (Stream stream = File.OpenRead(filePath))
                {
                    byte[] buffer1 = new byte[stream.Length];
                    stream.Read(buffer1, 0, buffer1.Length);
                    var fileinfo1 = new FileInfo(filePath);
                    string extn1 = fileinfo1.Extension;
                    string DocName = fileinfo1.Name;
                    sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                    sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
                    sqlCmd.Parameters.Add("@ارشفة_المستندات", SqlDbType.NVarChar).Value = DocName;
                }
            }
            sqlCmd.ExecuteNonQuery();

            sqlCon.Close();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (dataGridView2.Rows.Count > 1) {
            //    OpenFile(Convert.ToInt32(dataGridView2.CurrentRow.Cells[0].Value.ToString()));
            //    textTaskDesc.Text = dataGridView2.CurrentRow.Cells[3].Value.ToString();
            //}
        }

        
        private void button3_Click(object sender, EventArgs e)
        {
            //FillTaskDocs(رقم_المهمة.Text);            
        }

        private void mangerArch_CheckedChanged(object sender, EventArgs e)
        {

        }

        //private void creatPic(string id, int horizon, int vert, string text, string icon) {
        //    //hor = 191
        //    //ver = 142
        //    panel1.Visible = false;
        //    panelPics.Visible = true;
        //    PictureBox pictureBox = new PictureBox();
        //    switch (icon) {
        //        case ".docx":
        //            pictureBox.Image = global::PersAhwal.Properties.Resources.doc;
        //            break;
        //        case ".doc":
        //            pictureBox.Image = global::PersAhwal.Properties.Resources.doc;
        //            break;
        //        case ".pdf":
        //            pictureBox.Image = global::PersAhwal.Properties.Resources.pdf;
        //            break;
        //        case ".png":
        //            pictureBox.Image = global::PersAhwal.Properties.Resources.png;
        //            break;
        //        default:
        //            pictureBox.Image = global::PersAhwal.Properties.Resources.file;
        //            break;
        //    }
        //    pictureBox.Location = new System.Drawing.Point(10 + (300*horizon), 18+(240*vert));
        //    pictureBox.Name = id;
        //    pictureBox.Size = new System.Drawing.Size(88, 67);
        //    pictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
        //    pictureBox.TabIndex = 441;
        //    pictureBox.TabStop = false;
        //    pictureBox.Click += new System.EventHandler(this.pictClick);

        //    Label label = new Label();
        //    label.AutoSize = true;
        //    label.Font = new System.Drawing.Font("Arabic Typesetting", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        //    label.ForeColor = System.Drawing.Color.Black;
        //    label.Location = new System.Drawing.Point(10 + (300 * horizon), 91 + (240 * vert));
        //    label.Name = id;
        //    label.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
        //    label.Size = new System.Drawing.Size(300, 150);
        //    label.TabIndex = 646;
        //    label.Text = text;
        //    label.Click += new System.EventHandler(this.labelClick);

        //    panelPics.Controls.Add(pictureBox);
        //    panelPics.Controls.Add(label);
        //}
        private void pictClick(object sender, EventArgs e)
        {
            PictureBox picbox = (PictureBox)sender;
            OpenFile(Convert.ToInt32(picbox.Name));

        }

        private void labelClick(object sender, EventArgs e)
        {
            Label label= (Label)sender;
            OpenFile(Convert.ToInt32(label.Name));
        }
        

        
        private void comboTaskDuration_SelectedIndexChanged(object sender, EventArgs e)
        {
            وصف_المهمة.Items.Clear();
        }

        private void UpdateState(int id, string updateSay)
        {
            // MessageBox.Show(id.ToString() +"-"+ table + "-"+column + "-"+text);
            //string qurey = "update "+table+" set "+ column + "=@"+ column + " where ID=@id";
            string qurey = "update TableTasks set تاريخ_التحديث=@تاريخ_التحديث,الملف=@الملف where ID=@id";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@الملف", "0");
            sqlCmd.Parameters.AddWithValue("@تاريخ_التحديث", updateSay);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        private void UpdateDocID(int id, string docid)
        {
            // MessageBox.Show(id.ToString() +"-"+ table + "-"+column + "-"+text);
            //string qurey = "update "+table+" set "+ column + "=@"+ column + " where ID=@id";
            string qurey = "update TableTasks set رقم_المهمة=@رقم_المهمة where ID=@id";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@رقم_المهمة", docid);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //CultureInfo arSA = new CultureInfo("ar-SA");
            //arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            //Thread.CurrentThread.CurrentCulture = arSA;
            //new System.Globalization.GregorianCalendar();
            //GregorianDate = DateTime.Now.ToString("dd-MM-yyyy");

            if (passed || FellewUpState.CheckState == CheckState.Unchecked) 
                return;
            passed = true;
            //for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            //{
            //    string lastUpdate = dataGridView1.Rows[i].Cells[9].Value.ToString();
            //    if (lastUpdate == "")
            //        lastUpdate = GregorianDate;
                
            //    string id = dataGridView1.Rows[i].Cells[0].Value.ToString();
            //    string name = dataGridView1.Rows[i].Cells[1].Value.ToString();
            //    //getPhone(name);
            //    string taskId = dataGridView1.Rows[i].Cells[2].Value.ToString();
            //    string taskComment = dataGridView1.Rows[i].Cells[3].Value.ToString();
            //    string days = Days(lastUpdate, GregorianDate);
            //    //MessageBox.Show(lastUpdate+" - "+GregorianDate+" - "+days);
            //    string text = "عزيزي " + name + " يرجى موافاتنا بمستجدات التكليف بالرقم " + taskId; 
            //    if (Convert.ToInt32(days) > 2 && lastUpdate != GregorianDate && taskComment != "تعليق إضافي" && taskComment != "")
            //    {
            //        //SendSms("966" + PhoneNo, text);
            //        UpdateState(Convert.ToInt32(id), GregorianDate);
            //    }
            //}
            Console.WriteLine(passed);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
        }

        private void textTaskDesc_TextChanged(object sender, EventArgs e)
        {
            
        }
        






        private void ColorFulGrid1()
        {
            int count = -1;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                switch (dataGridView1.Rows[i].Cells["تحديث_الحالة"].Value.ToString())
                {
                    case "جديدة":

                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                        break;
                    case "غير محدثة":
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Orange;
                        break;
                    case "محدثة":
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Cyan;
                        break;

                    case "منتهية":
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                        break;

                    case "مؤجلة":
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                        break;

                    case "ملغية":
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Violet;
                        break;
                }
            }
        }

        
       private void TaskStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            btnSave.BackColor = System.Drawing.Color.LightGreen;
            switch (تحديث_الحالة.SelectedIndex) {
                case 0:
                    تعليق_جديد_Off.Text = التعليقات_السابقة_Off.Text = نوع_الاستفسار.Text = وصف_المهمة.Text = "";
                    المشرف_المناوب.Text = "إختر من القائمة";
                    الملف.Text = "0";
                    نوع_الحالة.SelectedIndex = 0;
                    DocIDGenerator();
                    break;  
                case 1:
                    btnSave.BackColor = System.Drawing.Color.Orange; 
                    break;                    
                case 2:
                    btnSave.BackColor = System.Drawing.Color.Cyan; 
                    break;

                case 3:
                    btnSave.BackColor = System.Drawing.Color.Red; 
                    break;

                case 4:
                    btnSave.BackColor = System.Drawing.Color.Yellow; 
                    break;

                case 5:
                    btnSave.BackColor = System.Drawing.Color.Violet; 
                    break;
            }

        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            FillDataGridView(dataGridView1,"");
            if (dataGridView1.Visible == true)
            {
                MainPanel.Visible = true;
                rescan.Visible = getScan.Visible = loadPic.Visible = reLoadPic.Visible = dataGridView1.Visible = false;
                button5.Text = "قائمة المهام";
            }
            else
            {
                ColorFulGrid1();
                rescan.Visible = getScan.Visible = loadPic.Visible = reLoadPic.Visible = MainPanel.Visible = false;
                dataGridView1.Visible = true;
                button5.Text = "إخفاء قائمة المهام";
            }
            
        }

        void FillDatafromGenArch(string id, string table, bool show, string nameDoc)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableGeneralArch where  رقم_المرجع='" + id + "' and docTable = '" + table + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow reader in dtbl.Rows)
            {
                var name = reader["المستند"].ToString();
                var Data = (byte[])reader["Data1"];
                var ext = reader["Extension1"].ToString();
                if (show && nameDoc == name.Replace(ext, ""))
                {
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }
                else if (!show && nameDoc == "")
                {
                    المستندات_المتعلقة.Items.Add(name.Replace(ext, ""));
                }
            }
            sqlCon.Close();
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                taskIDNo = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                التعليقات_السابقة_Off.Text = dataGridView1.CurrentRow.Cells["تعليق"].Value.ToString();
                foreach (Control control in MainPanel.Controls)
                {
                    if ((control is TextBox || control is ComboBox || control is CheckBox) && !control.Name.Contains("Off"))
                    {                        
                            control.Text = dataGridView1.CurrentRow.Cells[control.Name].Value.ToString();
                    }

                }
                if (dataGridView1.CurrentRow.Cells["مستندات"].Value.ToString() == "yes")
                {
                    FillDatafromGenArch(taskIDNo.ToString(), "TableTasks", false,""); 
                }
                dataGridView1.Visible = false;
                label13.Visible = missionDate.Visible = اسم_الموظف.Visible = MainPanel.Visible = true;
            }

            //if (dataGridView1.Rows.Count > 1)
            //{
            //    btnSave.Text = "تعديل";
            //    btnSave.Enabled = true;

            //    taskIDNo = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            //    Console.WriteLine("task id " + dataGridView1.CurrentRow.Cells[0].Value.ToString());
            //    رقم_المهمة.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                
            //    comment = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            //    string[] comments = dataGridView1.CurrentRow.Cells[3].Value.ToString().Split('*');
                
            //    //Panelapp_Paint(string text)
            //    for (int comID = 0; comID < comments.Length; comID++)
            //    {
            //        المستندات_المتعلقة.Items.Add(comments[comID]);
            //        //MessageBox.Show(comments[comID]);
            //    }
            //    اسم_المتصل.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            //    تاريخ_التكليف.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                
            //    if (Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString()) == 0)
            //        تحديث_الحالة.SelectedIndex = 1;
            //    else if (Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString()) == 1)
            //        تحديث_الحالة.SelectedIndex = 2;
            //    else
            //        تحديث_الحالة.SelectedIndex = Convert.ToInt32(dataGridView1.CurrentRow.Cells[4].Value.ToString());
            //    lastUpdate = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            //    if (string.IsNullOrEmpty(lastUpdate))
            //        missionDate.Text = "اخر تحديث منذ " + Days(تاريخ_التكليف.Text, GregorianDate) + " يوم";
            //    else
            //        missionDate.Text = "اخر تحديث منذ " + Days(تاريخ_التكليف.Text, lastUpdate) + " يوم";
            //    dontCheck = true;
            //    رقم_الهاتف.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();

            //    المشرف.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            //    //نوع_الاستفسار//.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
            //    النوع.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();


            //    //FillTaskDocs(رقم_المهمة.Text);
            //    dataGridView1.Visible = false;
            //    MainPanel.Visible = true;
            //    وصف_المهمة.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            //    dontCheck = true;
            //}

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

            return Months[month - 1];
        }

        private int save2DataBase(int taskID)
        {
            commentInfo();
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(updateAll, sqlConnection);
            if (taskID == 0) sqlCommand = new SqlCommand(insertAll, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            
                sqlCommand.Parameters.AddWithValue("@id", taskID);
            
            for (int i = 0; i < allList.Length; i++)
            {
                foreach (Control control in MainPanel.Controls)
                {
                    if (control.Name == allList[i])
                    {
                        sqlCommand.Parameters.AddWithValue("@" + allList[i], control.Text);
                        break;
                    }
                }
            }
            var reader = sqlCommand.ExecuteReader();
            if (reader.Read())
            {
                if (taskID == 0) 
                    taskID = Convert.ToInt32(reader["lastid"].ToString());
            }
            return taskID;
        }

        //private string ReportEntry(string dataSource, int id)
        //{
            
        //    SqlConnection sqlCon = new SqlConnection(dataSource);
        //    if (sqlCon.State == ConnectionState.Closed)
        //        sqlCon.Open();
        //    SqlCommand sqlCmd = new SqlCommand("TaskAddorEdit", sqlCon);
        //    sqlCmd.CommandType = CommandType.StoredProcedure;
        //    if (id == 0)
        //    {
        //        sqlCmd.Parameters.AddWithValue("@ID", 1);
        //        sqlCmd.Parameters.AddWithValue("@mode", "Add");
        //        sqlCmd.Parameters.AddWithValue("@تاريخ_التكليف", تاريخ_التحديث.Text);
        //        if (رقم_المهمة.Text != "")
        //            while (checkISUnique(رقم_المهمة.Text)) { رقم_المهمة.Text = (Convert.ToInt32(رقم_المهمة.Text) + 1).ToString(); }
        //        else if (رقم_المهمة.Text == "") رقم_المهمة.Text = (selectTopTask() + 1).ToString();
        //    }
        //    else {
        //        sqlCmd.Parameters.AddWithValue("@ID", id);
        //        sqlCmd.Parameters.AddWithValue("@mode", "Edit");
        //        //sqlCmd.Parameters.AddWithValue("@تاريخ_التكليف", تاريخ_التكليف.Text);
        //    }
        //    sqlCmd.Parameters.AddWithValue("@تاريخ_التحديث", تاريخ_التحديث.Text);
        //    sqlCmd.Parameters.AddWithValue("@رقم_المهمة", رقم_المهمة.Text);
        //    sqlCmd.Parameters.AddWithValue("@اسم_الموظف", اسم_الموظف.Text);
        //    sqlCmd.Parameters.AddWithValue("@تحديث_الحالة", تحديث_الحالة.Text);
        //    if (وصف_المهمة.Text == "وصف الإجراء") وصف_المهمة.Text = "";
        //    sqlCmd.Parameters.AddWithValue("@وصف_المهمة", وصف_المهمة.Text);
        //    commentInfo();
        //    sqlCmd.Parameters.AddWithValue("@تعليق", تعليق.Text);
        //    sqlCmd.Parameters.AddWithValue("@اسم_المتصل", اسم_المتصل.Text);
        //    sqlCmd.Parameters.AddWithValue("@رقم_الهاتف", رقم_الهاتف.Text);
        //    sqlCmd.Parameters.AddWithValue("@النوع", النوع.Text);
        //    sqlCmd.Parameters.AddWithValue("@المشرف", المشرف.Text);


        //    sqlCmd.ExecuteNonQuery();
        //    sqlCon.Close();
        //    return id.ToString();
        //}

        private void definColumn(string dataSource)
        {
            DataSource = dataSource;
            for (int index = 0; index < 100; index++)
                forbidDs[index] = "";
            foreach (Control control in MainPanel.Controls)
            {
                if ((control is TextBox || control is ComboBox || control is CheckBox) && !control.Name.Contains("Off"))
                {
                    if (!checkColumnName(control.Name, DataSource))
                    {
                        CreateColumn(control.Name, DataSource);
                    }
                }
            }            
        }

        private void CreateColumn(string Columnname, string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("alter table TableTasks add " + Columnname.Replace(" ", "_") + " nvarchar(150)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private bool checkColumnName(string colNo, string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableTasks", sqlCon);
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

        private void commentInfo()
        {
            if ((تعليق_جديد_Off.Text == "" || تعليق_جديد_Off.Text == "تعليق") && التعليقات_السابقة_Off.Text == "")
                تعليق.Text = "";

            if ((تعليق_جديد_Off.Text == "" || تعليق_جديد_Off.Text == "تعليق") && التعليقات_السابقة_Off.Text != "")
                تعليق.Text = التعليقات_السابقة_Off.Text;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text == "")
                تعليق.Text = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + اسم_الموظف.Text + " - "+ تاريخ_التحديث .Text+ Environment.NewLine + "--------------" + Environment.NewLine;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text != "")
                تعليق.Text = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + اسم_الموظف.Text + " - " + تاريخ_التحديث.Text + Environment.NewLine + "--------------" + Environment.NewLine + "*" + التعليقات_السابقة_Off.Text.Trim();
            
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (firstTime < 60)
            {
                ColorFulGrid1();
                firstTime++;
            }
            if (imagecount > 0)
            {
                rescan.Visible =reLoadPic.Visible =true;
                loadPic.Size = getScan.Size = new System.Drawing.Size(155, 42);
                loadPic.Location = new System.Drawing.Point(1070, 625);
                getScan.Location = new System.Drawing.Point(1070, 579);
            }
            else {
                rescan.Visible = reLoadPic.Visible = false;
                loadPic.Size = getScan.Size = new System.Drawing.Size(328, 42);
                loadPic.Location = new System.Drawing.Point(898, 625);
                getScan.Location = new System.Drawing.Point(898, 579);

            }
        }

        //private void Panelapp_Paint(string text)
        //{
        //    int newLine = 0;
        //    if (text.Contains('\n'))
        //        newLine = text.Split('\n').Length - 1;
        //    else newLine = 0;
            
        //    int hieght =  (30 * text.Length / 80) + newLine*30;
        //    if(hieght < 30)
        //        hieght = 30;
            

        //    TextBox textBoxDocNo = new TextBox();
        //    textBoxDocNo.Dock = System.Windows.Forms.DockStyle.Top;
        //    textBoxDocNo.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        //    textBoxDocNo.Location = new System.Drawing.Point(0, 0);
        //    textBoxDocNo.Multiline = true;
        //    textBoxDocNo.Name = "textTaskDesc";
        //    textBoxDocNo.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
        //    textBoxDocNo.Size = new System.Drawing.Size(575, hieght);
        //    textBoxDocNo.TabIndex = 663;
        //    textBoxDocNo.Text = text;
        //    textBoxDocNo.Click += new System.EventHandler(this.txtComment_Click);
        //    //panel2.Controls.Add(textBoxDocNo);

        //}

        

        private void txtNewComment_TextChanged(object sender, EventArgs e)
        {
            if (JobPosition.Contains("قنصل"))
            {
                if (تحديث_الحالة.SelectedIndex == 0)
                {
                    sms.Text = "عزيزي " + المشرف_المناوب.Text + " تم تكليفكم لإداء المهمة رقم " + رقم_المهمة.Text + " نثق في أداءكم المتميز دوما";
                }
                else
                {
                    sms.Text = "عزيزي " + المشرف_المناوب.Text + " يرجى التكرم بمتابعة العمل على المهمة بالرقم (" + رقم_المهمة.Text + ") " + تحديث_الحالة.Text;

                }
            }
        }

        private void btnChangeEmp_Click(object sender, EventArgs e)
        {
            ChangeEmp = true;
        }

        private void getFellowUpState()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select متابعة_تلقائية from TableSettings", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();


            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["متابعة_تلقائية"].ToString() == "1")
                {
                    FellewUpState.CheckState = CheckState.Checked;
                }else
                    FellewUpState.CheckState = CheckState.Unchecked;
            }

        }

        
       

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            panel1.Size = new System.Drawing.Size(748, 441);//729, 615//
            panel1.Location = new System.Drawing.Point(10, 201);//10, 3
        }

        private void FellewUpState_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (النوع.CheckState == CheckState.Unchecked)
                النوع.Text = "ذكر";
            else النوع.Text = "أنثى";
        }

        private void txtCallerPhone_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {
            panel1.Size = new System.Drawing.Size(748, 441);//729, 615//
            panel1.Location = new System.Drawing.Point(10, 201);//10, 3
        }

        private void panel1_Click(object sender, EventArgs e)
        {
            panel1.Size = new System.Drawing.Size(748, 441);//729, 615//
            panel1.Location = new System.Drawing.Point(10, 201);//10, 3
        }

        private void txtCallerName_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void txtCallerPhone_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            }


        private void txtPhoneNo_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            //MessageBox.Show(e.KeyChar.ToString());
            //if (e.KeyChar == (char)13)
            //{
            //    //checkWithPhone(indexGrid2);
                
            //}
        }

        private void txtPhoneNo_TextChanged(object sender, EventArgs e)
        {
            if (رقم_الهاتف.Text.Length != 12  || dontCheckG2)
            {
                return;
            }
            //else if (رقم_الهاتف.Text.Length == 12)
            //    checkWithPhone(0);            
        }


        private void FormTimeLine_FormClosed(object sender, FormClosedEventArgs e)
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

        private void button1_Click_1(object sender, EventArgs e)
        {
            panel1.SendToBack();
        }

        private void MainPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void المستندات_المتعلقة_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillDatafromGenArch(taskIDNo.ToString(), "TableTasks", true, المستندات_المتعلقة.Text);
        }

        private void المشرف_المناوب_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (JobPosition.Contains("قنصل"))
            {
                if (تحديث_الحالة.SelectedIndex == 0)
                {
                    sms.Text = "عزيزي " + المشرف_المناوب.Text + " تم تكليفكم لإنجاز المهمة رقم " + رقم_المهمة.Text + " نثق في أداءكم المتميز دوما";
                }
                else
                {
                    sms.Text = "عزيزي " + المشرف_المناوب.Text + " يرجى التكرم بمتابعة العمل على المهمة بالرقم (" + رقم_المهمة.Text + ") " + تحديث_الحالة.Text + " نثق في أداءكم المتميز دوما";
                }
            }
            if (ChangeEmp)
            {
                ChangeEmp = false;
                return;
            }
            getPhone(المشرف_المناوب.Text);
            fileComboTask(DataSource, المشرف_المناوب.Text);
        }

        private void riskyCase()
        {
            string text = "تم رصد حالة " + نوع_الحالة.Text + " برقم المتابعة " + رقم_المهمة.Text + " من المواطن/" + اسم_المتصل.Text + "برقم الهاتف: " + رقم_الهاتف.Text;
            if (!JobPosition.Contains("قنصل") && (نوع_الحالة.Text.Contains("حرجة") || نوع_الحالة.Text.Contains("خطرة")))
            {
                var selectedOption = MessageBox.Show("", "تنبيه المشرف برسالة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                    SendSms(getPhone(المشرف_المناوب.Text), text);
            }
        }
        private void نوع_الحالة_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (genReportTable.Rows.Count <= 0) {
                button2.Visible = false;
                MessageBox.Show("لا يوجد بيانات ليوم " + dateTimeTo.Text.Split(' ')[0].Replace("/", "-"));
                
                return;
            }
            button2.Enabled = false;
            string reportName = "Report" + DateTime.Now.ToString("mmss") + ".docx"; 
            dateReport(reportName, dateTimeTo.Text.Split(' ')[0].Replace("/", "-"), genReportTable);
            button2.Visible = false;
        }

        private void dateTimeTo_ValueChanged(object sender, EventArgs e)
        {
            button2.Visible = true;
            button2.Enabled= true;
            string currentDate = dateTimeTo.Text.Split(' ')[0].Replace("/", "-");


            string query = "select * from TableTasks where تاريخ_التكليف='" + currentDate + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;

            sqlDa.Fill(genReportTable);
            sqlCon.Close();
        }

        private void مصدر_الموضوع_CheckedChanged(object sender, EventArgs e)
        {
            if (مصدر_الموضوع.Checked)
                مصدر_الموضوع.Text = "صادر";
            else مصدر_الموضوع.Text = "صادر";
        }

        private void picVerify_Click(object sender, EventArgs e)
        {
            checkWithPhone(0, "رقم_الهاتف", رقم_الهاتف.Text);
        }

        private string callCountFun(int count)
        {
            if (count == 2)
                return "مرتين";
            else if (count > 2 && count < 11)
                return count.ToString() + " مرات";
            else if (count > 10)
                return count.ToString() + " مرة";
            else return "مرة";
        }

        private void checkWithPhone(int index, string col, string text)
        {
            FillDataGridView(dataGridView2, "where "+col+"=N'" + text + "'");
            callCount.Text = "إتصال لأول مرة";
            if (dataGridView2.Rows.Count > 1 && index < (dataGridView2.Rows.Count+1) && index >= 0)
            {
                callCount.Text = callCountFun(dataGridView2.Rows.Count-1);

                taskIDNo = Convert.ToInt32(dataGridView2.Rows[index].Cells[0].Value.ToString());
                التعليقات_السابقة_Off.Text = dataGridView2.Rows[index].Cells["تعليق"].Value.ToString();
                foreach (Control control in MainPanel.Controls)
                {
                    if ((control is TextBox || control is ComboBox || control is CheckBox) && !control.Name.Contains("Off"))
                    {
                        control.Text = dataGridView2.Rows[index].Cells[control.Name].Value.ToString();
                    }

                }
                if (dataGridView2.Rows[index].Cells["مستندات"].Value.ToString() == "yes")
                {
                    FillDatafromGenArch(taskIDNo.ToString(), "TableTasks", false, "");
                }
                picVerify.Visible = false;
                picVerified.Visible = true;
                dataGridView2.Visible = false;
                label13.Visible = missionDate.Visible = اسم_الموظف.Visible = MainPanel.Visible = true;
            }

           

        }

        private void رقم_الهاتف_KeyUp(object sender, KeyEventArgs e)
        {
            if (dataGridView2.Rows.Count > 1 && indexGrid2 < (dataGridView2.Rows.Count - 2))
            {
                indexGrid2++;
                Console.WriteLine("dataGridView2.Rows.Count =" + (dataGridView2.Rows.Count - 1).ToString() + " - indexGrid2 = " + indexGrid2.ToString());
                dontCheckG2 = true;
            }
            //checkWithPhone(indexGrid2);


        }

        private void رقم_الهاتف_KeyDown(object sender, KeyEventArgs e)
        {
            if (dataGridView2.Rows.Count > 1 && indexGrid2 > 0)
            {
                indexGrid2--;
                dontCheckG2 = true;
                Console.WriteLine("dataGridView2.Rows.Count =" + (dataGridView2.Rows.Count - 1).ToString() + " - indexGrid2 = " + indexGrid2.ToString());
            }
            //checkWithPhone(indexGrid2); 

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            checkWithPhone(0, "رقم_المهمة", رقم_المهمة.Text);
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

        

    }
}
