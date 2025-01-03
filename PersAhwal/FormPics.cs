﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
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
using Aspose.Words;
using System;
using System.Runtime.InteropServices;
using WIA;
using PersAhwal;
using System.Net;
using Image = System.Drawing.Image;
using Microsoft.Reporting.WinForms;
using Paragraph = Xceed.Document.NET.Paragraph;
using System.Xml;
using DocumentFormat.OpenXml.Office2010.Excel;
using ZXing;
using System.Data.SqlTypes;
using DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;
using System.Runtime.InteropServices.ComTypes;
using Aspose.Words.Settings;
using DocumentFormat.OpenXml.Office2016.Drawing.Charts;
using System.Text.RegularExpressions;
using static Azure.Core.HttpHeader;
using System.Security.Cryptography.X509Certificates;
using IronBarCode;
using DocumentFormat.OpenXml.Vml;
using ImageFile = WIA.ImageFile;
using SixLabors.ImageSharp.Memory;
using SautinSoft.Document;
using SixLabors.ImageSharp;
using Rectangle = System.Drawing.Rectangle;
using Color = System.Drawing.Color;
using System.Windows.Controls.Primitives;
using static System.Windows.Forms.AxHost;

namespace PersAhwal
{
    public partial class FormPics : Form
    {
        string Payment = "";
        string AppNameText = "";
        string Combo1Text = "";
        string Combo2Text = "";
        string refNo = "";
        string relatedRefNo = "";
        DeviceInfo AvailableScanner = null;
        string relatedPro = "";
        string[] PathImage = new string[100];
        string[] ImageName = new string[100];
        string rowCount = "";
        int imagecount = 0;
        string DataSource = "";
        string FilespathIn;
        string FilespathOut;
        int FormType;
        int Wafid13TableId = 0;
        bool finalArch = false;
        string archstat = "";
        string AuthNoPart1;
        string AuthNoPart2;
        int MandoubDoc = 0;
        string TableList;//= new string[20];
        string ColNumList;//= new string[20];
        string columnList;// = new string[20];
        string archCol;// = new string[20];
        string GenQuery;// = new string[20];
        bool readyToRemove = false;
        //static string queryUpdate;//= new string[20];
        //static string[] queryEntry = new string[15];
        static string[] MandouList= new string[100];
        string FileIDNo = "";
        
        static string[] GriDateM = new string[20000];
        
        string GregorianDate = "";
        bool PreArchieved = false;
        private static List<Stream> m_streams;
        private static int m_currentPageIndex = 0;
        bool ArchiveState;
        int Combo1Index = 0;
        string docIDNumber = "";
        string smsDocIDNumber = "";
        string JobPosition = "";
        string mandoubState = "";
        string CurrentFile = "";
        string PRimariFiles = @"D:\", PrimariFiles = @"D:\PrimariFiles\";
        string EmpName = "";
        string AVcName = "";
        string ServerType= "56";
        //int documentid = 0;
        bool newEntry = true;
        string genColName = "DocID";
        bool smsActiviated = false;
        string smsPhoneNo = "";
        string smsName = "";
        string Labdate = "";
        int drawBoxesindex = 0;
        string textButton = "";
        string picPath = "";
        string noForm = "01";
        
        string[] allInsertList;
        string[] allUpdateList;
        
        string[] allInsertNamesList;
        string[] allUpdateNamesList;
        string[] paraValues;
        string[] comboCol = new string[3];
        //string[] insertList;// = new string[100];
        //string[] updateList;// = new string[dtbl.Rows.Count];
        int archCase = 0;
        
        string proID =  "0";        
        string proForm2Val = "";
        string proForm1Val = "";
        bool dateofBirthcheced = false;
        string reqFile = "";
        bool ArchMandoub = false;
        string ArchMandoubID = "";
        string getComment = "";
        string getProcedName= "";
        string mainProNo= "";
        string altColName = "";
        string altSubColName = "";
        bool dontCheck = false;
        string txtMissionCode = "ق س ج/80/";
        bool MandArch = false;
        public FormPics( string serverType, string gregorianDate, string empName, string aVcName,string jobPosition,string dataSource, int index, string filespathIn, string filespathOut, int formType, string[] strData, string[] strSubData, bool archiveState, string[] mandounList, string[] griDate)
        {
            InitializeComponent();
            Console.WriteLine(1);
            ServerType = serverType;
            DataSource = dataSource;
            FormType = formType;
            GregorianDate = gregorianDate;
            button4.BringToFront();
            button4.BringToFront();
            button3.BringToFront();
            Console.WriteLine(2);
            btnAuth.Visible = true;
            FilespathIn = filespathIn;
            FilespathOut = filespathOut + @"\";
            ArchiveState = archiveState;
            EmpName = empName;
            AVcName = aVcName;
            GriDateM = griDate;
            Combo1Index = index;
            JobPosition = jobPosition;            
            genPreparation(strData, strSubData, index);
            Console.WriteLine(3);
            Combo1.SelectedIndex = Combo1Index;
            Console.WriteLine(4);
            Console.WriteLine(5);
            //correctNo();
            //checkRenew();
            try
            {
                txtMissionCode = missionBasicInfo().Split('*')[3];
                Console.WriteLine(7);
            }
            catch (Exception ex)
            {

            }
            //OpenFile("استمارات المناديب", true);
        }

        private bool checkRenew()
        {
            string query = "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            try
            {
                query = "SELECT DATEDIFF(year, '" + getRenew() + "', '" + GregorianDate + "') as yearDiff";
            }
            catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
                sqlCon.Close();

                foreach (DataRow row in dtbl.Rows)
                {
                    int age = Convert.ToInt32(row["yearDiff"].ToString());
                    Console.WriteLine(query);
                    //MessageBox.Show(row["yearDiff"].ToString());
                    if (age != 0)
                    {
                        this.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(query);
                MessageBox.Show("تاريخ ميلاد غير صالح");
            }
            return false;

        }

        private string getRenew()
        {
            string query = "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            try
            {
                query = "select  renew from TableSettings";
            }
            catch (Exception ex) { return ""; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
                sqlCon.Close();

                foreach (DataRow row in dtbl.Rows)
                {
                    return row["renew"].ToString();
                }
            }
            catch (Exception ex)
            {
                return "";
            }
            return "";

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


        private void fileColComboBox(ComboBox combbox, string source, string comlumnName, string colRight)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct " + comlumnName + " from TableAddContext where " + comlumnName + " is not null and ColRight " + colRight + " order by " + comlumnName + " asc";
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

        private void genPreparation(string[] strData, string[] strSubData,int index)
        {
            docId.Select();
            if(ArchiveState)
                panelFinalArch.Visible = true;

            noForm = FormType.ToString();
            if (FormType < 10) noForm = "0" + noForm;

            if (ArchiveState)
            {
                panel2.Visible = panel3.Visible = false;
                lalProType.Visible = true;
                if (ServerType == "56")
                    getColList(noForm, ArchiveState, Combo1Index.ToString());
                else getColList(noForm, ArchiveState, "-1");
            }
            mandoubState = "عن طريق أحد مندوبي القنصلية";
            //
            if (!ArchiveState)
            {
                label1.Text = "اسم مقدم الطلب";
                docId.Text = DateTime.Now.Year.ToString().Replace("20", "");
                button3.Visible = true;
                DocType.Visible = button4.Visible = true;
                docId.Height = 46;
            }
            else
            {
                //panelFinalArch.Visible = false;
                DocType.Visible = button4.Visible = false;
            }
            if (FormType == 12)
            {
                fileColComboBox(Combo1, DataSource, "altColName", "<> ''");

            }
            else if (FormType == 10)
            {
                fileColComboBox(Combo1, DataSource, "altColName", "= ''");
            }
            else
            {
                
                for (int x = 0; x < strData.Length; x++)
                {
                    if (strData[0] == "")
                    {

                        docId.Enabled = true;
                        Combo1.Visible = false;
                        break;
                    }
                    Combo1.Items.Add(strData[x]);
                }

                if (index != 12 && Combo1.Items.Count > 0)
                    Combo1.SelectedIndex = index;//error

                for (int x = 0; x < 100; x++)
                    PathImage[x] = ImageName[x] = "";
                if (strSubData.Length > 0 && strSubData[0] != "")
                {
                    Combo2.Visible = true;
                    Combo2.Items.Clear();
                    if (!checkColumnName(Combo1.Text.Replace(" ", "_")))
                    {
                        CreateColumn(Combo1.Text.Replace(" ", "_"));
                    }

                    for (int x = 0; x < strSubData.Length; x++)
                    {
                        if (!checkItemName(strSubData[x], Combo1.Text.Replace(" ", "_")))
                        {
                            int id = lastValidID(Combo1.Text.Replace(" ", "_"));
                            addItem(id, Combo1.Text.Replace(" ", "_"), strSubData[x]);
                        }
                        Combo2.Items.Add(strSubData[x]);
                    }
                }
            }
            if (ArchiveState)
            {
                label5.Visible = تاريخ_الميلاد.Visible = true;
                txtIDNo.Visible = false;
                noForm = DocIDGenerator(FormType);                
            }            
        }

        private void insertRow(string source, string[] data)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string[] colList = new string[11];
            colList[1] = "رقم_المعاملة";
            colList[0] = "المعاملة";
            colList[2] = "المطلوب_رقم1";
            colList[3] = "المطلوب_رقم2";
            colList[4] = "المطلوب_رقم3";
            colList[5] = "المطلوب_رقم4";
            colList[6] = "المطلوب_رقم5";
            colList[7] = "المطلوب_رقم6";
            colList[8] = "المطلوب_رقم7";
            colList[9] = "المطلوب_رقم8";
            colList[10] = "المطلوب_رقم9";
            string item = "رقم_المعاملة";
            string value = "@رقم_المعاملة";
            for (int col = 1; col < 11; col++)
            {
                item = item + "," + colList[col];
                value = value + ",@" + colList[col];
            }

            string query = "INSERT INTO TableProcReq (" + item + ") values (" + value + ")";

            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            for (int col = 0; col < 11; col++)
            {
                //MessageBox.Show(colList[col] + ","+data[col]);

                sqlCmd.Parameters.AddWithValue(colList[col], data[col]);
            }
            try
            {
                sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
            }
            sqlCon.Close();
        }

        private void insertReqRow(string source, string strName, string form)
        {
            if (strName == "" || checkProReq(strName))
                return;

            string[] data =  new string[11];
            data[1] = strName;
            data[0] = form;
            data[2] = "غير مدرج";
            data[3] = "غير مدرج";
            data[4] = "غير مدرج";
            data[5] = "غير مدرج";
            data[6] = "غير مدرج";
            data[7] = "غير مدرج";
            data[8] = "غير مدرج";
            data[9] = "غير مدرج";
            data[10] = "غير مدرج";
            
            SqlConnection sqlCon = new SqlConnection(source);
            string[] colList = new string[11];
            colList[0] = "رقم_المعاملة";
            colList[1] = "المعاملة";
            colList[2] = "المطلوب_رقم1";
            colList[3] = "المطلوب_رقم2";
            colList[4] = "المطلوب_رقم3";
            colList[5] = "المطلوب_رقم4";
            colList[6] = "المطلوب_رقم5";
            colList[7] = "المطلوب_رقم6";
            colList[8] = "المطلوب_رقم7";
            colList[9] = "المطلوب_رقم8";
            colList[10] = "المطلوب_رقم9";
            string item = "رقم_المعاملة";
            string value = "@رقم_المعاملة";
            for (int col = 1; col < 11; col++)
            {
                item = item + "," + colList[col];
                value = value + ",@" + colList[col];
            }

            string query = "INSERT INTO TableProcReq (" + item + ") values (" + value + ")";

            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            for (int col = 0; col < 11; col++)
            {
                //MessageBox.Show(colList[col] + ","+data[col]);

                sqlCmd.Parameters.AddWithValue(colList[col], data[col]);
            }
            try
            {
                sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
            }
            sqlCon.Close();
        }

        private bool checkProReq(string proName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query = "SELECT * FROM TableProcReq where المعاملة=N'" + proName + "'";
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
            sqlCon.Close();
            if (dtbl.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }
        private void CreateColumn(string Columnname)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("alter table TableListCombo add " + Columnname + " nvarchar(150)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        } 
        
        private void addItem(int id, string Columnname, string text)
        {
            string qurey = "update TableListCombo set " + Columnname + "=@" + Columnname + " where ID=@id";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@" + Columnname, text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        private string SpecificDigit(string text, int Firstdigits, int Lastdigits)
        {
            char[] characters = text.ToCharArray();
            string firstNchar = "";
            int z = 0;
            for (int x = Firstdigits - 1; x < Lastdigits && x<text.Length; x++)
            {
                firstNchar = firstNchar + characters[x];
                
            }
            return firstNchar;
        }
        
        private void drawBoxes(string text,bool hide, string id)
        {if (text == "" || text == "غير مدرج") return;
            //MessageBox.Show(text.Length % 16);
            PictureBox picAddReq1 = new PictureBox();
            PictureBox picRemReq1 = new PictureBox();
            PictureBox picUplReq1 = new PictureBox();
            //Button label = new Button();
            

            // 
            // picAddReq1
            // 
            int hieght = drawBoxesindex;
            picAddReq1.Image = global::PersAhwal.Properties.Resources.scan;
            picAddReq1.Location = new System.Drawing.Point(68, 3 + (32* hieght));
            picAddReq1.Name = "picAddReq_" + drawBoxesindex.ToString();
            picAddReq1.Size = new System.Drawing.Size(28, 30);
            picAddReq1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            picAddReq1.TabIndex = 507;
            picAddReq1.TabStop = false;
            picAddReq1.Click += new System.EventHandler(this.scanPic);
            picAddReq1.Visible = true;

            

            //
            //picUplReq1
            //
            picUplReq1.Image = global::PersAhwal.Properties.Resources.upload;
            picUplReq1.Location = new System.Drawing.Point(34, 3 + (32 * hieght));
            picUplReq1.Name = "picUplReq_" + drawBoxesindex.ToString();
            picUplReq1.Size = new System.Drawing.Size(28, 30);
            picUplReq1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            picUplReq1.TabIndex = 507;
            picUplReq1.TabStop = false;
            picUplReq1.Click += new System.EventHandler(this.uploadFile);
            picUplReq1.Visible = true;
            // 
            // picRemReq1
            // 
            picRemReq1.Image = global::PersAhwal.Properties.Resources.remove;
            picRemReq1.Location = new System.Drawing.Point(0, 3 + (32 * hieght));
            picRemReq1.Name = "picRemReq_" + drawBoxesindex.ToString();
            picRemReq1.Size = new System.Drawing.Size(28, 30);
            picRemReq1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            picRemReq1.TabIndex = 508;
            picRemReq1.TabStop = false;
            picRemReq1.Click += new System.EventHandler(this.removeFile);
            picRemReq1.Visible = true;
            // 
            // req1
            // 
            Button req1 = new Button();
            req1.Enabled = false;
            req1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            req1.Location = new System.Drawing.Point(101, 3 + (32 * hieght));
            req1.Name = "req_" + drawBoxesindex.ToString();
            req1.Size = new System.Drawing.Size(170, 32);
            
            req1.TabIndex = 475;
            req1.Text = text;
            req1.UseVisualStyleBackColor = true;
            req1.Click += new System.EventHandler(this.showFiles);
            req1.Visible = true;
            if (!hide)
            {
                req1.Location = new System.Drawing.Point(5, 3 + (32 * hieght));
                req1.Size = new System.Drawing.Size(270, 32);
                req1.Enabled = true;
                req1.Name= id;

            }

            drawPicPanel.Controls.Add(req1);
            if (hide)
            {
                drawPicPanel.Controls.Add(picAddReq1);
                drawPicPanel.Controls.Add(picUplReq1);
                drawPicPanel.Controls.Add(picRemReq1);
                
            }
            drawBoxesindex++;
        }
        
        private void drawNameBox()
        {
            PictureBox picUplReq1 = new PictureBox();
            System.Windows.Forms.TextBox AppName = new System.Windows.Forms.TextBox();

            AppName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            AppName.Location = new System.Drawing.Point(3, 38);
            AppName.Multiline = true;
            AppName.Name = "AppName";
            AppName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            AppName.Size = new System.Drawing.Size(290, 77);
            AppName.TabIndex = 847;
            AppName.Visible = true;
            AppName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            AppName.Text = "اكتب اسم مقدم الطلب";
            AppName.MouseClick += new System.Windows.Forms.MouseEventHandler(this.AppName_MouseClick); 
            AppName.TextChanged += new System.EventHandler(this.AppName_TextChanged);
            drawPicPanel.Controls.Add(AppName);
        }
        private void AppName_TextChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox textbox = (System.Windows.Forms.TextBox)sender;
            AppNameText = textbox.Text; 
        }
        
        private void AppName_MouseClick(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox textbox = (System.Windows.Forms.TextBox)sender;
            if (textbox.Text == "اكتب اسم مقدم الطلب")
                textbox.Text = "";
        }

        private void drawBoxesTitle(string text, int xLoc)
        {
            Button label = new Button();
            label.BackColor = System.Drawing.SystemColors.Control;
            label.FlatAppearance.BorderSize = 0;
            label.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            label.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label.Location = new System.Drawing.Point(xLoc, 3 + (32 * drawBoxesindex));
            label.Name = "label";
            label.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            label.Size = new System.Drawing.Size(295, 35);
            label.TabIndex = 593;
            label.Text = text;
            label.UseVisualStyleBackColor = false;


           
            
            //label.AutoSize = true;
            //label.Font = new System.Drawing.Font("Arabic Typesetting", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            //label.ForeColor = System.Drawing.Color.Red;
            //label.Location = new System.Drawing.Point(xLoc, 3 + (32 * drawBoxesindex));
            //label.Name = "label";
            //label.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            //label.Size = new System.Drawing.Size(111, 31);
            //label.TabIndex = 614;
            //label.Text = text;

            drawPicPanel.Controls.Add(label);
            drawBoxesindex++;
        }


        private void loadPreReq(string formNo, string proName, bool archiveState)
        {
            //MessageBox.Show("loadPreReq");
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query = "SELECT * FROM TableProcReq where المعاملة=N'" + proName + "' and رقم_المعاملة='" + formNo + "'";
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
            sqlCon.Close();
            panelFinalArch.Visible = true;
            Console.WriteLine(query);
            if (dtbl.Rows.Count > 0)
            {
                btnSaveEnd.Location = new System.Drawing.Point(736, 668);
                drawBoxesindex = 0;
                foreach (Control control in drawPicPanel.Controls)
                {
                    if (control.Name != "DocType" && control.Name != "button1" && control.Name != "button6" && control.Name != "button5" && control.Name != "checkPrint" && control.Name != "jpgFile" && control.Name != "wordFile")
                        control.Visible = false;

                }
                drawBoxesTitle("المستندات المطلوبة للإجراء", 60);
                توضيح_المعاملة.Visible = false;
                foreach (DataRow row in dtbl.Rows)
                {
                    proForm1Val = row["proForm1"].ToString();
                    proForm2Val = row["proForm2"].ToString();


                    توضيح_المعاملة.Text = row["توضيح_المعاملة"].ToString();
                    relatedPro = row["المعاملات_المرتبطة"].ToString();
                    توضيح_المعاملة.Visible = true;
                    توضيح_المعاملة.BringToFront();
                    proID = row["ID"].ToString();
                    for (int index = 1; index <= 9; index++)
                    {
                        string req = "المطلوب_رقم" + index.ToString();
                        if (row[req].ToString() != "غير مدرج")
                        {
                            drawBoxes(row[req].ToString(), archiveState, "");
                        }
                    }
                    drawBoxes("أخرى", archiveState, "");
                    if (ServerType != "56") panelFinalArch.Visible = false;
                    return;
                }
                return;
            }

        }

        private int lastValidID(string colName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT ID from TableListCombo WHERE " + colName + " is not null", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            int id = 0;
            if (dtbl.Rows.Count > 0)
            {                
                foreach (DataRow row in dtbl.Rows)
                {
                    if (Convert.ToInt32(row["ID"].ToString()) > id)
                        id = Convert.ToInt32(row["ID"].ToString());
                }
            }
            return id + 1;
        }

        private void ListofTables()
        {
//            TableList[0] = "TableDocIqrar";
//            TableList[1] = "TableTravIqrar";
//            TableList[2] = "TableMultiIqrar";
//            TableList[3] = "TableVisaApp";
//            TableList[4] = "TableFamilySponApp";
//            TableList[5] = "TableForensicApp";
//            TableList[6] = "TableTRName";
//            TableList[7] = "TableStudent";
//            TableList[8] = "TableMarriage";
//            TableList[9] = "TableFreeForm";
//            TableList[10] = "";
//            TableList[11] = "TableAuth";
//            TableList[12] = "TableWafid";
//            TableList[13] = "TableSuitCase";
//            TableList[14] = "TableMerrageDoc";
//            TableList[15] = "TablePassAway";
//          TableList[16] = "TableDivorce";
////            unfounddata(string[] tableList)
//            columnList[0] = "AppName";
//            columnList[1] = "AppName";
//            columnList[2] = "AppName";
//            columnList[3] = "AppName";
//            columnList[4] = "AppName";
//            columnList[5] = "AppName";
//            columnList[6] = "AppName";
//            columnList[7] = "AppName";
//            columnList[8] = "AppName";
//            columnList[9] = "AppName";
//            columnList[10] = "AppName";
//            columnList[11] = "مقدم_الطلب";
//            columnList[12] = "مقدم_الطلب";
//            columnList[14] = "اسم_الزوج";
//            columnList[15] = "اسم_المتوفى";
//            columnList[16] = "اسم_الزوج";
            

//            query[1] = "INSERT INTO TableTravIqrar (AppName,ProType,DocID,GriDate,DataMandoubName,DataInterType) values (@AppName,@ProType,@DocID,@GriDate,@DataMandoubName,@DataInterType);SELECT @@IDENTITY as lastid";
//            query[2] = "INSERT INTO TableMultiIqrar (AppName,IqrarPurpose,DocID,GriDate,DataMandoubName,DataInterType) values (@AppName,@IqrarPurpose,@DocID,@GriDate,@DataMandoubName,@DataInterType);SELECT @@IDENTITY as lastid";
//            query[3] = "INSERT INTO TableVisaApp (AppName,DocID,GriDate,DataMandoubName,DataInterType) values (@AppName,@DocID,@GriDate,@DataMandoubName,@DataInterType);SELECT @@IDENTITY as lastid";
//            query[4] = "INSERT INTO TableFamilySponApp (AppName,ProCase,DocID,GriDate,DataMandoubName,DataInterType) values (@AppName,@ProCase,@DocID,@GriDate,@DataMandoubName,@DataInterType);SELECT @@IDENTITY as lastid";
//            query[5] = "INSERT INTO TableForensicApp (AppName,DocID,GriDate,DataMandoubName,DataInterType,purpose) values (@AppName,@DocID,@GriDate,@DataMandoubName,@DataInterType,@purpose);SELECT @@IDENTITY as lastid";
//            query[6] = "INSERT INTO TableTRName (AppName,IqrarType,DocID,GriDate,DataMandoubName,DataInterType) values (@AppName,@IqrarType,@DocID,@GriDate,@DataMandoubName,@DataInterType);SELECT @@IDENTITY as lastid";
//            query[7] = "INSERT INTO TableStudent (AppName,DocID,GriDate,DataMandoubName,DataInterType) values (@AppName,@DocID,@GriDate,@DataMandoubName,@DataInterType);SELECT @@IDENTITY as lastid";
//            query[8] = "INSERT INTO TableMarriage (AppName,IqamaNo,DocID,GriDate,DataMandoubName,DataInterType) values (@AppName,@IqamaNo,@DocID,@GriDate,@DataMandoubName,@DataInterType);SELECT @@IDENTITY as lastid";
//            query[9] = "INSERT INTO TableFreeForm (AppName,SpecType,DocID,GriDate,DataMandoubName,DataInterType) values (@AppName,@SpecType,@DocID,@GriDate,@DataMandoubName,@DataInterType);SELECT @@IDENTITY as lastid";
//            query[11] = "INSERT INTO TableAuth (مقدم_الطلب,رقم_التوكيل,إجراء_التوكيل,نوع_التوكيل,التاريخ_الميلادي,اسم_المندوب,طريقة_الطلب) values (@مقدم_الطلب,@رقم_التوكيل,@إجراء_التوكيل,@نوع_التوكيل,@التاريخ_الميلادي,@اسم_المندوب,@طريقة_الطلب);SELECT @@IDENTITY as lastid";
//            query[12] = "INSERT INTO TableWafid (مقدم_الطلب,رقم_المعاملة,نوع_المعاملة,التاريخ_الميلادي,اسم_المندوب,طريقة_الطلب,جهة_العمل,رقم_الملف) values (@مقدم_الطلب,@رقم_المعاملة,@نوع_المعاملة,@التاريخ_الميلادي,@اسم_المندوب,@طريقة_الطلب,@جهة_العمل,@رقم_الملف);SELECT @@IDENTITY as lastid";
//            query[13] = "INSERT INTO TableSuitCase (مقدم_الطلب,التاريخ_الميلادي,رقم_لبرقية) values (@مقدم_الطلب,@التاريخ_الميلادي,@رقم_لبرقية);SELECT @@IDENTITY as lastid";
//            query[14] = "INSERT INTO TableMerrageDoc (التاريخ_الميلادي,رقم_المعاملة) values (@التاريخ_الميلادي,@رقم_المعاملة);SELECT @@IDENTITY as lastid";
//            query[15] = "INSERT INTO TablePassAway (التاريخ_الميلادي,رقم_اذن_الدفن) values (@التاريخ_الميلادي,@رقم_اذن_الدفن);SELECT @@IDENTITY as lastid";
//            query[16] = "INSERT INTO TableDivorce (التاريخ_الميلادي,رقم_المعاملة) values (@التاريخ_الميلادي,@رقم_المعاملة);SELECT @@IDENTITY as lastid";
            
//            //queryEntry[1] = "UPDATE TableTravIqrar SET Data1=@Data1,Extension1=@Extension1,FileName1=@FileName1 WHERE DocID=@DocID";
//            //queryEntry[2] = "UPDATE TableMultiIqrar SET Data1=@Data1,Extension1=@Extension1,FileName1=@FileName1 WHERE DocID=@DocID";
//            //queryEntry[3] = "UPDATE TableVisaApp SET Data1=@Data1,Extension1=@Extension1,FileName1=@FileName1 WHERE DocID=@DocID";
//            //queryEntry[4] = "UPDATE TableFamilySponApp SET Data1=@Data1,Extension1=@Extension1,FileName1=@FileName1 WHERE DocID=@DocID";
//            //queryEntry[5] = "UPDATE TableForensicApp SET Data1=@Data1,Extension1=@Extension1,FileName1=@FileName1 WHERE DocID=@DocID";
//            //queryEntry[6] = "UPDATE TableTRName SET Data1=@Data1,Extension1=@Extension1,FileName1=@FileName1 WHERE DocID=@DocID";
//            //queryEntry[7] = "UPDATE TableStudent SET Data1=@Data1,Extension1=@Extension1,FileName1=@FileName1 WHERE DocID=@DocID";
//            //queryEntry[8] = "UPDATE TableMarriage SET Data1=@Data1,Extension1=@Extension1,FileName1=@FileName1 WHERE DocID=@DocID";
//            //queryEntry[9] = "UPDATE TableFreeForm SET Data1=@Data1,Extension1=@Extension1,FileName1=@FileName1 WHERE DocID=@DocID";
//            //queryEntry[11] = "UPDATE TableAuth SET Data1=@Data1,Extension1=@Extension1,data1=@data1 WHERE رقم_التوكيل=@رقم_التوكيل";
//            //queryEntry[12] = "UPDATE TableWafid SET Data1=@Data1,Extension1=@Extension1,data1=@data1 WHERE رقم_المعاملة=@رقم_المعاملة";
//            //queryEntry[13] = "UPDATE TableSuitCase SET Data1=@Data1,Extension1=@Extension1,data1=@data1 WHERE رقم_لبرقية=@رقم_لبرقية";

//            queryUpdate[1] = "UPDATE TableTravIqrar SET ArchivedState=@ArchivedState WHERE DocID=@DocID"; 
//            queryUpdate[2] = "UPDATE TableMultiIqrar SET ArchivedState=@ArchivedState,GriDate=@GriDate WHERE DocID=@DocID";
//            queryUpdate[3] = "UPDATE TableVisaApp SET ArchivedState=@ArchivedState,GriDate=@GriDate WHERE DocID=@DocID";
//            queryUpdate[4] = "UPDATE TableFamilySponApp SET ArchivedState=@ArchivedState,GriDate=@GriDate WHERE DocID=@DocID"; 
//            queryUpdate[5] = "UPDATE TableForensicApp SET ArchivedState=@ArchivedState,GriDate=@GriDate WHERE DocID=@DocID";
//            queryUpdate[6] = "UPDATE TableTRName SET ArchivedState=@ArchivedState,GriDate=@GriDate WHERE DocID=@DocID"; 
//            queryUpdate[7] = "UPDATE TableStudent SET ArchivedState=@ArchivedState,GriDate=@GriDate WHERE DocID=@DocID"; 
//            queryUpdate[8] = "UPDATE TableMarriage SET ArchivedState=@ArchivedState,GriDate=@GriDate WHERE DocID=@DocID"; 
//            queryUpdate[9] = "UPDATE TableFreeForm SET ArchivedState=@ArchivedState,GriDate=@GriDate WHERE DocID=@DocID";
//            queryUpdate[11] = "UPDATE TableAuth SET حالة_الارشفة=@حالة_الارشفة,التاريخ_الميلادي=@التاريخ_الميلادي WHERE رقم_التوكيل=@رقم_التوكيل";
//            //queryUpdate[12] = Wafid13query();
//            queryUpdate[13] = "UPDATE TableSuitCase SET حالة_الارشفة=@حالة_الارشفة,التاريخ_الميلادي=@التاريخ_الميلادي WHERE رقم_لبرقية=@رقم_لبرقية";
//            queryUpdate[14] = "UPDATE TableMerrageDoc SET حالة_الارشفة=@حالة_الارشفة,التاريخ_الميلادي=@التاريخ_الميلادي WHERE رقم_المعاملة=@رقم_المعاملة";
//            queryUpdate[15] = "UPDATE TablePassAway SET حالة_الارشفة=@حالة_الارشفة,التاريخ_الميلادي=@التاريخ_الميلادي WHERE رقم_اذن_الدفن=@رقم_اذن_الدفن";
//        queryUpdate[16] = "UPDATE TableDivorce SET حالة_الارشفة=@حالة_الارشفة,التاريخ_الميلادي=@التاريخ_الميلادي WHERE رقم_المعاملة=@رقم_المعاملة";
        }

        private void getColList(string formType, bool archiveState, string index)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('TableFileArch')", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            //insertList = new string[dtbl.Rows.Count];
            //updateList = new string[dtbl.Rows.Count];
            
            paraValues = new string[dtbl.Rows.Count];
            
            int insexIndex = 0;
            int updateIndex = 0;
            int comboIndsex = 0;
            
            string insertItems = "";
            string insertValues = "";
            string updateValues = "";
            foreach (DataRow row in dtbl.Rows)
            {

                if (row["name"].ToString().Contains("insert")) 
                {
                    //MessageBox.Show(row["name"].ToString());
                    insexIndex++;
                }
                else if (row["name"].ToString().Contains("update")) 
                {
                    updateIndex++;
                }
                //else if (row["name"].ToString().Contains("combo")) 
                //{
                //    comboCol[comboIndsex] = row["name"].ToString();
                //    comboIndsex++;
                //}
            }
            
            allInsertList = new string[insexIndex];
            allUpdateList = new string[updateIndex];
            
            allInsertNamesList = new string[insexIndex];
            allUpdateNamesList = new string[updateIndex];
            for (int x = 0; x < insexIndex; x++)
            {
                allInsertNamesList[x] = allInsertList[x] = "";
            }
            for (int x = 0; x < updateIndex; x++)
            {
                allUpdateNamesList[x] = allUpdateList[x]  = "";
            }
            insexIndex = 0;
            updateIndex = 0;

            foreach (DataRow row in dtbl.Rows)
            {
                if (row["name"].ToString().Contains("insert")) 
                {
                    allInsertList[insexIndex] = row["name"].ToString(); 
                    insexIndex++;
                }
                else if (row["name"].ToString().Contains("update")) 
                {
                    allUpdateList[updateIndex] = row["name"].ToString(); 
                    updateIndex++;
                }
            }
            
            string query = "SELECT * FROM TableFileArch WHERE indexValue ='" + index + "' and FormType='" + formType + index+"'";
            if ((ServerType == "56" && !ArchiveState) ||(ServerType == "57"))
                query = "SELECT * FROM TableFileArch WHERE indexValue ='" + index + "' and FormType='" + formType + "'";
            Console.WriteLine(query);
            //MessageBox.Show(query);
            //MessageBox.Show(ServerType);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    sqlDa = new SqlDataAdapter(query, sqlCon);
                    sqlDa.SelectCommand.CommandType = CommandType.Text;
                    dtbl = new DataTable();
                    sqlDa.Fill(dtbl);
                    sqlCon.Close();
                }
                catch (Exception ex) { 
                }

           
            foreach (DataRow row in dtbl.Rows)
            {                
                TableList = row["TableList"].ToString();
                altColName = row["insert3"].ToString();
                //MessageBox.Show(altColName);
                altSubColName = row["insert4"].ToString();
                ColNumList = row["insert2"].ToString();
                getComment = row["getComment"].ToString();
                columnList = row["columnList"].ToString();
                archCol = row["ArchCol"].ToString();
                comboCol[0] = row["combo1"].ToString();
                comboCol[1] = row["combo2"].ToString();
                comboCol[2] = row["indexCombo"].ToString();
                for (int rows = 0; rows < allInsertList.Length; rows++)
                {
                    if (allInsertList[rows].Contains("insert") )
                    {
                        allInsertNamesList[rows] = row[allInsertList[rows]].ToString();
                        //MessageBox.Show("insert " + allInsertNamesList[fullRows]);
                    }                    
                }

                for (int rows = 0; rows < allUpdateList.Length; rows++)
                {
                    if (allUpdateList[rows].Contains("update"))
                    {
                        allUpdateNamesList[rows] = row[allUpdateList[rows]].ToString();
                        //MessageBox.Show(allUpdateNamesList[rows]);
                    }
                }
            }

            if (archiveState)
            {
                foreach (DataRow row in dtbl.Rows)
                {
                    for (int rows = 0; rows < allInsertList.Length; rows++)
                    {
                        if (allInsertList[rows] != "")
                        {
                            if (row[allInsertList[rows]].ToString() != "")
                            {
                                if (rows == 0)
                                {
                                    insertItems = row[allInsertList[rows]].ToString();
                                    insertValues = "@" + row[allInsertList[rows]].ToString();
                                }
                                else
                                {
                                    insertItems = insertItems + "," + row[allInsertList[rows]].ToString();
                                    insertValues = insertValues + "," + "@" + row[allInsertList[rows]].ToString();
                                }
                            }
                        }
                    }
                }
                GenQuery = "INSERT INTO " + TableList + "(" + insertItems + ") values (" + insertValues + ");SELECT @@IDENTITY as lastid";
                
            }
            else {
                foreach (DataRow row in dtbl.Rows)
                {
                    for (int rows = 0; rows < allUpdateList.Length; rows++)
                    {
                        if (allUpdateList[rows] != "")
                        {
                            if (row[allUpdateList[rows]].ToString() != "")
                            {
                                if (rows == 0)
                                {
                                    updateValues = row[allUpdateList[rows]].ToString() + "=@" + row[allUpdateList[rows]].ToString();
                                }
                                else
                                {
                                    updateValues = updateValues + ", " + row[allUpdateList[rows]].ToString() + "=@" + row[allUpdateList[rows]].ToString();
                                }
                            }
                        }
                    }
                }
                GenQuery = "UPDATE " + TableList + " SET " + updateValues + " where ID = @id";                
            }

            Console.WriteLine("GenQuery " + GenQuery);
            //MessageBox.Show("GenQuery " + GenQuery);
        }
        private string getTableList(string formType)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT TableList FROM TableFileArch WHERE FormType='" + formType + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow row in dtbl.Rows)
            {
                TableList = row["TableList"].ToString();
            }
            return TableList;
        }
        
        

        private string getColumnList(string text,string colName)
        {
            string col = "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * FROM TableFileArch WHERE "+ colName + "='" + text + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
            }
            catch (Exception ex) { }
            sqlCon.Close();
            foreach (DataRow row in dtbl.Rows)
            {
                col = row["columnList"].ToString();                
            }  
            return col;
        }

        private void FinalDataArch(string dataSource, string documentID)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(GenQuery, sqlCon);            
            sqlCmd.CommandType = CommandType.Text;
            archstat = "مؤرشف نهائي";
            if (mandoubName.Text != "حضور مباشرة إلى القنصلية" && (archCase == 2 || archCase == 3))
                archstat = "مؤرشف نهائي_" + mandoubName.Text.Split('-')[0];
            paraValues[0] = archstat;
            paraValues[1] = GregorianDate;
            paraValues[3] = تاريخ_الميلاد.Text;
            paraValues[4] = DateTime.Now.ToString("hh:mm:ss");
            sqlCmd.Parameters.AddWithValue("@id", FileIDNo);
            for (int rows = 0; rows < allUpdateNamesList.Length; rows++)
            {
                if (allUpdateNamesList[rows] != "")
                {
                    //MessageBox.Show(allUpdateNamesList[rows] +" - "+ paraValues[rows]);
                    sqlCmd.Parameters.AddWithValue("@" + allUpdateNamesList[rows], paraValues[rows]);
                }
            }
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        
        private bool checkMessageNo(string no)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string settingData = "select رقم_لبرقية from TableSuitCase";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["رقم_لبرقية"].ToString() == no)
                {
                    return true;
                }
            }
            return false;
        }

        private int getID(string text)
        {
            bool found = false;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string settingData = "select ID,EnterySheet from TableListCombo";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["EnterySheet"].ToString() == text)
                {
                    found = true;
                }

            }
            if (!found)
            {
                foreach (DataRow dataRow in dtbl.Rows)
                {
                    if (dataRow["EnterySheet"].ToString() == "")
                    {
                        return Convert.ToInt32(dataRow["ID"].ToString());
                    }
                }
            }
            return 0;
        }

        private void editForms()
        {
            bool found = false;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string settingData = "select ID,EnterySheet from TableListCombo";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);

            if (mandoubName.Text == "حضور مباشرة إلى القنصلية") { MessageBox.Show("يرجى إختيار اسم المندوب"); return; }
            int count1 = 0;
            int count2 = 1;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                string wordInFile = FilespathIn + dataRow["EnterySheet"].ToString().Trim();

                if (File.Exists(wordInFile) && dataRow["EnterySheet"].ToString() != "")
                {
                    count1++;
                }
            }
                    foreach (DataRow dataRow in dtbl.Rows)
            {
                string wordInFile = FilespathIn + dataRow["EnterySheet"].ToString().Trim();
                
                    if (File.Exists(wordInFile) && dataRow["EnterySheet"].ToString() != "")
                {
                    var selectedOption = MessageBox.Show("طباعة (" + count1.ToString() + "/" + count2.ToString() + ")", dataRow["EnterySheet"].ToString(), MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (selectedOption == DialogResult.Yes)
                    {
                        CreateMandoubfile(dataRow["EnterySheet"].ToString(), wordInFile, mandoubLine(mandoubName.Text.Split('-')[0].Trim()), true);
                    }
                    count2++;
                }

            }
            
        }

        private string getNumber()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string settingData = "select SudAffNo from TableSettings where ID='1'";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            string ver = txtMissionCode + "22/13/0";
            foreach (DataRow dataRow in dtbl.Rows)
            {                
                ver = (Convert.ToInt32(dataRow["SudAffNo"].ToString().Split('/')[4]) + 1).ToString();
            }
            return ver;
        }
        

        private int NewReportEntry(string dataSource,string query, int form, string com1, string com2)
        {
            //GenQuery
            if (mandoubName.SelectedIndex == 0)
                mandoubState = "حضور مباشرة إلى القنصلية";            
            else mandoubState = "عن طريق أحد مندوبي القنصلية";

            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;

            paraValues[0] = "";
            paraValues[1] = GregorianDate;
            noForm = DocIDGenerator(form);
            //if(AuthNoPart1)
            //MessageBox.Show(GregorianDate);
            paraValues[2] = AuthNoPart1;
            paraValues[3] = com1;
            paraValues[4] = com2;
            if (mandoubName.SelectedIndex == 0) 
                paraValues[5] = "";
            else 
            paraValues[5] = mandoubName.Text;
            paraValues[6] = mandoubState;
            paraValues[7] = "غير مؤرشف";
            if(ServerType == "56")
                paraValues[8] = "1";
            else paraValues[8] = Combo2.SelectedIndex.ToString();
            paraValues[9] = تاريخ_الميلاد.Text;
            paraValues[10] = DateTime.Now.ToString("hh:mm:ss");
            for (int rows = 0; rows < allInsertNamesList.Length; rows++)
            {
                if (allInsertNamesList[rows] != "")
                {
                    Console.WriteLine(allInsertNamesList[rows] +" - "+paraValues[rows]);
                    sqlCmd.Parameters.AddWithValue("@" + allInsertNamesList[rows], paraValues[rows]); 
                }
            }
            try
            {
                var reader = sqlCmd.ExecuteReader();
                if (reader.Read())
                {
                    return Convert.ToInt32(reader["lastid"].ToString());
                }
                sqlCon.Close();
            }
            catch (Exception ex) {
                printOut = false;
                MessageBox.Show(GenQuery);
            }
            return 0;
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
                Console.WriteLine("www.hisms.ws" + result);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }
        private void updateReportEntry(string dataSource, string filePath, string authNo)
        {
            
            //SqlConnection sqlCon = new SqlConnection(dataSource);
            //if (sqlCon.State == ConnectionState.Closed)
            //    sqlCon.Open();
            //SqlCommand sqlCmd = new SqlCommand(queryEntry[FormType - 1], sqlCon);
            //sqlCmd.CommandType = CommandType.Text;
            //if (FormType == 12)
            //{
            //    sqlCmd.Parameters.AddWithValue("@رقم_التوكيل", authNo);
                
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
            //            sqlCmd.Parameters.Add("@data1", SqlDbType.NVarChar).Value = DocName1;                        
            //        }
            //    }
            //}
            //else
            //{
            //    sqlCmd.Parameters.AddWithValue("@DocID", authNo);
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
            //            sqlCmd.Parameters.Add("@FileName1", SqlDbType.NVarChar).Value = DocName1;                        
            //        }

            //    }

            //}
            //sqlCmd.ExecuteNonQuery();

            //sqlCon.Close();
        }


        private void FormPics_Load(object sender, EventArgs e)
        {
            loadScanner();
            fileComboBoxMandoub(mandoubName, DataSource, "TableMandoudList");
            mandoublist();
        }
        private void mandoublist()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select distinct mandoubName from archives", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            MandouList = new string[dtbl.Rows.Count];
            int i = 0;
            foreach (DataRow dataRow in dtbl.Rows) 
            {
                MandouList[i] = dataRow["mandoubName"].ToString();
                    i++;
            }
        }
        private int todayList(string mandoubName, string date)
        {
            int found = 0;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select docDate,docID,appName from archives where mandoubName=@mandoubName and docDate <> N'"+date+"'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@mandoubName", mandoubName);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows) 
            {
                if (dataRow["docDate"].ToString() != date)
                {
                    //MessageBox.Show(dataRow["docID"].ToString() + " - "+ dataRow["appName"].ToString());
                    found++;
                }
            }
            return found;
        }
        
        private bool checkMandounbPro(string docID)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select appOldNew from archives where docID=@docID", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@docID", docID);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows) 
            {
                if (dataRow["appOldNew"].ToString() == "في انتظار نسخة المواطن")
                {
                    return true;
                }
            }
            return false;
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
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex) { }
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                try
                {
                    dataAdapter.Fill(table);
                    foreach (DataRow dataRow in table.Rows)
                    {
                        if (dataRow["MandoubNames"].ToString() != "")
                            combbox.Items.Add(dataRow["MandoubNames"].ToString() + " - " + dataRow["MandoubAreas"].ToString());
                    }
                }catch (Exception ex) { }
                saConn.Close();
            }
            if (combbox.Items.Count > 0)
                combbox.SelectedIndex = 0;
        }

        private string mandoubLine(string name)
        {
            string str = "";
            using (SqlConnection saConn = new SqlConnection(DataSource))
            {
                saConn.Open();

                string query = "select MandoubNames,MandoubAreas,MandoubPhones,مواعيد_الحضور from TableMandoudList";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;


                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                foreach (DataRow dataRow in table.Rows)
                {
                    if (dataRow["MandoubNames"].ToString() == name)
                    {
                        str = " خاصة بالسيد " + dataRow["MandoubNames"].ToString() + " مندوب جالية منطقة " + dataRow["MandoubAreas"].ToString() + Environment.NewLine + " موعد المواجعة  " + dataRow["مواعيد_الحضور"].ToString() + " رقم الهاتف " + dataRow["MandoubPhones"].ToString();
                        if (dataRow["MandoubAreas"].ToString() == "القنصلية العامة لجمهورية السودان - جدة")
                            str = " خاصة بالسيد " + dataRow["MandoubNames"].ToString() + " مندوب القنصلية العامة لجمهورية السودان - جدة رقم الهاتف " + dataRow["MandoubPhones"].ToString();
                    }
                }
                saConn.Close();
            }
            //MessageBox.Show(str);
            return str;
        }

        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName)
        {
            combbox.Visible = true;
            combbox.Items.Clear();            
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;

                //MessageBox.Show(query);
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex) { return; }
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                foreach (DataRow dataRow in table.Rows)
                {
                    if (dataRow[comlumnName].ToString() != "") combbox.Items.Add(dataRow[comlumnName].ToString());
                }
                saConn.Close();
            }
            //if (combbox.Items.Count > 0) 
            //    combbox.SelectedIndex = 0;
            //MessageBox.Show(combbox.Text);
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
        private void btnAuth_Click(object sender, EventArgs e)
        {
            dataGridView2.Visible = false;
            panel1.Visible = true;

            loadPic.Enabled = btnSaveEnd.Visible = btnAuth.Enabled = false;
            
            try

            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.

                    var ScanerItem = device.Items[1]; // select the scanner.


                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);

                    PathImage[imagecount] = FilespathOut + "ScanImg" + rowCount + imagecount.ToString() + ".jpg";
                    ImageName[imagecount]= "ScanImg" + rowCount + imagecount.ToString();
                    Console.WriteLine(imagecount.ToString() + " = "+PathImage[imagecount]);
                    //MessageBox.Show (PathImage[imagecount]);
                    if (File.Exists(PathImage[imagecount]))
                    {
                        File.Delete(PathImage[imagecount]);
                    }
                    imgFile.SaveFile(PathImage[imagecount]);
                    pictureBox1.ImageLocation = PathImage[imagecount];
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
            loadPic.BackColor = btnAuth.BackColor = System.Drawing.Color.LightGreen;
            loadPic.Text = btnAuth.Text = "اضافة مستند آخر (" + imagecount.ToString() + ")";

            loadPic.Enabled = btnSaveEnd.Visible = btnAuth.Enabled = true;
            btnAuth.Select();
        }

        private void drawTempPics(string location)
        {
            PictureBox picTemp = new PictureBox();
            picTemp.Dock = System.Windows.Forms.DockStyle.Top;
            picTemp.Location = new System.Drawing.Point(0, 0);
            picTemp.Name = "picTemp_"+ imagecount.ToString();
            picTemp.Size = new System.Drawing.Size(123, 137);
            picTemp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            picTemp.TabIndex = 841;
            picTemp.TabStop = false;
            picTemp.Click += new System.EventHandler(this.viewDeletePic); 
            picTemp.ImageLocation = location;
            panelpicTemp.Controls.Add(picTemp);
            panelpicTemp.Visible = true;
            توضيح_المعاملة.SendToBack();
        }
        private void viewDeletePic(object sender, EventArgs e)
        {
            PictureBox pictureBox = (PictureBox)sender;
            pictureBox1.ImageLocation = PathImage[Convert.ToInt32(pictureBox.Name.Split('_')[1])];
            //MessageBox.Show(pictureBox.Name.Split('_')[1]);
            var selectedOption = MessageBox.Show("حذف المستند من قائمة الأرشفة؟", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                pictureBox.Visible = false;
                ImageName[Convert.ToInt32(pictureBox.Name.Split('_')[1])] = PathImage[Convert.ToInt32( pictureBox.Name.Split('_')[1])] = "";
                
            }
        }
        private void scanPic(object sender, EventArgs e)
        {


            توضيح_المعاملة.SendToBack();
            dataGridView2.Visible = false;
            panel1.Visible = drawPicPanel.Visible = true;
            string btnName = "";
            
            Button button = new Button();
            PictureBox pictureBox = (PictureBox)sender;
            //picIndex = Convert.ToInt32(pictureBox.Name.Split('_')[1]);
            foreach (Control control in drawPicPanel.Controls) 
            {
                if (control.Name.Contains("req_") && control.Name.Split('_')[1] == pictureBox.Name.Split('_')[1])
                {
                    btnName = control.Text;
                    button.Name = control.Name;
                }
            }

            loadPic.Enabled = btnSaveEnd.Visible = btnAuth.Enabled = false;
            //MessageBox.Show(btnName + "_"+ picIndex.ToString());
            try

            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.
                    var ScanerItem = device.Items[1]; // select the scanner.
                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);
                    PathImage[imagecount] = FilespathOut + btnName + "_" + rowCount + imagecount.ToString() + ".jpg";
                    ImageName[imagecount] = btnName + "_" + rowCount + imagecount.ToString();
                    
                    if (File.Exists(PathImage[imagecount]))
                    {
                        File.Delete(PathImage[imagecount]);
                    }
                    imgFile.SaveFile(PathImage[imagecount]);
                    pictureBox1.ImageLocation = PathImage[imagecount];


                    //Bitmap bmp2 = new Bitmap(pictureBox1.Width, pictureBox1.Height);
                    //pictureBox1.DrawToBitmap(bmp2, pictureBox1.ClientRectangle);
                    //int rectW = 370, rectH = 307;
                    //int crpX = 370, crpY = 307;
                    //Bitmap crpImg = new Bitmap(rectW, rectH);

                    //for (int i = 0; i < rectW; i++)
                    //{
                    //    for (int y = 0; y < rectH; y++)
                    //    {
                    //        Color pxlclr = bmp2.GetPixel(crpX + i, crpY + y);
                    //        crpImg.SetPixel(i, y, pxlclr);
                    //    }
                    //}

                    //pictureBox2.Image = (Image)crpImg;
                    //pictureBox2.SizeMode = PictureBoxSizeMode.Zoom;


                    drawTempPics(PathImage[imagecount]);
                    imagecount++;
                    try
                    {
                        foreach (Control control in drawPicPanel.Controls)
                        {
                            if (control.Text == btnName)
                            {
                                control.BackColor = System.Drawing.Color.LightGreen;
                                control.Enabled = true;
                                if (reqFile == "")
                                {
                                    OpenFile(Combo1.Text + "-" + Combo2.Text, false); 
                                    if (reqFile == "")
                                        OpenFile(Combo1.Text, false);
                                }
                            }
                        }
                        if(btnSaveEnd.Visible)
                            btnSaveEnd.Select();
                    }
                    catch (Exception ex)
                    {
                        pictureBox1.Image = PersAhwal.Properties.Resources.noImage;
                    }
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


        private static Image cropImage(Image img, Rectangle cropArea)
        {
            Bitmap bmpImage = new Bitmap(img);
            Bitmap bmpCrop = bmpImage.Clone(cropArea,
            bmpImage.PixelFormat);
            return (Image)(bmpCrop);
        }

        private void removeFile(object sender, EventArgs e)
        {
            dataGridView2.Visible = false;
            panel1.Visible = true;
            int picIndex = 0;
            PictureBox pictureBox = (PictureBox)sender;
            picIndex = Convert.ToInt32(pictureBox.Name.Split('_')[1]);

            foreach (Control control in drawPicPanel.Controls)
            {
                if (control.Name == "req_" + pictureBox.Name.Split('_')[1])
                {
                    control.Enabled = false;
                    control.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                }
            }
            loadPic.Enabled = btnSaveEnd.Visible = btnAuth.Enabled = false;
            pictureBox1.Image = PersAhwal.Properties.Resources.noImage;
            PathImage[picIndex] = ImageName[picIndex] = "";            
        }
        
        private void uploadFile(object sender, EventArgs e)
        {
            dataGridView2.Visible = false;
            drawPicPanel.Visible = true;
            توضيح_المعاملة.SendToBack();    
            panel1.Visible = true;
            string btnName = "";
            PictureBox pictureBox = (PictureBox)sender;
            Button button = new Button();
            foreach (Control control in drawPicPanel.Controls)
            {
                if (control.Name.Contains("req_") && control.Name.Split('_')[1] == pictureBox.Name.Split('_')[1])
                {
                    btnName = control.Text;
                    button.Name = control.Name;
                }
            }
            string fileName = "";
            try
            {
                int picIndex = Convert.ToInt32(pictureBox.Name.Split('_')[1]);
                fileName = loadDocxFile();
                var fileinfo = new FileInfo(fileName);
                string extn = fileinfo.Extension;
            }
            catch (Exception ex) { return; }
            if (fileName != "")
            {
                //fileName = fileName.Replace(extn, btnName) + extn; ;
                PathImage[imagecount] = FilespathOut + btnName + "_" + rowCount + imagecount.ToString();
                PathImage[imagecount] = fileName;
                ImageName[imagecount] = button.Name;
                drawTempPics(PathImage[imagecount]);
                imagecount++;
                if (reqFile == "")
                {
                    OpenFile(Combo1.Text + "-" + Combo2.Text, false);
                    if (reqFile == "")
                        OpenFile(Combo1.Text, false);
                    
                        
                }
                foreach (Control control in drawPicPanel.Controls)
                {
                    if (control.Text == btnName)
                    {
                        control.BackColor = System.Drawing.Color.LightGreen;
                        control.Enabled = true;
                    }
                }
                try
                {
                    pictureBox1.Image = PersAhwal.Properties.Resources.noImage;
                    System.Diagnostics.Process.Start(fileName);
                }
                catch (Exception ex) {
                    
                }
                if (btnSaveEnd.Visible)
                    btnSaveEnd.Select();
            }
        }


        private string loadRerNo(int id, int form, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT DocID from " + table + " where ID=@ID", sqlCon); 
            if(form == 12) sqlDa = new SqlDataAdapter("SELECT رقم_التوكيل from TableAuth where ID=@ID", sqlCon);
            else if (form == 13) sqlDa = new SqlDataAdapter("SELECT رقم_المعاملة from TableWafid where ID=@ID", sqlCon);
            else if (form == 15) sqlDa = new SqlDataAdapter("SELECT رقم_المعاملة from TableMerrageDoc where ID=@ID", sqlCon);
            else if (form == 16) sqlDa = new SqlDataAdapter("SELECT رقم_اذن_الدفن from TablePassAway where ID=@ID", sqlCon);
            else if (form == 17) sqlDa = new SqlDataAdapter("SELECT رقم_المعاملة from TableDivorce where ID=@ID", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ID", id);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string rowCnt = "0";

            if (form == 12)
            {
                genColName = "رقم_التوكيل";
            }

            else if (form < 12)
            {
                genColName = "DocID";
            }

            else if (form == 13)// && row["رقم_المعاملة"].ToString().Split('/').Length == 5)
            {
                genColName = "رقم_المعاملة";
            }
            else if (form == 15 ||form == 17)// && row["رقم_المعاملة"].ToString().Split('/').Length == 5)
            {
                 genColName = "رقم_المعاملة";
            }
            else if (form == 16)// && row["رقم_المعاملة"].ToString().Split('/').Length == 5)
            {
                 genColName = "رقم_اذن_الدفن";
            }

            foreach (DataRow row in dtbl.Rows)
            {
                if (form == 12)
                {
                    rowCnt = (Convert.ToInt32(row[genColName].ToString().Split('/')[4]) + 1).ToString();
                }

                else if (form < 12)
                {
                    if (row[genColName].ToString().Split('/').Length == 4)
                        rowCnt = (Convert.ToInt32(row[genColName].ToString().Split('/')[3]) + 1).ToString();
                    else if (row[genColName].ToString().Split('/').Length == 5)
                        rowCnt = (Convert.ToInt32(row[genColName].ToString().Split('/')[4]) + 1).ToString();
                }

                else if (form == 13)// && row["رقم_المعاملة"].ToString().Split('/').Length == 5)
                {
                    rowCnt = getNumber();
                }
                else if (form == 15)// && row["رقم_المعاملة"].ToString().Split('/').Length == 5)
                {
                    rowCnt = (Convert.ToInt32(row[genColName].ToString().Split('/')[4]) + 1).ToString();
                }
                
                else if (form == 16)// && row["رقم_المعاملة"].ToString().Split('/').Length == 5)
                {
                    rowCnt = (Convert.ToInt32(row[genColName].ToString().Split('/')[4]) + 1).ToString();
                }
                else if (form == 17)// && row["رقم_المعاملة"].ToString().Split('/').Length == 5)
                {
                    rowCnt = (Convert.ToInt32(row[genColName].ToString().Split('/')[4]) + 1).ToString();
                }

                
                
            }
            return rowCnt;

        }


        private string OpenFile(string documenNo, int fileNo, string table)
        {
            string str = "";
            string query;


            SqlConnection Con = new SqlConnection(DataSource);
            query = "SELECT ID, Data1, Extension1,data1,طريقة_الطلب,اسم_المندوب from TableAuth where رقم_التوكيل=@رقم_التوكيل";
            SqlCommand sqlCmd1 = new SqlCommand(query, Con); 
            if (fileNo == 12)
            {
                
                sqlCmd1.Parameters.Add("@رقم_التوكيل", SqlDbType.NVarChar).Value = documenNo;
            }
            else
            {
                query = "SELECT ID, Data1, Extension1,FileName1,DataInterType,DataMandoubName  from " + table + " where DocID=@DocID";
                sqlCmd1 = new SqlCommand(query, Con);
                sqlCmd1.Parameters.Add("@DocID", SqlDbType.NVarChar).Value = documenNo;
            }
            
            
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                if (fileNo == 12)
                {
                    if (reader["طريقة_الطلب"].ToString() == "حضور مباشرة إلى القنصلية")
                    {
                        mandoubName.SelectedIndex = 0;
                    }else 
                        mandoubName.Text = reader["اسم_المندوب"].ToString();
                    str = reader["data1"].ToString();
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                    CurrentFile = PrimariFiles +  str.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                    //File.WriteAllBytes(CurrentFile, Data);
                    FileIDNo =reader["ID"].ToString();
                    //System.Diagnostics.Process.Start(CurrentFile);
                }
                else {
                    if (reader["DataInterType"].ToString() == "حضور مباشرة إلى القنصلية")
                    {
                        mandoubName.SelectedIndex = 0;
                    }
                    else mandoubName.Text = reader["DataMandoubName"].ToString();
                    str = reader["FileName1"].ToString();
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                    CurrentFile = PrimariFiles + str.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                    //File.WriteAllBytes(CurrentFile, Data);
                    FileIDNo = reader["ID"].ToString();
                    //System.Diagnostics.Process.Start(NewFileName);
                }
                
            }
            Con.Close();
            return str;
        }
        
        
        private void getZeroID(string colName, string table, string text)
        {
            string str = "";
            string query;

            if (colName == "") return;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "select ID from "+ table+" where "+ colName+" = N'" + text + "'";
            //MessageBox.Show(query);
            SqlCommand sqlCmd1 = new SqlCommand(query, Con); 
            
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                FileIDNo = reader["ID"].ToString();
            }
            Con.Close();
        }

        private void deleteRowsData(string v1, string v2, string colName)
        {
            string query;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "DELETE FROM " + v2 + " where "+ colName+" = N'"+ v1+"'";
            Console.WriteLine(query); 
            //MessageBox.Show(query);
            
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            Con.Close();
        }
         private void deleteArchRowsData(string v1, string v2, string colName)
        {
            string dataSource = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");
            string query;
            SqlConnection Con = new SqlConnection(dataSource);
            query = "DELETE FROM " + v2 + " where "+ colName+" = N'"+ v1+"'";
            Console.WriteLine(query); 
            //MessageBox.Show(query);
            
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            Con.Close();
        }


       
        private string checkBasicInfo(string documenNo)
        {
            string appName = "";
            string query;
            PreArchieved = false;
            smsActiviated = false;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "Select ID," + allInsertNamesList[0] + "," + allInsertNamesList[5] + "," + allInsertNamesList[6] + "," + allInsertNamesList[7] + " from " + TableList + " where " + allInsertNamesList[2] + "=@" + allInsertNamesList[2];
            if(allInsertNamesList[5] == ""||allInsertNamesList[6] == "")
                query = "Select ID," + allInsertNamesList[0] + "," + allInsertNamesList[7] + " from " + TableList + " where " + allInsertNamesList[2] + "=@" + allInsertNamesList[2];
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@" + allInsertNamesList[2], SqlDbType.NVarChar).Value = documenNo;
            //MessageBox.Show("query "+query);
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            try
            {
                var reader = sqlCmd1.ExecuteReader();
                if (reader.Read())
                {
                    FileIDNo = reader["ID"].ToString();
                    if (allInsertNamesList[5] != "" && allInsertNamesList[6] != "")
                    {
                        if (reader[allInsertNamesList[6]].ToString() == "حضور مباشرة إلى القنصلية")
                        {
                            mandoubName.SelectedIndex = 0;
                        }
                        else
                        {
                            mandoubName.Text = reader[allInsertNamesList[5]].ToString();
                            if (reader[allInsertNamesList[7]].ToString() == "حضور مباشرة إلى القنصلية")
                                finalArch = false;
                        }
                    }
                    appName = reader[allInsertNamesList[0]].ToString();
                    updateGenName(appName, documenNo, TableList);
                    PreArchieved = true;
                    Con.Close();
                    //MessageBox.Show(appName);
                    return appName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(query);
            }
            return appName;
        }
        //private string OpenFile2(string documenNo, int fileNo, string table)
        //{
        //    string str = "";
        //    string query;
        //    PreArchieved = false;
        //    smsActiviated = false;
        //    SqlConnection Con = new SqlConnection(DataSource);
        //    query = "SELECT ID,مقدم_الطلب,طريقة_الطلب,اسم_المندوب from TableAuth where رقم_التوكيل=@رقم_التوكيل";
        //    SqlCommand sqlCmd1 = new SqlCommand(query, Con);
        //    if (fileNo == 12)
        //    {

        //        sqlCmd1.Parameters.Add("@رقم_التوكيل", SqlDbType.NVarChar).Value = documenNo;
        //    }
        //    else if (fileNo == 13)
        //    {
        //        query = "SELECT ID, مقدم_الطلب  ,طريقة_الطلب  ,اسم_المندوب,نوع_المعاملة,رقم_الملف,رقم_هاتف1,sms from " + table + " where رقم_المعاملة=@رقم_المعاملة";
        //        sqlCmd1 = new SqlCommand(query, Con);
        //        sqlCmd1.Parameters.Add("@رقم_المعاملة", SqlDbType.NVarChar).Value = documenNo;
               
        //    }
        //    else if (fileNo == 15)
        //    {
        //        query = "SELECT ID, اسم_الزوج,sms,هاتف_الزوج  from " + table + " where رقم_المعاملة=@رقم_المعاملة";
        //        sqlCmd1 = new SqlCommand(query, Con);
        //        sqlCmd1.Parameters.Add("@رقم_المعاملة", SqlDbType.NVarChar).Value = documenNo;
               
        //    }
        //    else if (fileNo == 16)
        //    {
        //        query = "SELECT ID, اسم_المتوفى,sms  from " + table + " where رقم_اذن_الدفن=@رقم_اذن_الدفن";
        //        sqlCmd1 = new SqlCommand(query, Con);
        //        sqlCmd1.Parameters.Add("@رقم_اذن_الدفن", SqlDbType.NVarChar).Value = documenNo;
               
        //    }
        //    else if (fileNo == 17)
        //    {
        //        query = "SELECT ID, اسم_الزوج,sms  from " + table + " where رقم_المعاملة=@رقم_المعاملة";
        //        sqlCmd1 = new SqlCommand(query, Con);
        //        sqlCmd1.Parameters.Add("@رقم_المعاملة", SqlDbType.NVarChar).Value = documenNo;
               
        //    }
        //    else
        //    {
        //        query = "SELECT ID, AppName,DataInterType,DataMandoubName from " + table + " where DocID=@DocID";
        //        sqlCmd1 = new SqlCommand(query, Con);
        //        sqlCmd1.Parameters.Add("@DocID", SqlDbType.NVarChar).Value = documenNo;
                
        //    }

            
        //    if (Con.State == ConnectionState.Closed)
        //        Con.Open();

        //    var reader = sqlCmd1.ExecuteReader();
        //    if (reader.Read())
        //    {                
        //        FileTableID = fileNo;
        //        FileIDNo = reader["ID"].ToString();
        //        FileTable = table;
        //        if (fileNo == 12)
        //        {
        //            if (reader["طريقة_الطلب"].ToString() == "حضور مباشرة إلى القنصلية")
        //            {
        //                mandoubName.SelectedIndex = 0;
        //            }
        //            else mandoubName.Text = reader["اسم_المندوب"].ToString();
        //            str = reader["مقدم_الطلب"].ToString();
        //            PreArchieved = true;
        //            CurrentFile = "";
        //            return str;
        //        }
        //        else if (fileNo == 13)
        //        {
        //            //MessageBox.Show(reader["رقم_الملف"].ToString());
        //            if (reader["رقم_الملف"].ToString() == "99" && reader["sms"].ToString() != "done")
        //            {
        //                //MessageBox.Show(reader["رقم_هاتف1"].ToString());
        //                smsActiviated = true;
        //                smsPhoneNo = reader["رقم_هاتف1"].ToString();
        //            }
        //            if (reader["طريقة_الطلب"].ToString() == "حضور مباشرة إلى القنصلية")
        //            {
        //                mandoubName.SelectedIndex = 0;
        //            }
        //            else mandoubName.Text = reader["اسم_المندوب"].ToString();
        //            smsName = str = reader["مقدم_الطلب"].ToString();
                    
        //            CurrentFile = "";
        //            PreArchieved = true;
        //            return str;
        //        }
                
        //        else if (fileNo == 15)
        //        {
        //            smsActiviated = false; 
        //            //if (reader["sms"].ToString() != "done")
        //            //{
        //            //    smsActiviated = true;
        //            //    smsPhoneNo = reader["هاتف_الزوج"].ToString();
        //            //    smsName = reader["اسم_الزوج"].ToString();
        //            //}
        //            str = reader["اسم_الزوج"].ToString();
        //            PreArchieved = true;
        //        }
                
        //        else if (fileNo == 16)
        //        {                    
        //            str = reader["اسم_المتوفى"].ToString();                    
        //        }
                
        //        else if (fileNo == 17)
        //        {
        //            smsActiviated = false;
        //            //if (reader["sms"].ToString() != "done")
        //            //{
        //            //    smsActiviated = true;
        //            //    smsPhoneNo = reader["هاتف_الزوج"].ToString();
        //            //    smsName = reader["اسم_الزوج"].ToString();
        //            //}
        //            str = reader["اسم_الزوج"].ToString();
        //            PreArchieved = true;
        //        }
        //        else
        //        {
        //            if (reader["DataInterType"].ToString() == "حضور مباشرة إلى القنصلية")
        //            {
        //                mandoubName.SelectedIndex = 0;
        //            }
        //            else mandoubName.Text = reader["DataMandoubName"].ToString();

        //            str = reader["AppName"].ToString();
        //            CurrentFile = "";
        //            PreArchieved = true;
        //            return str;
        //        }

        //    }
        //    Con.Close();
        //    if (CurrentFile.Contains("text")) CurrentFile = "";
        //    return str;
        //}

        private string loadFile(string documenNo, int form, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT Data1, Extension1,data1 from TableAuth where رقم_التوكيل=@رقم_التوكيل", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_التوكيل", documenNo);

            if (form != 12)
            {
                sqlDa = new SqlDataAdapter("SELECT Data1, Extension1,FileName1  from " + table + " where DocID=@DocID", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                sqlDa.SelectCommand.Parameters.AddWithValue("@DocID", documenNo);
            }
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string rowCnt = "معاملة غير موجودة";

            foreach (DataRow row in dtbl.Rows)
            {
                if (form == 12) rowCnt = row["مقدم_الطلب"].ToString();

                else rowCnt = row["AppName"].ToString();
            }
            return rowCnt;

        }
        private string checkArchCase(string documenNo)
        {
            string text = "";
            string query = "SELECT " + allUpdateNamesList[0] +" FROM " + TableList + " WHERE " + allUpdateNamesList[2] + "=N'" + documenNo + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            //M
            Console.WriteLine(query);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
                sqlCon.Close();
                foreach (DataRow row in dtbl.Rows)
                {
                    text = row[allUpdateNamesList[0]].ToString();
                }
            }
            catch (Exception ex) { }
            return text;
        }
        private void getCom(string documenNo)
        {
            try
            {
                string text = "";
            //MessageBox.Show(getComment);
            string query = "SELECT " + getComment+" FROM " + TableList + " WHERE " + allUpdateNamesList[2] + "=N'" + documenNo + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            //M
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            
                sqlDa.Fill(dtbl);
                sqlCon.Close();
                foreach (DataRow row in dtbl.Rows)
                {
                    التعليقات_السابقة_Off.Text = row[getComment].ToString();
                    //MessageBox.Show(row[getComment].ToString());
                }
            }
            catch (Exception ex) { }
        }
        



        private string loadName(string documenNo, int form, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT ID,مقدم_الطلب,طريقة_الطلب,اسم_المندوب from TableAuth where رقم_التوكيل=@رقم_التوكيل", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_التوكيل", documenNo);

            if (form == 13) {
                sqlDa = new SqlDataAdapter("SELECT ID, مقدم_الطلب,طريقة_الطلب,اسم_المندوب from " + table + " where رقم_المعاملة=@رقم_المعاملة", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_المعاملة", documenNo);
            }
            else if (form == 15) {
                sqlDa = new SqlDataAdapter("SELECT ID, اسم_الزوج from " + table + " where رقم_المعاملة=@رقم_المعاملة", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_المعاملة", documenNo);
            }
            else if(form < 12)
            {
                sqlDa = new SqlDataAdapter("SELECT ID, AppName,DataInterType,DataMandoubName from " + table + " where DocID=@DocID", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                sqlDa.SelectCommand.Parameters.AddWithValue("@DocID", documenNo);
            }
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string rowCnt = "معاملة غير موجودة";

            foreach (DataRow row in dtbl.Rows)
            {
                if (form == 12)
                {
                    rowCnt = row["مقدم_الطلب"].ToString();
                    if (row["طريقة_الطلب"].ToString() == "حضور مباشرة إلى القنصلية")
                    {
                        mandoubName.SelectedIndex = 0;
                    }
                    else mandoubName.Text = row["اسم_المندوب"].ToString();
                    FileIDNo = row["ID"].ToString();
                }
                else if (form == 13) {
                    rowCnt = row["مقدم_الطلب"].ToString();
                    if (row["طريقة_الطلب"].ToString() == "حضور مباشرة إلى القنصلية")
                    {
                        mandoubName.SelectedIndex = 0;
                    }
                    else mandoubName.Text = row["اسم_المندوب"].ToString();
                    FileIDNo = row["ID"].ToString();
                }
                else if (form == 15) {
                    rowCnt = row["اسم_الزوج"].ToString();
                    FileIDNo = row["ID"].ToString();
                }
                else
                {
                    rowCnt = row["AppName"].ToString();
                    if (row["DataInterType"].ToString() == "حضور مباشرة إلى القنصلية")
                    {
                        mandoubName.SelectedIndex = 0;
                    }
                    else mandoubName.Text = row["DataMandoubName"].ToString();
                    FileIDNo = row["ID"].ToString();
                }
            }
            if (string.IsNullOrEmpty(rowCnt))
            {
                rowCnt = "معاملة غير موجودة";
                FileIDNo = "0";
            }
            return rowCnt;

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


        bool printOut = true;
        private void CreatePic(string[] location,string[] name)
        {
            printOut = true;
            if (ArchMandoubID != "")
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

                            insertDocMand(ArchMandoubID, GregorianDate, EmpName, DataSource, extn1, name[x], "data2", buffer1);
                            //Console.WriteLine(docid);
                        }
                    }
                }
                this.Close();
            }
            if(MandArch) return;
            if (ArchiveState && newEntry)
            {
                int docid = NewReportEntry(DataSource, GenQuery, FormType, Combo1.Text.Trim(), Combo2.Text.Trim());
                if (docid == 0) {
                    MessageBox.Show("عملية غير صالحة .. تعذر المتابعة، في حالة تكرار الرسالة يرجى إخطار مشغل البرنامج"); 
                    return; 
                }
                //Console.WriteLine(-1);
                for (int x = 0; x < imagecount; x++)
                {
                    //MessageBox.Show(location[x]);
                    if (location[x] != "")
                    {
                        using (Stream stream = File.OpenRead(location[x]))
                        {
                            byte[] buffer1 = new byte[stream.Length];
                            stream.Read(buffer1, 0, buffer1.Length);
                            var fileinfo1 = new FileInfo(location[x]);
                            string extn1 = fileinfo1.Extension;
                            string DocName1 = fileinfo1.Name;
                            
                            insertDoc("", docid.ToString(), GregorianDate, EmpName, DataSource, extn1, name[x], AuthNoPart1, "data1", buffer1);
                            //Console.WriteLine(docid);
                        }
                    }
                }
                UpdateComment(AuthNoPart1);
                if (Payment != "")
                    skipPayment(TableList, Payment, docid.ToString());
            }
            else if (ArchiveState && !newEntry)
            {
                UpdateComment(docIDNumber);
                if (docIDNumber == "")
                {
                    
                    return;
                }

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
                            insertDoc("",FileIDNo, GregorianDate, EmpName, DataSource, extn1, name[x], docIDNumber, "data1", buffer1);
                            //Console.WriteLine(docIDNumber);
                        }
                    }
                }
                
            }
            else if (!ArchiveState && (archCase == 2 || archCase == 3 || (AppNameText != "" && AppNameText != "اكتب اسم مقدم الطلب"))) {
                //MessageBox.Show("archCase = " + archCase.ToString());
                //"تم إصدارالمكاتبة النهائية باسم ولكن لم تتم أرشفتها بعد";
                UpdateComment(docIDNumber);
                if (docIDNumber == "")
                {
                    
                    return;
                }
                if (FileIDNo == "0")
                    if (FormType != 13)
                        getZeroID(columnList, TableList, docIDNumber);

                FinalDataArch(DataSource, docIDNumber);
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
                            insertDoc(relatedRefNo,getProcedName, FileIDNo, GregorianDate, EmpName, DataSource, extn1, name[x], docIDNumber, "data2", buffer1);
                        }
                    }
                }
                
                updateNames();

                if (mandoubName.Text == "حضور مباشرة إلى القنصلية" || ServerType == "56" || SpecificDigit(docId.Text, 3, 4) == "06")
                {
                    deleteRowsData(docIDNumber, "archives", "docID");
                    MessageBox.Show("تمت الأرشفة النهائية " + mandoubName.Text+" - "+ ServerType +" - "+ SpecificDigit(docId.Text, 3, 4));
                }
                
                else if (mandoubName.Text != "حضور مباشرة إلى القنصلية" && mandoubName.Text.Contains( "المندوب"))
                {
                    UpdateMandoubState(txtIDNo.Text, "appOldNew", "في انتظار نسخة المواطن");
                    MessageBox.Show("تمت الأرشفة وفي انتظار نسخة المواطن بعد البصمة");
                }
                
            }
            else if ((archCase == 4||archCase == 5) && !ArchiveState)
            {
                //MessageBox.Show("archCase = " + archCase.ToString());
                UpdateComment(docIDNumber);
                if (docIDNumber == "")
                {
                    
                    return;
                }
                if (FileIDNo == "0")
                    if (FormType != 13)
                        getZeroID(columnList, TableList, docIDNumber);
                FinalDataArch(DataSource, docIDNumber);
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
                            insertDoc(relatedRefNo,getProcedName, FileIDNo, GregorianDate, EmpName, DataSource, extn1, name[x], docIDNumber, "data2", buffer1);
                        }
                    }
                }
                

                CurrentFile = "";
                updateNames();

                deleteRowsData(docIDNumber, "archives", "docID");
                MessageBox.Show("تمت الأرشفة النهائية");
                if (mandoubName.Text != "حضور مباشرة إلى القنصلية" && mandoubName.Text != "")
                {
                    int found = todayList(mandoubName.Text.Trim(), Labdate);
                    if (found > 0)
                        MessageBox.Show("المندوب لدية عدد " + found.ToString() + " مكاتبات غير مكتملة،،، في انتظار المتبقي من المعاملات...");
                }
            }
        }

        

        private string checkArch(string documenNo, string name)
        {
            string dataSource = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select ID,نوع_المستند,التاريخ,الاسم,المستند,رقم_المرجع,رقم_المرجع_المرتبط_off from TableGeneralArch where رقم_معاملة_القسم=@رقم_معاملة_القسم or رقم_المرجع_المرتبط_off=@رقم_المرجع_المرتبط_off", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_معاملة_القسم", documenNo);
            sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_المرجع_المرتبط_off", documenNo);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            bool data1check = false;
            bool data2check = false;
            bool data3check = false;
            string[] id1List = new string[100];
            string[] id2List = new string[100];
            string[] data1List = new string[100];
            int index1 = 0;
            string[] data2List = new string[100];
            int index2 = 0;
            //MessageBox.Show(documenNo);
            foreach (DataRow row in dtbl.Rows)
            {
                data3check = true;
                //name = row["الاسم"].ToString();
                refNo = row["رقم_المرجع"].ToString();
                relatedRefNo = row["رقم_المرجع_المرتبط_off"].ToString();
                //MessageBox.Show(name);MessageBox.Show(relatedRefNo);
                //if (name != "")
                //{

                //}

                if (row["نوع_المستند"].ToString() == "data1")
                {
                    data1check = true;
                    data1List[index1] = row["المستند"].ToString();
                    id1List[index1] = row["ID"].ToString();
                    index1++;
                }

                if (row["نوع_المستند"].ToString() == "data2")
                {
                    data2check = true;
                    data2List[index2] = row["المستند"].ToString();
                    //MessageBox.Show(id2List[index2]);
                    id2List[index2] = row["ID"].ToString();
                    index2++;
                }
            }

            if (data1check)
            {
                drawBoxesTitle("المستندات الأولية للإجراء", 60);
                for (int index = 0; index < index1; index++)
                    drawBoxes(data1List[index], false, id1List[index]);
            }

            if (data2check)
            {
                drawBoxesTitle("------------------------", 60);
                drawBoxesTitle("المكاتبات النهائية من طرف القنصلية العامة", 20);
                for (int index = 0; index < index2; index++)
                    drawBoxes(data2List[index], false, id2List[index]);
                drawBoxes("أرشفة مستندات أخرى", true, "");
                btnSaveEnd.Visible = panelFinalArch.Visible = false;
                btnSaveEnd.Location = new System.Drawing.Point(736, 668);//759, 660
            }
            else if (name != "")
            {
                btnSaveEnd.Visible = panelFinalArch.Visible = false;
                btnSaveEnd.Location = new System.Drawing.Point(736, 668);
                drawBoxesTitle("أرشفة المكاتبات النهائية", 20);
                drawBoxes("استمارة الطلب بعد البصمة (أخرى)", true, "");
                if(noForm == "15")
                        drawBoxes("وثيقة عقد النكاح", true, "");
                if(noForm == "17")
                        drawBoxes("وثيقة اثبات الطلاق", true, "");
                else
                        drawBoxes("المكاتبة النهائية ", true, "");
                if (noForm == "15" || noForm == "17")
                    drawBoxes("الإيصال المالي (أخرى)", true, "");
                drawBoxes("أرشفة مستندات أخرى", true, "");
            }
            else if (name == "" && data3check)
            {
                var selectedOption = MessageBox.Show("تجاوز؟", "بيانات الطلب غير مكتملة", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                {
                    btnSaveEnd.Visible = panelFinalArch.Visible = false;
                    drawNameBox();

                    btnSaveEnd.Location = new System.Drawing.Point(736, 668);
                    drawBoxesTitle("أرشفة المكاتبات النهائية", 20);
                    drawBoxes("استمارة الطلب بعد البصمة", true, "");
                    if (noForm == "15")
                        drawBoxes("وثيقة عقد النكاح", true, "");
                    if (noForm == "17")
                        drawBoxes("وثيقة اثبات الطلاق", true, "");
                    else
                        drawBoxes("المكاتبة النهائية ", true, "");
                    if (noForm == "15" || noForm == "17")
                        drawBoxes("الإيصال المالي", true, "");
                    drawBoxes("أرشفة مستندات أخرى", true, "");
                }
            }

            //if (name == "مؤرشف نهائي")
            //{
            //    //
            //    requiredDocument.Size = new System.Drawing.Size(308, 85);
            //    requiredDocument.Enabled = true; nameSave.Visible = true; 
            //    return name; 
            //}

            if (ServerType == "56")
            {
                btnSaveEnd.Visible = panelFinalArch.Visible = true;
                commentPanel.Visible = false;
            }
            else commentPanel.Visible = true;

            //MessageBox.Show(name);
            //MessageBox.Show(checkArchCase(documenNo));

            getCom(documenNo);
            if (name == "" && data3check)
            {
                archCase = 1;
                return "المستندات مؤرشفة مبدئيا ولكن لم يتم إدخال بيانات مقدم الطلب برقم المعاملة " + documenNo;
            }
            else if ((data1check && !data2check) && name != "")
            {
                archCase = 2; return "تم إصدارالمكاتبة النهائية باسم /" + name + " ولكن لم تتم أرشفتها بعد";
            }
            else if (data2check && checkArchCase(documenNo).Contains("مؤرشف نهائي_"))
            {
                archCase = 5; return "تم إصدارالمكاتبة باسم /" + name + " وقد تمت أرشفتها بصورة نهائية وفي انتظار نسخة المواطن بعد البصمة";
            }
            else if (data2check && checkArchCase(documenNo) != "مؤرشف نهائي")
            {
                archCase = 3; return "تم إصدارالمكاتبة باسم /" + name + " وقد تمت أرشفتها بصورة نهائية وقد تمت إعادة طباعتها مجددا";
            }
            else if (data2check && checkArchCase(documenNo) == "مؤرشف نهائي")
            {
                archCase = 4; return "تم إصدارالمكاتبة باسم /" + name + " وقد تمت أرشفتها بصورة نهائية";
            }
            
            else
            {
                archCase = 0; return "لا يوجد بالنظام معاملة بالرقم " + documenNo;
            }
            
        }
        private void UpdateMandoubState(string id, string col,string text)
        {
            //sqlCmd.Parameters.AddWithValue("@appOldNew", "في انتظار نسخة المواطن");
            string qurey = "update archives set appOldNew=@appOldNew where docID=@docID";

            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@comment ", id);
            sqlCmd.Parameters.AddWithValue("@"+ col, text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void updateGenName(string name, string idDoc)
        {
            string dataSource = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");
            SqlConnection sqlCon = new SqlConnection(dataSource);

            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string query = "update TableGeneralArch set الاسم=N'" + name + "', Pass = N'تجاوز الادخال' where رقم_معاملة_القسم = '" + idDoc + "'";
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        
        private void updateAppName(string name, string idDoc, string table)
        {
            string dataSource = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");
            SqlConnection sqlCon = new SqlConnection(dataSource);

            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string query = "update TableGeneralArch set الاسم=N'" + name + "', Pass = N'تجاوز الادخال' where رقم_المرجع = '" + idDoc + "' and docTable = '" + table + "'";
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void UpdateComment(string documenNo)
        {
            //sqlCmd.Parameters.AddWithValue("@appOldNew", "في انتظار نسخة المواطن");
            commentInfo();
            string query = "SELECT " + getComment + " FROM " + TableList ;

            string qurey = "update " + TableList + " set " + getComment + "=@comment WHERE " + allUpdateNamesList[2] + " = N'" + documenNo + "'";
            //MessageBox.Show(qurey);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@comment", تعليق.Text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        private void UpdateState(int id, string col,string text, string table)
        {
            //sqlCmd.Parameters.AddWithValue("@appOldNew", "في انتظار نسخة المواطن");
            string qurey = "update "+ table+" set "+ col+"=@"+ col+" where ID=@id";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@"+ col, text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void deleteRowsData(string v1, string colName)
        {
            string query;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "DELETE FROM archives where "+ colName+" = @docID";
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@"+ colName, v1);
            sqlCmd.ExecuteNonQuery();
            Con.Close();
        }

        private void saveToDatabase(string filePath1)
        {
            //MessageBox.Show(DataSource);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableAuth (Data1, Extension1, data1 ,إجراء_التوكيل,نوع_التوكيل) values (@Data1, @Extension1, @data1,@إجراء_التوكيل,@نوع_التوكيل) ", sqlCon);
            //SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableAuth (Data2, Extension2, data2) values(@Data2, @Extension2, @data2) ", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", 1);
            sqlCmd.Parameters.AddWithValue("@إجراء_التوكيل", Combo2.Text);
            sqlCmd.Parameters.AddWithValue("@نوع_التوكيل", Combo1.Text.Trim());

            using (Stream stream = File.OpenRead(filePath1))
            {
                byte[] buffer1 = new byte[stream.Length];
                stream.Read(buffer1, 0, buffer1.Length);
                var fileinfo1 = new FileInfo(filePath1);
                string extn1 = fileinfo1.Extension;
                string DocName1 = fileinfo1.Name;
                sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
                sqlCmd.Parameters.Add("@data1", SqlDbType.NVarChar).Value = DocName1;
            }
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        string rowCountstr(String DataSource, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select ID from " + table, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            return dtbl.Rows.Count.ToString();

        }
        private void SMS(int id, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select MandoubPhones,الصفة from TableMandoudList", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string[] phoneNo = new string[10] { "","", "", "" , "", "" , "", "" , "", "" };
            int i = 0;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["الصفة"].ToString().Contains("قسم شؤون الرعايا"))
                {
                    if (!dataRow["الصفة"].ToString().Contains("*"))
                    {
                        string smsText = "تم إصدار خطاب حالة خاصة لإجراء تسهيل سفر بالرقم " + smsDocIDNumber + " للمواطن/ " + smsName + " بتاريخ:" + GregorianDate + " يرجى استلام المعاملة من القسم وإجراء ما يلزم مع الشكر";
                        SendSms(dataRow["MandoubPhones"].ToString(), smsText);
                        UpdateState(id, "sms", "sent", table);
                    }
                    else
                    {
                        string smsText = "تم إصدار خطاب حالة خاصة لإجراء تسهيل سفر بالرقم " + smsDocIDNumber + " للمواطن/ " + smsName + " بتاريخ:" + GregorianDate;
                        //MessageBox.Show(dataRow["MandoubNames"].ToString());
                        SendSms(dataRow["MandoubPhones"].ToString(), smsText);
                        UpdateState(id, "sms", "sent", table);
                    }
                }
                else if (dataRow["الصفة"].ToString().Contains("قسم الأحوال الشخصية"))
                {
                    //MessageBox.Show(dataRow["MandoubPhones"].ToString() + " - قسم الأحوال الشخصية");
                    string smsText = "تم إنهاء معاملة قسيمة زواج برقم معاملة " + smsDocIDNumber + " للمواطن/ " + smsName + " بتاريخ:" + GregorianDate;
                    
                    SendSms(dataRow["MandoubPhones"].ToString(), smsText);

                    SendSms(smsPhoneNo, smsText);
                    UpdateState(id, "sms", "sent", table);
                }
            }
            
        }
        public void FillDataGridView(String DataSource)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("AuthViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@مقدم_الطلب", "");
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            rowCount = dataGridView1.Rows.Count.ToString();
            sqlCon.Close();

        }

        public void mandoubFiles()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select EnterySheet from TableListCombo", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView2.DataSource = dtbl;
            dataGridView2.Columns[0].Width = 400;
            sqlCon.Close();

        }

        private void insertDoc(string relatedRefNo, string name, string id, string date, string employee, string dataSource, string extn1, string DocName1, string messNo, string docType, byte[] buffer1)
        {
            string query = "INSERT INTO TableGeneralArch (رقم_المرجع_المرتبط_off,Data1,Extension1,نوع_المستند,رقم_معاملة_القسم,المستند,الموظف,التاريخ,رقم_المرجع,docTable,الاسم) values (@رقم_المرجع_المرتبط_off,@Data1,@Extension1,@نوع_المستند,@رقم_معاملة_القسم,@المستند,@الموظف,@التاريخ,@رقم_المرجع,@docTable,@الاسم)";
            dataSource = dataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");
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
            sqlCmd.Parameters.AddWithValue("@رقم_المرجع_المرتبط_off", relatedRefNo);
            sqlCmd.Parameters.AddWithValue("@الاسم", name);
            sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
            sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
            sqlCmd.Parameters.Add("@المستند", SqlDbType.NVarChar).Value = DocName1;
            sqlCmd.Parameters.Add("@docTable", SqlDbType.NVarChar).Value = TableList;            
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        private void insertDoc(string relatedRefNo,string id, string date, string employee, string dataSource, string extn1, string DocName1, string messNo, string docType, byte[] buffer1)
        {
            dataSource = dataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");
            string query = "INSERT INTO TableGeneralArch (رقم_المرجع_المرتبط_off,Data1,Extension1,نوع_المستند,رقم_معاملة_القسم,المستند,الموظف,التاريخ,رقم_المرجع,docTable) values (@رقم_المرجع_المرتبط_off,@Data1,@Extension1,@نوع_المستند,@رقم_معاملة_القسم,@المستند,@الموظف,@التاريخ,@رقم_المرجع,@docTable)";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            Console.WriteLine();
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", messNo);
            sqlCmd.Parameters.AddWithValue("@نوع_المستند", docType);
            sqlCmd.Parameters.AddWithValue("@الموظف", employee);
            sqlCmd.Parameters.AddWithValue("@التاريخ", date);
            sqlCmd.Parameters.AddWithValue("@رقم_المرجع", id);
            sqlCmd.Parameters.AddWithValue("@رقم_المرجع_المرتبط_off", relatedRefNo);
            sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
            sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
            sqlCmd.Parameters.Add("@المستند", SqlDbType.NVarChar).Value = DocName1;
            sqlCmd.Parameters.Add("@docTable", SqlDbType.NVarChar).Value = TableList;            
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        
        private void insertDocMand(string id, string date, string employee, string dataSource, string extn1, string DocName1, string docType, byte[] buffer1)
        {
            dataSource = dataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");
            string query = "INSERT INTO TableGeneralArch (Data1,Extension1,نوع_المستند,رقم_معاملة_القسم,المستند,الموظف,التاريخ,رقم_المرجع,docTable,الاسم) values (@Data1,@Extension1,@نوع_المستند,@رقم_معاملة_القسم,@المستند,@الموظف,@التاريخ,@رقم_المرجع,@docTable,@الاسم)";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@الاسم", mandoubName.Text.Split('-')[0]);
            sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", mandoubName.Text.Split('-')[1]);
            sqlCmd.Parameters.AddWithValue("@نوع_المستند", docType);
            sqlCmd.Parameters.AddWithValue("@الموظف", employee);
            sqlCmd.Parameters.AddWithValue("@التاريخ", date);
            sqlCmd.Parameters.AddWithValue("@رقم_المرجع", id);
            sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
            sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
            sqlCmd.Parameters.Add("@المستند", SqlDbType.NVarChar).Value = DocName1;
            sqlCmd.Parameters.Add("@docTable", SqlDbType.NVarChar).Value = "TableMandoudList";            
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        
        private void updateName(string name,  string messNo,  string table)
        {
            string dataSource = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");
            string query = "update TableGeneralArch set الاسم=@الاسم where رقم_المرجع=@رقم_المرجع and docTable=@docTable";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_المرجع", messNo);
            sqlCmd.Parameters.AddWithValue("@docTable", table);
            sqlCmd.Parameters.AddWithValue("@الاسم", name);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private bool checkFormName(string itemName, string col)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select " + col + " from TableListCombo where " + col + "=N'" + itemName + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count > 0) return true;
            return false;
        }

        private void updatetproFormRow(string id, string source, string filePath)
        {
            var fileinfo1 = new FileInfo(filePath);
            
            SqlConnection sqlCon = new SqlConnection(source);
            
            string qurey = "UPDATE TableProcReq SET proForm1=N'"+ fileinfo1.Name + "' WHERE ID=@ID";
            if(proForm1Val != "" && proForm2Val == "" && proForm1Val!= proForm2Val)
                qurey = "UPDATE TableProcReq SET proForm2=N'" + fileinfo1.Name + "' WHERE ID=@ID";
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", id);            
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!checkActiv(mandoubName.Text.Split('-')[0].Trim()) && !MandArch)
            {
                MessageBox.Show("حساب المندوب عير مفعل، يرجى التواصل مع مدير القسم للتفعيل");
                return;
            }
            string date = DateTime.Now.Day.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Year.ToString();
            CreatePic(PathImage, ImageName);
            if (!printOut)return;
            if (ArchiveState)
            {
                if (تاريخ_الميلاد.Text.Length != 10)
                {
                    MessageBox.Show("يرجى إدخال تاريخ ميلاد مقدم الطلب أولا");
                    return;
                }
                if (docId.Text == "")
                    return;

                if (FormType != 6)
                {
                    string imageUri = PrimariFiles + @"FormData\" + Combo1.Text.Trim() + ".jpg";
                    string wordInFile = FilespathIn + Combo1.Text.Trim() + ".docx";
                    string wordOutFile = FilespathOut + Combo1.Text.Trim() + DateTime.Now.ToString("ssmm") + ".docx";
                    
                    if (Combo2.Visible && !Combo1.Text.Contains("جامعية") && !Combo1.Text.Contains("ميلاد"))
                    {
                        imageUri = PrimariFiles + @"FormData\" + Combo2.Text.Trim() + ".jpg";
                        wordInFile = FilespathIn + Combo2.Text.Trim() + ".docx";
                        wordOutFile = FilespathOut + Combo2.Text.Trim() + DateTime.Now.ToString("ssmm") + ".docx";
                    }
                    if (FormType >= 12 && FormType < 16)
                    {
                        imageUri = PrimariFiles + @"FormData\" + Combo1.SelectedIndex.ToString() + "-" + Combo2.Text.Trim() + ".jpg";
                        wordInFile = FilespathIn + Combo2.Text.Trim() + "-" + Combo1.SelectedIndex.ToString() + ".docx";
                        wordOutFile = FilespathOut + Combo2.Text.Trim() + DateTime.Now.ToString("ssmm") + ".docx";
                    }
                    else if (FormType == 16)
                    {
                        imageUri = PrimariFiles + @"FormData\" + Combo1.SelectedIndex.ToString() + "-" + Combo2.Text.Trim() + ".jpg";
                        wordInFile = FilespathIn + Combo1.Text.ToString() + ".docx";
                        wordOutFile = FilespathOut + Combo1.Text.Trim() + DateTime.Now.ToString("ssmm") + ".docx";
                    }

                    string SubNo = "02";
                    if (FormType < 10) SubNo = "0" + FormType.ToString();
                    else SubNo = FormType.ToString();


                    if (File.Exists(reqFile))
                    {
                        //updatetproFormRow(proID, DataSource, reqFile);
                        CreateAuth(date.Split('-')[2].Replace("20", "") + SubNo + rowCount + Environment.NewLine + date, reqFile, wordOutFile);
                    }

                    else if (File.Exists(wordInFile))
                    {
                        updatetproFormRow(proID, DataSource, wordInFile);
                        CreateAuth(date.Split('-')[2].Replace("20", "") + SubNo + rowCount + Environment.NewLine + date, wordInFile, wordOutFile);
                    }
                    
                }

            }            
            this.Close();            
            if (AppNameText != "" && AppNameText != "اكتب اسم مقدم الطلب")
                updateAppName(AppNameText, refNo, TableList);

            finalArch = false;
            this.Close();
            btnSaveEnd.Enabled = true;
        }

        private void getPaymentState(string mainCol, string subCol)
        {
            string query = "select payment from TableAddContext where altColName=N'" + mainCol + "' and altSubColName=N'" + subCol + "'";
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
                Payment = reader["payment"].ToString();                
            }
            //MessageBox.Show(Payment);
        }

        private void skipPayment(string table, string state, string intID)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("UPDATE " + table + " SET حالة_السداد =N'" + state + "' where ID = " + intID, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", intID);
            sqlCmd.ExecuteNonQuery();
        }

        private void insertDocMand(string dataSource, string id)
        {
            dataSource = dataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");

            string query = "INSERT INTO TableGeneralArch (رقم_معاملة_القسم) values (@رقم_معاملة_القسم)";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", id);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }


        private void addRelatedArch(string dataSource, string otherID, string relatedDocNo)
        {
            dataSource = dataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");

            string query = "update TableGeneralArch set رقم_المرجع_المرتبط_off=@رقم_المرجع_المرتبط_off where رقم_معاملة_القسم=@رقم_معاملة_القسم";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", relatedDocNo);
            sqlCmd.Parameters.AddWithValue("@رقم_المرجع_المرتبط_off", otherID);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void addRelatedInfo(string dataSource, string table, string otherID, string relatedDocNo, string relatedDocName)
        {
            string query = "update " + table + " set رقم_المرجع_المرتبط_off = @رقم_المرجع_المرتبط_off where "+ relatedDocName+"=@" + relatedDocName+"";
            Console.WriteLine(query);
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@"+ relatedDocName, relatedDocNo);
            sqlCmd.Parameters.AddWithValue("@رقم_المرجع_المرتبط_off", otherID);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void commentInfo()
        {
            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text == "")
                تعليق.Text = "";

            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text != "")
                تعليق.Text = التعليقات_السابقة_Off.Text;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text == "")
                تعليق.Text = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + DateTime.Now.ToString("g") + Environment.NewLine + "--------------" + Environment.NewLine;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text != "")
                تعليق.Text = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + DateTime.Now.ToString("g") + Environment.NewLine + "--------------" + Environment.NewLine + "*" + التعليقات_السابقة_Off.Text.Trim();

        }

        private void Report(string referenceNo, string refNumber, string imageUr)
        {
            LocalReport localReport = new LocalReport();
            string fullpath = PrimariFiles + @"pers\PersAhwal\PersAhwal\Report2.rdlc";
            localReport.ReportPath = fullpath;
            ReportParameterCollection reports = new ReportParameterCollection();
            reports.Add(new ReportParameter("number", referenceNo));
            reports.Add(new ReportParameter("refNumber", refNumber));
            reports.Add(new ReportParameter("image", new Uri(imageUr).AbsoluteUri));
            localReport.EnableExternalImages = true;
            localReport.SetParameters(reports);
            PrintToPrinter(localReport);
            
        }

        public static void PrintToPrinter(LocalReport report)
        {
            Export(report);

        }

        public static void Export(LocalReport report, bool print = true)
        {
            string deviceInfo =
             @"<DeviceInfo>
                <OutputFormat>EMF</OutputFormat>
                <PageWidth>8.3in</PageWidth>
                <PageHeight>11.70in</PageHeight>
                <MarginTop>0in</MarginTop>
                <MarginLeft>0in</MarginLeft>
                <MarginRight>0in</MarginRight>
                <MarginBottom>0in</MarginBottom>
            </DeviceInfo>";
            Warning[] warnings;
            m_streams = new List<Stream>();
            report.Render("Image",
                deviceInfo,
                CreateStream,
                out warnings);
            foreach (Stream stream in m_streams)
                stream.Position = 0;

            if (print)
            {
                Print();
            }
        }


        public static void Print()
        {
            if (m_streams == null || m_streams.Count == 0)
                throw new Exception("Error: no stream to print.");
            PrintDocument printDoc = new PrintDocument();
            if (!printDoc.PrinterSettings.IsValid)
            {
                throw new Exception("Error: cannot find the default printer.");
            }
            else
            {
                printDoc.PrintPage += new PrintPageEventHandler(PrintPage);
                m_currentPageIndex = 0;
                printDoc.Print();
            }
        }

        public static void PrintPage(object sender, PrintPageEventArgs ev)
        {
            //Metafile pageImage = new
            //   Metafile(m_streams[m_currentPageIndex]);

            //// Adjust rectangular area with printer margins.
            //Rectangle adjustedRect = new Rectangle(
            //    ev.PageBounds.Left - (int)ev.PageSettings.HardMarginX,
            //    ev.PageBounds.Top - (int)ev.PageSettings.HardMarginY,
            //    ev.PageBounds.Width,
            //    ev.PageBounds.Height
            //    );

            //// Draw a white background for the report
            //ev.Graphics.FillRectangle(Brushes.White, adjustedRect);

            //// Draw the report content
            //ev.Graphics.DrawImage(pageImage, adjustedRect);

            //// Prepare for the next page. Make sure we haven't hit the end.
            //m_currentPageIndex++;
            //ev.HasMorePages = (m_currentPageIndex < m_streams.Count);
        }


        public static Stream CreateStream(string name, string fileNameExtension, Encoding encoding, string mimeType, bool willSeek)
        {
            Stream stream = new MemoryStream();
            m_streams.Add(stream);
            return stream;
        }

        private void CombAuthType_SelectedIndexChanged(object sender, EventArgs e)
        {
            reqFile = "";
            string[] data = new string[2];
            data[0] = noForm;
            Console.WriteLine(11);
            fileComboBoxSubContext(Combo2, DataSource, "altColName", "altSubColName");
            Console.WriteLine(12);
            if (ArchiveState) DocIDGenerator(FormType);
            Console.WriteLine(13);
            //if (FormType == 10)
            //{
            //    autoCompleteTextBox(Combo2, DataSource, "نوع_الإجراء", "TableCollection");
            //    Console.WriteLine(14);
            //}
            //else if (FormType == 12)
            //{
            //    autoCompleteTextBox(Combo2, DataSource, "إجراء_التوكيل", "TableAuth");
            //    Console.WriteLine(14);
            //}
            return;
           
        }
        private void fileComboBoxSubContext(ComboBox combbox, string source, string comlumnName, string SubComlumnName)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct " + SubComlumnName + " from TableAddContext where " + SubComlumnName + " is not null and " + comlumnName + " =N'" + Combo1.Text + "' order by " + SubComlumnName + " asc";
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
        private void fileComboBoxSub(ComboBox combbox, string source, string comlumnName, string SubComlumnName)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct " + SubComlumnName + " from TableAddContext where " + SubComlumnName + " is not null and " + comlumnName + " =N'" + Combo1.Text + "' order by " + SubComlumnName + " asc";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;

                Console.WriteLine(query);
               // MessageBox.Show(query);
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

        private bool checkaltColName(string colNo)
        {
            string query = "select * from TableAddContext where altColName = N'" + colNo + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            Console.WriteLine(query);
           //MessageBox.Show(query);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count != 0)
                return true;
            else
                return false;
        }

        private string OpenFile(string documenNo, bool printOut)
        {
            string query = "SELECT ID, proForm1,توضيح_المعاملة,Data1, Extension1,المعاملات_المرتبطة from TableProcReq where المعاملة=N'" + documenNo + "'";
            reqFile = "";
            Console.WriteLine(query);
            //MessageBox.Show(query);
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);

            if (Con.State == ConnectionState.Closed)
                Con.Open();
            var reader = sqlCmd1.ExecuteReader();
            توضيح_المعاملة.Visible = false;
            if (reader.Read())
            {
                string str = reader["proForm1"].ToString();
                //MessageBox.Show(str);
                توضيح_المعاملة.Text = reader["توضيح_المعاملة"].ToString();
                relatedPro = reader["المعاملات_المرتبطة"].ToString();
                توضيح_المعاملة.Visible = true;
                توضيح_المعاملة.BringToFront();
                try
                {
                    var Data = (byte[])reader["Data1"];
                    reqFile = FilespathOut + str + DateTime.Now.ToString("ssmm") + ".docx";

                    while (File.Exists(reqFile))
                        reqFile = FilespathOut + str + DateTime.Now.ToString("ssmm") + ".docx";


                    File.WriteAllBytes(reqFile, Data);

                    FileInfo fileInfo = new FileInfo(reqFile);
                    if (fileInfo.IsReadOnly) fileInfo.IsReadOnly = false;
                    Console.WriteLine("reqFile " + reqFile);
                    if (printOut)
                    {
                        System.Diagnostics.Process.Start(reqFile);
                    }
                    return reqFile;

                }
                catch (Exception ex)
                {
                    return "";
                }
            }
            Con.Close();
            return reqFile;
        }
        private void insertRow( string[] data)
        {
            //MessageBox.Show(data[1]);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            //string[] colList = new string[11];
            //colList[0] = "رقم_المعاملة";
            //colList[1] = "المعاملة";
            
            //string item = "رقم_المعاملة";
            string value = "@رقم_المعاملة";
            //for (int col = 1; col < 11; col++)
            //{
            //    item = item + "," + colList[col];
            //    value = value + ",@" + colList[col];
            //}

            string query = "INSERT INTO TableProcReq (المعاملة,رقم_المعاملة) values (@المعاملة,@رقم_المعاملة)";

            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@المعاملة", data[1]);
            sqlCmd.Parameters.AddWithValue("@رقم_المعاملة", data[0]);
            try
            {
                sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
            }
            sqlCon.Close();
        }

        private void fillFormDocx(string formNo, string comb12)
        {
            string wordInFile = FilespathIn + comb12  + ".docx";
            //MessageBox.Show(formNo +" - "+ wordInFile);
            uploadForms(wordInFile, formNo);
        }

        private void uploadForms(string location, string formNo)
        {
            if (location != "" && File.Exists(location))
            {
                using (Stream stream = File.OpenRead(location))
                {
                    byte[] buffer1 = new byte[stream.Length];
                    stream.Read(buffer1, 0, buffer1.Length);
                    var fileinfo1 = new FileInfo(location);
                    string extn1 = fileinfo1.Extension;
                    string DocName1 = fileinfo1.Name;

                    string query = "UPDATE TableProcReq SET Data1=@Data1,Extension1=@Extension1,proForm1=@proForm1 WHERE المعاملة=N'" + formNo + "'";
                    SqlConnection sqlCon = new SqlConnection(DataSource);
                    if (sqlCon.State == ConnectionState.Closed)
                        sqlCon.Open();
                    SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                    sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
                    sqlCmd.Parameters.Add("@proForm1", SqlDbType.NVarChar).Value = formNo;
                    sqlCmd.ExecuteNonQuery();
                    sqlCon.Close();
                }
            }
        }
        private bool PreReqFound(string proName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query = "SELECT رقم_المعاملة FROM TableProcReq where المعاملة=N'" + proName + "'";
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
            sqlCon.Close();
            if(dtbl.Rows.Count > 0)
                return true;            
            else return false;
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
        
        private bool checkItemName(string itemName, string col)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select "+col+" from TableListCombo where " + col+ "=N'"+ itemName + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if(dtbl.Rows.Count > 0) return true;
            return false;
        }
        private string DocIDGenerator(int formT)
        {
            string formtype = "0" + formT.ToString();
            if (formT > 9 && formT < 100)
                formtype = FormType.ToString();
            if (ServerType == "56") formtype = formT + Combo1Index.ToString();
            
            string year = DateTime.Now.Year.ToString().Replace("20", "");
            
            string query = "select max(cast (right("+ ColNumList + ",LEN("+ ColNumList + ") - 15) as int)) as newDocID from "+ TableList+" where " + ColNumList + " like N'"+ txtMissionCode + year + "/" + formtype + "%'";  
            if(ServerType == "56") 
                query = "select max(cast (right("+ ColNumList + ",LEN("+ ColNumList + ") -16) as int)) as newDocID from "+ TableList + " where "+ ColNumList + " like N'"+ txtMissionCode + year + "/" + formtype + "%'";
            Console.WriteLine(query);
            //MessageBox.Show(query);
            rowCount = getUniqueID(query);
            docId.Text = "ق س ج" + "/" + rowCount + "/" + formtype + "/" + year + "/80";
            AuthNoPart1 = txtMissionCode  + year + "/" + formtype + "/" + rowCount;
            AuthNoPart2 = year + formtype + rowCount;            
            return formtype;
        }

        
        private void timer1_Tick(object sender, EventArgs e)
        {
            bool view = true;
            foreach (Control control in drawPicPanel.Controls)
            {
                if (control is Button)
                {
                    if (!control.Enabled && control.Text!= "أرشفة مستندات أخرى" && !control.Text.Contains("أخرى") && control.Visible)
                    {
                        view = false;
                        break;
                    }
                }
            }
            if (ServerType != "56")
                btnSaveEnd.Visible = view;


            if (imagecount > 0)
            {
                btnAuth.Size = new System.Drawing.Size(153, 59);
                btnAuth.Location = new System.Drawing.Point(164, 3);
                loadPic.Size = new System.Drawing.Size(153, 59);
                loadPic.Location = new System.Drawing.Point(164, 69);
                reLoadPic.Visible = button2.Visible = true;
                if (ServerType == "56")
                    btnSaveEnd.Visible = true;
                توضيح_المعاملة.SendToBack();
            }
            else
            {
                
                btnAuth.Location = new System.Drawing.Point(3, 3);
                loadPic.Location = new System.Drawing.Point(3, 69);
                loadPic.Width = btnAuth.Width = 311;
                button2.Visible = false;
                if (ServerType != "56")
                {
                    btnSaveEnd.Visible = view;
                    if (ArchiveState )
                        توضيح_المعاملة.BringToFront();
                }
                else btnSaveEnd.Visible = false;
            }
            if (تاريخ_الميلاد.Text.Length != 10 && ArchiveState)
                btnSaveEnd.Visible = false;
            else if (تاريخ_الميلاد.Text.Length != 10 && !ArchiveState && تاريخ_الميلاد.Visible)
                btnSaveEnd.Visible = false;
        }

        private void CreateAuth(string AuthID, string DocxInFile, string DocxOutFile)
        {
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object objCurrentCopy = DocxInFile;
            Word.Document oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);
            oBMicroWord.Selection.Find.ClearFormatting();
            oBMicroWord.Selection.Find.Replacement.ClearFormatting();           
            
            object ParaAuthIDNo = "MarkAuthIDNo";
            Word.Range BookAuthIDNo = oBDoc.Bookmarks.get_Item(ref ParaAuthIDNo).Range;
            BookAuthIDNo.Text = AuthID;
            object rangeAuthIDNo = BookAuthIDNo;
            oBDoc.Bookmarks.Add("AuthAuthIDNo", ref rangeAuthIDNo);

            string proName = Combo1.Text + " بصيغة عامة";
            if(Combo2.Text != "")
                proName = Combo1.Text + " - " + Combo2.Text;

            try
            {
                object ParaProType = "نوع_الإجراء";
                Word.Range BookProType = oBDoc.Bookmarks.get_Item(ref ParaProType).Range;
                BookProType.Text = proName;
                object rangeProType = BookProType;
                oBDoc.Bookmarks.Add("نوع_الإجراء", ref rangeProType);
            }
            catch (Exception ex) { }
            
            oBDoc.SaveAs2(DocxOutFile);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);

            var selectedOption = MessageBox.Show("تضمين الملفات المؤرشفة بالاستمارة؟", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                using (DocX document = DocX.Load(DocxOutFile))
                {
                    Paragraph p1 = document.InsertParagraph();
                    
                    // Append content to the Paragraph
                    for (int x = 0; x < imagecount; x++)
                    {
                        if (PathImage[x] == "") continue;
                        try
                        {
                            Console.WriteLine (x.ToString() + " - "+PathImage[x]);
                            var image = document.AddImage(PathImage[x]);

                            // Set Picture Height and Width.
                            var picture = image.CreatePicture(610, 510);//ratio 0.615
                            
                            p1.AppendPicture(picture);
                        }catch (Exception ex) {  }
                    }
                    document.Save();
                }
            }
            System.Diagnostics.Process.Start(DocxOutFile);
        }
        

        private void CreateMandoubfile( string text,string DocxInFile, string mandounName,bool copypast)
        {
            string wordOutFile = FilespathOut + Combo1.Text.Trim() + DateTime.Now.ToString("ssmm") + ".docx";            
            string activeCopy = FilespathIn + "\\" + text;
            try
            {
                System.IO.File.Copy(DocxInFile, activeCopy);
            }
            catch (Exception ex) { 
                //
            }
            using (var document = DocX.Load(activeCopy))
            {
                document.AddFooters();
                document.Footers.Odd.InsertParagraph(mandounName).Bold();
                document.Save();
            }
            System.Diagnostics.Process.Start(activeCopy);            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView2.Visible = false;
            panel1.Visible = true;
            try


            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.

                    var ScanerItem = device.Items[1]; // select the scanner.


                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);

                    PathImage[imagecount-1] = PrimariFiles + "ScanImg" + rowCount + (imagecount-1).ToString() + ".jpg";
                    ImageName[imagecount-1] = "ScanImg" + rowCount + (imagecount-1).ToString();


                    if (File.Exists(PathImage[imagecount-1]))
                    {
                        File.Delete(PathImage[imagecount-1]);
                    }
                    imgFile.SaveFile(PathImage[imagecount-1]);
                    pictureBox1.ImageLocation = PathImage[imagecount-1];
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
        
        private void showFiles(object sender, EventArgs e)
        {
            dataGridView2.Visible = false;
            panel1.Visible = true;
            Button button = (Button)sender;
            توضيح_المعاملة.SendToBack();
            if (ArchiveState)
            {
                int picIndex = Convert.ToInt32(button.Name.Split('_')[1]);
                picPath = PathImage[picIndex];
                pictureBox1.ImageLocation = PathImage[picIndex];
            }
            else
            {
                if (!button.Name.All(char.IsDigit))
                {
                    //MessageBox.Show(button.Name);
                    reqFile = "";
                    OpenFile(Combo1Text + "-" + Combo2Text, false); 
                    if (reqFile == "")
                        OpenFile(Combo1Text, false);

                    string wordOutFile = FilespathOut + Combo1.Text.Trim() + DateTime.Now.ToString("ssmm") + ".docx";
                    string date = DateTime.Now.Day.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Year.ToString();
                    string SubNo = "02";
                    if (FormType < 10) SubNo = "0" + FormType.ToString();
                    else SubNo = FormType.ToString();

                    if (File.Exists(reqFile)) 
                        CreateAuth(date.Split('-')[2].Replace("20", "") + SubNo + rowCount + Environment.NewLine + date, reqFile, wordOutFile);
                }
                else
                {
                    picPath = FillDatafromGenArch(button.Name);
                    pictureBox1.ImageLocation = picPath;
                }
            }
           // MessageBox.Show(pictureBox1.ImageLocation.ToString());
        }

        private void Combo2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (noForm != "" &&noForm != "06" && Combo2.Visible)
            {
                //MessageBox.Show("noForm= " + noForm + "_" + Combo2.SelectedIndex.ToString());
                loadPreReq(noForm, Combo1.Text + "-" + Combo2.Text, ArchiveState);
                reqFile = "";
                OpenFile(Combo1.Text + "-" + Combo2.Text, false); 
                if (reqFile == "")
                    OpenFile(Combo1.Text, false);
            }
            getPaymentState(Combo1.Text, Combo2.Text);
        }
       

        private string FillDatafromGenArch( string id)
        {
            //MessageBox.Show(id);
            string dataSource = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");

            string NewFileName = "";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableGeneralArch where  ID='" + id + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow reader in dtbl.Rows)
            {
                var name = reader["المستند"].ToString();
                if (name == "") return "";
                try
                {
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                    NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    //System.Diagnostics.Process.Start(NewFileName);
                }
                catch (Exception ex) { }
            }
            return NewFileName;
        }

        private string getTableWafid(string docid)
        {
            string[] qr = new string[6];
            qr[0] = "TableWafid";
            qr[1] = "TableWafidJed";
            qr[2] = "TableWafidMekkah";
            qr[3] = "TableTarheel";
            qr[4] = "TableTransfer";
            qr[5] = "TableCommity";            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            for (int x = 0; x < 6; x++)
            {
                TableList = qr[x];
                string query = "select نوع_المعاملة from " + TableList + " WHERE رقم_المعاملة=@رقم_المعاملة ";
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_المعاملة", docid);
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);
                sqlCon.Close();
                foreach (DataRow dataRow in dtbl.Rows)
                {
                    if (dataRow["نوع_المعاملة"].ToString() != "")
                    {
                        return x.ToString();// dataRow["نوع_المعاملة"].ToString();
                    }
                }
            }
            TableList = "";
            Console.WriteLine(qr);
            return "-1";

        }

        private void updateNames()
        {
            string dataSource = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");
            SqlConnection sqlCon = new SqlConnection(dataSource);
            //string query = "select رقم_المرجع,docTable from TableGeneralArch WHERE الاسم=@الاسم";
            string query = "select رقم_المرجع,docTable from TableGeneralArch where الاسم is null and رقم_المرجع <> 0";
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            //sqlDa.SelectCommand.Parameters.AddWithValue("@الاسم", "");
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["docTable"].ToString() != "")
                {
                    string name = getNames(dataRow["رقم_المرجع"].ToString(), dataRow["docTable"].ToString());
                    if (name != "" && name != "مؤرشف نهائي")
                    {
                        updateGenName(name, dataRow["رقم_المرجع"].ToString(), dataRow["docTable"].ToString());
                    }
                }
            }

        }

        private void updateGenName(string name, string idDoc, string table)
        {
            string dataSource = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");

            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string query = "update TableGeneralArch set الاسم=N'" + name + "' where رقم_المرجع = '" + idDoc + "' and docTable=N'" + table + "'";
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        
        private void updateGenNameError(string name, string idDoc)
        {
            string dataSource = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");

            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string query = "update TableGeneralArch set الاسم=N'" + name + "' where رقم_معاملة_القسم = N'" + idDoc + "'";
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        private void correctNo()
        {
            string dataSource = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");

            SqlConnection sqlCon = new SqlConnection(dataSource);
            //string query = "select رقم_المرجع,docTable from TableGeneralArch WHERE الاسم=@الاسم";
            string query = "select ID,رقم_معاملة_القسم from TableGeneralArch where ID >= 18213 and نوع_المستند = 'data1'";
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            //sqlDa.SelectCommand.Parameters.AddWithValue("@الاسم", "");
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            foreach (DataRow dataRow in dtbl.Rows)
            {
                string[] info = dataRow["رقم_معاملة_القسم"].ToString().Split('/');
                //MessageBox.Show("info " + info[0] + Environment.NewLine + "info " + info[1] + Environment.NewLine+ "info " + info[2] + Environment.NewLine+ "info " + info[3] + Environment.NewLine+ "info " + info[4]);
                if (info.Length == 5)
                {
                    string newInfo = info[0] + "/" + info[4] + "/" + info[3] + "/" + info[2] + "/" + info[1];

                    //MessageBox.Show(dataRow["رقم_معاملة_القسم"].ToString());
                    //MessageBox.Show(newInfo);
                    if (info[1] != "80")
                    {
                        sqlCon = new SqlConnection(dataSource);
                        if (sqlCon.State == ConnectionState.Closed)
                            sqlCon.Open();
                        SqlCommand sqlCmd = new SqlCommand("update TableGeneralArch set رقم_معاملة_القسم=@رقم_معاملة_القسم where ID=@id", sqlCon);
                        sqlCmd.CommandType = CommandType.Text;
                        sqlCmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataRow["ID"].ToString()));
                        sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", newInfo);
                        sqlCmd.ExecuteNonQuery();
                        sqlCon.Close();
                    }
                }
            }

        }

        private void unfounddata(string[] tableList)
        {
            string queryList = "";
            for (int table = 1; table < 15; table++)
            {
                if (tableList[table] == "") continue;
                for (int data = 1; data <= 2; data++)
                {
                    string query = "insert into TableGeneralArch (Data1,Extension1,المستند,نوع_المستند,رقم_معاملة_القسم,الموظف,التاريخ,رقم_المرجع,الاسم) " +
                                   "select Data" + data.ToString() + ",Extension" + data.ToString() + ",FileName" + data.ToString() + ",'data" + data.ToString() + "', DocID,DataInterName,GriDate,ID,AppName " +
                                   " from " + tableList[table] + " where  ID in (" +
                                   "select ID from " + tableList[table] + " where Extension" + data.ToString() + " is not null and not exists  (" +
                                   "select رقم_المرجع from TableGeneralArch " +
                                   "where docTable = '" + tableList[table] + "' and نوع_المستند = 'data" + data.ToString() + "' and " + tableList[table] + ".ID = TableGeneralArch.رقم_المرجع) )";
                    queryList = queryList + Environment.NewLine + query;
                    //MessageBox.Show(query);
                    //SqlConnection sqlCon = new SqlConnection(DataSource.Replace("AhwalDataBase", "ArchFilesDB"));
                    //SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
                    //if (sqlCon.State == ConnectionState.Closed)
                    //    sqlCon.Open();
                    //sqlCmd.CommandType = CommandType.Text;
                    //sqlCmd.ExecuteNonQuery();
                    //sqlCon.Close();
                    //Console.WriteLine(tableList[table]);
                }
            }
            dataSourceWrite("D:\\list.txt", queryList);
        }

        
        private string getNames(string id, string table)
        {
            string col = getColumnList(table, "TableList");
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query = "select " + col + " from " + table + " where ID='" + id + "'";            
            Console.WriteLine("query " + query);
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
                //MessageBox.Show(col +" - "+table);
            }
            sqlCon.Close();

            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow[col].ToString() != "")
                {
                   return dataRow[col].ToString();
                }
            }
            return "";
        }

        

        private void button3_Click(object sender, EventArgs e)
        {
            finalArch = false;
            btnAuth.Select();
            string AppName = "";
            resetItems();
            docId.Text = docId.Text.TrimStart().TrimEnd();
            
            string year= SpecificDigit(docId.Text, 1, 2).Trim();
            
            if (ServerType == "56")
            {
                FormType = Convert.ToInt32(SpecificDigit(docId.Text, 3, 5));
                noForm = SpecificDigit(docId.Text, 3, 5);
                rowCount = SpecificDigit(docId.Text, 6, docId.Text.Length);
                
            }
            else
            {
                
                FormType = Convert.ToInt32(SpecificDigit(docId.Text, 3, 4));
                noForm = SpecificDigit(docId.Text, 3, 4);
                rowCount = SpecificDigit(docId.Text, 5, docId.Text.Length);
            }
            txtIDNo.Text = docIDNumber = txtMissionCode + year + "/" + noForm + "/" + rowCount;
            smsDocIDNumber = "ق س ج/" + rowCount + "/" + noForm + "/" + year;

            //MerrageAndDivProc();



            string index = "-1";
            if (ServerType == "56")
                index = SpecificDigit(noForm, 3, 3);
            //MessageBox.Show(noForm);
            getColList(noForm, ArchiveState, index);
            //Console.Write("1 " + DateTime.Now.ToString("T"));
            if (noForm == "10" || noForm == "12")
            {
                getComboText(docIDNumber);
              //  Console.Write("2 " + DateTime.Now.ToString("T"));
            }
            else if (noForm == "15")
            {
                Combo1Text = "وثيقة زواج";
                Combo2Text = "";
            }
            else if (noForm == "17")
            {
                Combo2Text = "";
                Combo1Text = "وثيقة طلاق";
            }
            loadPreReq(noForm, Combo1Text + "-" + Combo2Text, ArchiveState);
            string formName = Combo1Text + "-" + Combo2Text;

            getTableList(noForm);
            string wordInFile = Combo1Text.Trim() + "-" + Combo2Text;
            if (Combo1Text != "")
            {
                drawBoxesTitle("استمارة الطلب الأولية", 80);
                drawBoxes(formName, false, wordInFile);
            }
            //Console.Write("4 " + DateTime.Now.ToString("T"));
            getProcedName = checkBasicInfo(docIDNumber);
            //Console.Write("5 " + DateTime.Now.ToString("T"));
            string CheckState = checkArch(docIDNumber, getProcedName);
            requiredDocument.Text = CheckState;
            paraValues[2] = docIDNumber;
            docId.Clear();
            docId.Text = DateTime.Now.Year.ToString().Replace("20", "");
            //Console.Write("6 " + DateTime.Now.ToString("T"));
        }

        private void MerrageAndDivProc()
        {
            if (noForm == "15" || noForm == "17")
            {
                if (noForm == "15")
                    label5.Text = "تاريخ إجراء عقد الزواج:";
                else if (noForm == "17")
                    label5.Text = "تاريخ إصدار وثيقة الطلاق:";
                تاريخ_الميلاد.Width = 90;
                تاريخ_الميلاد.Visible = label5.Visible = true;
                تاريخ_الميلاد.Location = new System.Drawing.Point(1044, 288);
                label5.Location = new System.Drawing.Point(1141, 288);
                تاريخ_الميلاد.Text = getProDate(txtIDNo.Text, getTableList(noForm));
                if (تاريخ_الميلاد.Text == "")
                    تاريخ_الميلاد.Text = "عام-شهر-يوم";
            }
            else
                تاريخ_الميلاد.Visible = label5.Visible = false;
        }

        private string getProDate(string docID, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string settingData = "select تاريخ_الاجراء from " + table + " where رقم_المعاملة = N'" + docID + "'";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            foreach (DataRow dataRow in dtbl.Rows)
            {
                return dataRow["تاريخ_الاجراء"].ToString();                
            }
            return "";
        }
        private void resetItems()
        {
            foreach (Control control in drawPicPanel.Controls)
            {
                control.Visible = false;
            }
            
            drawBoxesindex = 0;
            btnSaveEnd.Visible = false;
            PathImage = new string[100];
            rowCount = "";
            imagecount = 0;
            picPath = "";
            pictureBox1.Image = PersAhwal.Properties.Resources.noImage;
            proID = "0";
            reqFile = "";
        } 
        private string getComboIndex(string col, string combo1)
        {
            string query = "SELECT " + col+ " FROM TableListCombo WHERE " + col + " is not null";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            //MessageBox.Show(query);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            int x = 0;
            foreach (DataRow row in dtbl.Rows)
            {
                if (row[col].ToString() == combo1) return x.ToString();
                x++;
            }
            return "-1";
        }

        private void getComboText(string docIDNum)
        {
            dontCheck = true;
            string query = "SELECT " + altColName + "," + altSubColName + " FROM " + TableList + " WHERE " + allUpdateNamesList[2] + "=N'" + docIDNum + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            Console.WriteLine(query);
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow row in dtbl.Rows)
            {
                Console.WriteLine(row[altColName].ToString() + " - "+ row[altSubColName].ToString());
                try
                {

                    Combo1Text = row[altColName].ToString();
                    Combo2Text = row[altSubColName].ToString();
                    reqFile = "";
                    comboCol[0] = Combo1Text;
                    comboCol[1] = Combo2Text;
                }
                catch (Exception ex) 
                {
                    
                }
            }
            dontCheck = false;
            Console.WriteLine(query);
            //MessageBox.Show(query);

        }

        private void docId_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                button3.PerformClick();
                //MessageBox.Show(requiredDocument.Text );
                if (requiredDocument.Text != "معاملة غير موجودة" && JobPosition.Contains("قنصل"))
                {
                    btnDelete.Visible = btnArchived.Visible = btnExten.Visible = true;
                }
                else
                {
                    btnDelete.Visible = btnArchived.Visible = btnExten.Visible = false;
                    
                }
            }
        }

        

        private void printPreviewDialog1_Load(object sender, EventArgs e)
        {

        }

        private void docId_TextChanged(object sender, EventArgs e)
        {

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            التاريخ_الميلادي.Text = Labdate = GregorianDate = DateTime.Now.ToString("MM-dd-yyyy");
            timer2.Enabled = false;
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

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView2.Visible = false;
            panel1.Visible = true;
            string fileName = loadDocxFile();
            if (fileName != "") {
                pictureBox1.ImageLocation = PathImage[imagecount] = fileName;
                imagecount++;
                btnAuth.BackColor = System.Drawing.Color.LightGreen;
                btnAuth.Text = "اضافة مستند آخر (" + (imagecount + 1).ToString() + ")";
                
            }            
        }

        private void reLoadPic_Click(object sender, EventArgs e)
        {
            dataGridView2.Visible = false;
            panel1.Visible = true;
            
            string fileName = loadDocxFile();
            if (fileName != "")
            {
                pictureBox1.ImageLocation = PathImage[imagecount - 1] = fileName;
            }
        }

        private void mandoubName_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (mandoubName.SelectedIndex == 0 || ServerType == "56") return;
                int found = todayList(mandoubName.Text.Trim(), Labdate);            
            if(found != 0 && ArchiveState )
            {
                //btnSaveEnd.Enabled = false;
                MessageBox.Show("المندوب لدية عدد " + found.ToString() + " مكاتبات غير مكتملة،،، يرجى الحصول على النسخ الاصلية أو صورها مع البصمة");                
            }
            else
                btnSaveEnd.Enabled = true;
            if (mandoubName.Text != "حضور مباشرة إلى القنصلية") 
                panel3.Visible = button5.Visible = true;
            else 
                panel3.Visible = button5.Visible = false;

            if (!ArchiveState && mandoubName.SelectedIndex != 0)
            {
                
                mandoubName.Size = new System.Drawing.Size(256, 35);
                mandoubName.Location = new System.Drawing.Point(1098, 247);
            }
            else {
                mandoubName.Size = new System.Drawing.Size(312, 35);
                mandoubName.Location = new System.Drawing.Point(1042, 247);
            }


        }


        private bool checkActiv(string mandoub)
        {
            if (mandoub == "حضور مباشرة إلى القنصلية")
                return true;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select وضع_المندوب from TableMandoudList where MandoubNames=N'" + mandoub.Split('-')[0].Trim() + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow row in dtbl.Rows)
            {
                if(row["وضع_المندوب"].ToString() == "الحساب مفعل")
                    return true;
            }
           return false;
        }
        private void getMandoubID(string mandoub)
        {
            //MessageBox.Show(mandoub);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select ID from TableMandoudList where MandoubNames=N'" + mandoub.Split('-')[0].Trim() + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow row in dtbl.Rows)
            {
                ArchMandoubID = row["ID"].ToString();
                //MessageBox.Show(ArchMandoubID);
            }
           
        }

                private void DocType_CheckedChanged(object sender, EventArgs e)
        {
            if (DocType.CheckState == CheckState.Checked)
            {
                DocType.Text = "أصل المكاتبة";
            }
            else {
                DocType.Text = "صورة نهائية";
            }
            }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            deleteRowsData(FileIDNo,TableList,"ID");
            deleteArchRowsData(docIDNumber, "TableGeneralArch", "رقم_معاملة_القسم");
            deleteRowsData(docIDNumber, "archives", "docID");
            btnDelete.Visible = btnArchived.Visible = btnExten.Visible = false;
            
        }

        private void UpdateArchState(int id, string table)
        {            
            string qurey = "update " + table + " set " + archCol + "=@" + archCol + " where ID=@id";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@"+ archCol, "مؤرشف نهائي");
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        private void UpdateNameState(int id, string table, string name)
        {            
            string qurey = "update " + table + " set " + columnList + "=@" + columnList + " where ID=@id";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@"+ columnList, name);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void UpdateDateState(int id, int FileID, string table)
        {
            string archdate = "GriDate";
            if (FileID >= 11) archdate = "التاريخ_الميلادي";

            string qurey = "update " + table + " set " + archdate + "=@" + archdate + " where ID=@id";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@" + archdate, GregorianDate);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void btnArchived_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(TableList);
            UpdateArchState(Convert.ToInt32(FileIDNo), TableList);
            deleteRowsData(docIDNumber, "archives", "docID");
            btnDelete.Visible = btnArchived.Visible = btnExten.Visible = false;
            this.Close();
        }

        private void btnExten_Click(object sender, EventArgs e)
        {
            UpdateMandoubState(txtIDNo.Text, "docDate", GregorianDate);
            //UpdateDateState(Convert.ToInt32(FileIDNo), FileTableID, FileTable);
            btnDelete.Visible = btnArchived.Visible = btnExten.Visible = false;
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            //editForms();
            if (!checkActiv(mandoubName.Text.Split('-')[0].Trim()))
            {
                MessageBox.Show("حساب المندوب عير مفعل، يرجى التواصل مع مدير القسم للتفعيل");
                return;
            }

            if (!mandoubName.Text.Contains("-")) return;
            button5.Enabled = false;
            reqFile = "";
            OpenFile("استمارات المناديب", false);
            string DocxOutFile = FilespathOut + mandoubName.Text.Trim() + DateTime.Now.ToString("ssmm") + ".docx";
            string PDFOutFile = FilespathOut + mandoubName.Text.Trim() + DateTime.Now.ToString("ssmm") + ".pdf";

            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object objCurrentCopy = reqFile;
            Word.Document oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);
            oBMicroWord.Selection.Find.ClearFormatting();
            oBMicroWord.Selection.Find.Replacement.ClearFormatting();
            object ParaAuthIDNo = "المندوب";
            Word.Range BookAuthIDNo = oBDoc.Bookmarks.get_Item(ref ParaAuthIDNo).Range;
            BookAuthIDNo.Text = mandoubLine(mandoubName.Text.Split('-')[0].Trim());
            object rangeAuthIDNo = BookAuthIDNo;
            oBDoc.Bookmarks.Add("المندوب", ref rangeAuthIDNo);
            oBDoc.SaveAs2(DocxOutFile);

            oBDoc.ExportAsFixedFormat(PDFOutFile, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(PDFOutFile);
            File.Delete(DocxOutFile);
            //MessageBox.Show(التاريخ_الميلادي.Text);
            reqFile = "";
            restrictions();
            button5.Enabled = true;
        }
        private void restrictions() {
            reqFile = "";
            OpenFile("ضوابط استمارات المناديب", false);
            string DocxOutFile = FilespathOut + Combo1.Text.Trim() + DateTime.Now.ToString("ssmm") + ".docx";
            string PDFOutFile = FilespathOut + Combo1.Text.Trim() + DateTime.Now.ToString("ssmm") + ".pdf";

            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object objCurrentCopy = reqFile;
            Word.Document oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);
            oBMicroWord.Selection.Find.ClearFormatting();
            oBMicroWord.Selection.Find.Replacement.ClearFormatting();
            
            object ParaAuthIDNo1 = "التاريخ_الميلادي";
            Word.Range BookAuthIDNo1 = oBDoc.Bookmarks.get_Item(ref ParaAuthIDNo1).Range;
            BookAuthIDNo1.Text = التاريخ_الميلادي.Text;
            object rangeAuthIDNo1 = BookAuthIDNo1;
            oBDoc.Bookmarks.Add("التاريخ_الميلادي", ref rangeAuthIDNo1);
            
            object ParaAuthIDNo2 = "بيانات_المندوب";
            Word.Range BookAuthIDNo2 = oBDoc.Bookmarks.get_Item(ref ParaAuthIDNo2).Range;
            BookAuthIDNo2.Text = بيانات_المندوب.Text;
            object rangeAuthIDNo2 = BookAuthIDNo1;
            oBDoc.Bookmarks.Add("بيانات_المندوب", ref rangeAuthIDNo2);
            
            oBDoc.SaveAs2(DocxOutFile);


            oBDoc.ExportAsFixedFormat(PDFOutFile, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(PDFOutFile);
            File.Delete(DocxOutFile);

            reqFile = "";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            mandoubFiles();
            dataGridView2.Visible = true;
            panel1.Visible = false;
            //MessageBox.Show(id.ToString());
            string fileName = loadDocxFile();
            if (fileName != "")
            {
                int id = getID(fileName);

                if (id == 0) return;
                SqlConnection sqlCon = new SqlConnection(DataSource);
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlCommand sqlCmd = new SqlCommand("update TableListCombo set EnterySheet=@EnterySheet where ID=@id", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@id", id);
                sqlCmd.Parameters.AddWithValue("@EnterySheet", fileName.Split('\\')[7]);
                sqlCmd.ExecuteNonQuery();
                sqlCon.Close();
            }
        }

        private void FormPics_FormClosed(object sender, FormClosedEventArgs e)
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
            
            
            dataSourceWrite(primeryLink + @"\refresh.txt", "Allowed");
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

        private void txtIDNo_TextChanged(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from archives where docID=@docID", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@docID", txtIDNo.Text);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows) {
                readyToRemove = true;
                if (dataRow["mandoubName"].ToString() != "")
                {
                    mandoubName.Text = dataRow["mandoubName"].ToString();
                    mandoubName.Visible = true;
                }
                else {
                   // mandoubName.Visible = false;
                    mandoubName.Text = "حضور مباشرة إلى القنصلية";
                }
            }
        }

        private int getMaxDocNo(string table, string docid, string colName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select "+ colName+" from "+ table+" where " + colName+" like N'"+ txtMissionCode + DateTime.Now.Year.ToString().Replace("20", "") + "/%'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@" + colName, docid);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            int maxID = 0;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow[colName].ToString().Contains('/'))
                {
                    try
                    {
                        string newInfo = dataRow[colName].ToString().Split('/')[4];
                        int id = Convert.ToInt32(newInfo);
                        if (id > maxID) maxID = id;
                    }
                    catch (Exception ex) {
                        maxID = 1;
                    }
                }

            }
            return maxID;
        }
        
        private string getUniqueID(string query)
        {
            //string dataSource = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            //dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            Console.WriteLine(query);
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string maxID = "1";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                try
                {
                    maxID = (Convert.ToInt32(dataRow["newDocID"].ToString())+1).ToString();
                }
                catch (Exception ex)
                {
                    return maxID;
                }
            }
            return maxID;
        }
        private bool checkISUnique(string docid)
        {
            string dataSource = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            dataSource = dataSource.Replace("SudaneseAffairs", "ArchFilesDB");

            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            //SqlDataAdapter sqlDa = new SqlDataAdapter("select " + docName + " from " + table + " where " + docName + "=@" + docName, sqlCon);

            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableGeneralArch where رقم_معاملة_القسم =@col", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@col", docid);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            Console.WriteLine("uniqueness " + dtbl.Rows.Count.ToString()); 
            if (dtbl.Rows.Count != 0) return true;
            else return false;
            
        }

        private void jpgFile_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox1_LoadCompleted(object sender, AsyncCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                pictureBox1.Image = PersAhwal.Properties.Resources.noImage;
                try
                {
                    System.Diagnostics.Process.Start(picPath);
                }
                catch (Exception ex) { }
            }
            
        }

        private void تاريخ_الميلاد_ValueChanged(object sender, EventArgs e)
        {
            //تاريخ_الميلاد
            
        }
        string lastInput2 = "";
        private void تاريخ_الميلاد_TextChanged(object sender, EventArgs e)
        {
            if (تاريخ_الميلاد.Text.Length  == 10)
            {
                try
                {
                    int month = Convert.ToInt32(SpecificDigit(تاريخ_الميلاد.Text, 4, 5));
                
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //تاريخ_الميلاد.Text = "";
                    تاريخ_الميلاد.Text = SpecificDigit(تاريخ_الميلاد.Text, 7, 10);
                    return;
                }
                }
                catch (Exception ex) { return; }
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

        private void nameSave_Click(object sender, EventArgs e)
        {
            UpdateNameState(Convert.ToInt32(FileIDNo), TableList, requiredDocument.Text);
            updateGenNameError(requiredDocument.Text, txtIDNo.Text);
            this.Close();   
        }

        private void Combo1_TextChanged(object sender, EventArgs e)
        {
            if (dontCheck) return;
            Console.WriteLine(21);
            loadPreReq(noForm, Combo1.Text, ArchiveState);
            Console.WriteLine(22);
            if (FormType == 10 ) {
                fileComboBox(Combo2, DataSource, Combo1.Text.Replace(" ","_"), "TableListCombo");
                Console.WriteLine(23);
            }
            if (FormType == 10)
            {
                autoCompleteTextBox(Combo2, DataSource, "نوع_الإجراء", "TableCollection");
                Console.WriteLine(24);
            }
            else if (FormType == 12)
            {
                autoCompleteTextBox(Combo2, DataSource, "إجراء_التوكيل", "TableAuth");
                Console.WriteLine(25);
            }

            }

        private void autoCompleteTextBox(ComboBox textbox, string source, string comlumnName, string tableName)
        {

            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct " + comlumnName + " from " + tableName + " where " + comlumnName + "<> ''";
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
                    if (!string.IsNullOrEmpty(dataRow[comlumnName].ToString()))
                    {
                        //for (int x = 0; x < Textboxtable.Rows.Count; x++)
                        //    if (dataRow[comlumnName].ToString().Equals(Textboxtable.Rows[x]))
                        //        newSrt = false;

                        //if (newSrt)
                            autoComplete.Add(dataRow[comlumnName].ToString());
                    }
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }
        public static bool IsRtl(string input)
        {
            return Regex.IsMatch(input, @"\p{IsArabic}");
        }

        private void Combo2_TextChanged(object sender, EventArgs e)
        {
            if (dontCheck) return;
            //MessageBox.Show(Combo2.Text);
                int index = 0;
            for (; index < Combo2.Items.Count; index++) {
                if (Combo2.Text != Combo2.Items[index].ToString()) continue;
                else break;
            }
            if (Combo2.Items.Count > 0 && Combo2.Visible && noForm != "06")
            {
                Combo2.Visible = true;
                //Combo2.SelectedIndex = index;
                //MessageBox.Show("noForm= " + noForm + "_" + Combo1.Text + "-" + Combo2.Text);
                loadPreReq(noForm, Combo1.Text + "-" + Combo2.Text, ArchiveState);
            }
            if(Combo2.Enabled) reqFile = "";
            reqFile = "";
            OpenFile(Combo1.Text, false);
            if (reqFile == "")
                OpenFile(Combo1.Text + "-" + Combo2.Text, false);

            //MessageBox.Show("noForm= " +noForm +"_"+index.ToString());
            //1if (noForm != "" && Combo2.Visible)
            //{
            //    loadPreReq(noForm, Combo1.Text + "-" + Combo2.Text, ArchiveState);
            //}
        }

        private void mandoubName_TextChanged(object sender, EventArgs e)
        {
            try
            {
                بيانات_المندوب.Text = mandoubName.Text.Split('-')[0]+" مندوب جالية "+ mandoubName.Text.Split('-')[1];
                //MessageBox.Show(بيانات_المندوب.Text);
            }
            catch (Exception ex) { 

            }
        }

        private void btnArchMandoub_Click(object sender, EventArgs e)
        {
            MandArch = true;
            drawBoxes("أرشفة خطاب تفويض المندوب ", true, "");
            drawBoxes("أرشفة جواز سفر المندوب ", true, "");
            ArchMandoubID = "";
            getMandoubID(mandoubName.Text);                
            btnSaveEnd.Location = new System.Drawing.Point(736, 668);
        }

        private void تعليق_جديد_Off_Click(object sender, EventArgs e)
        {
            تعليق_جديد_Off.Height = 290;
        }

        private void التعليقات_السابقة_Off_Click(object sender, EventArgs e)
        {
            تعليق_جديد_Off.Height = 45;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show( pictureBox1. );
            try
            {
                System.Diagnostics.Process.Start(pictureBox1.ImageLocation);
            }
            catch (Exception ex) { }
        }

        private void تاريخ_الميلاد_Click(object sender, EventArgs e)
        {
            if (تاريخ_الميلاد.Text == "عام-شهر-يوم")
                تاريخ_الميلاد.Text = "";
        }

        private void pictureBox1_LocationChanged(object sender, EventArgs e)
        {
           

        }

        private void pictureBox1_BackgroundImageChanged(object sender, EventArgs e)
        {
            MessageBox.Show(pictureBox1.ImageLocation.ToString());
        }

        private void Combo1_TextUpdate(object sender, EventArgs e)
        {
            if (dontCheck) return;
            Console.WriteLine(31);
            if (checkaltColName(Combo1.Text))
            {
                reqFile = "";
                string[] data = new string[2];

                data[0] = noForm;
                Console.WriteLine(32);
                fileComboBoxSub(Combo2, DataSource, "altColName", "altSubColName");
                Console.WriteLine(33);
                if (ArchiveState) DocIDGenerator(FormType);
                Console.WriteLine(34);
                loadPreReq(noForm, Combo1.Text, ArchiveState);
                Console.WriteLine(35);
                if (FormType == 10)
                {
                    autoCompleteTextBox(Combo2, DataSource, "نوع_الإجراء", "TableCollection");
                    Console.WriteLine(36);
                }
                else if (FormType == 12)
                {
                    autoCompleteTextBox(Combo2, DataSource, "إجراء_التوكيل", "TableAuth");
                    Console.WriteLine(37);
                }

                }
            
        }

        private void تاريخ_الميلاد_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                if (btnSaveEnd.Visible)
                    btnSaveEnd.Select();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            finalArch = false;
            if (docId.Text.Length < 5)
            {
                MessageBox.Show("يرجى كتابة الرقم المرجعي كاملا");
                return;
            }
            if (docId.Text.Length == 8)
                rowCount = SpecificDigit(docId.Text, 5, 8);
            getTableList((Convert.ToInt32(SpecificDigit(docId.Text, 3, 4)) - 1).ToString());
            docIDNumber = "ق س ج/160/" + SpecificDigit(docId.Text, 3, 4) + "/" + rowCount;
            requiredDocument.Text = OpenFile(docIDNumber, Convert.ToInt32(SpecificDigit(docId.Text, 3, 4)), TableList);
            if (requiredDocument.Text == "معاملة غير موجودة")
            {
                docIDNumber = "CGSJ/160/" + SpecificDigit(docId.Text, 3, 4) + "/" + rowCount;

                requiredDocument.Text = OpenFile(docIDNumber, Convert.ToInt32(SpecificDigit(docId.Text, 3, 4)), TableList);
            }

            if (requiredDocument.Text != "معاملة غير موجودة" && requiredDocument.Text != "")
            {
                btnAuth.Visible = true;
                FormType = Convert.ToInt32(SpecificDigit(docId.Text, 3, 4));

                ArchiveState = true;
                newEntry = false;
            }
        }
    }
}
