
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
using System.Runtime.InteropServices.ComTypes;

namespace PersAhwal
{
    public partial class FormSudAffairs : Form
    {
        DataTable dtblMain;
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        string[] PathImage = new string[100];
        DeviceInfo AvailableScanner = null;
        int imagecount = 0;
        string FormDataFile;
        Excel.Range range;
        bool NewEntrey = true;
        string gregorianDate = "";
        int iqrarType = 0;
        int[] SearchfileList = new int[1000];
        string fileVersio;
        bool Pers_Peope = true;
        int cl = 0, rw=0;
        string Docfile = "";
        string smsText = "";
        int[] ides = new int[10];
        int ALLSmsIndex = 0;
        int Nobox = 0;
        bool grdiFill = false;
        int FileIDNo = 0;
        string FilesPathIn, FilesPathOut;
        string DataSource;
        string[,] preffix = new string[20,20];
        string[] checkBoxSex = new string[10];
        string[] Names  = new string[10];
        string[] WorkOffices = new string[15];
        string[] DocumentNo = new string[10];
        string[] Relativity = new string[10];
        int[] fileList;
        string[] AppNames = new string[10000];
        string[] PhoneNo = new string[10000];
        string[] Status = new string[10000];
        string[] AppDocumentNo = new string[10000];
        string[] straffair = new string[3] { "مخاطبات إدارة مكتب العمل", "مخاطبات إدارة الوافدين", "مخاطبات اللجنة العمالية" };
        string Jobposition;
        string RefDocument = "";
        int VCIndex = 0;
        int AppIndex = 1;
        int ApplicantID = -1;
        bool newData = false;
        int AffairIndex = 0;
        int CityIndex = 0;
        bool ModifyPermit = true;
        string contractState = "0";
        string EmpName = "";
        public FormSudAffairs(bool modifyPermit,int affairIndex,int vcIndex,string dataSource,string filesPathIn, string filesPathOut, string jobposition, string empName, string datasource57)
        {
            InitializeComponent();
            FilesPathIn = filesPathIn;
            FilesPathOut = filesPathOut;
            Jobposition = jobposition;
            DataSource = dataSource;
            VCIndex = vcIndex;
            AffairIndex = affairIndex;
            ModifyPermit = modifyPermit;
            EmpName = empName;
            if (ModifyPermit)
                finalPro.Visible = true;
            else 
                finalPro.Visible = false;
            
            //MessageBox.Show(FilesPathIn);
            WorkOffices[0] = "مدير إدارة الوافدين";
            WorkOffices[1] = "مدير إدارة الوافدين مكة المكرمة";
            WorkOffices[2] = "مدير إدارة الوافدين الطائف";
            WorkOffices[3] = "مدير إدارة الوافدين المدينة المنورة";
            WorkOffices[4] = "مدير إدارة الوافدين تبوك";
            WorkOffices[5] = "مدير إدارة الوافدين محايل عسير";
            WorkOffices[6] = "مدير إدارة الوافدين الباحة";
            WorkOffices[7] = "مدير إدارة الوافدين نجران";
            WorkOffices[8] = "مدير إدارة الوافدين جازان";
            WorkOffices[9] = "مدير إدارة الوافدين ابها";
            WorkOffices[10] = "مدير إدارة الوافدين عسير";
            WorkOffices[11] = "مدير إدارة الوافدين خميس مشيط";
            WorkOffices[12] = "مدير إدارة الوافدين بيشة";
            WorkOffices[13] = "مدير إدارة الوافدين ينبع";
            WorkOffices[14] = "مدير إدارة الوافدين القنفذة";
            dtblMain = new DataTable();
            fillFileBox(AffairIndex,dataGridView1);

            //combFileNo.Text = getFileNo(DataSource);
            clear_All();
            if (Jobposition.Contains("قنصل"))
                deleteRow.Visible =  true;
            else deleteRow.Visible = false;
            
            //button4_Click();
            if(combFileNo.Items.Count > 0 && AffairIndex < 5) 
                combFileNo.SelectedIndex = 0;

            //MessageBox.Show();
            //MessageBox.Show(fileCount(getFileNo(0)).ToString());
            //deleteDuplicate(AffairIndex);

            //btnTarListUpload();
            labelEmp.Text = "اسم الموظف:" + empName;
            comFinalaPro.SelectedIndex = 0;
            ListSearch.Select();
            //getColList(getTable(AffairIndex));
            
        }


        private bool checkUnique(string table, string text)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from " + table + " WHERE رقم_الهوية = '" + text+"'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count > 0)
                return true;
            else
                return false;

        }

        private void FillFilesView1()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("FileViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@FileNo", "0");
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);

            sqlCon.Close();
            //dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[5].Width = dataGridView1.Columns[6].Width = dataGridView1.Columns[4].Width = 150;
            dataGridView1.Columns[7].Width = dataGridView1.Columns[8].Width = dataGridView1.Columns[9].Width = 150;
            dataGridView1.Columns[10].Width = dataGridView1.Columns[11].Width  = 150;
            

        }

        private void ColorFulGrid1(int index, string txt)
        {
            //24
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;
                if (dataGridView1.Rows[i].Cells[index].Value.ToString() != txt)
                {
                    Console.WriteLine(dataGridView1.Rows[i].Cells[index].Value.ToString());
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                }

                //else dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.;

            }
            //
        }

        private int  getFileInfo()
        {
            //SqlConnection sqlCon = new SqlConnection(DataSource);
            //if (sqlCon.State == ConnectionState.Closed)
            //    sqlCon.Open();
            //SqlDataAdapter sqlDa = new SqlDataAdapter("select مقدم_الطلب,رقم_الهوية,رقم_الملف,نوع_المعاملة from TableWafid", sqlCon);
            //sqlDa.SelectCommand.CommandType = CommandType.Text;
            //DataTable dtbl = new DataTable();
            //sqlDa.Fill(dtbl);           
            //sqlCon.Close();
            int x = 0;

            AppNames = new string[dataGridView1.Rows.Count - 1];
            PhoneNo = new string[dataGridView1.Rows.Count - 1];
            Status = new string[dataGridView1.Rows.Count - 1];
            AppDocumentNo = new string[dataGridView1.Rows.Count - 1];
            int c = 0;

            foreach (DataGridViewRow dataRow in dataGridView1.Rows)
            {
                if (dataGridView1.Rows.Count <= 1) return 0;
                var cell = dataGridView1.Rows[c].Cells[2];                
                if (cell.Value != null )   //Check for null reference
                {
                    if (!string.IsNullOrEmpty(dataGridView1.Rows[c].Cells[2].Value.ToString()) && !string.IsNullOrEmpty(dataGridView1.Rows[c].Cells[5].Value.ToString()))
                    {
                        AppNames[x] = dataGridView1.Rows[c].Cells[2].Value.ToString();
                        AppDocumentNo[x] = dataGridView1.Rows[c].Cells[5].Value.ToString();
                        PhoneNo[x] = dataGridView1.Rows[c].Cells[19].Value.ToString();
                        Status[x] = dataGridView1.Rows[c].Cells[26].Value.ToString();
                        x++;
                    }
                }
                c++;
            }


            //foreach (DataRow dataRow in dtbl.Rows)
            //{
            //    if (!string.IsNullOrEmpty(dataRow["مقدم_الطلب"].ToString()) && !string.IsNullOrEmpty(dataRow["رقم_الهوية"].ToString()) && dataRow["رقم_الملف"].ToString()== combFileNo.Text && dataRow["نوع_المعاملة"].ToString() == AffairIndex.ToString())
            //    {
            //         dataRow["مقدم_الطلب"].ToString();
            //        [x] = dataRow["رقم_الهوية"].ToString()
            //        x++;
            //    }else dtbl.Rows[x].Delete();
            //}

            
            //MessageBox.Show(x.ToString());
            return c;
        }

        private void AllDataGridView1()
        {
            string str = "";
            DataTable dtbl = new DataTable();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            //for(int index = 0; index<5; index++)
            //{
            //    switch (index)
            //    {
            //        case 0:
            //            str = "WafidVOrS";
            //            break;
            //        case 1:
            //            str = "WafidJeddViewOrSearch";
            //            break;
            //        case 2:
            //            str = "WafidMekkahViewOrSearch";
            //            break;
            //        case 3:
            //            str = "TarheelViewOrSearch";
            //            break;
            //        case 4:
            //            str = "TransferViewOrSearch";
            //            break;
            //        case 5:
            //            str = "WafidComViewOrSearch";
            //            break;
            //        default: return;
            //    }
                SqlDataAdapter sqlDa = new SqlDataAdapter("unioonAllPro", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                
                sqlDa.Fill(dtbl);
            //}
            dataGridView1.DataSource = dtbl;
            int z = 0;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                bool found = false;

                for (int a = 0; a < combFileNo.Items.Count; a++)
                {
                    if (!string.IsNullOrEmpty(dataRow["رقم_الملف"].ToString()) && dataRow["رقم_الملف"].ToString() == combFileNo.Items[a].ToString())
                        found = true;
                }
                if (!found)
                {
                    Console.WriteLine("file no" + dataRow["رقم_الملف"].ToString());
                    combFileNo.Items.Add(dataRow["رقم_الملف"].ToString());
                }
                z++;
            }

            fileList = new int[combFileNo.Items.Count];
            for (int a = 0; a < combFileNo.Items.Count; a++)
            {
                if (combFileNo.Items[a].ToString() != "")
                    fileList[a] = Convert.ToInt32(combFileNo.Items[a].ToString());
                else fileList[a] = -1;

            }
            combFileNo.Items.Clear();
            for (int b = 0; b < fileList.Length; b++)
                for (int a = 1; a < fileList.Length; a++)
                {
                    int first = fileList[a - 1];
                    if (fileList[a - 1] < fileList[a])
                    {
                        fileList[a - 1] = fileList[a];
                        fileList[a] = first;
                    }
                }
            for (int a = 0; a < fileList.Length; a++)
                if (fileList[a] != -1) combFileNo.Items.Add(fileList[a]);

            z = 0;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["رقم_الملف"].ToString() == "-1" && combFileNo.Items.Count > 0)
                {
                    dataRow["رقم_الملف"] = combFileNo.Items[0].ToString();
                }

                //if (Convert.ToInt32(dataRow["رقم_الملف"].ToString()) == 100)
                //{
                //    UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), "1");
                //}

                //if (dataRow["جهة_العمل"].ToString() == "0")
                //{
                //    //TransferData(dataRow, "WafidJedAddorEdit");
                //    //deleteRowsData(Convert.ToInt32(dataRow["ID"].ToString()), "TableWafid", DataSource);
                //}
                z++;
            }

            lalCount.Text = dtbl.Rows.Count.ToString() + "/فرد";
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);

            sqlCon.Close();
            dataGridView1.Columns[7].Visible = false;
            dataGridView1.Columns[8].Visible = false;
            dataGridView1.Columns[9].Visible = false;
            //dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 200;
            dataGridView1.Columns[32].Width = 200;
        }

        private void UpdateState(int id, string col, string text, string table)
        {
            //sqlCmd.Parameters.AddWithValue("@appOldNew", "في انتظار نسخة المواطن");
            string qurey = "update " + table + " set " + col + "=@" + col + " where ID=@id";
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
        private void deleteDuplicate(int index)
        {
            //Console.WriteLine("index  " + index);
            string str = "WafidVOrS", table = "TableWafid";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            switch (index)
            {
                case 0:
                    str = "WafidVOrS";
                    table = "TableWafid";
                    break;
                case 1:
                    str = "WafidJeddViewOrSearch";
                    table = "TableWafidJed";
                    break;
                case 2:
                    str = "WafidMekkahViewOrSearch";
                    table = "TableWafidMekkah";
                    break;
                case 3:
                    str = "TarheelViewOrSearch";
                    table = "TableTarheel";
                    break;
                case 4:
                    str = "TransferViewOrSearch";
                    table = "TableTransfer";
                    break;
                case 5:
                    str = "WafidComViewOrSearch";
                    table = "TableCommity";
                    break;
                default: return;
            }

            SqlDataAdapter sqlDa = new SqlDataAdapter(str, sqlCon);

            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            //sqlDa.SelectCommand.Parameters.AddWithValue("@جهة_العمل", "");
            DataTable dtbl = new DataTable();

            sqlDa.Fill(dtbl);
            //dataGridView1.DataSource = dtbl;
            int z = 0;
            string appName;
            int appID;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                appName = dataRow["مقدم_الطلب"].ToString();
                appID =  Convert.ToInt32(dataRow["ID"].ToString());
                foreach (DataRow dataRow1 in dtbl.Rows)
                {
                    if (appName == dataRow1["مقدم_الطلب"].ToString() && appID != Convert.ToInt32(dataRow1["ID"].ToString()))
                    {
                        Console.WriteLine(dataRow1["ID"].ToString());
                        deleteRowsData(Convert.ToInt32(dataRow1["ID"].ToString()), table, DataSource);
                    }
                }
            }
        }

        private void GridViewSeach(int index, string fileN)
        {
            Console.WriteLine("index  " + index);
            string str = "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            switch (index)
            {
                case 0:
                    str = "WafidVOrS1";
                    break;
                case 1:
                    str = "WafidJeddVoS";
                    break;
                case 2:
                    str = "WafidMekkahVOS";
                    break;
                case 3:
                    str = "TarheelVoS";
                    break;
                case 4:
                    str = "TransferVoS";
                    break;
                case 5:
                    str = "WafidComVoS";
                    break;
                default: return;
            }
            SqlDataAdapter sqlDa = new SqlDataAdapter(str, sqlCon);

            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@FileNoID", fileN);
            DataTable dtbl = new DataTable();

            sqlDa.Fill(dtbl);
            dataGridView6.DataSource = dtbl;
            sqlCon.Close();
            
        }


        private void GridView(int index, string fileN)
        {
            Console.WriteLine("index  " + index);
            string str = "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            switch (index)
            {
                case 0:
                    str = "WafidVOrS1";
                    break;
                case 1:
                    str = "WafidJeddVoS";
                    break;
                case 2:
                    str = "WafidMekkahVOS";
                    break;
                case 3:
                    str = "TarheelVoS";
                    break;
                case 4:
                    str = "TransferVoS";
                    break;
                case 5:
                    str = "WafidComVoS";
                    break;
                default: return;
            }
            SqlDataAdapter sqlDa = new SqlDataAdapter(str, sqlCon);

            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@FileNoID", fileN);
            DataTable dtbl = new DataTable();

            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            //foreach (DataRow row in dtbl.Rows)
            //{
            //    string docID = row["رقم_المعاملة"].ToString();
            //    string tableID = row["ID"].ToString();
            //    updatetablesinGenArch(tableID, docID, getTable(index));
            //}
            lalCount.Text = dtbl.Rows.Count.ToString() + "/فرد";
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            dataGridView1.Sort(dataGridView1.Columns["ord"], System.ComponentModel.ListSortDirection.Ascending);
            sqlCon.Close();

            //foreach (DataRow dataRow in dtbl.Rows)
            //{
            //    if (dataRow["رقم_الملف"].ToString() == "-1" && combFileNo.Items.Count > 0)
            //    {
            //        dataRow["رقم_الملف"] = combFileNo.Items[0].ToString();
            //    }

            //    //if (Convert.ToInt32(dataRow["رقم_الملف"].ToString()) > 68)
            //    //{
            //    //    //UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), "1", getFileName(AffairIndex));
            //    //}

            //    if (dataRow["الحالة"].ToString() == "تغيب عن العمل")
            //    {
            //        //TransferData(dataRow, "TarheelAddorEdit");
            //        //deleteRowsData(Convert.ToInt32(dataRow["ID"].ToString()), "TableTarheel", DataSource);
            //    }
            //    z++;
            //}

            if (AffairIndex != 7)
            {
                //if (combFileNo.Items.Count > 0) combFileNo.SelectedIndex = combFileNo.Items.Count - 1;
                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[8].Visible = false;
                dataGridView1.Columns[9].Visible = false;

            }

            //dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 200;
            dataGridView1.Columns[32].Width = 200;
            combFileNo.Text = fileN;
        }
        
        private void filllExcelGrid(int index, string fileN)
        {
            Console.WriteLine("filllExcelGrid  " + index);
            if (index != 1) return;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string strQuery = "select مقدم_الطلب,رقم_الهوية,الحالة,ord from TableWafidJed where رقم_الملف = '" + fileN + "'";
            SqlDataAdapter sqlDa = new SqlDataAdapter(strQuery, sqlCon);

            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();

            sqlDa.Fill(dtbl);
            gridExcel.DataSource = dtbl;
            gridExcel.Sort(gridExcel.Columns["ord"], System.ComponentModel.ListSortDirection.Ascending);
            sqlCon.Close();
            string ReportName = DateTime.Now.ToString("mmss");
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xlsx" })
            {
                //sfd.FileName = FilesPathIn + "رقم الملف " + fileN +"_" +ReportName;
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var fileinfo = new FileInfo(sfd.FileName);
                        using (var package = new ExcelPackage(fileinfo))
                        {
                            ExcelWorksheet excelsheet = package.Workbook.Worksheets.Add("Rights");
                            excelsheet.Cells.LoadFromDataTable(dtbl);
                            package.Save();

                        }
                    }
                    catch (Exception ex)
                    {
                    }

                    System.Diagnostics.Process.Start(sfd.FileName);
                }
            }
        }

        private void updatetablesinGenArch(string tableId, string docID, string current)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            //string query = "select رقم_المرجع,docTable from TableGeneralArch WHERE الاسم=@الاسم";
            string query = "select ID,docTable from TableGeneralArch where رقم_المرجع=@رقم_المرجع and رقم_معاملة_القسم=@رقم_معاملة_القسم and docTable <> ' "+ current+ "' and  التاريخ  like '%-09-2022'";
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_المرجع", tableId);
            sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_معاملة_القسم", docID);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            if (dtbl.Rows.Count != 0)
            {
                foreach (DataRow dataRow in dtbl.Rows)
                {
                    if (current != dataRow["docTable"].ToString())
                    {
                        //MessageBox.Show(dataRow["ID"].ToString());
                        int id = Convert.ToInt32(dataRow["ID"].ToString());
                        UpdateState(id, "docTable", current, "TableGeneralArch");
                    }
                    //if (dataRow["docTable"].ToString() != "")
                    //{
                    //    //MessageBox.Show(dataRow["رقم_المرجع"].ToString() + " - " + dataRow["docTable"].ToString());
                    //    string name = getNames(dataRow["رقم_المرجع"].ToString(), dataRow["docTable"].ToString());
                    //    if (name != "")
                    //    {

                    //        sqlCon = new SqlConnection(DataSource);
                    //        query = "update TableGeneralArch set الاسم=N'" + name + "' where رقم_المرجع = '" + dataRow["رقم_المرجع"].ToString() + "'";
                    //        SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
                    //        if (sqlCon.State == ConnectionState.Closed)
                    //            sqlCon.Open();
                    //        sqlCmd.CommandType = CommandType.Text;
                    //        sqlCmd.ExecuteNonQuery();
                    //        sqlCon.Close();
                    //    }
                    //}
                    //MessageBox.Show(dataRow["رقم_المرجع"].ToString());

                }

            }
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
            string list = "";
            foreach (DataRow row in dtbl.Rows)
            {
                allList[i] = row["name"].ToString();
                list = list + Environment.NewLine +allList[i];
                i++;
            }
            //MessageBox.Show(list);
            return allList;

        }


        private void fillFileBox(int index, DataGridView dataGridView)
        {
            //Console.WriteLine("index  " + index);
            string str = "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            switch (index) {
                case 0:
                    str = "WafidVOrS";
                    break;
                case 1:
                    str = "WafidJeddViewOrSearch";
                    break;
                case 2:
                    str = "WafidMekkahViewOrSearch";
                    break;
                case 3:
                    str = "TarheelViewOrSearch";
                    break;
                case 4:
                    str = "TransferViewOrSearch";
                    break;
                case 5:
                    str = "WafidComViewOrSearch";
                    break;
                default: return;
            }
            SqlDataAdapter sqlDa = new SqlDataAdapter(str, sqlCon);
            
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            
            sqlDa.Fill(dtblMain);
            dataGridView.DataSource = dtblMain;
            dataGridView1.Sort(dataGridView1.Columns["ord"], System.ComponentModel.ListSortDirection.Ascending);
            int z = 0;
            string appName;
            int appID;
            //foreach (DataRow dataRow in dtbl.Rows)
            //{
                //appName = "محمد حامد محمد بلال";// dataRow["مقدم_الطلب"].ToString();
                //appID = 4226;// Convert.ToInt32(dataRow["ID"].ToString());
                //foreach (DataRow dataRow1 in dtbl.Rows)
                //{
                //    if (appName == dataRow1["مقدم_الطلب"].ToString() && appID != Convert.ToInt32(dataRow1["ID"].ToString()))
                //    {
                //        Console.WriteLine(dataRow1["ID"].ToString());
                //    }
                //}

            //}
            foreach (DataRow dataRow in dtblMain.Rows)
            {
                //UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), AffairIndex.ToString(), getFileName(AffairIndex)); ;

                bool found = false;

                for (int a = 0; a < combFileNo.Items.Count; a++)
                {
                    if (!string.IsNullOrEmpty(dataRow["رقم_الملف"].ToString()) && dataRow["رقم_الملف"].ToString() == combFileNo.Items[a].ToString())
                        found = true;
                }
                if (!found)
                {
                    Console.WriteLine("file no" + dataRow["رقم_الملف"].ToString());
                    combFileNo.Items.Add(dataRow["رقم_الملف"].ToString());
                }
                z++;
            }
            
            fileList = new int[combFileNo.Items.Count];
            for (int a = 0; a < combFileNo.Items.Count; a++)
            {
                if (combFileNo.Items[a].ToString() != "")
                    fileList[a] = Convert.ToInt32(combFileNo.Items[a].ToString());

            }
            combFileNo.Items.Clear();
            for (int b = 0; b < fileList.Length; b++)
                for (int a = 1; a < fileList.Length; a++)
                {
                    int first = fileList[a - 1];
                    if (fileList[a - 1] < fileList[a])
                    {
                        fileList[a - 1] = fileList[a];
                        fileList[a] = first;
                    }
                }

            int fileCo = 0;
            for (int a = 0; a < fileList.Length; a++)
                if (fileList[a] >=0)
                {
                    combFileNo.Items.Add(fileList[a]);
                    fileCo++;
                }

            z = 0;
            
            sqlCon.Close();
            
            
        }
        
        private int fillFileBox8(string table, string text)
        {            
            string str = "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();

            str = "SELECT رقم_الهوية from " + table + " WHERE رقم_الهوية = '" + text + "' and رقم_الملف = '" + combFileNo.Text+ "'";

            SqlDataAdapter sqlDa = new SqlDataAdapter(str, sqlCon);            
            sqlDa.SelectCommand.CommandType = CommandType.Text;            
            sqlDa.Fill(dtblMain);
            dataGridView8.DataSource = dtblMain;            
            sqlCon.Close();
            for (int rowIndex = 0; rowIndex < dataGridView8.Rows.Count - 1; rowIndex++)
            {
                if (dataGridView8.Rows[rowIndex].Cells["رقم_الهوية"].Value.ToString() == text)
                    return rowIndex;
            }
            return -1;
        }

        private int fillFileBoxSerach(int index)
        {
            //Console.WriteLine("index  " + index);
            string str = "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            switch (index)
            {
                case 0:
                    str = "WafidVOrS";
                    break;
                case 1:
                    str = "WafidJeddViewOrSearch";
                    break;
                case 2:
                    str = "WafidMekkahViewOrSearch";
                    break;
                case 3:
                    str = "TarheelViewOrSearch";
                    break;
                case 4:
                    str = "TransferViewOrSearch";
                    break;
                case 5:
                    str = "WafidComViewOrSearch";
                    break;
                default: return 0;
            }
            SqlDataAdapter sqlDa = new SqlDataAdapter(str, sqlCon);

            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;

            sqlDa.Fill(dtblMain);
            dataGridView6.DataSource = dtblMain;
            int z = 0;
            string appName;
            int appID;
            //foreach (DataRow dataRow in dtbl.Rows)
            //{
            //appName = "محمد حامد محمد بلال";// dataRow["مقدم_الطلب"].ToString();
            //appID = 4226;// Convert.ToInt32(dataRow["ID"].ToString());
            //foreach (DataRow dataRow1 in dtbl.Rows)
            //{
            //    if (appName == dataRow1["مقدم_الطلب"].ToString() && appID != Convert.ToInt32(dataRow1["ID"].ToString()))
            //    {
            //        Console.WriteLine(dataRow1["ID"].ToString());
            //    }
            //}

            //}
            foreach (DataRow dataRow in dtblMain.Rows)
            {
                //UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), AffairIndex.ToString(), getFileName(AffairIndex)); ;

                bool found = false;

                for (int a = 0; a < combFileNo.Items.Count; a++)
                {
                    if (dataRow["رقم_الملف"].ToString() != "" && dataRow["رقم_الملف"].ToString() != "0" && dataRow["رقم_الملف"].ToString() == combFileNo.Items[a].ToString())
                        found = true;
                }
                if (!found)
                {
                    combFileNo.Items.Add(dataRow["رقم_الملف"].ToString());
                }
                z++;
            }

            fileList = new int[combFileNo.Items.Count];
            for (int a = 0; a < combFileNo.Items.Count; a++)
            {
                if (combFileNo.Items[a].ToString() != "")
                    fileList[a] = Convert.ToInt32(combFileNo.Items[a].ToString());

            }
            combFileNo.Items.Clear();
            for (int b = 0; b < fileList.Length; b++)
                for (int a = 1; a < fileList.Length; a++)
                {
                    int first = fileList[a - 1];
                    if (fileList[a - 1] < fileList[a])
                    {
                        fileList[a - 1] = fileList[a];
                        fileList[a] = first;
                    }
                }

            int fileCo = 0;
            int aCount = 0;
            for (; aCount < fileList.Length; aCount++)
                if (fileList[aCount] != 0)
                {

                    comboBox3.Items.Add(fileList[aCount]);
                    Console.WriteLine("fileCo " + fileCo .ToString()+" - "+ SearchfileList[fileCo]);
                    fileCo++;
                }

            z = 0;

            sqlCon.Close();

            return fileCo;
        }

        private void FillCount1(int index)
        {
            //Console.WriteLine("index  " + index);
            string str = "WafidVOrS";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            switch (index)
            {
                case 0:
                    str = "WafidVOrS";
                    break;
                case 1:
                    str = "WafidJeddViewOrSearch";
                    break;
                case 2:
                    str = "WafidMekkahViewOrSearch";
                    break;
                case 3:
                    str = "TarheelViewOrSearch";
                    break;
                case 4:
                    str = "TransferViewOrSearch";
                    break;
                case 5:
                    str = "WafidComViewOrSearch";
                    break;
                default: return;
            }
            SqlDataAdapter sqlDa = new SqlDataAdapter(str, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            //sqlDa.SelectCommand.Parameters.AddWithValue("@جهة_العمل", index);
            DataTable dtbl = new DataTable();

            sqlDa.Fill(dtbl);
            
            dataGridView3.DataSource = dtbl;
            sqlCon.Close();

        }

        //private DataTable Table(this BindingSource bs)
        //{
        //    var bsFirst = bs;
        //    while (bsFirst.DataSource is BindingSource)
        //        bsFirst = (BindingSource)bsFirst.DataSource;

        //    DataTable dt;
        //    if (bsFirst.DataSource is DataSet)
        //        dt = ((DataSet)bsFirst.DataSource).Tables[bsFirst.DataMember];
        //    else if (bsFirst.DataSource is DataTable)
        //        dt = (DataTable)bsFirst.DataSource;
        //    else
        //        return null;

        //    if (bsFirst != bs)
        //    {
        //        if (dt.DataSet == null) return null;
        //        dt = dt.DataSet.Relations[bs.DataMember].ChildTable;
        //    }

        //    return dt;
        //}
        private void PersToServed_CheckedChanged(object sender, EventArgs e)
        {
            if (PersToServed.CheckState == CheckState.Unchecked) { 
                PersToServed.Text = "مقدم الطلب";
                pictureBox13.Visible = pictureBox11.Visible = false;
                //labelStatus.Visible = comboStatus.Visible = true;
            }
            else { 
                PersToServed.Text = "شخص إخر";
                radioButton3.Checked = true;
                //labelStatus.Visible = comboStatus.Visible = false;
                pictureBox13.Visible = pictureBox11.Visible = true;
            }
        }

        private void Panelapp_Paint(object sender, PaintEventArgs e)
        {
                         
        }

        private void Panelapp_Paint(string name, string relativity, string docno)
        {
            Label labelName = new Label();
            labelName.AutoSize = true;
            labelName.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labelName.Location = new System.Drawing.Point(650, 0);
            labelName.Name = "labelName_" + Nobox.ToString();
            labelName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labelName.Size = new System.Drawing.Size(98, 27);
            labelName.TabIndex = 94;
            labelName.Text = "الاسم: " + (Nobox+1).ToString() + ":";
            // 
            // AppName1
            // 
            TextBox AppName1 = new TextBox();
            AppName1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            AppName1.Location = new System.Drawing.Point(390, 3);
            AppName1.Name = "txtAppName_" + Nobox.ToString();
            AppName1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            AppName1.Size = new System.Drawing.Size(230, 35);
            AppName1.TabIndex = 93;
            AppName1.Tag = "valid";
            AppName1.Text = name;
            // 
            // labeltitle1
            // 
            Label labeltitle1 = new Label();
            labeltitle1.AutoSize = true;
            labeltitle1.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labeltitle1.Location = new System.Drawing.Point(344, 0);
            labeltitle1.Name = "labeltitle1_" + Nobox.ToString();
            labeltitle1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labeltitle1.Size = new System.Drawing.Size(40, 27);
            labeltitle1.TabIndex = 176;
            labeltitle1.Text = "النوع:";
            
            // 
            // checkSexType1
            // 
            CheckBox checkBoxSex = new CheckBox();
            checkBoxSex.AutoSize = true;
            checkBoxSex.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            checkBoxSex.Location = new System.Drawing.Point(289, 3);
            checkBoxSex.Name = "checkBoxSex_" + Nobox.ToString();
            checkBoxSex.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            checkBoxSex.Size = new System.Drawing.Size(49, 31);
            checkBoxSex.TabIndex = 177;
            checkBoxSex.Text = "ذكر";
            checkBoxSex.CheckedChanged += new System.EventHandler(this.checkBoxSexe_CheckedChanged);
            checkBoxSex.UseVisualStyleBackColor = true;
            checkBoxSex.Tag = "valid";
            checkBoxSex.CheckState = CheckState.Unchecked;
            if (relativity == "ابنة" || relativity == "زوجة" || relativity == "اخت" || relativity == "ام") {
                checkBoxSex.Text = "انثى";
                checkBoxSex.CheckState = CheckState.Checked;
            }
            // 
            // label8
            // 
            Label labeldoctype = new Label();
            labeldoctype.AutoSize = true;
            labeldoctype.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labeldoctype.Location = new System.Drawing.Point(165, 0);
            labeldoctype.Name = "labeldoctype_" + Nobox.ToString();
            labeldoctype.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labeldoctype.Size = new System.Drawing.Size(118, 27);
            labeldoctype.TabIndex = 117;
            labeldoctype.Text = "صلة القرابة:";
            
            // 
            // comboBoxDocType
            // 
            ComboBox comboBoxDocType = new ComboBox();
            comboBoxDocType.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            comboBoxDocType.FormattingEnabled = true;
            comboBoxDocType.Items.AddRange(new object[] {"ابن","ابنة","زوج","زوجة","اخ","اخت","اب","ام"});
            comboBoxDocType.Location = new System.Drawing.Point(10, 3);
            comboBoxDocType.Name = "comboBoxDocType_" + Nobox.ToString();
            comboBoxDocType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            comboBoxDocType.Size = new System.Drawing.Size(120, 35);
            comboBoxDocType.Tag = "valid";
            comboBoxDocType.Text = relativity;
            // 
            // labeldoctype1
            // 
            Label labeldocNo = new Label();
            labeldocNo.AutoSize = true;
            labeldocNo.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            labeldocNo.Location = new System.Drawing.Point(633, 41);
            labeldocNo.Name = "labeldocNo_" + Nobox.ToString();
            labeldocNo.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            labeldocNo.Size = new System.Drawing.Size(115, 27);
            labeldocNo.TabIndex = 119;
            labeldocNo.Text = "رقم اثبات الشخصية:";
           
            labeldocNo.TabIndex = 122;
            
            // 
            // DocNo1
            // 
            TextBox textBoxDocNo = new TextBox();
            textBoxDocNo.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            textBoxDocNo.Location = new System.Drawing.Point(444, 44);
            textBoxDocNo.Name = "textBoxDocNo_" + Nobox.ToString();
            textBoxDocNo.RightToLeft = System.Windows.Forms.RightToLeft.No;
            textBoxDocNo.Size = new System.Drawing.Size(102, 35);
            textBoxDocNo.TabIndex = 120;
            textBoxDocNo.Tag = "pass";
            textBoxDocNo.Tag = "valid";
            textBoxDocNo.Text = docno;

            PictureBox picadd = new PictureBox();
            picadd.Image = global::PersAhwal.Properties.Resources.add;
            picadd.Location = new System.Drawing.Point(872, 3);
            picadd.Name = "picadd_" + Nobox.ToString();
            picadd.Size = new System.Drawing.Size(28, 30);
            picadd.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            picadd.TabIndex = 819;
            picadd.Visible = true;
            picadd.Click += new System.EventHandler(this.pictureBoxadd_Click);
            // 
            // pictremove
            // 
            PictureBox pictremove = new PictureBox();
            pictremove.Image = global::PersAhwal.Properties.Resources.remove;
            pictremove.Location = new System.Drawing.Point(838, 3);
            pictremove.Name = "pictremove_" + Nobox.ToString();
            pictremove.Size = new System.Drawing.Size(28, 30);
            pictremove.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            pictremove.TabIndex = 820;
            pictremove.Visible = true;
            pictremove.Click += new System.EventHandler(this.pictureBoxremove_Click);

            PaneTransfer.Controls.Add(labelName);
            PaneTransfer.Controls.Add(AppName1);
            PaneTransfer.Controls.Add(labeltitle1);
            PaneTransfer.Controls.Add(checkBoxSex);
            PaneTransfer.Controls.Add(labeldoctype);
            PaneTransfer.Controls.Add(comboBoxDocType);
            PaneTransfer.Controls.Add(labeldocNo);
            PaneTransfer.Controls.Add(textBoxDocNo);
            PaneTransfer.Controls.Add(picadd);
            PaneTransfer.Controls.Add(pictremove);
            Nobox++;
            
        }
        public void pictureBoxadd_Click(object sender, EventArgs e)
        {
            Panelapp_Paint("", "", "");
        }
        public void pictureBoxremove_Click(object sender, EventArgs e)
        {
            PictureBox pictureBox = (PictureBox)sender;
            string index = pictureBox.Name.Split('_')[1];
            //MessageBox.Show(index);
            foreach (Control control in PaneTransfer.Controls)
            {
                if (control.Name.Split('_')[1] == index)
                {
                    control.Visible = false;
                    control.Tag = "Unvalid";
                    control.Text = "";
                }
            }
            //Authcases();
        }
        private void checkBoxSexe_CheckedChanged(object sender, EventArgs e)
        {
            
            CheckBox checkBoxSexe = (CheckBox)sender;
            if (checkBoxSexe.CheckState == CheckState.Unchecked) checkBoxSexe.Text = "ذكر";
            else checkBoxSexe.Text = "أنثى";
        }
        private void pictureBox11_Click(object sender, EventArgs e)
        {
            Panelapp_Paint("", "", ""); getText(DataSource);
        }

        private void pictureBox13_Click(object sender, EventArgs e)
        {
            
            //foreach (Control control in Panelapp.Controls)
            //{
            //    if (control.Name.Split('_')[1] == (Nobox-1).ToString())
            //    {
            //        control.Visible = false;
            //        control.Tag = "Unvalid";
            //        Nobox--;
            //    }
            //}
            //getText(DataSource);
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
        private void FormType_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void comPurpose_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        
            private void CreateGroupFileTarheel(string ActiveCopy, string fileNo)
        {
            string fileCount = (dataGridView1.RowCount - 1).ToString();
            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilesPathIn + AffairIndex.ToString() + ".docx";
            route = FilesPathIn + "قائمة الترحيل.docx";
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
            Word.Range BookFileNo = oBDoc.Bookmarks.get_Item(ref ParaFileNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookGreData1 = oBDoc.Bookmarks.get_Item(ref ParaGreData1).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            Word.Range BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;
            Word.Range BookPurposeText = oBDoc.Bookmarks.get_Item(ref ParaPurposeText).Range;
            Word.Range BookPurText = oBDoc.Bookmarks.get_Item(ref ParaPurText).Range;
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;
            Word.Range BookIndivNo = oBDoc.Bookmarks.get_Item(ref ParaIndivNo).Range;

            BookDestin.Text = " مدير إدارة الوافدين";
            



            BookGreData1.Text = GregorianDate.Text;
            BookGreData.Text = GregorianDate.Text;
            BookHijriData.Text = HijriDate.Text;            
            BookvConsul.Text = AttendViceConsul.Text;

           
            int indexNo = 0;
            //deleteRowsDataTicket();
            Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
            table.Rows[1].Cells[4].Range.Text = "الحالة";
            if (PersToServed.CheckState == CheckState.Unchecked)
            {
                int i = 0;
                for (int x = 0; x < dataGridView1.RowCount-1; x++)
                {
                    string AppNames = dataGridView1.Rows[x].Cells[2].Value.ToString();
                    string isReady =  dataGridView1.Rows[x].Cells["ready"].Value.ToString();
                    if (AppNames != "" && isReady == "مكتمل")
                    {
                        table.Rows.Add();
                        
                        table.Rows[indexNo + 2].Cells[1].Range.Text = (indexNo + 1).ToString() + ".";
                        table.Rows[indexNo + 2].Cells[2].Range.Text = AppNames;
                        table.Rows[indexNo + 2].Cells[3].Range.Text = dataGridView1.Rows[x].Cells[5].Value.ToString(); //AppDocumentNo[x];
                        table.Rows[indexNo + 2].Cells[4].Range.Text = dataGridView1.Rows[x].Cells["الحالة"].Value.ToString(); //status[x];
                        table.Rows[indexNo + 2].Cells[5].Range.Text = dataGridView1.Rows[x].Cells["رقم_هاتف1"].Value.ToString(); //phone[x];
                        indexNo++;
                    }

                }
            }

            //string FileNo = indexNo.ToString();
            BookIndivNo.Text = indexNo.ToString();
            BookPurText.Text = "(" + fileNo + ")";
            BookFileNo.Text = fileCount + fileNo;
            BookPurposeText.Text = fileNo;
            
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

            string docxouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(docxouput);
            //File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
        }private void CreateGroupFileTrans(string ActiveCopy, string fileNo, bool alldata)
        {
            string fileCount = (dataGridView1.RowCount - 1).ToString();
            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilesPathIn + AffairIndex.ToString() + ".docx";
            route = FilesPathIn + "قائمة الترحيل.docx";
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
            Word.Range BookFileNo = oBDoc.Bookmarks.get_Item(ref ParaFileNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookGreData1 = oBDoc.Bookmarks.get_Item(ref ParaGreData1).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            Word.Range BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;
            Word.Range BookPurposeText = oBDoc.Bookmarks.get_Item(ref ParaPurposeText).Range;
            Word.Range BookPurText = oBDoc.Bookmarks.get_Item(ref ParaPurText).Range;
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;
            Word.Range BookIndivNo = oBDoc.Bookmarks.get_Item(ref ParaIndivNo).Range;

            BookDestin.Text = " مدير إدارة الوافدين";
            



            BookGreData1.Text = GregorianDate.Text;
            BookGreData.Text = GregorianDate.Text;
            BookHijriData.Text = HijriDate.Text;            
            BookvConsul.Text = AttendViceConsul.Text;

           
            int indexNo = 0;
            //deleteRowsDataTicket();
            Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
            table.Rows[1].Cells[4].Range.Text = "الوصف/العلاقة";
            if (PersToServed.CheckState == CheckState.Unchecked)
            {
                int lines = 0;
                for (int x = 0; ; x++)
                {
                    string AppNames = dataGridView1.Rows[lines].Cells[2].Value.ToString();
                    string status = dataGridView1.Rows[lines].Cells["الحالة"].Value.ToString();                    
                    if (AppNames != "")
                    {
                        table.Rows.Add();
                        indexNo++;
                        table.Rows[x + 2].Cells[1].Range.Text = (x + 1).ToString() + ".";
                        table.Rows[x + 2].Cells[2].Range.Text = AppNames;
                        table.Rows[x + 2].Cells[3].Range.Text = dataGridView1.Rows[lines].Cells[5].Value.ToString(); //AppDocumentNo[x];
                        table.Rows[x + 2].Cells[4].Range.Text = "إنتهاء الإقامة"; //status[x];
                        table.Rows[x + 2].Cells[5].Range.Text = dataGridView1.Rows[lines].Cells["رقم_هاتف1"].Value.ToString(); //phone[x];
                        if (alldata)
                        {
                            string famillist = dataGridView1.Rows[lines].Cells[7].Value.ToString();
                            string Relativity = dataGridView1.Rows[lines].Cells[8].Value.ToString();
                            string DocumentNo = dataGridView1.Rows[lines].Cells[9].Value.ToString();
                            //MessageBox.Show(AppNames + " - "+ famillist +" - "+ famillist.Split('_').Length.ToString());
                            if (famillist != "")
                            {
                                if (famillist.Contains("_"))
                                {
                                    
                                    for (int i = 0; i < famillist.Split('_').Length; i++)
                                    {
                                        if (famillist.Split('_')[i] != "")
                                        {
                                            table.Rows.Add();
                                            x++;
                                            indexNo++;
                                            table.Rows[x + 2].Cells[1].Range.Text = (x + 1).ToString() + ".";
                                            table.Rows[x + 2].Cells[2].Range.Text = famillist.Split('_')[i];
                                            table.Rows[x + 2].Cells[3].Range.Text = DocumentNo.Split('_')[i];
                                            table.Rows[x + 2].Cells[4].Range.Text = Relativity.Split('_')[i];
                                            table.Rows[x + 2].Cells[5].Range.Text = dataGridView1.Rows[lines].Cells["رقم_هاتف1"].Value.ToString();
                                        }
                                    }
                                }
                                else
                                {
                                    if (famillist != "")
                                    {
                                        table.Rows.Add();
                                        x++;
                                        indexNo++;
                                        table.Rows[x + 2].Cells[1].Range.Text = (x + 1).ToString() + ".";
                                        table.Rows[x + 2].Cells[2].Range.Text = famillist;
                                        table.Rows[x + 2].Cells[3].Range.Text = DocumentNo;
                                        table.Rows[x + 2].Cells[4].Range.Text = Relativity;
                                        table.Rows[x + 2].Cells[5].Range.Text = dataGridView1.Rows[lines].Cells["رقم_هاتف1"].Value.ToString();
                                    }
                                }
                            }
                        }
                    }                    
                    if (lines == dataGridView1.RowCount - 2) break;
                    lines++;
                }
            }

            //string FileNo = indexNo.ToString();
            BookIndivNo.Text = indexNo.ToString();
            BookPurText.Text = "(" + fileNo + ")";
            BookFileNo.Text = fileCount + fileNo;
            BookPurposeText.Text = fileNo;
            
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

            string docxouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(docxouput);
            //File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
        }
        private void CreateGroupFile(string ActiveCopy, string fileNo)
        {
            string fileCount = (dataGridView1.RowCount - 1).ToString();
            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilesPathIn + AffairIndex.ToString()+ ".docx";
            if (AffairIndex == 1 || AffairIndex == 2)
                route = FilesPathIn + "قائم للوافدين.docx";
                System.IO.File.Copy(route, ActiveCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

            object ParaGreData = "MarkGreData";
            object ParaGreData1 = "MarkGreDa1";
            object ParaHijriData = "MarkHijriData";
            object ParaFileNo= "MarkFileNo";
            object ParaDest = "MarkDest";
            object ParaPurpose = "MarkPurpose";//message لعناية ... note verbal فارغ
            object ParaPurposeText = "MarkPurposeText";
            object ParaPurText = "MarkPurText";
            object ParavConsul = "MarkViseConsul";  //note verbal لتعرب       


            Word.Range BookDestin = oBDoc.Bookmarks.get_Item(ref ParaDest).Range;
            Word.Range BookFileNo = oBDoc.Bookmarks.get_Item(ref ParaFileNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookGreData1 = oBDoc.Bookmarks.get_Item(ref ParaGreData1).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            Word.Range BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;
            Word.Range BookPurposeText = oBDoc.Bookmarks.get_Item(ref ParaPurposeText).Range;
            Word.Range BookPurText = oBDoc.Bookmarks.get_Item(ref ParaPurText).Range;
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;            
            
            BookDestin.Text = " مدير إدارة الوافدين";

            

            BookGreData1.Text= GregorianDate.Text; 
            BookGreData.Text = GregorianDate.Text;
            BookHijriData.Text = HijriDate.Text;
            BookPurText.Text = fileNo; 
            BookPurposeText.Text = fileNo;            
            BookvConsul.Text = AttendViceConsul.Text;

            object rangeDesin = BookDestin;
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
            for (int x = 0; x < dataGridView1.RowCount - 1; x++)
            {
                string AppNames = dataGridView1.Rows[x].Cells[2].Value.ToString();
                if (AppNames != "")
                {
                    table.Rows.Add();
                    indexNo++;
                    table.Rows[x + 2].Cells[1].Range.Text = (x + 1).ToString() + ".";
                    table.Rows[x + 2].Cells[2].Range.Text = AppNames;// AppNames[x];
                    table.Rows[x + 2].Cells[3].Range.Text = dataGridView1.Rows[x].Cells[5].Value.ToString();

                }

            }

            if (AffairIndex == 1)
            {
                string FileNo = indexNo.ToString() + fileNo + " (" + "جدة" + ")";
                BookFileNo.Text = FileNo;
            }
            else if (AffairIndex == 2)
            {
                string FileNo = indexNo.ToString() + fileNo + " (" + "مكة المكرمة" + ")";
                BookFileNo.Text = FileNo;
            }
            else if (AffairIndex == 3)
            {
                string FileNo = indexNo.ToString() + fileNo + " (" + "الترحيل" + ")";
                BookFileNo.Text = FileNo;
            }
            else 
                BookFileNo.Text = indexNo.ToString() + fileNo;

            BookPurpose.Text = indexNo.ToString();
            oBDoc.Bookmarks.Add("MarkDest", ref rangeDesin);
            oBDoc.Bookmarks.Add("MarkPurpose", ref rangePurpose);
            oBDoc.Bookmarks.Add("MarkFileNo", ref rangeFileNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkGreDa1", ref rangeGreData1);
            oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);
            oBDoc.Bookmarks.Add("MarkPurposeText", ref rangePurposeText);
            oBDoc.Bookmarks.Add("MarkPurText", ref rangePurText);
            oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);

            string docxouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(docxouput);
            //File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
        }
        
        private void CreateGroupFileJeddah1(string ActiveCopy, string fileNo, string destin)
        {
            string fileCount = (dataGridView1.RowCount - 1).ToString();
            string FileNo = fileCount + fileNo + " (" + "جدة" + ")";

            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilesPathIn + AffairIndex.ToString()+ ".docx";
            if (AffairIndex == 1 || AffairIndex == 2)
                route = FilesPathIn + "قائم للوافدين جدة1.docx";
                System.IO.File.Copy(route, ActiveCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

            object ParaGreData = "MarkGreData";
            object ParafileCount = "MarkfileCount";
            
            object ParaHijriData = "MarkHijriData";
            object ParaFileNo= "MarkFileNo";
            object ParaDest = "MarkDest";
            //object ParaPurpose = "MarkPurpose";//message لعناية ... note verbal فارغ
            object ParaPurposeText = "MarkPurposeText";
            
            object ParavConsul = "MarkViseConsul";  //note verbal لتعرب       


            Word.Range BookDestin = oBDoc.Bookmarks.get_Item(ref ParaDest).Range;
            Word.Range BookfileCount = oBDoc.Bookmarks.get_Item(ref ParafileCount).Range;
            Word.Range BookFileNo = oBDoc.Bookmarks.get_Item(ref ParaFileNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;            
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            //Word.Range BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;
            Word.Range BookPurposeText = oBDoc.Bookmarks.get_Item(ref ParaPurposeText).Range;
            
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;            
            
            BookDestin.Text = destin;
            BookfileCount.Text = fileCount;

            BookGreData.Text = GregorianDate.Text;
            BookHijriData.Text = HijriDate.Text;            
            BookPurposeText.Text = fileNo;            
            BookvConsul.Text = AttendViceConsul.Text;
            BookFileNo.Text = FileNo;
            //BookPurpose.Text = fileCount;

            object rangeDesin = BookDestin;
            //object rangePurpose = BookPurpose;
            object rangeFileNo = BookFileNo;
            object rangeGreData = BookGreData;            
            object rangeHijriData = BookHijriData;
            object rangePurposeText = BookPurposeText;            
            object rangevConsul = BookvConsul;
            object rangefileCount = BookfileCount;
            
            

            oBDoc.Bookmarks.Add("MarkDest", ref rangeDesin);
            //oBDoc.Bookmarks.Add("MarkPurpose", ref rangePurpose);
            oBDoc.Bookmarks.Add("MarkFileNo", ref rangeFileNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkfileCount", ref rangefileCount);
            
            oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);
            oBDoc.Bookmarks.Add("MarkPurposeText", ref rangePurposeText);
            oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);

            string docxouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(docxouput);
            //File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
        }
        private void CreateGroupFileJeddah2(string ActiveCopy, string fileNo)
        {
            string fileCount = (dataGridView1.RowCount - 1).ToString();
            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilesPathIn + AffairIndex.ToString()+ ".docx";
            if (AffairIndex == 1 || AffairIndex == 2)
                route = FilesPathIn + "قائم للوافدين جدة2.docx";
                System.IO.File.Copy(route, ActiveCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

            object ParaGreData1 = "MarkGreDa1";
            object ParaPurText = "MarkPurText";     


            Word.Range BookGreData1 = oBDoc.Bookmarks.get_Item(ref ParaGreData1).Range;
            Word.Range BookPurText = oBDoc.Bookmarks.get_Item(ref ParaPurText).Range;           
            
            BookGreData1.Text= GregorianDate.Text; 
           
            BookPurText.Text = fileNo; 
            
            object rangeGreData1 = BookGreData1;
            object rangePurText = BookPurText;
            int indexNo = 0;
            
            Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
            for (int x = 0; x < dataGridView1.RowCount - 1; x++)
            {
                string AppNames = dataGridView1.Rows[x].Cells[2].Value.ToString();
                if (AppNames != "")
                {
                    table.Rows.Add();
                    indexNo++;
                    table.Rows[x + 2].Cells[1].Range.Text = (x + 1).ToString() + ".";
                    table.Rows[x + 2].Cells[2].Range.Text = AppNames;// AppNames[x];
                    table.Rows[x + 2].Cells[3].Range.Text = dataGridView1.Rows[x].Cells[5].Value.ToString();

                }

            }

            if (AffairIndex == 1)
            {
                string FileNo = indexNo.ToString() + fileNo + " (" + "جدة" + ")";
               
            }
            else if (AffairIndex == 2)
            {
                string FileNo = indexNo.ToString() + fileNo + " (" + "مكة المكرمة" + ")";
                
            }
            else if (AffairIndex == 3)
            {
                string FileNo = indexNo.ToString() + fileNo + " (" + "الترحيل" + ")";
                
            }
            oBDoc.Bookmarks.Add("MarkGreDa1", ref rangeGreData1);
            oBDoc.Bookmarks.Add("MarkPurText", ref rangePurText);
            
            string docxouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(docxouput);
            //File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
        }

        private void CreateWordFile(string ActiveCopy, string status)
        {
            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilesPathIn + AffairIndex.ToString()+ ".docx";
            route = FilesPathIn + "الوافدين.docx";
            System.IO.File.Copy(route, ActiveCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

            object ParaGreData = "MarkGreData";
            object ParaHijriData = "MarkHijriData";
            object ParaNo = "MarkNo";
            object ParaDest = "MarkDest";
            object ParaPurpose = "MarkPurpose";//message لعناية ... note verbal فارغ
            object ParaPurposeText = "MarkPurposeText";
            object ParavConsul = "MarkViseConsul";  //note verbal لتعرب       
            

            Word.Range BookDestin = oBDoc.Bookmarks.get_Item(ref ParaDest).Range;
            Word.Range BookNo = oBDoc.Bookmarks.get_Item(ref ParaNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            Word.Range BookPurpose = oBDoc.Bookmarks.get_Item(ref ParaPurpose).Range;            
            Word.Range BookPurposeText = oBDoc.Bookmarks.get_Item(ref ParaPurposeText).Range;
            Word.Range BookvConsul = oBDoc.Bookmarks.get_Item(ref ParavConsul).Range;
            Word.Range BookApplicantName2;

            
            string Auth = "";
            string str = "";
            if (ApplicantSex.Text != "ذكر") str = "ة";

            //if (FormType.Text == "إفادة لمن يهمه الأمر")
            //{
            //    Auth = "قد حررت هذه الإفادة بناءً على طلب المذكور" + str + " أعلاه لاستخدامها على الوجه المشروع";
            //}
            //else if (FormType.Text == "شهادة عدم ممانعة" || FormType.Text == "شهادة لمن يهمه الأمر")
            //    Auth = "قد حررت هذه الشهادة بناءً على طلب المذكور" + str + "أعلاه لاستخدامها على الوجه المشروع";

            BookDestin.Text = DocDestin.Text;
            BookPurpose.Text = "إجراء تسهيل سفر";
            BookNo.Text = txtId.Text;
            BookGreData.Text = GregorianDate.Text;
            BookHijriData.Text = HijriDate.Text;
            BookPurposeText.Text = text.Text;
            BookvConsul.Text = AttendViceConsul.Text;

            object rangeDesin = BookDestin;
            object rangePurpose = BookPurpose;
            object rangeIqrarNo = BookNo;
            object rangeGreData = BookGreData;
            object rangeHijriData = BookHijriData;
            object rangePurposeText = BookPurposeText;
            object rangevConsul = BookvConsul;

            Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
            table.Rows[1].Cells[4].Range.Text = status;
            int startIndex = 0;
            if (PersToServed.CheckState == CheckState.Unchecked)
            {

                table.Rows.Add();
                table.Rows[2].Cells[1].Range.Text = "1.";
                table.Rows[2].Cells[2].Range.Text = ApplicantName.Text;
                table.Rows[2].Cells[3].Range.Text = DocNo.Text;
                table.Rows[2].Cells[4].Range.Text = comboStatus.Text;
            }
            else
            {
                var selectedOption = MessageBox.Show("", "إضافة رب الأسرة إلى القائمة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                {
                    table.Rows.Add();
                    table.Rows[2].Cells[1].Range.Text = "1.";
                    table.Rows[2].Cells[2].Range.Text = ApplicantName.Text;
                    table.Rows[2].Cells[3].Range.Text = DocNo.Text;
                    table.Rows[2].Cells[4].Range.Text = comboStatus.Text;
                    startIndex++;
                }
                for (int x = 0; x < Nobox; x++)
                {
                    if (Names[x] != "")
                    {
                        table.Rows.Add();
                        table.Rows[startIndex + 2].Cells[1].Range.Text = (startIndex + 1).ToString() + ".";
                        table.Rows[startIndex + 2].Cells[2].Range.Text = Names[x];
                        table.Rows[startIndex + 2].Cells[3].Range.Text = DocumentNo[x];
                        table.Rows[startIndex + 2].Cells[4].Range.Text = Relativity[x];
                        startIndex++;
                    }
                }

            }
            oBDoc.Bookmarks.Add("MarkDest", ref rangeDesin);
            oBDoc.Bookmarks.Add("MarkPurpose", ref rangePurpose);
            oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);
            oBDoc.Bookmarks.Add("MarkPurposeText", ref rangePurposeText);
            oBDoc.Bookmarks.Add("MarkViseConsul", ref rangevConsul);

            string docxouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(docxouput);
            //File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
        }

        private void IqrarTravel(string ActiveCopy)
        {
            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilesPathIn + AffairIndex.ToString()+ ".docx";
            route = FilesPathIn + "إقرار ترحيل.docx";
            System.IO.File.Copy(route, ActiveCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);
            
            object ParaGreData = "MarkGreData";
            object ParaSex = "MarkSex";
            object ParaHijriData = "MarkHijriData";
            object ParaNo = "MarkIqrarNo";
            object ParaApplicantName1 = "MarkApplicantName1";
            object ParaApplicantName2 = "MarkApplicantName2";
            object ParaPassIqama = "MarkPassIqama";
            object ParaAppliigamaNo = "MarkAppliigamaNo";
            

            Word.Range BookSex = oBDoc.Bookmarks.get_Item(ref ParaSex).Range;
            Word.Range BookApplicantName1 = oBDoc.Bookmarks.get_Item(ref ParaApplicantName1).Range;
            Word.Range BookApplicantName2 = oBDoc.Bookmarks.get_Item(ref ParaApplicantName2).Range;
            Word.Range BookNo = oBDoc.Bookmarks.get_Item(ref ParaNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            Word.Range BookPassIqama = oBDoc.Bookmarks.get_Item(ref ParaPassIqama).Range;
            Word.Range BookAppliigamaNo = oBDoc.Bookmarks.get_Item(ref ParaAppliigamaNo).Range;
            
            


            string Auth = "";
            string str = "";
            if (ApplicantSex.Text != "ذكر") str = "ة";
            BookSex.Text = str;
            BookApplicantName2.Text = BookApplicantName1.Text = ApplicantName.Text;
            BookPassIqama.Text = DocType.Text;
            BookNo.Text = txtId.Text;
            BookGreData.Text = GregorianDate.Text;
            BookHijriData.Text = HijriDate.Text;
            BookAppliigamaNo.Text = DocNo.Text;

            object rangeSex = BookSex;
            object rangeApplicantName1 = BookApplicantName1;
            object rangeApplicantName2 = BookApplicantName2;
            object rangePassIqama = BookPassIqama;
            object rangeIqrarNo = BookNo;
            object rangeGreData = BookGreData;
            object rangeHijriData = BookHijriData;
            object rangeAppliigamaNo = BookAppliigamaNo;

            oBDoc.Bookmarks.Add("MarkSex", ref rangeSex);
            oBDoc.Bookmarks.Add("MarkApplicantName1", ref rangeApplicantName1);
            oBDoc.Bookmarks.Add("MarkApplicantName2", ref rangeApplicantName2);
            oBDoc.Bookmarks.Add("MarkPassIqama", ref rangePassIqama);
            oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);
            oBDoc.Bookmarks.Add("MarkAppliigamaNo", ref rangeAppliigamaNo);
            
            string docxouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(docxouput);
            //File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
        }

        private void IqrarFinalExit(string ActiveCopy)
        {
            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilesPathIn + AffairIndex.ToString() + ".docx";
            route = FilesPathIn + "إقرار خروج نهائي.docx";
            System.IO.File.Copy(route, ActiveCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

            object ParaGreData = "MarkGreData";
            object ParaSex = "MarkSex";
            object ParaHijriData = "MarkHijriData";
            object ParaNo = "MarkIqrarNo";
            object ParaApplicantName1 = "MarkApplicantName1";
            object ParaApplicantName2 = "MarkApplicantName2";
            object ParaPassIqama = "MarkPassIqama";
            object ParaAppliigamaNo = "MarkAppliigamaNo";
            object ParaAppIssSource = "MarkAppIssSource";
            object ParaAuthorization = "MarkAuthorization";
            object ParaACV = "MarkACV";

            Word.Range BookACV = oBDoc.Bookmarks.get_Item(ref ParaACV).Range;
            Word.Range BookSex = oBDoc.Bookmarks.get_Item(ref ParaSex).Range;
            Word.Range BookApplicantName1 = oBDoc.Bookmarks.get_Item(ref ParaApplicantName1).Range;
            Word.Range BookApplicantName2 = oBDoc.Bookmarks.get_Item(ref ParaApplicantName2).Range;
            Word.Range BookNo = oBDoc.Bookmarks.get_Item(ref ParaNo).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            Word.Range BookPassIqama = oBDoc.Bookmarks.get_Item(ref ParaPassIqama).Range;
            Word.Range BookAppliigamaNo = oBDoc.Bookmarks.get_Item(ref ParaAppliigamaNo).Range;
            Word.Range BookAuthorization = oBDoc.Bookmarks.get_Item(ref ParaAuthorization).Range;


            string str = "";
            string Auth= "";
            if (comboBox2.SelectedIndex == 0)
            {
                if (ApplicantSex.CheckState == CheckState.Unchecked)
                    Auth = "أشهد أنا/" + AttendViceConsul.Text + " نائب قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكور أعلاه قد حضر للقنصلية ووقع بتوقيعه على هذا الإقرار بعد تلاوته عليه وبعد أن فهم مضمونه ومحتواه. ";
                else
                {
                    Auth = "أشهد أنا/ " + AttendViceConsul.Text + " نائب قنصل بالقنصلية العامة لجمهورية السودان بجدة، بأن المذكورة أعلاه قد حضرت للقنصلية ووقعت بتوقيعها على هذا الإقرار بعد تلاوته عليها وبعد أن فهمت مضمونه ومحتواه. ";                    
                }
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                string[] strmandoub = new string[2];
                strmandoub = mandoubName.Text.Split('-');
                if (ApplicantSex.Text == "ذكر")

                    if (strmandoub[1].Trim() != "القنصلية العامة لجمهورية السودان بجدة")
                    {
                        Auth = "المواطن المذكور أعلاه حضر ووقع بتوقيعه على هذا الإقرار أمام مندوب جالية منطقة" + strmandoub[1] + " السيد/ " + strmandoub[0] + "، وذلك بموجب التفويض الممنوح له من القنصلية العامة، ";
                    }
                    else { Auth = "المواطن المذكور أعلاه حضر ووقع بتوقيعه على هذا الإقرار أمام مندوب " + strmandoub[1] + " السيد/ " + strmandoub[0] + "، وذلك بموجب التفويض الممنوح له من القنصلية العامة، "; }
                if (ApplicantSex.Text == "أنثى")
                {
                    if (strmandoub[1].Trim() != "القنصلية العامة لجمهورية السودان بجدة")
                    {
                        Auth = "المواطنة المذكورة أعلاه حضرت ووقعت بتوقيعها على هذا الإقرار أمام مندوب جالية منطقة" + strmandoub[1] + " السيد/ " + strmandoub[0] + "، وذلك بموجب التفويض الممنوح له من القنصلية العامة، ";
                    }
                    else { Auth = "المواطنة المذكورة أعلاه حضرت ووقعت بتوقيعها على هذا الإقرار أمام مندوب " + strmandoub[1] + " السيد/ " + strmandoub[0] + "، وذلك بموجب التفويض الممنوح له من القنصلية العامة، "; }
                }



            }


            
            if (ApplicantSex.Text != "ذكر") str = "ة";
            BookAuthorization.Text = Auth;
            BookACV.Text = AttendViceConsul.Text;
            BookSex.Text = str;
            BookApplicantName2.Text = BookApplicantName1.Text = ApplicantName.Text;
            BookPassIqama.Text = DocType.Text;
            BookNo.Text = txtId.Text;
            BookGreData.Text = GregorianDate.Text;
            BookHijriData.Text = HijriDate.Text;
            BookAppliigamaNo.Text = DocNo.Text;

            object rangeAuthorization = BookAuthorization;
            object rangeSex = BookSex; 
            object rangeACV = BookACV;
            object rangeApplicantName1 = BookApplicantName1;
            object rangeApplicantName2 = BookApplicantName2;
            object rangePassIqama = BookPassIqama;
            object rangeIqrarNo = BookNo;
            object rangeGreData = BookGreData;
            object rangeHijriData = BookHijriData;
            object rangeAppliigamaNo = BookAppliigamaNo;

            oBDoc.Bookmarks.Add("MarkAuthorization", ref rangeAuthorization);
            oBDoc.Bookmarks.Add("MarkACV", ref rangeACV);
            oBDoc.Bookmarks.Add("MarkSex", ref rangeSex);
            oBDoc.Bookmarks.Add("MarkApplicantName1", ref rangeApplicantName1);
            oBDoc.Bookmarks.Add("MarkApplicantName2", ref rangeApplicantName2);
            oBDoc.Bookmarks.Add("MarkPassIqama", ref rangePassIqama);
            oBDoc.Bookmarks.Add("MarkIqrarNo", ref rangeIqrarNo);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);
            oBDoc.Bookmarks.Add("MarkAppliigamaNo", ref rangeAppliigamaNo);

            string docxouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(docxouput);
            //File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
        }
        private void CreateWokOffice(string ActiveCopy)
        {
            string ReportName = DateTime.Now.ToString("mmss");
            string route = FilesPathIn + AffairIndex.ToString()+ ".docx";
            route = FilesPathIn + "استمارة مكتب العمل.docx";
            System.IO.File.Copy(route, ActiveCopy);
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object Routseparameter = ActiveCopy;
            Word.Document oBDoc = oBMicroWord.Documents.Open(Routseparameter, oBMiss);

            
            object ParaName1 = "name1";
            object ParaName2 = "name2";
            object ParaDate1 = "date1";
            object ParaDate2 = "date2";
            object ParaJob = "job";
            object Parareasons = "reasons";
            object ParaWorkPlace = "workPlace";
            object ParaDocNo = "DocNo";
            
            Word.Range BookeName1 = oBDoc.Bookmarks.get_Item(ref ParaName1).Range;
            Word.Range BookName2 = oBDoc.Bookmarks.get_Item(ref ParaName2).Range;
            Word.Range BookDate1 = oBDoc.Bookmarks.get_Item(ref ParaDate1).Range;
            Word.Range BookDate2 = oBDoc.Bookmarks.get_Item(ref ParaDate2).Range;
            Word.Range BookJob = oBDoc.Bookmarks.get_Item(ref ParaJob).Range;
            Word.Range Bookreasons = oBDoc.Bookmarks.get_Item(ref Parareasons).Range;
            Word.Range BookWorkPlace = oBDoc.Bookmarks.get_Item(ref ParaWorkPlace).Range;
            Word.Range BookDocNo = oBDoc.Bookmarks.get_Item(ref ParaDocNo).Range;
           // MessageBox.Show(GregorianDate.Text);

            BookName2.Text = BookeName1.Text = ApplicantName.Text;             
            BookDate2.Text = BookDate1.Text = GregorianDate.Text;
            BookJob.Text = txtJob.Text;
            Bookreasons.Text = comboStatus.Text;
            BookWorkPlace.Text = txtWorkPlace.Text;
            BookDocNo.Text = DocNo.Text;

            object rangeName1 = BookeName1;
            object rangeName2 = BookName2;
            object rangeDate1 = BookDate1;
            object rangeDate2 = BookDate2;
            object rangeJob = BookJob;
            object rangereasons = Bookreasons;
            object rangeWorkPlace = BookWorkPlace;
            object rangeDocNo = BookDocNo;

           
            oBDoc.Bookmarks.Add("name1", ref rangeName1);
            oBDoc.Bookmarks.Add("name2", ref rangeName2);
            oBDoc.Bookmarks.Add("date1", ref rangeDate1);
            oBDoc.Bookmarks.Add("date2", ref rangeDate2);
            oBDoc.Bookmarks.Add("job", ref rangeJob);
            oBDoc.Bookmarks.Add("reasons", ref rangereasons);
            oBDoc.Bookmarks.Add("workPlace", ref rangeWorkPlace);
            oBDoc.Bookmarks.Add("DocNo", ref rangeDocNo);

            string docxouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(docxouput);
            //File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
            return;
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
                        if (dataRow[comlumnName].ToString() != "" )
                            combbox.Items.Add(dataRow[comlumnName].ToString());
                    }
                }
                saConn.Close();
            }
        }
        private void fileComboBoxMandoub(ComboBox combbox, string source, string comlumnName, string tableName)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + ",MandoubAreas from " + tableName;
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
                        if (dataRow[comlumnName].ToString() != "" && dataRow["MandoubAreas"].ToString() != "القنصلية العامة لجمهورية السودان - جدة")
                            combbox.Items.Add(dataRow[comlumnName].ToString() +"-"+ dataRow["MandoubAreas"].ToString());
                    }
                }
                saConn.Close();
            }
        }
        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName, string index)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName  + " where division='" + index + "'";
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
        
        private void loadComments(ComboBox combbox, string source, string comlumnName, string tableName, string index)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName  + " where division='" + index + "'";
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

        private string getFileNo(int id)
        {
            string str = "1";
            using (SqlConnection saConn = new SqlConnection(DataSource))
            {
                saConn.Open();

                string query = "select رقم_ملف_جدة,رقم_ملف_مكة,رقم_ملف_اللجنة,رقم_ملف_الوافدين,عدد_الأفراد,عدد_الأفراد_مكة,عدد_الأفراد_الوافدين,عدد_الأفراد_اللجنة from TableSettings";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;


                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                
                foreach (DataRow dataRow in table.Rows)
                {
                    if (!string.IsNullOrEmpty(dataRow["رقم_ملف_جدة"].ToString()))
                    {
                        switch (id)
                        {
                            case 0:
                                if (dataRow["رقم_ملف_جدة"].ToString() != "")
                                    str = dataRow["رقم_ملف_جدة"].ToString();
                                break;
                            case 1:
                                if (dataRow["رقم_ملف_مكة"].ToString() != "")
                                    str = dataRow["رقم_ملف_مكة"].ToString();
                                break;
                            
                            case 2:
                                if (dataRow["رقم_ملف_الوافدين"].ToString() != "")
                                    str = dataRow["رقم_ملف_الوافدين"].ToString();
                                break;
                            
                            case 3:
                                if (dataRow["رقم_ملف_اللجنة"].ToString() != "")
                                    str = dataRow["رقم_ملف_اللجنة"].ToString();
                                break;
                            case 4:
                                if (dataRow["عدد_الأفراد"].ToString() != "")
                                    str = dataRow["عدد_الأفراد"].ToString();
                                break;
                            case 5:
                                if (dataRow["عدد_الأفراد_مكة"].ToString() != "")
                                    str = dataRow["عدد_الأفراد_مكة"].ToString();
                                break;
                            case 6:
                                if (dataRow["عدد_الأفراد_الوافدين"].ToString() != "")
                                    str = dataRow["عدد_الأفراد_الوافدين"].ToString();
                                break;
                            case 7:
                                if (dataRow["عدد_الأفراد_اللجنة"].ToString() != "")
                                    str = dataRow["عدد_الأفراد_اللجنة"].ToString();
                                break;

                        }
                    }
                }
                saConn.Close();
            }
            return str;
        }


        private int getDocInfo(int id)
        {
            string query;
            string NewFileName = "";
            SqlConnection Con = new SqlConnection(DataSource);
            query = "select Data1, Extension1,FileName1 from TableFiles  where ID=@id";
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@Id", SqlDbType.Int).Value = id;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            imagecount = 0;
            if (reader.Read())
            {
                var name = reader["FileName1"].ToString();
                if (string.IsNullOrEmpty(name)) return 0;
                var ext = reader["Extension1"].ToString();
                if (string.IsNullOrEmpty(ext)) return 0;
                var Data = (byte[])reader["Data1"];
                PathImage[0] = FilesPathOut + DateTime.Now.ToString("mmss") + name;
                if(ext != "docx") 
                    pictureBox3.ImageLocation = PathImage[0];
                //NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                File.WriteAllBytes(PathImage[0], Data);
                imagecount = 1;
                loadPic.BackColor = btnAuth.BackColor = System.Drawing.Color.LightGreen;
                loadPic.Text = btnAuth.Text = "اضافة مستند آخر (" + imagecount.ToString() + ")";

                loadPic.Enabled = button10.Visible = btnAuth.Enabled = true;
                //MessageBox.Show(NewFileName);
                //System.Diagnostics.Process.Start(NewFileName);
            }
            Con.Close();
            return imagecount;
        }
        private void setFileNo(string source, string id,string col)
        {
            string colRef = "@" + col;
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();
                SqlCommand cmd = new SqlCommand("UPDATE TableSettings SET "+col+ "="+colRef + " WHERE ID = @ID", saConn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.AddWithValue("@ID", 1);
                cmd.Parameters.AddWithValue(colRef, id.ToString());
                cmd.ExecuteNonQuery();                
            }
            
        }

        private void getText(string source)
        {
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();
                string query = "select TextModel,TextTitle from TableAddContextAffair";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                Console.WriteLine(Authcases());
                foreach (DataRow dataRow in table.Rows)
                {
                    if (!string.IsNullOrEmpty(dataRow["TextTitle"].ToString()))
                    {
                        if (dataRow["TextTitle"].ToString() == comPurpose.Text)
                        {
                            text.Text = dataRow["TextModel"].ToString();
                            for (int x = 0; x < 10; x++)
                                text.Text = SuffPrefReplacements(text.Text, Authcases());
                            return;
                        }
                        
                    }
                }
                saConn.Close();
            }
        }
        

        private void FormSudAffairs_Load(object sender, EventArgs e)
        {
            autoCompleteTextBox(txtJob, DataSource, "jobs", "TableListCombo");
            fileComboBoxMandoub(mandoubName, DataSource, "MandoubNames", "TableMandoudList");
            fileComboBox(comPurpose, DataSource, "TextTitle", "TableAddContextAffair","0");
            fileComboBox(AttendViceConsul, DataSource, "ArabicAttendVC", "TableListCombo");
            //CountingData(); 
            //if (combFileNo.Items.Count > 0)
            //    combFileNo.SelectedIndex = 0;
            clear_All();
            loadScanner();
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
        private void autoCompleteTextBox(string source, string tableName)
        {
            AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select Questionare from " + tableName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                DataTable Textboxtable = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(Textboxtable);

                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    if (dataRow["Questionare"].ToString().Contains("*"))
                    {
                        autoComplete.Add(dataRow["Questionare"].ToString().Split('*')[1]);
                    }
                }
                txtDiffculties.AutoCompleteMode = AutoCompleteMode.Suggest;
                txtDiffculties.AutoCompleteSource = AutoCompleteSource.CustomSource;
                txtDiffculties.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
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

        private string SuffPrefReplacements(string text, int index)
        {
            string title = "";
            if (ApplicantSex.CheckState == CheckState.Checked)
                title = "ة";
                Suffex_preffixList();
            if (text.Contains("#6"))
                return text.Replace("#6", "السيد" + title + "/ "+ApplicantName.Text +" حامل"  +title + " " +DocType.Text + " بالرقم: " + DocNo.Text);
            if (text.Contains("#5"))
                return text.Replace("#5", preffix[index, 3]);
            if (text.Contains("$$$"))
                return text.Replace("$$$", preffix[index, 7]);
            if (text.Contains("###"))
                return text.Replace("###", preffix[index, 1]);
            if (text.Contains("***"))
                return text.Replace("***", preffix[index, 6]);
            if (text.Contains("%%%"))
                return text.Replace("%%%", preffix[index, 2]);
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

            preffix[0, 6] = "هو";//***
            preffix[1, 6] = "هي";
            preffix[2, 6] = "هما";
            preffix[3, 6] = "هما";
            preffix[4, 6] = "هن";
            preffix[5, 6] = "هم";

            preffix[0, 7] = "يكون";//$$$
            preffix[1, 7] = "تكون";
            preffix[2, 7] = "يكونا";
            preffix[3, 7] = "تكونا";
            preffix[4, 7] = "يكن";
            preffix[5, 7] = "يكونو";
        }

        private void ApplicantSex_CheckedChanged(object sender, EventArgs e)
        {
            if (ApplicantSex.CheckState == CheckState.Unchecked) 
                ApplicantSex.Text = "ذكر";
            else ApplicantSex.Text = "أنثى";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*
             خروج نهائي نظامي عام
خروج نهائي نظامي جدة
خروج نهائي نظامي مكة
خروج نهائي بالترحيل
مخاطبات اللجنة العمالية
تحويل المقابل المالي
قضايا المواطنين السودانيين
جميع القوائم
ملخص الملفات
             */
            Authcases();
            if (picVerify.Visible && newData)
            {
                Console.WriteLine("checkDataInfo");
                int id = checkDataInfo(false);
                if (picVerify.Visible)
                {
                    MessageBox.Show("يوجد إجراء سابق متطابق مع رقم الهوية");
                    return;

                }
            }
            /* 
وثيقة سفر أضطرارية
اقامة
رقم حدود
             */
            Save2DataBase(Convert.ToInt32(combFileNo.Text), contractState);
            int caseID= 1;
            if (comboStatus.Text == "تغيب عن العمل" || DocType.Text == "اقامة") 
                caseID = 1;
            if (comboStatus.Text == "تغيب عن العمل" || DocType.Text == "رقم حدود")
                caseID = 2;
            if (comboStatus.Text == "تغيب عن العمل" || DocType.Text == "جواز سفر")
                caseID = 3;
            if (comboStatus.Text == "تغيب عن العمل" || DocType.Text == "وثيقة سفر أضطرارية")
                caseID = 4; 
            if (comboStatus.Text == "عمرة"|| comboStatus.Text == "حج") 
                caseID = 5; 
            else if (comboStatus.Text == "مجهول") 
                caseID = 6;
            updatecase(ApplicantID, getTable(AffairIndex), caseID);
            string activeCopy = FilesPathOut + ApplicantName.Text +"خطاب الوافدين" + DateTime.Now.ToString("mmss") + ".docx";
            if (DocDestin.SelectedIndex > 1)
                CreateWordFile(activeCopy,"الحالة");
            activeCopy = FilesPathOut + " استمارة "+ApplicantName.Text + DateTime.Now.ToString("mmss") + ".docx";

            if (AffairIndex == 3)
            {
                while (File.Exists(activeCopy))
                    activeCopy = FilesPathOut + "خطابات الوافدين " + DateTime.Now.ToString("mmss") + ".docx";
                IqrarTravel(activeCopy);
            }
            else if (AffairIndex < 3)
            {
                while (File.Exists(activeCopy))
                    activeCopy = FilesPathOut + "خطابات الوافدين " + DateTime.Now.ToString("mmss") + ".docx";
                CreateWokOffice(activeCopy);
                activeCopy = FilesPathOut + "إقرار الوافدين " + DateTime.Now.ToString("mmss") + ".docx";
                IqrarFinalExit(activeCopy);
            }

            else if (AffairIndex == 4)
            {
                while (File.Exists(activeCopy))
                    activeCopy = FilesPathOut + "خطابات الوافدين " + DateTime.Now.ToString("mmss") + ".docx";
                CreateWordFile(activeCopy, "صلة القرابة"); 
                activeCopy = FilesPathOut + "إقرار الوافدين " + DateTime.Now.ToString("mmss") + ".docx";
                IqrarFinalExit(activeCopy);
            }

            clear_All();
           // FillDataGridView1(DocDestin.SelectedIndex);
            if (!dataGridView1.Visible)
            {
                labdate.Visible = dataGridView1.Visible = true;
                PanelMain.Visible = false;
            }
            else if (Jobposition.Contains("قنصل") && dataGridView1.Visible)
            {
                labdate.Visible = dataGridView1.Visible = false;
               
                PanelMain.Visible = true;
            }
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

        private void timer2_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();
            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            GregorianDate.Text = DateTime.Now.ToString("MM-dd-yyyy");
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == 1)

                mandoubName.Visible = mandoubLabel.Visible = true;

            else
            {
                mandoubName.Visible = mandoubLabel.Visible = false;
                mandoubName.Text = "";
            }
            }

        private void button2_Click(object sender, EventArgs e)
        {
            clear_All();
        }

        private void clear_All()
        {
            DocDestin.SelectedIndex = CityIndex;
            if (AffairIndex != 4)
            {
                pictureBox11.Visible = pictureBox13.Visible = PersToServed.Visible = label5.Visible = false;
            }
            else
            {
                pictureBox11.Visible = pictureBox13.Visible = PersToServed.Visible = label5.Visible = true;
            }
            if (AffairIndex < 7)
            {
                GridView(AffairIndex, combFileNo.Text);
                ColorFulGrid1(24,"لا تعليق");

            }
            else if (AffairIndex == 7)
            {
                AllDataGridView1();
                
            }
            else
            {
                FillFilesView1();
                ColorFulGrid1(12,"");
                btnFileUpload.Visible = btnFileDownload.Visible = true;
            }
            CountingData();

            int boxNo = 0;
            //foreach (Control control in Panelapp.Controls)
            //{
            //    if (control.Name.Split('_')[1] == boxNo.ToString())
            //    {
            //        control.Visible = false;
            //        control.Tag = "Unvalid";
            //        boxNo++;
            //    }
            //}

            foreach (Control control in PanelFiles.Controls)
            {
                if (control is TextBox)
                {
                    control.Text = "";
                }
            }

            foreach (Control control in this.Controls)
            {

                if (control is TextBox && control.Name != "txtId")
                {
                    control.Text = "";
                }
                if (control is CheckBox) ((CheckBox)control).CheckState = CheckState.Unchecked;
            }
            if (AttendViceConsul.Items.Count > VCIndex) AttendViceConsul.SelectedIndex = VCIndex;
            if (comboBox2.Items.Count >= 0) comboBox2.SelectedIndex = 0;
            
            if (DocType.Items.Count >= 1) DocType.SelectedIndex = 1;
            if (comboStatus.Items.Count >= 0) 
                comboStatus.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            gregorianDate = mandoubName.Text = "";
            
            labdate.Visible = dataGridView1.Visible = true;
            Suddanese_Affair.SelectedIndex = AffairIndex;
            newData = true;
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void Savecollective()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                string query = "insert into TableCollective (المهن_المعدلة,المهنة,نوع_المعاملة,جهة_العمل,رقم_الهوية,التاريخ_الميلادي,التاريخ_الهجري,الحالة,العمر) values (@المهن_المعدلة,@المهنة,@نوع_المعاملة,@جهة_العمل,@رقم_الهوية,@التاريخ_الميلادي,@التاريخ_الهجري,الحالة,@العمر)";
                
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
                sqlCmd.CommandType = CommandType.Text;

                sqlCmd.Parameters.AddWithValue("@جهة_العمل", DocDestin.SelectedIndex.ToString());                
                sqlCmd.Parameters.AddWithValue("@رقم_الهوية", DocNo.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", gregorianDate);                
                sqlCmd.Parameters.AddWithValue("@التاريخ_الهجري", HijriDate.Text.Trim());               
                sqlCmd.Parameters.AddWithValue("@الحالة", comboStatus.Text);
               
                sqlCmd.Parameters.AddWithValue("@المهنة", txtJob.Text);
                sqlCmd.Parameters.AddWithValue("@نوع_المعاملة", AffairIndex.ToString());
                sqlCmd.Parameters.AddWithValue("@العمر", txtBirth.Text);
                sqlCmd.ExecuteNonQuery();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
            }
            finally
            {
                sqlCon.Close();
            }
        }


        private void Save2DataBase(int fileNo, string cont)
        {
            string AppGender, namelist = "", doclist = "", relativelsit = "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            

            if (Names.Length == 1)
            {
                namelist = namelist + "_" + Names;
                doclist = doclist + "_" + DocumentNo;
                relativelsit = relativelsit + "_" + Relativity;
            }
            else
            {
                for (int x = 0; x < Names.Length; x++)
                {
                    namelist = namelist + "_" + Names[x];
                    doclist = doclist + "_" + DocumentNo[x];
                    relativelsit = relativelsit + "_" + Relativity[x];
                }
            }
            try
            {
                string str = "";
                switch (AffairIndex)
                {
                    case 0:
                        str = "WafidAddorEdit";
                        break;
                    case 1:
                        str = "WafidJedAddorEdit";
                        break;
                    case 2:
                        str = "WafidMekkahAddorEdit";
                        break;
                    case 3:
                        str = "TarheelAddorEdit";
                        break;
                    case 4:
                        str = "TransferAddorEdit";
                        break;
                    case 5:
                        str = "WafidComAddorEdit";
                        break;
                    default: return;
                }
                Console.WriteLine(str);
                
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlCommand sqlCmd = new SqlCommand(str, sqlCon);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                
                    Console.WriteLine("تعديل " + ApplicantName.Text);
                    sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                    sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                    sqlCmd.Parameters.AddWithValue("@المقصود_بالإجراء", PersToServed.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@صلة_القرابة", relativelsit.Trim());
                    sqlCmd.Parameters.AddWithValue("@افراد_الاسرة", namelist.Trim());
                    sqlCmd.Parameters.AddWithValue("@رقم_هوية_الافراد", doclist.Trim());
                    sqlCmd.Parameters.AddWithValue("@جهة_العمل", DocDestin.SelectedIndex.ToString());
                    sqlCmd.Parameters.AddWithValue("@موقع_العاملة", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@نوع_المعالجة", checkBoxes());
                    sqlCmd.Parameters.AddWithValue("@رقم_هاتف1", txtPhone1.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@رقم_هاتف2", txtPhone2.Text.Trim());

                if (txtComment.Text == "" && txtOldComment.Text == "")
                    sqlCmd.Parameters.AddWithValue("@تعليق", "");
                
                if(txtComment.Text == "" && txtOldComment.Text != "")
                sqlCmd.Parameters.AddWithValue("@تعليق", txtOldComment.Text.Trim() + Environment.NewLine + EmpName +" - "+GregorianDate.Text);

                if (txtComment.Text != "" && txtOldComment.Text == "")
                    sqlCmd.Parameters.AddWithValue("@تعليق", txtComment.Text.Trim() + Environment.NewLine + EmpName + " - " + GregorianDate.Text + Environment.NewLine + "--------------" + Environment.NewLine);

                if (txtComment.Text != "" && txtOldComment.Text != "")
                sqlCmd.Parameters.AddWithValue("@تعليق", txtComment.Text.Trim() + Environment.NewLine + EmpName + " - " + GregorianDate.Text + Environment.NewLine + "--------------" + Environment.NewLine + "*"+ txtOldComment.Text.Trim());
                
                if (NewEntrey)
                    sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", "new");
                else sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", "old");

                sqlCmd.Parameters.AddWithValue("@رقم_المعاملة", txtId.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@مقدم_الطلب", ApplicantName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@النوع", ApplicantSex.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@نوع_الهوية", DocType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@رقم_الهوية", DocNo.Text.Trim());
                    if (gregorianDate != "")
                        sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", gregorianDate);
                    else
                        sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@التاريخ_الهجري", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@طريقة_الطلب", comboBox2.Text);
                    sqlCmd.Parameters.AddWithValue("@اسم_الموظف", EmpName + "-" + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@اسم_المندوب", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@معاملة_مرجعية", RefDocument);
                    sqlCmd.Parameters.AddWithValue("@الحالة", comboStatus.Text);
                    sqlCmd.Parameters.AddWithValue("@نص_المكاتبة", text.Text);
                    sqlCmd.Parameters.AddWithValue("@المهنة", txtJob.Text);
                    sqlCmd.Parameters.AddWithValue("@مكان_العمل", txtWorkPlace.Text);
                    sqlCmd.Parameters.AddWithValue("@رقم_الملف", fileNo.ToString());
                    sqlCmd.Parameters.AddWithValue("@نوع_المعاملة", AffairIndex.ToString());
                sqlCmd.Parameters.AddWithValue("@sms", smsText);
                sqlCmd.Parameters.AddWithValue("@Questionare", cont + "*"+ txtDiffculties.Text + txtJobGroup.Text);
                sqlCmd.Parameters.AddWithValue("@العمر", txtBirth.Text);
                
                sqlCmd.Parameters.AddWithValue("@الغرض", comPurpose.Text);
                sqlCmd.ExecuteNonQuery();

                //Savecollective();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
            }
            finally
            {
                sqlCon.Close();
            }
        }

        private void updatecase(int id, string table, int caseID)
        {

            string query = "update "+ table+" set ord = "+caseID.ToString()+" where ID = @id";
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@id", id);
                
            sqlCommand.ExecuteNonQuery();
        }
        private void updateReadyCase(int id, string table, string  ready )
        {

            string query = "update "+ table+ " set ready  = N'" + ready+"' where ID = @id";
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@id", id);
                
            sqlCommand.ExecuteNonQuery();
        }

        private void unfounddata(string tableList)
        {
            for (int table = 0; table < 6; table++)
            {
                if (tableList == "") continue;

                string query = "insert into TableCollective (المهن_المعدلة,المهنة,نوع_المعاملة,جهة_العمل,رقم_الهوية,التاريخ_الميلادي,التاريخ_الهجري,الحالة) " +
                    "select المهنة, المهنة, نوع_المعاملة, جهة_العمل, رقم_الهوية, التاريخ_الميلادي, التاريخ_الهجري, الحالة from " + tableList+
                    "where رقم_الهوية in (" +
                    "select رقم_الهوية from " + tableList+ " where رقم_الهوية is not null and not exists(" +
                    "select رقم_الهوية from TableCollective where TableCollective.رقم_الهوية = " + tableList + ".رقم_الهوية) )";
                SqlConnection sqlCon = new SqlConnection(DataSource);
                SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.ExecuteNonQuery();
                sqlCon.Close();

            }

        }


        private void Save2DataBaseList(int fileNo)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            int fileCount = filesCount(fileNo.ToString());
            string AppGender, namelist = "", doclist = "", relativelsit = "";            
            Console.WriteLine("fileNo " + fileNo.ToString());
            Console.WriteLine("fileCount " + fileCount.ToString());
            Console.WriteLine("filelimit " + getFileNo(4));
            MessageBox.Show("filelimit " + getFileNo(4));
            if (fileCount >= Convert.ToInt32(getFileNo(4)) && AffairIndex == 1)
            {
                if (Convert.ToInt32(getFileNo(0)) > Convert.ToInt32(getFileNo(1)))
                    fileNo = Convert.ToInt32(getFileNo(0)) + 1;
                else fileNo = Convert.ToInt32(getFileNo(1)) + 1;
                setFileNo(DataSource, fileNo.ToString(), "رقم_ملف_جدة ");
            }
            else if (fileCount >= Convert.ToInt32(getFileNo(5)) && AffairIndex == 2)
            {
                if (Convert.ToInt32(getFileNo(1)) > Convert.ToInt32(getFileNo(0)))
                    fileNo = Convert.ToInt32(getFileNo(1)) + 1;
                else fileNo = Convert.ToInt32(getFileNo(0)) + 1;
                setFileNo(DataSource, fileNo.ToString(), "رقم_ملف_مكة ");
            }
            else if (fileCount >= Convert.ToInt32(getFileNo(6)) && AffairIndex == 3)
            {
                fileNo = Convert.ToInt32(getFileNo(3)) + 1;
                setFileNo(DataSource, fileNo.ToString(), "رقم_ملف_الوافدين ");
            }

            if (Names.Length == 1)
            {
                namelist = namelist + "_" + Names;
                doclist = doclist + "_" + DocumentNo;
                relativelsit = relativelsit + "_" + Relativity;
            }
            else
            {
                for (int x = 0; x < Names.Length; x++)
                {
                    namelist = namelist + "_" + Names[x];
                    doclist = doclist + "_" + DocumentNo[x];
                    relativelsit = relativelsit + "_" + Relativity[x];
                }
            }
            try
            {
                string str = "";
                switch (AffairIndex)
                {
                    case 0:
                        str = "WafidAddorEdit";
                        break;
                    case 1:
                        str = "WafidJedAddorEdit";
                        break;
                    case 2:
                        str = "WafidMekkahAddorEdit";
                        break;
                    case 3:
                        str = "TarheelAddorEdit";
                        break;
                    case 4:
                        str = "TransferAddorEdit";
                        break;
                    case 5:
                        str = "WafidComAddorEdit";
                        break;
                    default: return;
                }

                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlCommand sqlCmd = new SqlCommand(str, sqlCon);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                
                    sqlCmd.Parameters.AddWithValue("@ID", ApplicantID);
                    sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                    sqlCmd.Parameters.AddWithValue("@المقصود_بالإجراء", PersToServed.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@صلة_القرابة", relativelsit.Trim());
                    sqlCmd.Parameters.AddWithValue("@افراد_الاسرة", namelist.Trim());
                    sqlCmd.Parameters.AddWithValue("@رقم_هوية_الافراد", doclist.Trim());
                    sqlCmd.Parameters.AddWithValue("@جهة_العمل", DocDestin.SelectedIndex.ToString());
                    sqlCmd.Parameters.AddWithValue("@موقع_العاملة", AttendViceConsul.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@نوع_المعالجة", "");
                    sqlCmd.Parameters.AddWithValue("@رقم_هاتف1", txtPhone1.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@رقم_هاتف2", txtPhone2.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@تعليق", txtComment.Text.Trim());
                    if(NewEntrey)
                    sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", "new");
                    else sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", "old");
                sqlCmd.Parameters.AddWithValue("@رقم_المعاملة", txtId.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@مقدم_الطلب", ApplicantName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@النوع", ApplicantSex.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@نوع_الهوية", DocType.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@رقم_الهوية", DocNo.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", GregorianDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@التاريخ_الهجري", HijriDate.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@طريقة_الطلب", comboBox2.Text);
                    sqlCmd.Parameters.AddWithValue("@اسم_الموظف", labelEmp.Text + "-" + DateTime.Now.ToString("hh:mm"));
                    sqlCmd.Parameters.AddWithValue("@اسم_المندوب", mandoubName.Text.Trim());
                    sqlCmd.Parameters.AddWithValue("@معاملة_مرجعية", RefDocument);
                    sqlCmd.Parameters.AddWithValue("@الحالة", comboStatus.Text);
                    sqlCmd.Parameters.AddWithValue("@نص_المكاتبة", text.Text);
                    sqlCmd.Parameters.AddWithValue("@المهنة", txtJob.Text);
                    sqlCmd.Parameters.AddWithValue("@مكان_العمل", txtWorkPlace.Text);
                    sqlCmd.Parameters.AddWithValue("@رقم_الملف", fileNo.ToString());
                    sqlCmd.Parameters.AddWithValue("@نوع_المعاملة", AffairIndex.ToString());
                    sqlCmd.Parameters.AddWithValue("@sms", smsText);
                    sqlCmd.ExecuteNonQuery();
                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
            }
            finally
            {
                sqlCon.Close();
            }
        }


        //private void getDataByID(string DocID) { 
        //ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
        //            newData = false;
        //            int newFiles = 0;
        //            string FileNo = getFileNo(AffairIndex).ToString();
        //            for (int s = 0; s < dataGridView1.Rows.Count - 1; s++)
        //            {
        //                if (dataGridView1.Rows[s].Cells[2].Value.ToString() == "")
        //                    newFiles++;
        //            }

        //            if (dataGridView1.CurrentRow.Cells[2].Value.ToString() == "")
        //            {
        //                int fileSet = dataGridView1.Rows.Count - 1 - newFiles;
        //                txtId.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
        //                //
        //                //else if (fileSet >= 15 && AffairIndex == 3)
        //                //{
        //                //    if (Convert.ToInt32(getFileNo(0)) > Convert.ToInt32(getFileNo(3)))
        //                //        FileNo = (Convert.ToInt32(getFileNo(0)) + 1).ToString();
        //                //    else FileNo = (Convert.ToInt32(getFileNo(3)) + 1).ToString();
        //                //    setFileNo(DataSource, FileNo, "رقم_ملف_مكة ");
        //                //}
        //                //else if (fileSet >= 50 && AffairIndex == 1)
        //                //{
        //                //    FileNo = (Convert.ToInt32(getFileNo(1)) + 1).ToString();
        //                //    setFileNo(DataSource, FileNo, "رقم_ملف_الوافدين ");
        //                //}
        //                combFileNo.Text = FileNo;
        //                NewEntrey = true;
        //                OpenFileDoc(ApplicantID, 1,AffairIndex);
        //                if (Jobposition.Contains("قنصل"))
        //                    deleteRow.Visible = true;
        //                return;
        //            }
                    
        //            NewEntrey = false;
        //            PersToServed.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
        //            Relativity = dataGridView1.CurrentRow.Cells[8].Value.ToString().Split('_');
        //            Names = dataGridView1.CurrentRow.Cells[7].Value.ToString().Split('_');
        //            DocumentNo = dataGridView1.CurrentRow.Cells[9].Value.ToString().Split('_');
                    
        //            txtPhone1.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
        //            txtPhone2.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();
        //            txtComment.Text = dataGridView1.CurrentRow.Cells[24].Value.ToString();
        //            ArchivedSt.Text = dataGridView1.CurrentRow.Cells[25].Value.ToString();
        //            if (ArchivedSt.Text != "غير مؤرشف")
        //                ArchivedSt.CheckState = CheckState.Checked;

                    
        //            txtId.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
        //            ApplicantName.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
        //            //
        //            ApplicantSex.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
        //            DocType.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
        //            DocNo.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                    
        //            HijriDate.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
        //            GregorianDate.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
        //            AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
        //            string str = GregorianDate.Text;
        //            if (!HijriDate.Text.Contains("1443"))
        //            {
        //                GregorianDate.Text = HijriDate.Text;
        //                HijriDate.Text = str;
        //            }
        //            comboBox2.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
        //            //txtFileNo.Text =  dataGridView1.CurrentRow.Cells[16].Value.ToString().Split('-')[0];
        //            mandoubName.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
        //            RefDocument = dataGridView1.CurrentRow.Cells[18].Value.ToString();
        //            comboStatus.Text = dataGridView1.CurrentRow.Cells[26].Value.ToString();
        //            text.Text = dataGridView1.CurrentRow.Cells[27].Value.ToString();
        //            txtJob.Text = dataGridView1.CurrentRow.Cells[28].Value.ToString();
        //            txtWorkPlace.Text = dataGridView1.CurrentRow.Cells[29].Value.ToString();
        //            combFileNo.Text = dataGridView1.CurrentRow.Cells[30].Value.ToString();
        //            AffairIndex = Convert.ToInt32(dataGridView1.CurrentRow.Cells[32].Value.ToString());
        //            if (AffairIndex <= 5)
        //                Suddanese_Affair.SelectedIndex = AffairIndex;
        //            Console.WriteLine("AffairIndex " + AffairIndex.ToString());
        //            smsText = dataGridView1.CurrentRow.Cells[31].Value.ToString();
        //            ALLSmsIndex = smsText.Split('-').Length;
        //            txtSMS.Text = dataGridView1.CurrentRow.Cells[31].Value.ToString().Split('-')[ALLSmsIndex - 1];
        //            if (ALLSmsIndex > 1) previous.Visible = next.Visible = true;
        //            else previous.Visible = next.Visible = false;
        //            //Suddanese_Affair.SelectedIndex = AffairIndex;
        //            //AffairIndex = Convert.ToInt32(dataGridView1.CurrentRow.Cells[32].Value.ToString());
        //            if (dataGridView1.CurrentRow.Cells[7].Value.ToString() != "")
        //            {
        //                for (int x = 0; x < Names.Length; x++)
        //                {
        //                    if (Names[x] != "") Panelapp_Paint(Names[x], Relativity[x], DocumentNo[x]);
        //                }
        //                int labourOffice = -1;
        //                if (!string.IsNullOrEmpty(dataGridView1.CurrentRow.Cells[10].Value.ToString()))
        //                    labourOffice = Convert.ToInt32(dataGridView1.CurrentRow.Cells[10].Value.ToString());
        //                if (labourOffice != -1)
        //                    DocDestin.SelectedIndex = labourOffice;
        //                else
        //                    DocDestin.Text = "ملف غير محدد";
        //                //MessageBox.Show(WorkOffices[10]);
        //                //if (!string.IsNullOrEmpty(dataGridView1.CurrentRow.Cells[13].Value.ToString()))
        //                //{
                            
        //                //}

        //                //نوع_المعالجة = dataGridView1.CurrentRow.Cells[14].Value.ToString();
        //            }
        //        }}

        //private void TransferData(DataRow row, string strQuery, string source56)
        //{
        //    SqlConnection sqlCon = new SqlConnection(source56);
        //    if (sqlCon.State == ConnectionState.Closed)
        //            sqlCon.Open();
        //    SqlCommand sqlCmd = new SqlCommand(strQuery, sqlCon);
        //    sqlCmd.CommandType = CommandType.StoredProcedure;
        //    sqlCmd.Parameters.AddWithValue("@ID", 0);
        //    sqlCmd.Parameters.AddWithValue("@mode", "Add");
        //    sqlCmd.Parameters.AddWithValue("@المقصود_بالإجراء", row["المقصود_بالإجراء"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@صلة_القرابة", row["صلة_القرابة"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@افراد_الاسرة", row["افراد_الاسرة"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@رقم_هوية_الافراد", row["رقم_هوية_الافراد"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@جهة_العمل", row["جهة_العمل"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@موقع_العاملة", row["موقع_العاملة"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@نوع_المعالجة", row["نوع_المعالجة"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@رقم_هاتف1", row["رقم_هاتف1"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@رقم_هاتف2", row["رقم_هاتف2"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@تعليق", row["تعليق"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", row["حالة_الارشفة"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@رقم_المعاملة", row["رقم_المعاملة"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@مقدم_الطلب", row["مقدم_الطلب"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@النوع", row["النوع"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@نوع_الهوية", row["نوع_الهوية"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@رقم_الهوية", row["رقم_الهوية"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", row["التاريخ_الميلادي"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@التاريخ_الهجري", row["التاريخ_الهجري"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@طريقة_الطلب", row["طريقة_الطلب"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@اسم_الموظف", row["اسم_الموظف"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@اسم_المندوب", row["اسم_المندوب"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@معاملة_مرجعية", row["معاملة_مرجعية"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@الحالة", row["الحالة"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@نص_المكاتبة", row["نص_المكاتبة"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@المهنة", row["المهنة"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@مكان_العمل", row["مكان_العمل"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@رقم_الملف", row["رقم_الملف"].ToString());
        //    sqlCmd.Parameters.AddWithValue("@نوع_المعاملة", Suddanese_Affair.SelectedIndex.ToString());
        //    sqlCmd.Parameters.AddWithValue("@sms", row["sms"].ToString());

        //    //sqlCmd.Parameters.AddWithValue("@Data2", row["Data2"]);
        //    //sqlCmd.Parameters.AddWithValue("@Extension2", row["Extension2"].ToString());
        //    //sqlCmd.Parameters.AddWithValue("@المكاتبة_النهائية", row["المكاتبة_النهائية"].ToString());

        //    //sqlCmd.Parameters.AddWithValue("@Data1", row["Data1"]);
        //    //sqlCmd.Parameters.AddWithValue("@Extension1", row["Extension1"].ToString());
        //    //sqlCmd.Parameters.AddWithValue("@ارشفة_المستندات", row["ارشفة_المستندات"].ToString());

        //    sqlCmd.ExecuteNonQuery();
        //    sqlCon.Close();
        //}

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
                if (name == "") return;
                var ext = reader["Extension1"].ToString();
                if (ext == "") return;
                var Data = (byte[])reader["Data1"];
                
                
                var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                File.WriteAllBytes(NewFileName, Data);
                System.Diagnostics.Process.Start(NewFileName);
            }


            sqlCon.Close();
        }


        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {            
            if (dataGridView1.CurrentRow.Index != -1)
            {
                
                grdiFill = true;
                if (AffairIndex == 8)
                {
                    btnFileSaveUpdate.Text = "تعديل";
                    panelFile.Visible = true;
                    labdate.Visible = dataGridView1.Visible = false;
                    PanelMain.Visible = false;
                    
                    FileIDNo = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                    Docfile = OpenFilesDoc(FileIDNo);
                    if (Docfile != "")
                    {
                        btnOpenFile.Visible = true;
                    }
                    else {
                        btnOpenFile.Visible = false;
                    }
                    txt1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    txt2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                    txt3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                    txt4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                    txt5.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                    txt6.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                    txt7.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                    txt8.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                    txt9.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                    txt10.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                    txt11.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                    getDocInfo(FileIDNo);
                }
                else
                {
                    
                    panelFile.Visible = false;
                    labdate.Visible = labdate.Visible = labdate.Visible = dataGridView1.Visible = false;
                    

                    PanelMain.Visible = true;
                    ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                    newData = false;
                    int newFiles = 0;
                    string FileNo = getFileNo(AffairIndex).ToString();
                    for (int s = 0; s < dataGridView1.Rows.Count - 1; s++)
                    {
                        if (dataGridView1.Rows[s].Cells[2].Value.ToString() == "")
                            newFiles++;
                    }

                    if (dataGridView1.CurrentRow.Cells[2].Value.ToString() == "")
                    {
                        int fileSet = dataGridView1.Rows.Count - 1 - newFiles;
                        txtId.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                        //
                        //else if (fileSet >= 15 && AffairIndex == 3)
                        //{
                        //    if (Convert.ToInt32(getFileNo(0)) > Convert.ToInt32(getFileNo(3)))
                        //        FileNo = (Convert.ToInt32(getFileNo(0)) + 1).ToString();
                        //    else FileNo = (Convert.ToInt32(getFileNo(3)) + 1).ToString();
                        //    setFileNo(DataSource, FileNo, "رقم_ملف_مكة ");
                        //}
                        //else if (fileSet >= 50 && AffairIndex == 1)
                        //{
                        //    FileNo = (Convert.ToInt32(getFileNo(1)) + 1).ToString();
                        //    setFileNo(DataSource, FileNo, "رقم_ملف_الوافدين ");
                        //}
                        combFileNo.Text = FileNo;
                        NewEntrey = true;
                        newData = true;
                        //OpenFileDoc(ApplicantID, 1,AffairIndex);
                        //MessageBox.Show(ApplicantID.ToString()+" - "+ getTable(AffairIndex));
                        FillDatafromGenArch( "data1", ApplicantID.ToString(), getTable(AffairIndex));
                        if (Jobposition.Contains("قنصل"))
                            deleteRow.Visible = true;
                        return;
                    }
                    
                    NewEntrey = false;
                    PersToServed.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                    Relativity = dataGridView1.CurrentRow.Cells[8].Value.ToString().Split('_');
                    Names = dataGridView1.CurrentRow.Cells[7].Value.ToString().Split('_');
                    DocumentNo = dataGridView1.CurrentRow.Cells[9].Value.ToString().Split('_');
                    
                    txtPhone1.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
                    txtPhone2.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();
                    txtComment.Text = "";
                    txtOldComment.Text = dataGridView1.CurrentRow.Cells["تعليق"].Value.ToString();
                    
                    //ArchivedSt.Text = dataGridView1.CurrentRow.Cells[25].Value.ToString();
                    //if (ArchivedSt.Text != "غير مؤرشف")
                    //    ArchivedSt.CheckState = CheckState.Checked;


                    txtId.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    ApplicantName.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                    //

                    if (dataGridView1.CurrentRow.Cells[3].Value.ToString().Contains("_"))
                    {
                        ApplicantSex.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString().Split('_')[0];
                        txtBirth.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString().Split('_')[1];
                    }
                    else
                    {
                        ApplicantSex.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                        txtBirth.Text = dataGridView1.CurrentRow.Cells["العمر"].Value.ToString();
                    }
                        string Questionare = dataGridView1.CurrentRow.Cells["Questionare"].Value.ToString();


                    if (ApplicantSex.Text == "ذكر")
                        ApplicantSex.CheckState = CheckState.Unchecked;
                    else ApplicantSex.CheckState = CheckState.Checked;

                    if (Questionare.Split('*').Length == 2)
                    {
                        if(Questionare.Split('*')[0] == "1") contract.CheckState = CheckState.Checked; 
                        else if(Questionare.Split('*')[0] == "0") contract.CheckState = CheckState.Unchecked;

                        txtDiffculties.Text = Questionare.Split('*')[1];
                    }
                    else if (Questionare.Split('*').Length == 3)
                    {
                        if(Questionare.Split('*')[0] == "1") contract.CheckState = CheckState.Checked; 
                        else if(Questionare.Split('*')[0] == "0") contract.CheckState = CheckState.Unchecked;
                        txtDiffculties.Text = Questionare.Split('*')[1];
                        txtJobGroup.Text = Questionare.Split('*')[2];
                    }


                    DocType.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                    DocNo.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                    
                    HijriDate.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();



                    labdate.Text = "تاريخ الإجراء " + dataGridView1.CurrentRow.Cells[12].Value.ToString();
                    labdate.Visible = true;
                    GregorianDate.Text = gregorianDate = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                    AttendViceConsul.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                    string str = GregorianDate.Text;
                    if (!HijriDate.Text.Contains("1443"))
                    {
                        GregorianDate.Text = gregorianDate = HijriDate.Text;
                        HijriDate.Text = str;
                    }
                    string[] strings = dataGridView1.CurrentRow.Cells[14].Value.ToString().Split('_');
                    setBoxes(strings);
                    comboBox2.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                    //txtFileNo.Text =  dataGridView1.CurrentRow.Cells[16].Value.ToString().Split('-')[0];
                    mandoubName.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
                    RefDocument = dataGridView1.CurrentRow.Cells[18].Value.ToString();
                    comboStatus.Text = dataGridView1.CurrentRow.Cells[26].Value.ToString();
                    text.Text = dataGridView1.CurrentRow.Cells[27].Value.ToString();
                    txtJob.Text = dataGridView1.CurrentRow.Cells[28].Value.ToString();
                    txtWorkPlace.Text = dataGridView1.CurrentRow.Cells[29].Value.ToString();
                    combFileNo.Text = dataGridView1.CurrentRow.Cells[30].Value.ToString();
                    
                    try
                    {
                        AffairIndex = Convert.ToInt32(dataGridView1.CurrentRow.Cells[32].Value.ToString());
                    }
                    catch (Exception ex) { AffairIndex = 0; }
                    if (AffairIndex <= 5)
                        Suddanese_Affair.SelectedIndex = AffairIndex;
                    Console.WriteLine("AffairIndex " + AffairIndex.ToString());
                    smsText = dataGridView1.CurrentRow.Cells[31].Value.ToString();
                    ALLSmsIndex = smsText.Split('-').Length;
                    //txtSMS.Text = dataGridView1.CurrentRow.Cells[31].Value.ToString().Split('-')[ALLSmsIndex - 1];
                    //if (ALLSmsIndex > 1) previous.Visible = next.Visible = true;
                    //else previous.Visible = next.Visible = false;
                    comFinalaPro.SelectedIndex = 0;
                    ready.Text = dataGridView1.CurrentRow.Cells["ready"].Value.ToString();
                    if (ready.Text == "مكتمل")
                        ready.Checked = true;
                    else ready.Checked= false;
                        

                    if (dataGridView1.CurrentRow.Cells[7].Value.ToString() != "")
                    {
                        for (int x = 0; x < Names.Length; x++)
                        {
                            if (Names[x] != "") Panelapp_Paint(Names[x], Relativity[x], DocumentNo[x]);
                        }
                        int labourOffice = -1;
                        if (!string.IsNullOrEmpty(dataGridView1.CurrentRow.Cells[10].Value.ToString()))
                            labourOffice = Convert.ToInt32(dataGridView1.CurrentRow.Cells[10].Value.ToString());
                        if (labourOffice != -1)
                            DocDestin.SelectedIndex = labourOffice;
                        else
                            DocDestin.Text = "ملف غير محدد";
                        //MessageBox.Show(WorkOffices[10]);
                        //if (!string.IsNullOrEmpty(dataGridView1.CurrentRow.Cells[13].Value.ToString()))
                        //{
                            
                        //}

                        //نوع_المعالجة = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                    }
                    
                }
            }
        }
        
        private void dataGridView1_index(int index)
        {

            MessageBox.Show(index.ToString());

            if (dataGridView1.RowCount > 1)
            {
                
                grdiFill = true;
                
                    
                    panelFile.Visible = false;
                    labdate.Visible = labdate.Visible = labdate.Visible = dataGridView1.Visible = false;
                    

                    PanelMain.Visible = true;
                    ApplicantID = Convert.ToInt32(dataGridView1.Rows[index].Cells[0].Value.ToString());
                    newData = false;
                    int newFiles = 0;
                    string FileNo = getFileNo(AffairIndex).ToString();
                    for (int s = 0; s < dataGridView1.Rows.Count - 1; s++)
                    {
                        if (dataGridView1.Rows[s].Cells[2].Value.ToString() == "")
                            newFiles++;
                    }

                    if (dataGridView1.Rows[index].Cells[2].Value.ToString() == "")
                    {
                        int fileSet = dataGridView1.Rows.Count - 1 - newFiles;
                        txtId.Text = dataGridView1.Rows[index].Cells[1].Value.ToString();
                        //
                        //else if (fileSet >= 15 && AffairIndex == 3)
                        //{
                        //    if (Convert.ToInt32(getFileNo(0)) > Convert.ToInt32(getFileNo(3)))
                        //        FileNo = (Convert.ToInt32(getFileNo(0)) + 1).ToString();
                        //    else FileNo = (Convert.ToInt32(getFileNo(3)) + 1).ToString();
                        //    setFileNo(DataSource, FileNo, "رقم_ملف_مكة ");
                        //}
                        //else if (fileSet >= 50 && AffairIndex == 1)
                        //{
                        //    FileNo = (Convert.ToInt32(getFileNo(1)) + 1).ToString();
                        //    setFileNo(DataSource, FileNo, "رقم_ملف_الوافدين ");
                        //}
                        combFileNo.Text = FileNo;
                        NewEntrey = true;
                        newData = true;
                        //OpenFileDoc(ApplicantID, 1,AffairIndex);
                        //MessageBox.Show(ApplicantID.ToString()+" - "+ getTable(AffairIndex));
                        FillDatafromGenArch( "data1", ApplicantID.ToString(), getTable(AffairIndex));
                        if (Jobposition.Contains("قنصل"))
                            deleteRow.Visible = true;
                        return;
                    }
                    
                    NewEntrey = false;
                    PersToServed.Text = dataGridView1.Rows[index].Cells[6].Value.ToString();
                    Relativity = dataGridView1.Rows[index].Cells[8].Value.ToString().Split('_');
                    Names = dataGridView1.Rows[index].Cells[7].Value.ToString().Split('_');
                    DocumentNo = dataGridView1.Rows[index].Cells[9].Value.ToString().Split('_');
                    
                    txtPhone1.Text = dataGridView1.Rows[index].Cells[19].Value.ToString();
                    txtPhone2.Text = dataGridView1.Rows[index].Cells[20].Value.ToString();
                    txtComment.Text = "";
                    txtOldComment.Text = dataGridView1.Rows[index].Cells[24].Value.ToString();
                    
                    //ArchivedSt.Text = dataGridView1.Rows[index].Cells[25].Value.ToString();
                    //if (ArchivedSt.Text != "غير مؤرشف")
                    //    ArchivedSt.CheckState = CheckState.Checked;


                    txtId.Text = dataGridView1.Rows[index].Cells[1].Value.ToString();
                    ApplicantName.Text = dataGridView1.Rows[index].Cells[2].Value.ToString();
                    //

                    if (dataGridView1.Rows[index].Cells[3].Value.ToString().Contains("_"))
                    {
                        ApplicantSex.Text = dataGridView1.Rows[index].Cells[3].Value.ToString().Split('_')[0];
                        txtBirth.Text = dataGridView1.Rows[index].Cells[3].Value.ToString().Split('_')[1];
                    }
                    else
                    {
                        ApplicantSex.Text = dataGridView1.Rows[index].Cells[3].Value.ToString();
                        txtBirth.Text = dataGridView1.Rows[index].Cells["العمر"].Value.ToString();
                    }
                        string Questionare = dataGridView1.Rows[index].Cells["Questionare"].Value.ToString();


                    if (ApplicantSex.Text == "ذكر")
                        ApplicantSex.CheckState = CheckState.Unchecked;
                    else ApplicantSex.CheckState = CheckState.Checked;

                    if (Questionare.Split('*').Length == 2)
                    {
                        if(Questionare.Split('*')[0] == "1") contract.CheckState = CheckState.Checked; 
                        else if(Questionare.Split('*')[0] == "0") contract.CheckState = CheckState.Unchecked;

                        txtDiffculties.Text = Questionare.Split('*')[1];
                    }
                    else if (Questionare.Split('*').Length == 3)
                    {
                        if(Questionare.Split('*')[0] == "1") contract.CheckState = CheckState.Checked; 
                        else if(Questionare.Split('*')[0] == "0") contract.CheckState = CheckState.Unchecked;
                        txtDiffculties.Text = Questionare.Split('*')[1];
                        txtJobGroup.Text = Questionare.Split('*')[2];
                    }


                    DocType.Text = dataGridView1.Rows[index].Cells[4].Value.ToString();
                    DocNo.Text = dataGridView1.Rows[index].Cells[5].Value.ToString();
                    
                    HijriDate.Text = dataGridView1.Rows[index].Cells[11].Value.ToString();



                    labdate.Text = "تاريخ الإجراء " + dataGridView1.Rows[index].Cells[12].Value.ToString();
                    labdate.Visible = true;
                    GregorianDate.Text = gregorianDate = dataGridView1.Rows[index].Cells[12].Value.ToString();
                    AttendViceConsul.Text = dataGridView1.Rows[index].Cells[13].Value.ToString();
                    string str = GregorianDate.Text;
                    if (!HijriDate.Text.Contains("1443"))
                    {
                        GregorianDate.Text = gregorianDate = HijriDate.Text;
                        HijriDate.Text = str;
                    }
                    string[] strings = dataGridView1.Rows[index].Cells[14].Value.ToString().Split('_');
                    setBoxes(strings);
                    comboBox2.Text = dataGridView1.Rows[index].Cells[15].Value.ToString();
                    //txtFileNo.Text =  dataGridView1.Rows[index].Cells[16].Value.ToString().Split('-')[0];
                    mandoubName.Text = dataGridView1.Rows[index].Cells[17].Value.ToString();
                    RefDocument = dataGridView1.Rows[index].Cells[18].Value.ToString();
                    comboStatus.Text = dataGridView1.Rows[index].Cells[26].Value.ToString();
                    text.Text = dataGridView1.Rows[index].Cells[27].Value.ToString();
                    txtJob.Text = dataGridView1.Rows[index].Cells[28].Value.ToString();
                    txtWorkPlace.Text = dataGridView1.Rows[index].Cells[29].Value.ToString();
                    combFileNo.Text = dataGridView1.Rows[index].Cells[30].Value.ToString();
                    AffairIndex = Convert.ToInt32(dataGridView1.Rows[index].Cells[32].Value.ToString());
                    if (AffairIndex <= 5)
                        Suddanese_Affair.SelectedIndex = AffairIndex;
                    Console.WriteLine("AffairIndex " + AffairIndex.ToString());
                    smsText = dataGridView1.Rows[index].Cells[31].Value.ToString();
                    ALLSmsIndex = smsText.Split('-').Length;
                    //txtSMS.Text = dataGridView1.Rows[index].Cells[31].Value.ToString().Split('-')[ALLSmsIndex - 1];
                    //if (ALLSmsIndex > 1) previous.Visible = next.Visible = true;
                    //else previous.Visible = next.Visible = false;
                    comFinalaPro.SelectedIndex = 0;
                    if (dataGridView1.Rows[index].Cells[7].Value.ToString() != "")
                    {
                        for (int x = 0; x < Names.Length; x++)
                        {
                            if (Names[x] != "") Panelapp_Paint(Names[x], Relativity[x], DocumentNo[x]);
                        }
                        int labourOffice = -1;
                        if (!string.IsNullOrEmpty(dataGridView1.Rows[index].Cells[10].Value.ToString()))
                            labourOffice = Convert.ToInt32(dataGridView1.Rows[index].Cells[10].Value.ToString());
                        if (labourOffice != -1)
                            DocDestin.SelectedIndex = labourOffice;
                        else
                            DocDestin.Text = "ملف غير محدد";
                        //MessageBox.Show(WorkOffices[10]);
                        //if (!string.IsNullOrEmpty(dataGridView1.Rows[index].Cells[13].Value.ToString()))
                        //{
                            
                        //}

                        //نوع_المعالجة = dataGridView1.Rows[index].Cells[14].Value.ToString();
                    }
                    
                
            }
        }

        private void setBoxes(string[] strings)
        {
            check1.CheckState = CheckState.Unchecked;
            check2.CheckState = CheckState.Unchecked;
            check3.CheckState = CheckState.Unchecked;
            check4.CheckState = CheckState.Unchecked;
            check5.CheckState = CheckState.Unchecked;
            check6.CheckState = CheckState.Unchecked;
            check7.CheckState = CheckState.Unchecked;
            if (strings.Length > 1)
            {
                if (strings[0] == "1") check1.CheckState = CheckState.Checked;
                else check1.CheckState = CheckState.Unchecked;
                if (strings[1] == "1") check2.CheckState = CheckState.Checked;
                else check2.CheckState = CheckState.Unchecked;
                if (strings[2] == "1") check3.CheckState = CheckState.Checked;
                else check3.CheckState = CheckState.Unchecked;
                if (strings[3] == "1") check4.CheckState = CheckState.Checked;
                else check4.CheckState = CheckState.Unchecked;
                if (strings[4] == "1") check5.CheckState = CheckState.Checked;
                else check5.CheckState = CheckState.Unchecked;
                if (strings[5] == "1") check6.CheckState = CheckState.Checked;
                else check6.CheckState = CheckState.Unchecked;
                if (strings[6] == "1") check7.CheckState = CheckState.Checked;
                else check7.CheckState = CheckState.Unchecked;
            }
        }

        private void selectedID(int Rowindex, DataGridView dataGridView) {
            ApplicantID = Convert.ToInt32(dataGridView.Rows[Rowindex].Cells[0].Value.ToString());
            newData = false;
            int newFiles = 0;
            string FileNo = getFileNo(AffairIndex).ToString();
            for (int s = 0; s < dataGridView.Rows.Count - 1; s++)
            {
                if (dataGridView.Rows[s].Cells[2].Value.ToString() == "")
                    newFiles++;
            }

            if (dataGridView.Rows[Rowindex].Cells[2].Value.ToString() == "")
            {
                int fileSet = dataGridView.Rows.Count - 1 - newFiles;
                txtId.Text = dataGridView.Rows[Rowindex].Cells[1].Value.ToString();
                //
                //else if (fileSet >= 15 && AffairIndex == 3)
                //{
                //    if (Convert.ToInt32(getFileNo(0)) > Convert.ToInt32(getFileNo(3)))
                //        FileNo = (Convert.ToInt32(getFileNo(0)) + 1).ToString();
                //    else FileNo = (Convert.ToInt32(getFileNo(3)) + 1).ToString();
                //    setFileNo(DataSource, FileNo, "رقم_ملف_مكة ");
                //}
                //else if (fileSet >= 50 && AffairIndex == 1)
                //{
                //    FileNo = (Convert.ToInt32(getFileNo(1)) + 1).ToString();
                //    setFileNo(DataSource, FileNo, "رقم_ملف_الوافدين ");
                //}
                combFileNo.Text = FileNo;
                NewEntrey = true;
                //OpenFileDoc(ApplicantID, 1, AffairIndex);
                FillDatafromGenArch("data1", ApplicantID.ToString(), getTable(AffairIndex));
                if (Jobposition.Contains("قنصل"))
                    deleteRow.Visible = true;
                return;
            }

            NewEntrey = false;
            PersToServed.Text = dataGridView.Rows[Rowindex].Cells[6].Value.ToString();
            Relativity = dataGridView.Rows[Rowindex].Cells[8].Value.ToString().Split('_');
            Names = dataGridView.Rows[Rowindex].Cells[7].Value.ToString().Split('_');
            DocumentNo = dataGridView.Rows[Rowindex].Cells[9].Value.ToString().Split('_');

            txtPhone1.Text = dataGridView.Rows[Rowindex].Cells[19].Value.ToString();
            txtPhone2.Text = dataGridView.Rows[Rowindex].Cells[20].Value.ToString();
            txtComment.Text = "";
            txtOldComment.Text = dataGridView1.Rows[Rowindex].Cells[24].Value.ToString();
            //ArchivedSt.Text = dataGridView.Rows[Rowindex].Cells[25].Value.ToString();
            //if (ArchivedSt.Text != "غير مؤرشف")
            //    ArchivedSt.CheckState = CheckState.Checked;


            txtId.Text = dataGridView.Rows[Rowindex].Cells[1].Value.ToString();
            ApplicantName.Text = dataGridView.Rows[Rowindex].Cells[2].Value.ToString();
            //
            ApplicantSex.Text = dataGridView.Rows[Rowindex].Cells[3].Value.ToString();
            DocType.Text = dataGridView.Rows[Rowindex].Cells[4].Value.ToString();
            DocNo.Text = dataGridView.Rows[Rowindex].Cells[5].Value.ToString();

            HijriDate.Text = dataGridView.Rows[Rowindex].Cells[11].Value.ToString();
            GregorianDate.Text = gregorianDate = dataGridView.Rows[Rowindex].Cells[12].Value.ToString();
            AttendViceConsul.Text = dataGridView.Rows[Rowindex].Cells[13].Value.ToString();
            string str = GregorianDate.Text;
            if (!HijriDate.Text.Contains("1443"))
            {
                GregorianDate.Text = gregorianDate = HijriDate.Text;
                HijriDate.Text = str;
            }
            comboBox2.Text = dataGridView.Rows[Rowindex].Cells[15].Value.ToString();
            //txtFileNo.Text =  dataGridView.Rows[Rowindex].Cells[16].Value.ToString().Split('-')[0];
            mandoubName.Text = dataGridView.Rows[Rowindex].Cells[17].Value.ToString();
            RefDocument = dataGridView.Rows[Rowindex].Cells[18].Value.ToString();
            comboStatus.Text = dataGridView.Rows[Rowindex].Cells[26].Value.ToString();
            text.Text = dataGridView.Rows[Rowindex].Cells[27].Value.ToString();
            txtJob.Text = dataGridView.Rows[Rowindex].Cells[28].Value.ToString();
            txtWorkPlace.Text = dataGridView.Rows[Rowindex].Cells[29].Value.ToString();
            combFileNo.Text = dataGridView.Rows[Rowindex].Cells[30].Value.ToString();
            AffairIndex = Convert.ToInt32(dataGridView.Rows[Rowindex].Cells[32].Value.ToString());
            if (AffairIndex <= 5)
                Suddanese_Affair.SelectedIndex = AffairIndex;
            Console.WriteLine("AffairIndex " + AffairIndex.ToString());
            smsText = dataGridView.Rows[Rowindex].Cells[31].Value.ToString();
            ALLSmsIndex = smsText.Split('-').Length;
            //txtSMS.Text = dataGridView.Rows[Rowindex].Cells[31].Value.ToString().Split('-')[ALLSmsIndex - 1];
            //if (ALLSmsIndex > 1) previous.Visible = next.Visible = true;
            //else previous.Visible = next.Visible = false;
            //Suddanese_Affair.SelectedIndex = AffairIndex;
            //AffairIndex = Convert.ToInt32(dataGridView.Rows[Rowindex].Cells[32].Value.ToString());
            if (dataGridView.Rows[Rowindex].Cells[7].Value.ToString() != "")
            {
                for (int x = 0; x < Names.Length; x++)
                {
                    if (Names[x] != "") Panelapp_Paint(Names[x], Relativity[x], DocumentNo[x]);
                }
                int labourOffice = -1;
                if (!string.IsNullOrEmpty(dataGridView.Rows[Rowindex].Cells[10].Value.ToString()))
                    labourOffice = Convert.ToInt32(dataGridView.Rows[Rowindex].Cells[10].Value.ToString());
                if (labourOffice != -1)
                    DocDestin.SelectedIndex = labourOffice;
                else
                    DocDestin.Text = "ملف غير محدد";
                //MessageBox.Show(WorkOffices[10]);
                //if (!string.IsNullOrEmpty(dataGridView.Rows[Rowindex].Cells[13].Value.ToString()))
                //{

                //}

                //نوع_المعالجة = dataGridView.Rows[Rowindex].Cells[14].Value.ToString();
            }
        }
        private string getTable(int index ) {
            string table = "";
            switch (index)
            {
                case 0:
                    table = "TableWafid";
                    break;
                case 1:
                    table = "TableWafidJed";
                    break;
                case 2:
                    table = "TableWafidMekkah";
                    break;
                case 3:
                    table = "TableTarheel";
                    break;
                case 4:
                    table = "TableTransfer";
                    break;
                case 5:
                    table = "TableCommity";
                    break;
            }
            return table;
        }

        private string getFileEditAdd(int index)
        {
            string table = "";
            switch (index)
            {
                case 0:
                    table = "WafidAddorEdit";
                    break;
                case 1:
                    table = "WafidJedAddorEdit";
                    break;
                case 2:
                    table = "WafidMekkahAddorEdit";
                    break;
                case 3:
                    table = "TarheelAddorEdit";
                    break;
                case 4:
                    table = "TransferAddorEdit";
                    break;
                case 5:
                    table = "CommityAddorEdit";
                    break;
            }

            return table;
        }
        private void OpenFileDoc(int id, int fileNo, int fileno)
        {
            string query, table;
            
            SqlConnection Con = new SqlConnection(DataSource);

            table = getTable(fileno);

            if (fileNo == 1)
            {
                query = "select Data1, Extension1,ارشفة_المستندات from "+table+"  where ID=@id";
            }
            else if (fileNo == 2)
            {
                query = "select Data2, Extension2,المكاتبة_النهائية from "+table+"  where ID=@id";
            }
            else
                query = "select Data3, Extension3,المكاتبة_الأولية from "+table+"  where ID=@id";

            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@Id", SqlDbType.Int).Value = id;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                if (fileNo == 1)
                {
                    string name = reader["ارشفة_المستندات"].ToString();
                    if (string.IsNullOrEmpty(name)) return;
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                    if (string.IsNullOrEmpty(ext)) return;
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("hhmmss")) + ext;
                    
                    File.WriteAllBytes(NewFileName, Data);
                    if (string.IsNullOrEmpty(NewFileName))
                        return;
                    try
                    {
                        System.Diagnostics.Process.Start(NewFileName);
                    }
                    catch (Exception ex) {
                        Console.WriteLine("FileName " + NewFileName);
                    }
                }
                else if (fileNo == 2)
                {
                    string name = reader["المكاتبة_النهائية"].ToString();
                    if (string.IsNullOrEmpty(name)) return;
                    var Data = (byte[])reader["Data2"];
                    var ext = reader["Extension2"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }
                else
                {
                    string name = reader["المكاتبة_الأولية"].ToString();
                    if (string.IsNullOrEmpty(name)) return;
                    var Data = (byte[])reader["Data3"];
                    var ext = reader["Extension3"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                   
                    System.Diagnostics.Process.Start(NewFileName);
                }

            }
            Con.Close();


        }

      
            private void ListSearch_TextChanged(object sender, EventArgs e)
        {
            //
            Console.WriteLine("Before Items count " + dataGridView1.Rows.Count.ToString());
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            if (ListSearch.Text.All(char.IsDigit) && ListSearch.Text.Length > 0)
            {
                textBox1.Text = "ق س ج/" + ListSearch.Text;
                bs.Filter = dataGridView1.Columns[5].HeaderText.ToString() + " LIKE '%" + ListSearch.Text + "%'";
                dataGridView1.DataSource = bs;
            }
            else if (!ListSearch.Text.All(char.IsDigit) && ListSearch.Text.Length > 0)
            {
                bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '%" + ListSearch.Text + "%'";
                dataGridView1.DataSource = bs;
            }
            Console.WriteLine("File No " + combFileNo.Text);
            Console.WriteLine("after Items count " + dataGridView1.Rows.Count.ToString());

        }

        private void btnListView_Click(object sender, EventArgs e)
        {
            if (!dataGridView1.Visible)
            {
                clear_All();
                //if (AffairIndex == 8)
                //    FillFilesView1(); 
                //else FillDataGridView1(AffairIndex); 

                labdate.Visible = labdate.Visible = dataGridView1.Visible = true;
                panelFile.Visible = false;
                PanelMain.Visible = false;
            }
            else if (Jobposition.Contains("قنصل") && dataGridView1.Visible)
            {
                clear_All();
                
                if (AffairIndex == 6)
                {
                    //FillFilesView1();
                    panelFile.Visible = true;
                    PanelMain.Visible = false;
                }
                else
                {
                    //FillDataGridView1(DocDestin.SelectedIndex);
                    panelFile.Visible = false;
                    PanelMain.Visible = true;
                }
                labdate.Visible = dataGridView1.Visible = false;
                
            }
        }

        private void text_TextChanged(object sender, EventArgs e)
        {
            if(text.Text.Contains("على إقامة رب الأسرة"))
            {
                PersToServed.CheckState = CheckState.Checked;                
            }
        }

        private void deleteRow_Click(object sender, EventArgs e)
        {
            string table = "";
            switch (AffairIndex)
            {
                case 0:
                    table = "TableWafid";
                    break;
                case 1:
                    table = "TableWafidJed";
                    break;
                case 2:
                    table = "TableWafidMekkah";
                    break;
                case 3:
                    table = "TableTarheel";
                    break;
                case 4:
                    table = "TableTransfer";
                    break;
                case 5:
                    table = "TableCommity";
                    break;
                default: return;
            }
            var selectedOption = MessageBox.Show("", "تأكيد عملية الحذف", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                deleteRowsData(ApplicantID, table, DataSource);

                clear_All();                     
                PanelMain.Visible = false;
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

        private void btnFile2_Click(object sender, EventArgs e)
        {
            //OpenFileDoc(ApplicantID, 2, AffairIndex);
            btnFile2.Enabled = false;
            FillDatafromGenArch("data2", ApplicantID.ToString(), getTable(AffairIndex));
            btnFile2.Enabled = true;
        }

        private void btnFile1_Click(object sender, EventArgs e)
        {
            //OpenFileDoc(ApplicantID, 1, AffairIndex);
            btnFile1.Enabled = false;
            FillDatafromGenArch("data1", ApplicantID.ToString(), getTable(AffairIndex));
            btnFile1.Enabled = true;
        }

        private void btnFile3_Click(object sender, EventArgs e)
        {
            OpenFileDoc(ApplicantID, 3, AffairIndex);
        }


        int CountingData()
        {
            string table = "";
            switch (AffairIndex)
            {
                case 0:
                    table = "TableWafid";
                    break;
                case 1:
                    table = "TableWafidJed";
                    break;
                case 2:
                    table = "TableWafidMekkah";
                    break;
                case 3:
                    table = "TableTarheel";
                    break;
                case 4:
                    table = "TableTransfer";
                    break;
                case 5:
                    table = "TableCommity";
                    break;
                default: return 0;
            }
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT رقم_الملف, COUNT(*) FROM "+table+" Group by رقم_الملف", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);            
            sqlCon.Close();
            string str = "";
            
            int z = 0;
            fileList = new int[dtbl.Rows.Count];
            foreach (DataRow row in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(row["رقم_الملف"].ToString()) && row["رقم_الملف"].ToString().All(char.IsDigit))
                {
                    fileList[z] = Convert.ToInt32(row["رقم_الملف"].ToString());
                    z++;
                    //str = row["رقم_الملف"].ToString();
                    ////if(fill)
                    //combFileNo.Items.Add(row["رقم_الملف"].ToString());
                }
                //z++;
            }
            return dtbl.Rows.Count;
        }

        
        int CountingCurrentData(string noID)
        {
            string table = "";
            switch (AffairIndex)
            {
                case 0:
                    table = "TableWafid";
                    break;
                case 1:
                    table = "TableWafidJed";
                    break;
                case 2:
                    table = "TableWafidMekkah";
                    break;
                case 3:
                    table = "TableTarheel";
                    break;
                case 4:
                    table = "TableTransfer";
                    break;
                case 5:
                    table = "TableCommity";
                    break;
                default: return 0;
            }
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT رقم_الملف,نوع_المعاملة,جهة_العمل,مقدم_الطلب FROM "+table+"", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            int x = 0;
            //foreach (DataRow dataRow in dtbl.Rows)
            //{
            //    if (dataRow["نوع_المعاملة"].ToString() != AffairIndex.ToString())
            //    {
            //        dtbl.Rows[x].Delete();
            //    }
            //    x++;
            //}

            sqlCon.Close();
            //MessageBox.Show(noID.ToString()); ;
            x = 0;int y = 0;
            foreach (DataRow row in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(row["رقم_الملف"].ToString()) && row["رقم_الملف"].ToString().All(char.IsDigit) && row["نوع_المعاملة"].ToString() == AffairIndex.ToString() && row["جهة_العمل"].ToString() == "0")
                    if (row["رقم_الملف"].ToString() == noID)
                    {
                        x++;
                        if (!string.IsNullOrEmpty(row["مقدم_الطلب"].ToString()))
                            y++;
                    }
            }
            return (x-y);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            
        }

        

        private void StaredColumns()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select Name,IdNo FROM TableTemp", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var fileinfo = new FileInfo(sfd.FileName);
                        using (var package = new ExcelPackage(fileinfo))
                        {
                            ExcelWorksheet excelsheet = package.Workbook.Worksheets.Add("Rights");
                            excelsheet.Cells.LoadFromDataTable(dtbl);

                            package.Save();

                        }
                    }
                    catch (Exception ex)
                    {
                    }

                }
            }  
            
        }

        private void deleteRowsDataTicket()
        {
            string query;
            SqlConnection Con = new SqlConnection(DataSource);            
            query = "DELETE FROM TableTemp";
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            Con.Close();
        }

        private void ListSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            int index = 0;
            if (e.KeyChar == (char)13)
            {
                if (dataGridView1.RowCount == 2)
                {
                    

                    labdate.Visible = dataGridView1.Visible = false;
                    
                    PanelMain.Visible = true;
                    ApplicantID = Convert.ToInt32(dataGridView1.Rows[index].Cells[0].Value.ToString());
                    newData = false;
                    if (dataGridView1.Rows[index].Cells[2].Value.ToString() == "")
                    {

                        txtId.Text = dataGridView1.Rows[index].Cells[1].Value.ToString();

                        //OpenFileDoc(ApplicantID, 1, AffairIndex);
                        FillDatafromGenArch( "data1", ApplicantID.ToString(), getTable(AffairIndex));
                        if (Jobposition.Contains("قنصل"))
                            deleteRow.Visible = true;
                        return;
                    }
                    PersToServed.Text = dataGridView1.Rows[index].Cells[6].Value.ToString();
                    Relativity = dataGridView1.Rows[index].Cells[8].Value.ToString().Split('_');
                    Names = dataGridView1.Rows[index].Cells[7].Value.ToString().Split('_');
                    DocumentNo = dataGridView1.Rows[index].Cells[9].Value.ToString().Split('_');
                    for (int x = 0; x < Names.Length; x++)
                    {
                        if (Names[x] != "") Panelapp_Paint(Names[x], Relativity[x], DocumentNo[x]);
                    }
                    DocDestin.Text = WorkOffices[Convert.ToInt32(dataGridView1.CurrentRow.Cells[10].Value.ToString())];
                    AttendViceConsul.Text = dataGridView1.Rows[index].Cells[13].Value.ToString();
                    //نوع_المعالجة = dataGridView1.Rows[index].Cells[14].Value.ToString();
                    txtPhone1.Text = dataGridView1.Rows[index].Cells[19].Value.ToString();
                    txtPhone2.Text = dataGridView1.Rows[index].Cells[20].Value.ToString();
                    txtComment.Text = dataGridView1.Rows[index].Cells[24].Value.ToString();
                    //ArchivedSt.Text = dataGridView1.Rows[index].Cells[25].Value.ToString();
                    //if (ArchivedSt.Text != "غير مؤرشف")
                    //    ArchivedSt.CheckState = CheckState.Checked;
                    txtId.Text = dataGridView1.Rows[index].Cells[1].Value.ToString();
                    ApplicantName.Text = dataGridView1.Rows[index].Cells[2].Value.ToString();
                    ApplicantSex.Text = dataGridView1.Rows[index].Cells[3].Value.ToString();
                    DocType.Text = dataGridView1.Rows[index].Cells[4].Value.ToString();
                    DocNo.Text = dataGridView1.Rows[index].Cells[5].Value.ToString();
                    HijriDate.Text = dataGridView1.Rows[index].Cells[11].Value.ToString();
                    GregorianDate.Text = gregorianDate = dataGridView1.Rows[index].Cells[12].Value.ToString();
                    
                    string str = GregorianDate.Text;
                    if (!HijriDate.Text.Contains("1443"))
                    {
                        gregorianDate = GregorianDate.Text = HijriDate.Text;
                        HijriDate.Text = str;
                    }
                    
                    comboBox2.Text = dataGridView1.Rows[index].Cells[15].Value.ToString();
                    mandoubName.Text = dataGridView1.Rows[index].Cells[17].Value.ToString();
                    RefDocument = dataGridView1.Rows[index].Cells[18].Value.ToString();
                    comboStatus.Text = dataGridView1.Rows[index].Cells[26].Value.ToString();
                    text.Text = dataGridView1.Rows[index].Cells[27].Value.ToString();
                    txtJob.Text = dataGridView1.Rows[index].Cells[28].Value.ToString();
                    txtWorkPlace.Text = dataGridView1.Rows[index].Cells[29].Value.ToString();
                    combFileNo.Text = dataGridView1.Rows[index].Cells[30].Value.ToString();

                }
            }
        }

       
        private void UpdateTableTrans(string id, string text, string previousID, string newTable)
        {
            // MessageBox.Show(id.ToString() +"-"+ table + "-"+column + "-"+text);
            //string qurey = "update "+table+" set "+ column + "=@"+ column + " where ID=@id";
            string qurey = "update TableGeneralArch set رقم_المرجع=@الرقم_السابق, docTable =@docTable where رقم_معاملة_القسم=@رقم_معاملة_القسم and رقم_المرجع=@الرقم_الجديد";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@الرقم_السابق", previousID);
            sqlCmd.Parameters.AddWithValue("@الرقم_الجديد", id);
            sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", text);
            sqlCmd.Parameters.AddWithValue("@docTable", newTable);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        
        private void updataState(int id, string text, string column, string table)
        {
            // MessageBox.Show(id.ToString() +"-"+ table + "-"+column + "-"+text);
            //string qurey = "update "+table+" set "+ column + "=@"+ column + " where ID=@id";
            string qurey = "update "+ table+" set " + column+"=@" + column+" where ID=@id";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@"+ column, text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }


        private void FillTable(string name,string idno)
        {
            // MessageBox.Show(id.ToString() +"-"+ table + "-"+column + "-"+text);
            //string qurey = "update "+table+" set "+ column + "=@"+ column + " where ID=@id";
            string qurey = "INSERT INTO TableTemp(Name,IdNo) VALUES (@Name,@IdNo)";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@Name", name);
            sqlCmd.Parameters.AddWithValue("@IdNo", idno);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }


        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            lalCount.Text = fileFiletering(combFileNo.Text).ToString();
            
        }

        private int filesCount(string text)
        {
            FillCount1(AffairIndex);
            if (string.IsNullOrEmpty(text)) return dataGridView3.Rows.Count - 1;
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView3.DataSource;
                bs.Filter = dataGridView3.Columns[30].HeaderText.ToString() + " LIKE '" + text + "%'";
                dataGridView3.DataSource = bs;
                
            return dataGridView3.Rows.Count - 1;
        }

        private int fileFiletering(string text)
        {
            int value = 0;
            GridView(AffairIndex, combFileNo.Text);
            ColorFulGrid1(24,"لا تعليق");
            if (!combFileNo.Text.All(char.IsDigit))
            {
                
                
                return dataGridView1.Rows.Count - 1;
            }
            else
            {

                if (AffairIndex > 5 ||string.IsNullOrEmpty(text)) return dataGridView1.Rows.Count - 1;
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns[30].HeaderText.ToString() + " LIKE '" + text + "%'";
                dataGridView1.DataSource = bs;
                dataGridView1.Sort(dataGridView1.Columns["ord"], System.ComponentModel.ListSortDirection.Ascending);
                value = dataGridView1.Rows.Count - 1;
                //
            }
            return dataGridView1.Rows.Count - 1;
        }
        private void button4_Click()
        {

        }

        private string checkBoxes()
        {
            string strings = "";
            if (check1.Checked) strings = "1";
            else strings = "0";
            if (check2.Checked) strings = strings + "_1";
            else strings = strings + "_0";
            if (check3.Checked) strings = strings + "_1";
            else strings = strings + "_0";
            if (check4.Checked) strings = strings + "_1";
            else strings = strings + "_0";
            if (check5.Checked) strings = strings + "_1";
            else strings = strings + "_0";
            if (check6.Checked) strings = strings + "_1";
            else strings = strings + "_0";
            if (check7.Checked) strings = strings + "_1";
            else strings = strings + "_0";
            
            return strings;
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            if (ApplicantName.Text == "") return;
            //Authcases();
            
            if (picVerify.Visible && newData) {
                Console.WriteLine("checkDataInfo");
                int id = checkDataInfo(false);
                if (picVerify.Visible)
                {
                    MessageBox.Show("يوجد إجراء سابق متطابق مع رقم الهوية .. يرجى فحص رقم اثبات الشخصية أولا");
                    return;
                    
                }
            }

            
            
                Save2DataBase(Convert.ToInt32(combFileNo.Text), contractState = "0");


            clear_All();
            fileFiletering("");
            labdate.Visible = dataGridView1.Visible = true;
            PanelMain.Visible = false;
        }

        private void button12_Click(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            //SendSms(txtPhone1.Text, txtSMS.Text);
            //if (smsText.Split('-').Length > 1) smsText = smsText + "-"+txtSMS.Text;
            //else smsText = txtSMS.Text;
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


        }

        private void combFileNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                fileFiletering(combFileNo.Text);
        }


        private void CreateSummary(string RouteFile,int rowCount, string [,] data)
        {
            string docxouput = RouteFile;
            string pdfouput = RouteFile.Replace(".docx",".pdf");
            System.IO.File.Copy(FilesPathIn + "تقرير ملخص الملفات.docx", RouteFile);

            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();

            object objCurrentCopy = RouteFile;

            Word.Document oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);
            oBMicroWord.Selection.Find.ClearFormatting();
            oBMicroWord.Selection.Find.Replacement.ClearFormatting();
            Microsoft.Office.Interop.Word.Table table = oBDoc.Tables[1];
            for (int x = 0; x < rowCount; x++)
            {
                table.Rows.Add();
                table.Rows[x + 3].Cells[12].Range.Text = data[x, 0];
                table.Rows[x + 3].Cells[11].Range.Text = data[x, 1];
                table.Rows[x + 3].Cells[10].Range.Text = data[x, 2];
                table.Rows[x + 3].Cells[9].Range.Text = data[x, 3];
                table.Rows[x + 3].Cells[8].Range.Text = data[x, 4];
                table.Rows[x + 3].Cells[7].Range.Text = data[x, 5];
                table.Rows[x + 3].Cells[6].Range.Text = data[x, 6];
                table.Rows[x + 3].Cells[5].Range.Text = data[x, 7];
                table.Rows[x + 3].Cells[4].Range.Text = data[x, 8];
                table.Rows[x + 3].Cells[3].Range.Text = data[x, 9];
                table.Rows[x + 3].Cells[2].Range.Text = data[x, 10];
                table.Rows[x + 3].Cells[1].Range.Text = data[x, 11];
            }



                object ParaAuthNo = "MarkAuthNo";
            object ParaHijriData = "MarkHijriData";
            object ParaGreData = "MarkGreData";
            object ParaAuthBody1part1 = "MarkAVC";
            

            Word.Range BookAuthNo = oBDoc.Bookmarks.get_Item(ref ParaAuthNo).Range;
            Word.Range BookHijriData = oBDoc.Bookmarks.get_Item(ref ParaHijriData).Range;
            Word.Range BookGreData = oBDoc.Bookmarks.get_Item(ref ParaGreData).Range;
            Word.Range BookAuthBody1part1 = oBDoc.Bookmarks.get_Item(ref ParaAuthBody1part1).Range;

            BookAuthNo.Text = "ق س ج/80/13/" + DateTime.Now.ToString("dd");
            BookHijriData.Text = HijriDate.Text;
            BookGreData.Text = GregorianDate.Text;
            BookAuthBody1part1.Text = AttendViceConsul.Text;
            
            object rangeAuthNo = BookAuthNo;
            object rangeHijriData = BookHijriData;
            object rangeGreData = BookGreData;
            object rangeAuthBody1part1 = BookAuthBody1part1;
            
            oBDoc.Bookmarks.Add("MarkAuthNo", ref rangeAuthNo);
            oBDoc.Bookmarks.Add("MarkHijriData", ref rangeHijriData);
            oBDoc.Bookmarks.Add("MarkGreData", ref rangeGreData);
            oBDoc.Bookmarks.Add("MarkAVC", ref rangeAuthBody1part1);
            
            oBDoc.SaveAs2(docxouput);

            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);


            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(docxouput);
            //File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;



        
        }



        private void btnFileDownload_Click(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select * from TableFiles", sqlCon);
            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa1.Fill(dtbl);
            dataGridView4.DataSource = dtbl;
            sqlCon.Close();
            

                    string[,] Celldata = new string[dtbl.Rows.Count,12];
            for (int y = 1; y < 12; y++)
                for (int x = 0; x < dtbl.Rows.Count; x++) {
                    if(dataGridView4.Rows[x].Cells[y].Value.ToString() != "غير محدد")
                Celldata[x, y-1] = dataGridView4.Rows[x].Cells[y].Value.ToString();
                    else Celldata[x, y - 1] = "";
                }
            string activeCopy = FilesPathOut + "تقرير ملخص الملفات" + DateTime.Now.ToString("mmss") + ".docx";
            while (File.Exists(activeCopy))
                activeCopy = FilesPathOut + "تقرير ملخص الملفات" + DateTime.Now.ToString("mmss") + ".docx";

            CreateSummary(activeCopy, dtbl.Rows.Count, Celldata);
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var fileinfo = new FileInfo(sfd.FileName);
                        using (var package = new ExcelPackage(fileinfo))
                        {
                            ExcelWorksheet excelsheet = package.Workbook.Worksheets.Add("Rights");
                            excelsheet.Cells.LoadFromDataTable(dtbl);

                            //excelsheet.Cells.LoadFromCollection(dataGridView7.DataSource); 
                            package.Save();

                        }
                    }
                    catch (Exception ex)
                    {
                    }


                }
            }
        }


        private void downloadData(int tableIndex) {
            string dataSource57 = "Data Source=192.168.100.57,49170;Network Library=DBMSSOCN;Initial Catalog=AhwalDataBase;User ID=sa;Password=1234";
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select * from "+ getTable(tableIndex), sqlCon);
            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa1.Fill(dtbl);
            dataGridView4.DataSource = dtbl;
            sqlCon.Close();
            int xx = 0;
            string fileLocation = "";
            string dataSource56 = "Data Source=192.168.100.56,49170;Network Library=DBMSSOCN;Initial Catalog=AhwalDataBase;User ID=ConsJeddahAdmin;Password=DataBC0nsJ49170";
            //MessageBox.Show(getFileEditAdd(Suddanese_Affair.SelectedIndex));
            foreach (DataRow dataRow in dtbl.Rows)
            {
                //TransferData(dataRow, getFileEditAdd(Suddanese_Affair.SelectedIndex), dataSource56);
                for (xx = 0;xx<3;xx++) {                    
                    fileLocation  = AllOpenFilesDoc(Convert.ToInt32(dataRow["ID"].ToString()), getTable(tableIndex), xx, dataSource57);
                    Console.WriteLine(fileLocation);
                    if (fileLocation != "") {
                        MessageBox.Show(fileLocation);
                        updateCellData(dataRow["رقم_الهوية"].ToString(), fileLocation, dataSource56, xx, getTable(tableIndex));
                    }
                }
                
            }

            //string[,] Celldata = new string[dtbl.Rows.Count, 12];
            //for (int y = 1; y < 12; y++)
            //    for (int x = 0; x < dtbl.Rows.Count; x++)
            //    {
            //        if (dataGridView4.Rows[x].Cells[y].Value.ToString() != "غير محدد")
            //            Celldata[x, y - 1] = dataGridView4.Rows[x].Cells[y].Value.ToString();
            //        else Celldata[x, y - 1] = "";
            //    }
            
            
            //using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xlsx" })
            //{
            //    if (sfd.ShowDialog() == DialogResult.OK)
            //    {
            //        try
            //        {
            //            var fileinfo = new FileInfo(sfd.FileName);
            //            using (var package = new ExcelPackage(fileinfo))
            //            {
            //                ExcelWorksheet excelsheet = package.Workbook.Worksheets.Add("Rights");
            //                excelsheet.Cells.LoadFromDataTable(dtbl);

            //                //excelsheet.Cells.LoadFromCollection(dataGridView7.DataSource); 
            //                package.Save();

            //            }
            //        }
            //        catch (Exception ex)
            //        {
            //        }


            //    }
            //}
        }
        private void btnFileUpload_Click(object sender, EventArgs e)
        {
            DeleteTable("TableFiles");
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            for (int rCnt = 1; ; rCnt++)
            {
                //MessageBox.Show((string)(range.Cells[rCnt, 2] as Excel.Range).Value2);
                if (string.IsNullOrEmpty((string)(range.Cells[rCnt, 2] as Excel.Range).Value2)) 
                    break;
                string FileNo = Convert.ToString((range.Cells[rCnt, 1] as Excel.Range).Value2);
                string FileDest = Convert.ToString((range.Cells[rCnt, 2] as Excel.Range).Value2);
                string IndividualNo = Convert.ToString((range.Cells[rCnt, 3] as Excel.Range).Value2);
                string ConToWafidDate = Convert.ToString((range.Cells[rCnt, 4] as Excel.Range).Value2);
                string ConsToWafidNo = Convert.ToString((range.Cells[rCnt, 5] as Excel.Range).Value2);
                string WafidToWorkDate = Convert.ToString((range.Cells[rCnt, 6] as Excel.Range).Value2);
                string WafidToWorkNo = Convert.ToString((range.Cells[rCnt, 7] as Excel.Range).Value2);
                string WorkToWafidDate = Convert.ToString((range.Cells[rCnt, 8] as Excel.Range).Value2);
                string WorkToWafidNo = Convert.ToString((range.Cells[rCnt, 9] as Excel.Range).Value2);
                string WorkNotesNo = Convert.ToString((range.Cells[rCnt, 10] as Excel.Range).Value2);
                string WorkFinsihed = Convert.ToString((range.Cells[rCnt, 11] as Excel.Range).Value2);
                UpdateFileList(0, FileNo, FileDest, IndividualNo, ConToWafidDate, ConsToWafidNo, WafidToWorkDate, WafidToWorkNo, WorkToWafidDate, WorkToWafidNo, WorkNotesNo, WorkFinsihed);
            }

            sqlCon.Close();
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            FillFilesView1();
            ColorFulGrid1(12,"");
        }

        private void DeleteTable(string table)
        {
            string sql = "delete from " + table;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(sql, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void UpdateColumn(string source, string strName, string IdNo, string phoneNo, string fileNo, string destination, string date, string comment, string phone2, string status, string docType)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            SqlCommand sqlCmd = new SqlCommand("TarAddorEdit", sqlCon);
            sqlCmd.CommandType = CommandType.StoredProcedure;
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.Parameters.AddWithValue("@ID", 1);
            sqlCmd.Parameters.AddWithValue("@mode", "Add"); 
            sqlCmd.Parameters.AddWithValue("@رقم_الملف", fileNo);
            sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", date);      
            sqlCmd.Parameters.AddWithValue("@تعليق", comment);
            sqlCmd.Parameters.AddWithValue("@رقم_هاتف1", phoneNo);
            sqlCmd.Parameters.AddWithValue("@رقم_هاتف2", phone2);
            sqlCmd.Parameters.AddWithValue("@نوع_الهوية", docType);
            sqlCmd.Parameters.AddWithValue("@رقم_الهوية", IdNo);
            sqlCmd.Parameters.AddWithValue("@مقدم_الطلب", strName);
            sqlCmd.Parameters.AddWithValue("@رقم_المعاملة", "ق س ج/80/22/13/1");
            sqlCmd.Parameters.AddWithValue("@الحالة", status);
            sqlCmd.Parameters.AddWithValue("@نوع_المعاملة", "3");

            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void btnTarListUpload()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            string col = "Col0";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            for (int rCnt = 1; rCnt < rw; rCnt++)
            {
                string strName = (string)(range.Cells[rCnt, 1] as Excel.Range).Value2;
                if (string.IsNullOrEmpty(strName))
                    strName = "غير محدد";
                string IdNo = Convert.ToString((range.Cells[rCnt, 2] as Excel.Range).Value2);
                if (string.IsNullOrEmpty(IdNo))
                    IdNo = "غير محدد";
                string phoneNo = Convert.ToString((range.Cells[rCnt, 3] as Excel.Range).Value2);
                if (string.IsNullOrEmpty(phoneNo))
                    phoneNo = "غير محدد"; 
                string status= Convert.ToString((range.Cells[rCnt, 4] as Excel.Range).Value2);
                if (string.IsNullOrEmpty(status))
                    status = "غير محدد";
                string Comment = Convert.ToString((range.Cells[rCnt, 5] as Excel.Range).Value2);
                if (string.IsNullOrEmpty(Comment))
                    Comment = "غير محدد";
                string ProDate = Convert.ToString((range.Cells[rCnt, 6] as Excel.Range).Value2);
                if (string.IsNullOrEmpty(ProDate))
                    ProDate = "غير محدد";
                string fileNo = Convert.ToString((range.Cells[rCnt, 7] as Excel.Range).Value2);
                if (string.IsNullOrEmpty(fileNo))
                    fileNo = "غير محدد";
                string destination = Convert.ToString((range.Cells[rCnt, 5] as Excel.Range).Value2);
                if (string.IsNullOrEmpty(destination))
                    destination = "غير محدد";
                string phoneNo2 = Convert.ToString((range.Cells[rCnt, 8] as Excel.Range).Value2);
                if (string.IsNullOrEmpty(phoneNo2))
                    phoneNo2 = "غير محدد";
                string docType = Convert.ToString((range.Cells[rCnt, 9] as Excel.Range).Value2);
                if (string.IsNullOrEmpty(docType))
                    docType = "غير محدد";
                UpdateColumn(DataSource, strName, IdNo, phoneNo, fileNo, destination, ProDate, Comment, phoneNo2, status, docType);
            }

            sqlCon.Close();
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            var selectedOption = MessageBox.Show("", "تأكيد عملية الحذف", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                deleteRowsData(FileIDNo, "TableFiles", DataSource);
                {
                    FillFilesView1();
                    ColorFulGrid1(12,"");
                }
                    labdate.Visible = dataGridView1.Visible = true;
                
                panelFile.Visible = false; 
            }
        }

        private void btnFileDown_Click(object sender, EventArgs e)
        {
            if (txt1.Text == "")
                return;
                if (btnFileSaveUpdate.Text == "تعديل")
            {
                UpdateFileList(FileIDNo, txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt6.Text, txt7.Text, txt8.Text, txt9.Text, txt10.Text, txt11.Text);
                
            }
            else
            {
                UpdateFileList(0, txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt6.Text, txt7.Text, txt8.Text, txt9.Text, txt10.Text, txt11.Text);
                
            }
            txt1.Text=txt2.Text=txt3.Text=txt4.Text=txt5.Text=txt6.Text=txt7.Text=txt8.Text=txt9.Text=txt10.Text=txt11.Text = txt1.Text = "";
            btnFileSaveUpdate.Text = "حفظ";
            FileIDNo = 0;
            FillFilesView1();
            ColorFulGrid1(12,"");
            labdate.Visible = dataGridView1.Visible = true;
            panelFile.Visible = false;
            PanelMain.Visible = false;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            txtFileSearch.Text = dlg.FileName;
        }

        private void UpdateFileList(int iD, string fileNo, string fileDest, string individualNo, string conToWafidDate, string consToWafidNo, string wafidToWorkDate, string wafidToWorkNo, string workToWafidDate, string workToWafidNo, string workNotesNo, string workFinsihed )
        {
            string sql = "UPDATE TableFiles SET FileNo = @FileNo, FileDest = @FileDest, IndividualNo = @IndividualNo, ConToWafidDate = @ConToWafidDate, ConsToWafidNo = @ConsToWafidNo, WafidToWorkDate = @WafidToWorkDate, WafidToWorkNo = @WafidToWorkNo, WorkToWafidDate = @WorkToWafidDate, WorkToWafidNo = @WorkToWafidNo, WorkNotesNo = @WorkNotesNo, WorkFinsihed = @WorkFinsihed WHERE ID = @ID";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableFiles (FileNo, FileDest, IndividualNo, ConToWafidDate, ConsToWafidNo, WafidToWorkDate, WafidToWorkNo, WorkToWafidDate, WorkToWafidNo, WorkNotesNo, WorkFinsihed) values(@FileNo, @FileDest, @IndividualNo, @ConToWafidDate, @ConsToWafidNo, @WafidToWorkDate, @WafidToWorkNo, @WorkToWafidDate, @WorkToWafidNo, @WorkNotesNo, @WorkFinsihed)", sqlCon);
            if(iD != 0)
                sqlCmd = new SqlCommand(sql, sqlCon);
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
            //if (txtFileSearch.Text != "")
            //{
                
                
                if (imagecount == 1)
                    UpdateFileList(iD, PathImage[0]);
                else if (imagecount >= 1 && PathImage[0].Split('.')[1] == "docx")
                {
                
                using (DocX document = DocX.Load(PathImage[0]))
                    {
                        Paragraph p1 = document.InsertParagraph();

                        for (int x = 1; x < imagecount; x++)
                        {
                            var image = document.AddImage(PathImage[x]);
                            // Set Picture Height and Width.
                            var picture = image.CreatePicture(600, 500);

                            p1.AppendPicture(picture);
                        }
                        document.Save();
                    }
                    imagecount = 0;
                    UpdateFileList(iD, PathImage[0]);
                }
                else if (imagecount >= 1 && PathImage[0].Split('.')[1] != "docx")
                {
                
                string fileLocationDocx = FilesPathOut + "fileLocation" + DateTime.Now.ToString("mmss") + ".docx";
                    using (DocX document = DocX.Create(fileLocationDocx))
                    {
                        Paragraph p1 = document.InsertParagraph();

                        for (int x = 0; x < imagecount; x++)
                        {
                            var image = document.AddImage(PathImage[x]);
                            // Set Picture Height and Width.
                            var picture = image.CreatePicture(600, 500);

                            p1.AppendPicture(picture);
                        }


                        document.Save();

                    }
                    UpdateFileList(iD, fileLocationDocx);
                }

           // }   
                
        }

        private string AllOpenFilesDoc(int id, string table, int fileNo, string source)
        {
            string query = "select Data1, Extension1,ارشفة_المستندات from " + table + " where ID=@id"; ;
            string NewFileName = "";
            SqlConnection Con = new SqlConnection(source);
            switch (fileNo)
            {               
                case 2:
                    query = "select Data2, Extension2,المكاتبة_النهائية from " + table + " where ID=@id";
                    break;
                case 3:
                    query = "select Data3, Extension3,المكاتبة_الأولية from " + table + " where ID=@id";
                    break;
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
                    var name = reader["ارشفة_المستندات"].ToString();
                    if (string.IsNullOrEmpty(name)) return "";
                    var ext = reader["Extension1"].ToString();
                    if (string.IsNullOrEmpty(ext)) return "";
                    var Data = (byte[])reader["Data1"];
                    NewFileName = name.Replace(ext, id.ToString()) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                }
                else if (fileNo == 2) {
                    var name = reader["المكاتبة_النهائية"].ToString();
                    if (string.IsNullOrEmpty(name)) return "";
                    var ext = reader["Extension2"].ToString();
                    if (string.IsNullOrEmpty(ext)) return "";
                    var Data = (byte[])reader["Data2"];
                    NewFileName = name.Replace(ext, id.ToString()) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                }
                else if (fileNo == 3) {
                    var name = reader["المكاتبة_الأولية"].ToString();
                    if (string.IsNullOrEmpty(name)) return "";
                    var ext = reader["Extension3"].ToString();
                    if (string.IsNullOrEmpty(ext)) return "";
                    var Data = (byte[])reader["Data3"];
                    NewFileName = name.Replace(ext, id.ToString()) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                }
                System.Diagnostics.Process.Start(NewFileName);
            }
            Con.Close();
            return NewFileName;
        }

        private void updateCellData(string id, string filePath, string source, int index, string table) {
            
            string query = "UPDATE " + table + " SET Data1 = @Data1,Extension1 = @Extension1,ارشفة_المستندات = @ارشفة_المستندات where رقم_الهوية=@id";
            switch (index) {
                case 2:
                    query = "UPDATE " + table + " SET Data2 = @Data2,Extension2 = @Extension2,المكاتبة_النهائية = @المكاتبة_النهائية where رقم_الهوية=@id";
                    break;
                case 3:
                    query = "UPDATE " + table + " SET Data3 = @Data3,Extension3 = @Extension3,المكاتبة_الأولية = @المكاتبة_الأولية where رقم_الهوية=@id";
                    break;

            }
            SqlConnection sqlCon = new SqlConnection(source);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.Add("@id", SqlDbType.VarBinary).Value = id;
            if (index == 1)
            {
                using (Stream stream = File.OpenRead(filePath))
                {
                    byte[] buffer1 = new byte[stream.Length];
                    stream.Read(buffer1, 0, buffer1.Length);
                    var fileinfo1 = new FileInfo(filePath);
                    string extn1 = fileinfo1.Extension;
                    string DocName1 = fileinfo1.Name;
                    sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                    sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
                    sqlCmd.Parameters.Add("@ارشفة_المستندات", SqlDbType.NVarChar).Value = DocName1;

                }

            }
            else if (index == 2)
            {
                using (Stream stream = File.OpenRead(filePath))
                {
                    byte[] buffer1 = new byte[stream.Length];
                    stream.Read(buffer1, 0, buffer1.Length);
                    var fileinfo1 = new FileInfo(filePath);
                    string extn1 = fileinfo1.Extension;
                    string DocName1 = fileinfo1.Name;
                    sqlCmd.Parameters.Add("@Data2", SqlDbType.VarBinary).Value = buffer1;
                    sqlCmd.Parameters.Add("@Extension2", SqlDbType.Char).Value = extn1;
                    sqlCmd.Parameters.Add("@المكاتبة_النهائية", SqlDbType.NVarChar).Value = DocName1;

                }

            }
            else if (index == 3)
            {
                using (Stream stream = File.OpenRead(filePath))
                {
                    byte[] buffer1 = new byte[stream.Length];
                    stream.Read(buffer1, 0, buffer1.Length);
                    var fileinfo1 = new FileInfo(filePath);
                    string extn1 = fileinfo1.Extension;
                    string DocName1 = fileinfo1.Name;
                    sqlCmd.Parameters.Add("@Data3", SqlDbType.VarBinary).Value = buffer1;
                    sqlCmd.Parameters.Add("@Extension3", SqlDbType.Char).Value = extn1;
                    sqlCmd.Parameters.Add("@المكاتبة_الأولية", SqlDbType.NVarChar).Value = DocName1;

                }

            }
            sqlCmd.ExecuteNonQuery();

            sqlCon.Close();
        }

        private string OpenFilesDoc(int id)
        {
            string query;
            string NewFileName = "";
            SqlConnection Con = new SqlConnection(DataSource);
            query = "select Data1, Extension1,FileName1 from TableFiles  where ID=@id";
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@Id", SqlDbType.Int).Value = id;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                var name = reader["FileName1"].ToString();
                if (string.IsNullOrEmpty(name)) return "";
                var ext = reader["Extension1"].ToString();
                if (string.IsNullOrEmpty(ext)) return "";
                var Data = (byte[])reader["Data1"];
                NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                File.WriteAllBytes(NewFileName, Data);
                //System.Diagnostics.Process.Start(NewFileName);
            }
            Con.Close();
            return NewFileName;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Docfile);
        }

        private void DocDestin_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Console.WriteLine("DocDestin " + AffairIndex.ToString());
            //if (AffairIndex == 7) return;
            //    combFileNo.Items.Clear();
            //combFileNo.Text = "";
            //FillDataGridView1(AffairIndex);
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

       

        private void comPurpose_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            getText(DataSource);
        }

        private void Suddanese_Affair_SelectedIndexChanged(object sender, EventArgs e)
        {            
            if (!Suddanese_Affair.Enabled) 
                return;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();

            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID( 'TableWafid') ", sqlCon);

            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.Fill(dtblMain);
            string tableList = "";
            int i = 0;
            foreach (DataRow dataRow in dtblMain.Rows)
            {
                if (dataRow["name"].ToString() != "ID" && !dataRow["name"].ToString().Contains("Data") && !dataRow["name"].ToString().Contains("Extension") && !dataRow["name"].ToString().Contains("Data"))
                    if (i == 0) tableList = dataRow["name"].ToString();
                    else tableList = tableList + ", " + dataRow["name"].ToString();
                i++;
            }
            //MessageBox.Show(tableList +" ID= " + ApplicantID);
            int id = transferData(ApplicantID.ToString(), tableList, getTable(AffairIndex), getTable(Suddanese_Affair.SelectedIndex));
            UpdateTableTrans(id.ToString(), txtId.Text, ApplicantID.ToString(), getTable(Suddanese_Affair.SelectedIndex));
            //MessageBox.Show(id.ToString() +" - " +getTable(Suddanese_Affair.SelectedIndex) + " - "+ Suddanese_Affair.SelectedIndex.ToString());
            updataState(id, Suddanese_Affair.SelectedIndex.ToString(), "نوع_المعاملة", getTable(Suddanese_Affair.SelectedIndex));

            string comment = EmpName + " قام بتحويل المعاملة من " + Suddanese_Affair.Items[AffairIndex] + " إلى " + Suddanese_Affair.Text;
            if (txtOldComment.Text != "") comment = comment + Environment.NewLine + txtOldComment.Text;
            updataState(id, comment, "تعليق", getTable(Suddanese_Affair.SelectedIndex));
            
            deleteRowsData(ApplicantID, getTable(AffairIndex), DataSource);

            if (Suddanese_Affair.SelectedIndex == 4)
            {

                radioButton3.Checked = true; radioButton2.Checked = false;
            }
            else
            {
                radioButton3.Text = "أفراد الأسرة";
                radioButton3.Checked = false; radioButton2.Checked = true;
            }
            Suddanese_Affair.Enabled = false;
        }
        
        private void btnAddToList_Click(object sender, EventArgs e)
        {
            Save2DataBaseList(Convert.ToInt32(getFileNo(AffairIndex - 1)));

            clear_All();
            fileFiletering("");
            labdate.Visible = dataGridView1.Visible = true;
            PanelMain.Visible = false;
        }

        private int transferData(string id, string columns, string tableOrig, string destin)
        {
            int newId = 0;
            // MessageBox.Show(id.ToString() +"-"+ table + "-"+column + "-"+text);
            //string qurey = "update "+table+" set "+ column + "=@"+ column + " where ID=@id";
            string qurey = "insert into "+ destin + " ("+ columns+ ") select " + columns + " from " + tableOrig + " where ID = @id;SELECT @@IDENTITY as lastid";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            var reader = sqlCmd.ExecuteReader();
            if (reader.Read())
            {

                newId = Convert.ToInt32( reader["lastid"].ToString());
            }
            sqlCon.Close();
            return newId;
        }
        private void DocNo_TextChanged(object sender, EventArgs e)
        {
            if (grdiFill) return;
            
        }

        private void getPreDocIDs()
        {
            if (DocNo.Text.Length == 6 || DocNo.Text.Length == 9 || DocNo.Text.Length == 10)
            {
                for (int index = 0; index < 6; index++)
                {
                    if (checkUnique(getTable(index), DocNo.Text))
                    {
                        int rowIndex = fillFileBox8(getTable(index), DocNo.Text);
                        var selectedOption = MessageBox.Show("معاينة الإجراء السابق؟", "يوجد إجراء مطابق بنافذة " + Suddanese_Affair.Items[index].ToString(), MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (selectedOption == DialogResult.Yes)
                        {


                            dataGridView1_index(rowIndex);
                        }
                        break;
                    }
                }
            }
        }
        private void button21_Click(object sender, EventArgs e)
        {

        }

        private void txt3_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnAuth_Click(object sender, EventArgs e)
        {
            loadPic.Enabled = button10.Visible = reLoadPic.Visible = btnAuth.Enabled = false;

            try

            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.

                    var ScanerItem = device.Items[1]; // select the scanner.


                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);
                    
                    PathImage[imagecount] = FilesPathOut + "ScanImg" + DateTime.Now.ToString("mmss") + "_" + imagecount.ToString() + ".jpg";


                    if (File.Exists(PathImage[imagecount]))
                    {
                        File.Delete(PathImage[imagecount]);
                    }
                    imgFile.SaveFile(PathImage[imagecount]);
                    pictureBox3.ImageLocation = PathImage[imagecount];
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

            loadPic.Enabled = button10.Visible = reLoadPic.Visible = btnAuth.Enabled = true;
        }

        private void loadPic_Click(object sender, EventArgs e)
        {
            string fileName = loadDocxFile();
            if (fileName != "")
            {
                pictureBox3.ImageLocation = PathImage[imagecount] = fileName;
                imagecount++;
                btnAuth.BackColor = System.Drawing.Color.LightGreen;
                loadPic.Text = btnAuth.Text = "اضافة مستند آخر (" + imagecount.ToString() + ")";

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

        private void button10_Click_1(object sender, EventArgs e)
        {
            try


            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.

                    var ScanerItem = device.Items[1]; // select the scanner.


                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);

                    PathImage[imagecount - 1] = FilesPathOut + "ScanImg" + DateTime.Now.ToString("mmss") + "_" + imagecount.ToString() + ".jpg";


                    if (File.Exists(PathImage[imagecount - 1]))
                    {
                        File.Delete(PathImage[imagecount - 1]);
                    }
                    imgFile.SaveFile(PathImage[imagecount - 1]);
                    pictureBox3.ImageLocation = PathImage[imagecount - 1];
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

        private void reLoadPic_Click(object sender, EventArgs e)
        {
            string fileName = loadDocxFile();
            if (fileName != "")
            {
                pictureBox3.ImageLocation = PathImage[imagecount - 1] = fileName;
            }
        }

        private void panelFile_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            if (imagecount > 0)
            {
                btnAuth.Size = new System.Drawing.Size(155, 34);
                btnAuth.Location = new System.Drawing.Point(1093, 543);
                loadPic.Size = new System.Drawing.Size(155, 34);
                loadPic.Location = new System.Drawing.Point(1093, 583); 
                button10.Visible = reLoadPic.Visible = true;


            }
            else
            {
                btnAuth.Size = new System.Drawing.Size(311, 34);
                btnAuth.Location = new System.Drawing.Point(938, 543);
                loadPic.Size = new System.Drawing.Size(311, 34);
                loadPic.Location = new System.Drawing.Point(938, 583);
                button10.Visible = reLoadPic.Visible = false;
            }
            //if (combFileNo.Text == "99")
            //{
            //    comFinalaPro.Size = new System.Drawing.Size(210, 35);
            //    comFinalaPro.Location = new System.Drawing.Point(191, 651);
            //}
            //else {
            //    comFinalaPro.Size = new System.Drawing.Size(300, 35);
            //    comFinalaPro.Location = new System.Drawing.Point(101, 651);
            //}
            if (dataGridView1.Visible) comFinalaPro.SelectedIndex = 4;
            }

        private void button14_Click(object sender, EventArgs e)
        {
            //if (Jobposition.Contains("قنصل")) 
                Suddanese_Affair.Enabled = true;
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            checkDataInfo(true);
            //getPreDocIDs();
        }

        private int checkDataInfo(bool show) {
            grdiFill = true;
            picVerify.Visible = true;
            picVerified.Visible = false;
            int count = -1;
            for (int x = 0; x < 6; x++)
            {
                count = fillFileBoxSerach(x);

                for (int y = 0; y < comboBox3.Items.Count; y++)
                {

                    GridViewSeach(x, comboBox3.Items[y].ToString());
                    for (int z = 0; z < dataGridView6.RowCount - 1; z++)
                    {
                        //Console.WriteLine("Main = " +x.ToString()+" FileNo = "+ SearchfileList[y].ToString() +z.ToString()+ " - DocNo  " + dataGridView1.Rows[z].Cells[5].Value.ToString());
                        if (dataGridView6.Rows[z].Cells[5].Value.ToString() == DocNo.Text)
                        {
                            if (show)
                            {
                                var selectedOption = MessageBox.Show("عرض؟","رقم الهوية تطابق مع إجرا سابق", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (selectedOption == DialogResult.Yes)
                                {
                                    selectedID(z, dataGridView6);
                                }
                            }

                            grdiFill = false;
                            return Convert.ToInt32(dataGridView6.Rows[z].Cells[0].Value.ToString());
                        }
                    }

                }
            }
            picVerify.Visible = false;
            picVerified.Visible = true;
            return -1;
        }

        private void picVerified_Click(object sender, EventArgs e)
        {
            checkDataInfo(true);
        }

        private void button16_Click_1(object sender, EventArgs e)
        {
            //if (flowLayoutPanel1.Visible)
            //{
            //    button16.Text = "عرض الملاحظات";
            //    flowLayoutPanel1.Visible = false;
            //    Panelapp.Visible = true;
            //}
            //else {
            //    button16.Text = "أفراد الأسرة";
            //    flowLayoutPanel1.Visible = true;
            //    Panelapp.Visible = false;
            //}
        }

        private bool PreRequesttoFinish() {
            if (Suddanese_Affair.Text.Contains("نظامي") && (DocType.Text.Contains("جواز") || DocType.Text.Contains("وثيقة"))) 
            {
                MessageBox.Show("إجراءات مكتب العمل لابد ان تتم عبر الإقامة او رقم الحدود"); 
                return false; 
            }
            if ((DocType.Text.Contains("اقامة") || DocType.Text.Contains("رقم حدود")) && DocNo.Text.Length != 10)
            {
                MessageBox.Show("رقم رقم اثبات الشخصية غير صحيح"); 
                return false; 
            }
            if (!ModifyPermit)
            {
                MessageBox.Show("حساب الموظف غير مخول بإجراء تعديلات بنافذة " + this.Text + Environment.NewLine + "يرجى التواصل  مع مدير النظام");
                return false;
            }

            if (contract.CheckState == CheckState.Indeterminate)
            {
                radioButton2.Checked = true;
                MessageBox.Show("يرجى توضيح حالة التعاقد");
                return false;
            }
            if (txtBirth.Text.Length != 4)
            {
                radioButton2.Checked = true;
                MessageBox.Show("يرجى إضافة عام الميلاد  المكون من اربعة أرقم فقط");
                return false;
            }
            //if (txtJobGroup.Text == "" )
            //{
            //panelStatist.Visible = true;
            //flowbasicInfo.Visible = PaneTransfer.Visible = false;
            //panelStatist.BringToFront();
            //    MessageBox.Show("يرجى توضيح تصنيف الوظيف");
            //    return;
            //}
            if (txtDiffculties.Text == "")
            {
                radioButton2.Checked = true;
                MessageBox.Show("يرجى توضيح سبب ايقاف العمل");
                return false;
            }
            return true;
        }
        private void finalPro_Click(object sender, EventArgs e)
        {
           


            if (comFinalaPro.SelectedIndex <=3 && newData && !picVerified.Visible)
                checkDataInfo(true);
            //getPreDocIDs();



            //            طباعة وحفظ
            //حفظ فقط
            //طباعة فقط
            //إضافة إلى القائمة
            //تصدير قائمة
            //استبعاد من القوائم النشطة
            //طباعة ملخص الملفات

            if (comFinalaPro.SelectedIndex <= 3)
            {
                if (!PreRequesttoFinish()) return;
                if (!contract.Checked)
                {
                    var selectedOption = MessageBox.Show("", "لدى مقدم الطلب عقد عمل موثق؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (selectedOption == DialogResult.Yes)
                    {
                        contractState = "1";
                    }
                    if (selectedOption == DialogResult.No)
                    {
                        contractState = "0";
                    }
                }
                Authcases();
            }
            string activeCopy = FilesPathOut + ApplicantName.Text + DateTime.Now.ToString("mmss") + ".docx";
            switch (comFinalaPro.SelectedIndex) {
                case 0:
                    
                    Save2DataBase(Convert.ToInt32(combFileNo.Text), contractState);
                    printDocs(activeCopy);
                    if (combFileNo.Text == "99") MessageBox.Show("سيتم تنبيه مندوب القنصلية بعد الأرشفة النهائية");
                    break;
                case 1:
                    Save2DataBase(Convert.ToInt32(combFileNo.Text), contractState);
                    break;
                case 2:
                    printDocs(activeCopy);
                    if (combFileNo.Text == "99") MessageBox.Show("سيتم تنبيه مندوب القنصلية العامة تلقائيا بعد الأرشفة النهائية");
                    break;
                case 3:
                    break;
                case 4:
                    string fileNo = "WJ" + combFileNo.Text;
                    string fileDest = "جدة";
                    if (AffairIndex == 2)
                    {
                        fileNo = "WM" + combFileNo.Text;
                        fileDest = "مكة";
                    }
                    else if (AffairIndex == 3)
                    {
                        fileNo = "TJ" + combFileNo.Text;
                        fileDest = "الوافدين الشميسي";
                    }
                    else if (AffairIndex == 4)
                    {
                        fileNo = "TM" + combFileNo.Text;
                        fileDest = "المقابل المالي";
                    }
                    //MessageBox.Show(AffairIndex.ToString());
                    activeCopy = FilesPathOut +"قائمة " + Suddanese_Affair.Text + DateTime.Now.ToString("mmss") + "1.docx";
                    if (AffairIndex == 1 ) {
                        CreateGroupFileJeddah2(activeCopy, combFileNo.Text);
                        activeCopy = FilesPathOut + "قائمة " + Suddanese_Affair.Text + DateTime.Now.ToString("mmss") + "2.docx";
                        CreateGroupFileJeddah1(activeCopy, combFileNo.Text, " مدير إدارة الوافدين");
                        activeCopy = FilesPathOut + "قائمة " + Suddanese_Affair.Text + DateTime.Now.ToString("mmss") + "3.docx";
                        CreateGroupFileJeddah1(activeCopy, combFileNo.Text, "مدير مكتب العمل جدة");
                        filllExcelGrid(AffairIndex, combFileNo.Text);

                    } 
                    
                    else if (AffairIndex == 2) {
                        CreateGroupFile(activeCopy, combFileNo.Text);
                    } else if(AffairIndex == 3)
                    {
                        CreateGroupFileTarheel(activeCopy, combFileNo.Text);
                    }
                    
                    else if(AffairIndex == 4)
                    {
                            //MessageBox.Show(fileDest + AffairIndex.ToString());
                            var selectedOption = MessageBox.Show("", "طباعة جميع القائمة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (selectedOption == DialogResult.Yes)
                        {
                            CreateGroupFileTrans(activeCopy, combFileNo.Text,true);
                        } else if (selectedOption == DialogResult.No)
                        {
                            CreateGroupFileTrans(activeCopy, combFileNo.Text, false);
                        }
                    }


                    if(AffairIndex > 0 && AffairIndex < 5)
                    {
                        var selectionMessage = MessageBox.Show("", "إضافة الملف إلى قائمة ملخص الملفات؟", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                        if (selectionMessage == DialogResult.OK)
                        {
                            if (!checkExist(fileNo, "TableFiles"))
                                UpdateFileList(0, fileNo, fileDest, (dataGridView1.RowCount - 1).ToString(), GregorianDate.Text, (dataGridView1.RowCount - 1).ToString() + combFileNo.Text, "", "", "", "", "", "");
                        }
                    }
                    this.Close();
                    break;

            }
            if (!dataGridView1.Visible)
            {
                labdate.Visible = dataGridView1.Visible = true;
                PanelMain.Visible = false;
            }
            else if (Jobposition.Contains("قنصل") && dataGridView1.Visible)
            {
                labdate.Visible = dataGridView1.Visible = false;

                PanelMain.Visible = true;
            }
            finalPro.Enabled = true;
        }


        private bool checkExist(string documenNo, string TableFiles)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * from " + TableFiles + " where FileNo = @FileNo", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@FileNo", documenNo);

            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count == 0)
                return false;
            else return true;

        }

        private void comFinalaPro_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
         private void printDocs(string activeCopy)
        {
            finalPro.Enabled = false;
            //MessageBox.Show(FilesPathIn + @"استمارة مكتب العمل.docx");
            switch (AffairIndex) {
                case 0://خروج نهائي نظامي عام

                    activeCopy = FilesPathOut + "استمارة مكتب العمل " + DateTime.Now.ToString("mmss") + ".docx";
                    CreateWokOffice(activeCopy);
                    
                    activeCopy = FilesPathOut + "إقرار الوافدين " + DateTime.Now.ToString("mmss") + ".docx";
                    IqrarFinalExit(activeCopy);

                    activeCopy = FilesPathOut + "خطاب إدارة الوافدين " + DateTime.Now.ToString("mmss") + ".docx";                     
                    CreateWordFile(activeCopy, "الحالة");//خطاب إدارة الوافدين

                    break;
                case 1://خروج نهائي نظامي جدة
                    activeCopy = FilesPathOut + "استمارة مكتب العمل " + DateTime.Now.ToString("mmss") + ".docx";
                    CreateWokOffice(activeCopy);
                    
                    activeCopy = FilesPathOut + "أقرار الوافدين " + DateTime.Now.ToString("mmss") + ".docx"; 
                    IqrarFinalExit(activeCopy);
                    if (combFileNo.Text == "99")
                    {
                        activeCopy = FilesPathOut + "خطاب إدارة الوافدين " + DateTime.Now.ToString("mmss") + ".docx";
                        CreateWordFile(activeCopy, "الحالة");//خطاب إدارة الوافدين
                    }
                    break;
                case 2://خروج نهائي نظامي مكة
                    
                    activeCopy = FilesPathOut + "استمارة مكتب العمل " + DateTime.Now.ToString("mmss") + ".docx";
                    CreateWokOffice(activeCopy);
                    
                    activeCopy = FilesPathOut + "أقرار الوافدين " + DateTime.Now.ToString("mmss") + ".docx"; 
                    IqrarFinalExit(activeCopy);

                    if (combFileNo.Text == "99")
                    {
                        activeCopy = FilesPathOut + "خطاب إدارة الوافدين " + DateTime.Now.ToString("mmss") + ".docx";
                        CreateWordFile(activeCopy, "الحالة");//خطاب إدارة الوافدين
                    }
                    break;
                case 3://خروج نهائي بالترحيل
                    activeCopy = FilesPathOut + "إقرار الوافدين " + DateTime.Now.ToString("mmss") + ".docx";
                    IqrarTravel(activeCopy);

                    if (DocDestin.SelectedIndex != 0 || combFileNo.Text == "99")
                    {
                        activeCopy = FilesPathOut + "خطابات الوافدين " + DateTime.Now.ToString("mmss") + ".docx";
                        CreateWordFile(activeCopy, "الحالة");//خطاب إدارة الوافدين
                    }
                    
                    activeCopy = FilesPathOut + "إقرار مستحقات خروج نهائي" + DateTime.Now.ToString("mmss") + ".docx";                    
                    IqrarFinalExit(activeCopy);
                    break;
                case 4://تحويل المقابل المالي
                    activeCopy = FilesPathOut + "خطابات الوافدين " + DateTime.Now.ToString("mmss") + ".docx";
                    CreateWordFile(activeCopy, "صلة القرابة");

                    activeCopy = FilesPathOut + "إقرار مستحقات خروج نهائي" + DateTime.Now.ToString("mmss") + ".docx"; 
                    IqrarFinalExit(activeCopy);

                    if (DocDestin.SelectedIndex != 0 || combFileNo.Text == "99")
                    {
                        activeCopy = FilesPathOut + "خطابات الوافدين " + DateTime.Now.ToString("mmss") + ".docx";
                        CreateWordFile(activeCopy, "الحالة");//خطاب إدارة الوافدين
                    }
                    break;
                case 5://مخاطبات اللجنة العمالية
                    MessageBox.Show("لم يتم إدراج ملف للقسم .. يرجى مراجعة مدير النظام");
                    break;

            }
            this.Close();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            autoCompleteTextBox(DataSource, getTable(AffairIndex));
            panelStatist.Visible = true;
            flowbasicInfo.Visible = PaneTransfer.Visible = false;
            panelStatist.BringToFront();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            PaneTransfer.BringToFront();
            PersToServed.Checked = true;
            PaneTransfer.Visible = true;
            flowbasicInfo.Visible = panelStatist.Visible = false;
            if(Nobox == 0) Panelapp_Paint("", "", "");
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            flowbasicInfo.BringToFront();
            label5.BringToFront();
            PersToServed.BringToFront();
            label5.Visible = PersToServed.Visible = flowbasicInfo.Visible = true;
            PaneTransfer.Visible = false;
            panelStatist.Visible = false;
        }

        private void contract_CheckedChanged(object sender, EventArgs e)
        {
            if (contract.CheckState == CheckState.Checked) 
            { 
                contract.Text = "يوجد لدي عقد عمل موثق من السلطات السعودية."; 
            }
            else contract.Text = "لا يوجد لدي عقد عمل.";
        }

        private void preComment_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                checkBox2.Text = "إعادة التنبيه";
                smsText = "";
            }
        }

        private void comboStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Suddanese_Affair.SelectedIndex == 4) {
                if (comboStatus.SelectedIndex != 4) MessageBox.Show("إجراء المقابل المالي يكون لحالة إنتهاء الإقامة فقط");
            }
        }

        private void btnfileCheck_Click(object sender, EventArgs e)
        {
            //if (dataGridView1.CurrentRow.Index != -1)
            //{

            //        btnFileSaveUpdate.Text = "تعديل";
            //        panelFile.Visible = true;
            //        labdate.Visible = dataGridView1.Visible = false;
            //        PanelMain.Visible = false;

            //        FileIDNo = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            //        Docfile = OpenFilesDoc(FileIDNo);
            //        if (Docfile != "")
            //        {
            //            btnOpenFile.Visible = true;
            //        }
            //        else
            //        {
            //            btnOpenFile.Visible = false;
            //        }
            //        txt1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            //        txt2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            //        txt3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            //        txt4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            //        txt5.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            //        txt6.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            //        txt7.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            //        txt8.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            //        txt9.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            //        txt10.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            //        txt11.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            //        getDocInfo(FileIDNo);
            //    }
            }

        private void ready_CheckedChanged(object sender, EventArgs e)
        {
            if (ready.Checked)
                ready.Text = "مكتمل";
            else
                ready.Text = "غير مكتمل";
            updateReadyCase(ApplicantID, getTable(AffairIndex), ready.Text);
        }

        private void UpdateFileList(int iD, string filePath)
        {
            string sql = "UPDATE TableFiles SET Data1 = @Data1, Extension1 = @Extension1, FileName1 = @FileName1 WHERE ID = @ID";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableFiles (Data1, Extension1, FileName1) values (@Data1, @Extension1, @FileName1)", sqlCon);
            if (iD != 0)
                sqlCmd = new SqlCommand(sql, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            if (iD != 0)
                sqlCmd.Parameters.AddWithValue("@ID", iD);

            using (Stream stream = File.OpenRead(filePath))
            {
                byte[] buffer1 = new byte[stream.Length];
                stream.Read(buffer1, 0, buffer1.Length);
                var fileinfo1 = new FileInfo(filePath);
                string extn1 = fileinfo1.Extension;
                string DocName1 = fileinfo1.Name;
                sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
                sqlCmd.Parameters.Add("@FileName1", SqlDbType.NVarChar).Value = DocName1;               
            }

            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        


        public int Authcases()
        {
            AppIndex = 0;
            int name = 0;
            string names = "";
            int docno = 0;
            int relativity = 0;
            int titleIndex = 0;
            for (int x = 0; x < 10; x++) 
            {
                Names[x] = DocumentNo[x] = Relativity[x] = checkBoxSex[x] = "";
            }
            foreach (Control control in PaneTransfer.Controls)
            {
                
                if (control is TextBox && control.Name.Contains("txtAppName"))
                {
                    Names[name] = control.Text;
                    names = names + "_" + Names[name];
                    name++;
                }

                if (control is TextBox && control.Name.Contains("textBoxDocNo") && control.Tag.ToString() == "valid")
                {
                    DocumentNo[docno] = control.Text;
                    docno++;
                }

                if (control is ComboBox && control.Name.Contains("comboBoxDocType_") && control.Tag.ToString() == "valid")
                {
                    Relativity[relativity] = control.Text;
                    relativity++;
                }

                if (control is CheckBox && control.Name.Contains("checkBoxSex") && control.Tag.ToString() == "valid")
                {
                    if (((CheckBox)control).CheckState == CheckState.Unchecked)
                        checkBoxSex[AppIndex] = "ذكر";
                    else checkBoxSex[AppIndex] = "أنثى";
                    AppIndex++;
                }
            }
            for (int x = 0; x < Nobox; x++)
            {
                if (x == 0) titleIndex =  0;
                else if (x == 1 && checkBoxSex[1] == "ذكر")
                    titleIndex =  0;
                else if (x == 1 && checkBoxSex[1] == "أنثى")
                    titleIndex =  1;
                else if (x == 2 && checkBoxSex[1] == "ذكر" && checkBoxSex[2] == "ذكر")
                    titleIndex =  2;
                else if (x == 2 && checkBoxSex[1] == "أنثى" && checkBoxSex[2] == "ذكر")
                    titleIndex =  2;
                else if (x == 2 && checkBoxSex[1] == "ذكر" && checkBoxSex[2] == "أنثى")
                    titleIndex =  2;
                else if (x == 2 && checkBoxSex[1] == "أنثى" && checkBoxSex[2] == "أنثى")
                    titleIndex =  3;
                else if (x == 3 && checkBoxSex[1] == "أنثى" && checkBoxSex[2] == "أنثى" && checkBoxSex[3] == "أنثى")
                    titleIndex =  4;
                else titleIndex =  5;
            }
            //MessageBox.Show(names);
            return titleIndex;
        }
    }
}
