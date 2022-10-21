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

namespace PersAhwal
{
    public partial class PassAway : Form
    {
        string DataSource = "";
        string insertAll = "";
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
        //string defMandoub = "أبوبكر الصديق";
        public PassAway( int vcIndex, string dataSource, string filesPathIn, string filesPathOut, string jobposition, string empName, string gregorianDate, string hijriDate)
        {
            InitializeComponent();
            definColumn(dataSource);
            DataSource = dataSource;
            FilespathIn = filesPathIn;
            FilespathOut = filesPathOut;
            
           
            AtVCIndex = vcIndex;
            allList = getColList("TablePassAway");

            التاريخ_الهجري.Text = HijriDate = hijriDate;
            التاريخ_الميلادي.Text = GregorianDate = gregorianDate;

            fillFileBox(DataSource);

            dataGridView1.Visible = true;
            PanelMain.Visible = false;
            dataGridView1.BringToFront();
            colIDs[4] = موظف_الإدخال.Text = الموظف.Text = empName;
            موظف_الإدخال.Text = الموظف.Text = empName;
            fileComboBox(موقع_الإذن, DataSource, "ArabicAttendVC", "TableListCombo");
            if (موقع_الإذن.Items.Count > AtVCIndex)
                موقع_الإذن.SelectedIndex = AtVCIndex;
            else موقع_الإذن.SelectedIndex = 0;
        }
        private void definColumn(string dataSource)
        {
            DataSource = dataSource;
            for (int index = 0; index < 100;index++) 
                forbidDs[index] = "";

            forbidDs[0] = "تعليق";
            forbidDs[1] = "حالة_الارشفة";
            forbidDs[2] = "sms";
            foreach (Control control in PanelMain.Controls)
            {
                if ((control is TextBox || control is ComboBox || control is CheckBox) && !control.Name.Contains("Off"))
                {
                    if (!checkColumnName(control.Name, DataSource))
                    {
                        CreateColumn(control.Name, DataSource);
                    }
                }
            }
            for (int index = 0; forbidDs[index] != ""; index++) 
            {
                if (!checkColumnName(forbidDs[index].Replace(" ","_"), DataSource))
                {
                    CreateColumn(forbidDs[index].Replace(" ", "_"), DataSource);
                }
            }
        }
        private bool ready()
        {
            for (int i = 0; i < allList.Length; i++)
            {
                foreach (Control control in PanelMain.Controls)
                {
                    if (control.Name == allList[i] && control.Visible)
                    {
                        
                        
                        if (control.Name == "الميلاد" && control.Text == "")
                        {
                            MessageBox.Show("يرجى إضافة عام الميلاد لخانة " + control.Name.Replace("_"," ")); return false;
                        }

                        else if (control.Visible && (control.Name.Contains("هاتف") && control.Text.Length != 12))
                        {
                            MessageBox.Show("يرجى إضافة رقم الهاتف لخانة " + control.Name.Replace("_", " ")); return false;
                        }
                        else if (control.Visible && (control.Name.Contains("قامة") && control.Text.Length != 10))
                        {
                            MessageBox.Show("يرجى إضافة رقم الإقامة بصورة صحيحة لخانة " + control.Name.Replace("_", " ")); return false;
                        }
                        else if (control.Visible && (control.Name.Contains("جواز") && control.Text.Length != 9))
                        {
                            MessageBox.Show("يرجى إضافة رقم الجواز بصورة صحيحة لخانة " + control.Name.Replace("_", " ")); return false;
                        }
                        else if (control.Text == "")
                        {
                            MessageBox.Show("يرجى إضافة بيانات خانة " + control.Name.Replace("_"," ")); return false;
                        }
                    }
                }
            }
            return true;
        }

        private void CreateColumn(string Columnname, string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("alter table TablePassAway add " + Columnname.Replace(" ", "_") + " nvarchar(150)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private bool checkColumnName(string colNo, string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TablePassAway", sqlCon);
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

                if (row["name"].ToString() != "ID" && row["name"].ToString() != "حالة_الارشفة")
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
            insertAll = "INSERT INTO " + table + "(" + insertItems + ") values (" + insertValues + ")";
            //MessageBox.Show(updateAll);
            return allList;

        }
        private void fillFileBox(string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TablePassAway", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtblMain = new DataTable();
            sqlDa.Fill(dtblMain);
            dataGridView1.DataSource = dtblMain;
            sqlCon.Close();
            dataGridView1.Columns[0].Visible = false;
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


        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void PassAway_Load(object sender, EventArgs e)
        {
            //autoCompleteTextBox(المهنة, DataSource, "jobs", "TableListCombo");
            fileComboBox(اسم_المندوب, DataSource, "MandoubNames", "TableListCombo");
            fileComboBox(اسم_المندوب_للتصدير, DataSource, "MandoubNames", "TableMandoudList");
            
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

        private string commentInfo()
        {
            string comment = "";
            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text == "")
                comment = "";

            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text != "")
                comment = التعليقات_السابقة_Off.Text;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text == "")
                comment = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + GregorianDate + Environment.NewLine + "--------------" + Environment.NewLine;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text != "")
                comment = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + GregorianDate + Environment.NewLine + "--------------" + Environment.NewLine + "*" + التعليقات_السابقة_Off.Text.Trim();

            return comment;
        }

        private void addarchives()
        {
            //colIDs[0] = رقم_اذن_الدفن.Text;
            //colIDs[1] = genIDNo.ToString();
            //colIDs[2] = GregorianDate;
            //colIDs[3] = اسم_المتوفى.Text;
            //colIDs[4] = الموظف.Text;
            //colIDs[5] = طريقة_تقديم_الطلب.Text;
            //colIDs[6] = اسم_المندوب.Text;
            //colIDs[7] = "new";

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
            sqlCommand.CommandType = CommandType.Text;
            for (int i = 1; i < allList.Length; i++)
            {
                
                sqlCommand.Parameters.AddWithValue("@" + allList[i], colIDs[i - 1]);
            }
            sqlCommand.ExecuteNonQuery();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            if (!ready()) return;
            button1.Enabled = false; 
            save2DataBase();
            if (newData)
            {
                colIDs[0] = رقم_اذن_الدفن.Text;
                colIDs[1] = genIDNo.ToString();
                colIDs[2] = GregorianDate;
                colIDs[3] = اسم_المتوفى.Text;
                colIDs[4] = الموظف.Text;
                colIDs[5] = طريقة_تقديم_الطلب.Text;
                colIDs[6] = اسم_المندوب.Text;
                colIDs[7] = "new";
                addarchives();
            }
            CreateAuth();
            this.Close();
        }
        private void save2DataBase()
        {
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(updateAll, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@id", genIDNo);
            for (int i = 0; i < allList.Length; i++)
            {
                if (allList[i] == "تعليق")
                {
                    sqlCommand.Parameters.AddWithValue("@" + allList[i], commentInfo());
                }
                else if (allList[i] == "sms")
                {
                    sqlCommand.Parameters.AddWithValue("@" + allList[i], "");
                }
                else
                    foreach (Control control in PanelMain.Controls)
                    {
                        if (control.Name == allList[i])
                        {
                            sqlCommand.Parameters.AddWithValue("@" + allList[i], control.Text);
                            break;
                        }
                    }
            }
            sqlCommand.ExecuteNonQuery();
        }
        private int CountUnsubmittedDoc(string Mandoub)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);

            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();

            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from  archives where mandoubName=@mandoubName", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@mandoubName", Mandoub+"*");
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            return dtbl.Rows.Count;
        }
        private void CreatePermitForm(string AuthID, string AuthName, string DocxInFile, string DocxOutFile, string pdfouput)
        {
            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();
            object objCurrentCopy = DocxInFile;
            Word.Document oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);
            oBMicroWord.Selection.Find.ClearFormatting();
            oBMicroWord.Selection.Find.Replacement.ClearFormatting();

            object ParaAuthIDNo = "MarkAuthIDNo";
            Word.Range BookAuthIDNo = oBDoc.Bookmarks.get_Item(ref ParaAuthIDNo).Range;
            BookAuthIDNo.Text = GregorianDate + Environment.NewLine + AuthID;
            object rangeAuthIDNo = BookAuthIDNo;
            oBDoc.Bookmarks.Add("AuthAuthIDNo", ref rangeAuthIDNo);
            
            object ParaAuthIDName = "MarkAuthIDName";
            Word.Range BookAuthIDName = oBDoc.Bookmarks.get_Item(ref ParaAuthIDName).Range;
            BookAuthIDName.Text = AuthName;
            object rangeAuthIDName = BookAuthIDName;
            oBDoc.Bookmarks.Add("AuthAuthIDName", ref rangeAuthIDName);

            oBDoc.SaveAs2(DocxOutFile);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(pdfouput);
            File.Delete(DocxOutFile);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;
        }

        private void CreateAuth()
        {
            string DocxInFile = FilespathIn + "إذن دفن.docx";
            //MessageBox.Show(DocxInFile);
            //string wordOutFile = FilespathOut + "إذن دفن.docx" + DateTime.Now.ToString("ssmm") + ".docx";
            نص_الشهادة.Text = "القنصـلية العـامة لجمهـورية الســودان بجـدة،  بأن المواطننوع_مقدم_الطلب السودانينوع_مقدم_الطلب السيدنوع_مقدم_الطلب/***  حاملنوع_مقدم_الطلب ^^^ ه### أقرب الأقربين للمواطننوع_المتوفى السوادنينوع_المتوفى المتوف@@@ المرحومنوع_المتوفى بإذن الله السيدنوع_المتوفى/!!! حاملنوع_المتوفى &&& $$$ وافته%%% المنية بمدينة *&*، ترجو القنصلية العامة من جهات الاختصاص بالمملكة العربية السعودية تسليم جثمانه%%% إلى المذكور نوع_مقدم_الطلب أعلاه وتسهيل مهمة دفنه%%% محليا";
            
            for (int x = 0; x < 20; x++)
                نص_الشهادة.Text = SuffPrefReplacements(نص_الشهادة.Text);

            string[] markLits = new string[10] { "","","","","","","","","",""};
            markLits[0] = "التاريخ_الميلادي";
            markLits[1] = "التاريخ_الهجري";
            markLits[2] = "رقم_اذن_الدفن";
            markLits[3] = "موقع_الإذن";
            markLits[4] = "نص_الشهادة";

            //

            object oBMiss = System.Reflection.Missing.Value;
            Word.Application oBMicroWord = new Word.Application();

            object objCurrentCopy = DocxInFile;

            Word.Document oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);
            oBMicroWord.Selection.Find.ClearFormatting();
            oBMicroWord.Selection.Find.Replacement.ClearFormatting();
            for (int index = 0; markLits[index] != ""; index++)
            {
                foreach (Control control in PanelMain.Controls)
                {
                    if (markLits[index] == control.Name)
                    {
                        object ParaAuthIDNo = markLits[index];
                        try
                        {
                            //MessageBox.Show(control.Name + " - "+control.Text);
                            Word.Range BookAuthIDNo = oBDoc.Bookmarks.get_Item(ref ParaAuthIDNo).Range;
                            BookAuthIDNo.Text = control.Text;
                            object rangeAuthIDNo = BookAuthIDNo;
                            oBDoc.Bookmarks.Add(markLits[index], ref rangeAuthIDNo);
                        }
                        catch (Exception ex) { MessageBox.Show(markLits[index]); }
                    }
                    
                }
            }
            string docxouput = FilespathOut + "إذن دفن رقم " + رقم_اذن_الدفن.Text.Split('/')[4] + DateTime.Now.ToString("ssmm") + ".docx";
            string pdfouput = FilespathOut + "إذن دفن رقم " + رقم_اذن_الدفن.Text.Split('/')[4] + DateTime.Now.ToString("ssmm") + ".pdf";
            oBDoc.SaveAs2(docxouput);
            oBDoc.ExportAsFixedFormat(pdfouput, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(pdfouput);
            File.Delete(docxouput);
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;

        }
        private string SuffPrefReplacements(string text)
        {
            //MessageBox.Show(text);
            string str1 = "";
            string str2 = "";
            string str3 = "و";
            string str4 = "ى";
            string str5 = "الذي";
            string str6 = "";
            string text1 ="", text2="";
            if (جواز_المتوفى.Text != "" && اقامة_المتوفى.Text != "") text1 = "جواز سفر بالرقم " + جواز_المتوفى.Text + " وإقامة بالرقم " + اقامة_المتوفى.Text;
            else if (جواز_المتوفى.Text == "" && اقامة_المتوفى.Text != "") text1 = "إقامة بالرقم " + اقامة_المتوفى.Text;
            else if (جواز_المتوفى.Text != "" && اقامة_المتوفى.Text == "") text1 = "جواز سفر بالرقم " + جواز_المتوفى.Text;

            if (جواز_اقرب_الاقربين.Text != "" && إقامة_اقرب_الاقربين.Text != "") text2 = "جواز سفر بالرقم " + جواز_اقرب_الاقربين.Text + " وإقامة بالرقم " + إقامة_اقرب_الاقربين.Text;
            else if (جواز_اقرب_الاقربين.Text == "" && إقامة_اقرب_الاقربين.Text != "") text2 = "إقامة بالرقم " + إقامة_اقرب_الاقربين.Text;
            else if (جواز_اقرب_الاقربين.Text != "" && إقامة_اقرب_الاقربين.Text == "") text2 = "جواز سفر بالرقم " + جواز_اقرب_الاقربين.Text;

            if (!نوع_المتوفى.Checked)
            {
                str1 = "ة";
               str4 = "ية";
                str5 = "التي";
                str6 = "ا";

            }
            if (!نوع_مقدم_الطلب.Checked)
            {
                str2 = "ة";
                str3 = "ي";
            }

            if (text.Contains("نوع_المتوفى"))
                return text.Replace("نوع_المتوفى", str1);
            else if (text.Contains("نوع_مقدم_الطلب"))
                return text.Replace("نوع_مقدم_الطلب", str2);
            else if (text.Contains("###"))
                return text.Replace("###", str3);
            else if (text.Contains("@@@"))
                return text.Replace("@@@", str4);
            else if (text.Contains("$$$"))
                return text.Replace("$$$", str5);
            else if (text.Contains("%%%"))
                return text.Replace("%%%", str6);
            else if (text.Contains("^^^"))
                return text.Replace("^^^", text2);
            else if (text.Contains("&&&"))
                return text.Replace("&&&", text1);
            else if (text.Contains("!!!"))
                return text.Replace("!!!", اسم_المتوفى.Text);
            else if (text.Contains("***"))
                return text.Replace("***", اسم_اقرب_الاقربين.Text);
            else if (text.Contains("*&*"))
                return text.Replace("*&*", مكان_الوفاة.Text);
            else 
                return text;
        }

        
        private void placeTextBookMark(Document oBDoc, string text, string strText)
        {
            
        }
        private void طريقة_تقديم_الطلب_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (طريقة_تقديم_الطلب.SelectedIndex != 0)
                اسم_المندوب.Visible = true;
            else
            {
                اسم_المندوب.Text = "";
                اسم_المندوب.Visible = false;
            }
        }

        private void btnListView_Click(object sender, EventArgs e)
        {
            if (PanelMain.Visible)
            {
                dataGridView1.Visible = true;
                PanelMain.Visible = false;
                dataGridView1.BringToFront();
            }
            else {
                dataGridView1.Visible = false;
                PanelMain.Visible = true;
                dataGridView1.SendToBack();
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                genIDNo = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                التعليقات_السابقة_Off.Text = dataGridView1.CurrentRow.Cells["تعليق"].Value.ToString();
                dataGridView1.Visible = false;
                PanelMain.Visible = true;
                foreach (Control control in PanelMain.Controls)
                {
                    if ((control is TextBox || control is ComboBox || control is CheckBox) && !control.Name.Contains("Off"))
                    {
                        if (!control.Name.Contains("التاريخ") && !control.Name.Contains("موظف"))
                            control.Text = dataGridView1.CurrentRow.Cells[control.Name].Value.ToString();
                    }
                }

                if (نوع_مقدم_الطلب.Text == "" || نوع_مقدم_الطلب.Text == "ذكر") نوع_مقدم_الطلب.Checked = true;
                if (نوع_المتوفى.Text == ""|| نوع_المتوفى.Text == "ذكر") نوع_المتوفى.Checked = true;

                if (dataGridView1.CurrentRow.Cells["اسم_المتوفى"].Value.ToString() == "")
                {
                    newData = true;
                    FillDatafromGenArch("data1", genIDNo.ToString(), "TablePassAway");
                }
                AddEdit = false;
                
            }
        }

        void FillDatafromGenArch(string doc, string id, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableGeneralArch where  رقم_المرجع='" + id + "' and نوع_المستند='" + doc + "' and docTable = '" + table + "'", sqlCon);
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

        private void btnFile1_Click(object sender, EventArgs e)
        {
            FillDatafromGenArch("data1", genIDNo.ToString(), "TablePassAway");
        }

        private void btnFile2_Click(object sender, EventArgs e)
        {
            FillDatafromGenArch("data1", genIDNo.ToString(), "TablePassAway");
        }

        private void deleteRow_Click(object sender, EventArgs e)
        {
            var selectedOption = MessageBox.Show("", "سيتم حذف المستند وجميع الملفات المتعلقة به؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                deleteRowsData(رقم_اذن_الدفن.Text, "TablePassAway");
                //deleteRowsData(رقم_اذن_الدفن.Text, "TableGeneralArch");
                //deleteRowsData(رقم_اذن_الدفن.Text, "archives");
            }
        }

        private void deleteRowsData(string v1, string table)
        {
            string query;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "DELETE FROM " + table + " where رقم_اذن_الدفن = @رقم_اذن_الدفن";
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_اذن_الدفن", v1);
            sqlCmd.ExecuteNonQuery();
            Con.Close();
        }

        private string DocIDGenerator()
        {
            string AuthNoPart1 = "ق س ج/80/" + DateTime.Now.Year.ToString().Replace("20", "") + "/16/1";
            while (checkISUnique(AuthNoPart1))
            {
                string rowCount = (getMaxDocNo("TablePassAway", AuthNoPart1, "رقم_اذن_الدفن") + 1).ToString();
                
                AuthNoPart1 = "ق س ج/80/" + DateTime.Now.Year.ToString().Replace("20", "") + "/16/" + rowCount;
            }
            return AuthNoPart1;
        }

        private int getMaxDocNo(string table, string docid, string colName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select " + colName + " from " + table + " where " + colName + " like N'ق س ج/80/" + DateTime.Now.Year.ToString().Replace("20", "") + "/%'", sqlCon);
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
                    string newInfo = dataRow[colName].ToString().Split('/')[4];
                    int id = Convert.ToInt32(newInfo);
                    if (id > maxID) maxID = id;
                    Console.WriteLine("maxID " + maxID);
                }

            }
            return maxID;
        }

        private bool checkISUnique(string docid)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select رقم_اذن_الدفن from TablePassAway where رقم_اذن_الدفن=N'" + docid + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count != 0) return true;
            else return false;
        }

        private int NewReportEntry(string dataSource, string authNo)
        {
            string query = "INSERT INTO TablePassAway (التاريخ_الميلادي,رقم_اذن_الدفن,التاريخ_الهجري,اسم_المندوب,طريقة_تقديم_الطلب) values (@التاريخ_الميلادي,@رقم_اذن_الدفن,@التاريخ_الهجري,@اسم_المندوب,@طريقة_تقديم_الطلب);SELECT @@IDENTITY as lastid";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_اذن_الدفن", authNo);
            sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", GregorianDate);
            sqlCmd.Parameters.AddWithValue("@التاريخ_الهجري", HijriDate);
            sqlCmd.Parameters.AddWithValue("@اسم_المندوب", اسم_المندوب_للتصدير.Text);
            sqlCmd.Parameters.AddWithValue("@طريقة_تقديم_الطلب", طريقة_تقديم_الطلب.Items[1].ToString());
            var reader = sqlCmd.ExecuteReader();
            if (reader.Read())
            {
                return Convert.ToInt32(reader["lastid"].ToString());
            }
            sqlCon.Close();
            return 0;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            اسم_المندوب_للتصدير.Visible = عدد_الاستمارات.Visible = true;
            
        }

        private void عدد_الاستمارات_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (اسم_المندوب_للتصدير.Text == "" || اسم_المندوب_للتصدير.Text == "اسم المندوب") {
                MessageBox.Show("يرجى إختيار اسم المندوب وتحديد عدد الاستمارات المطلوبة");return; 
            }

            int count = CountUnsubmittedDoc(اسم_المندوب_للتصدير.Text);
            if (count != 0)
            {
                MessageBox.Show("مندوب القنصلية لديه عدد(" + count.ToString() + ") من المكاتبات لم يقم بأرشفتها.. لا يمكن المتابعة");
                return;
            }
            var selectedOption = MessageBox.Show("", "تأكيد تصدير عدد  " + (عدد_الاستمارات.SelectedIndex + 1).ToString() + " استمارة للسيد/ " + اسم_المندوب_للتصدير.Text + "؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                for (int docNo = 0; docNo <= عدد_الاستمارات.SelectedIndex; docNo++)
                {
                    string wordInFile = FilespathIn.Replace("ModelFiles", "FormData") + "استمارة إذن دفن.docx";
                    string wordOutFile = FilespathOut + "استمارة إذن دفن" + docNo.ToString() + ".docx" + DateTime.Now.ToString("ssmm") + ".docx";
                    string pdfOutFile = FilespathOut + "استمارة إذن دفن" + docNo.ToString() + ".docx" + DateTime.Now.ToString("ssmm") + ".pdf";
                    string docID = DocIDGenerator();
                    CreatePermitForm(docID, اسم_المندوب_للتصدير.Text, wordInFile, wordOutFile, pdfOutFile);
                    int genIDNo = NewReportEntry(DataSource, DocIDGenerator());
                    colIDs[0] = docID;
                    colIDs[1] = genIDNo.ToString();
                    colIDs[2] = GregorianDate;
                    colIDs[3] = اسم_المتوفى.Text;
                    colIDs[4] = الموظف.Text;
                    colIDs[5] = طريقة_تقديم_الطلب.Items[1].ToString();
                    colIDs[6] = اسم_المندوب_للتصدير.Text + "*";
                    colIDs[7] = "new";
                    addarchives();
                }
                اسم_المندوب_للتصدير.Visible = عدد_الاستمارات.Visible = false;
            }
        }

        private void نوع_المتوفى_CheckedChanged(object sender, EventArgs e)
        {
            if (نوع_المتوفى.Checked) نوع_المتوفى.Text = "ذكر";
            else نوع_المتوفى.Text = "أنثى";
        }

        private void نوع_مقدم_الطلب_CheckedChanged(object sender, EventArgs e)
        {
            if (نوع_مقدم_الطلب.Checked) نوع_مقدم_الطلب.Text = "ذكر";
            else نوع_مقدم_الطلب.Text = "أنثى";
        }

        private void اسم_المندوب_للتصدير_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
