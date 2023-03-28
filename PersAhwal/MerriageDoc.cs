using DocumentFormat.OpenXml.Drawing.Diagrams;

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
using System.Data.SqlTypes;
using System.Xml.Linq;

namespace PersAhwal
{
    public partial class MerriageDoc : Form
    {
        string DataSource = "";
        string insertAll = "";
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
        string FilespathIn = "";
        string FilespathOut = "";
        Word.Document oBDoc;
        object oBMiss;
        Word.Application oBMicroWord;
        bool button1Enabel = false;
        string localCopy_off = "";
        public MerriageDoc(string dataSource, bool addEdit, string empName, int atVCIndex, string gregorianDate, string hijriDate, string filespathIn, string filespathOut)
        {
            InitializeComponent();
            DataSource = dataSource;
            AddEdit = addEdit;
            AtVCIndex = atVCIndex;
            allList = getColList("TableMerrageDoc");
            FilespathIn = filespathIn;
            FilespathOut = filespathOut;
            التاريخ_الهجري.Text = HijriDate = hijriDate;
            التاريخ_الميلادي.Text = GregorianDate = gregorianDate;
            fillFileBox(DataSource);
            if (AddEdit) {
                dataGridView1.Visible = false;
                PanelMain.Visible = true;
            }
            else {
                dataGridView1.Visible = true;
                PanelMain.Visible = false;
            }

            colIDs[4] = موظف_الإدخال .Text = الموظف.Text = empName;
            موظف_الإدخال.Text = الموظف.Text = empName;
            fileComboBox(المأذون, DataSource, "ArabicAttendVC", "TableListCombo");
            if(المأذون.Items.Count> AtVCIndex)
                المأذون.SelectedIndex = AtVCIndex;
            else المأذون.SelectedIndex = 0;
            
            طريقة_الإجراء.SelectedIndex = 0;
            
            definColumn(DataSource);
        }
        private void definColumn(string dataSource)
        {
            DataSource = dataSource;
            for (int index = 0; index < 100; index++)
                forbidDs[index] = "";

            forbidDs[0] = "تعليق";
            forbidDs[1] = "حالة_الارشفة";
            forbidDs[2] = "sms";
            foreach (Control control in PanelMain.Controls)
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

        private void CreateColumn(string Columnname, string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("alter table TableMerrageDoc add " + Columnname.Replace(" ", "_") + " nvarchar(150)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private bool checkColumnName(string colNo, string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableMerrageDoc", sqlCon);
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
                
                if (row["name"].ToString() != "ID" && row["name"].ToString() != "تاريخ_الارشفة1"&& row["name"].ToString() != "تاريخ_الاجراء" && row["name"].ToString() != "تاريخ_الارشفة2" && row["name"].ToString() != "حالة_الارشفة"&& row["name"].ToString() != "sms")
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
            updateAll = "UPDATE "+ table+" SET " + updateValues + " where ID = @id";
            insertAll = "INSERT INTO " + table + "(" + insertItems + ") values (" + insertValues + ")";
            
            return allList;

        }
        
        private bool getDate(TextBox textBox)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT DATEDIFF(day, '" + textBox.Text + "', '" + GregorianDate + "')/365 as daysDiff", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
                sqlCon.Close();

                foreach (DataRow row in dtbl.Rows)
                {
                    int age = Convert.ToInt32(row["daysDiff"].ToString());
                    if (age <= 18)
                    {
                        textBox.BackColor = System.Drawing.Color.MistyRose;
                        MessageBox.Show("عمر حرج ويحتاج إلى موافق صريحة من المأذون");

                        return true;
                    }
                    else
                    {
                        textBox.BackColor = System.Drawing.Color.White;
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("تاريخ ميلاد غير صالح"); }
            return false;

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (طريقة_الإجراء.SelectedIndex == 1) {
                labhusSideName.Visible = وكيل_الزوج.Visible = labhusSidePass.Visible = جواز_وكيل_الزوج.Visible = true;
            }
            else 
                labhusSideName.Visible = وكيل_الزوج.Visible = labhusSidePass.Visible = جواز_وكيل_الزوج.Visible = false;
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
            sqlCommand.CommandType = CommandType.Text;
            for (int i = 1; i < allList.Length; i++)
            {
                sqlCommand.Parameters.AddWithValue("@" + allList[i], text[i - 1]);
            }
            sqlCommand.ExecuteNonQuery();
        }

        private void fillFileBox(string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();            
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableMerrageDoc order by ID desc", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtblMain = new DataTable();
            sqlDa.Fill(dtblMain);
            dataGridView1.DataSource = dtblMain;
            sqlCon.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Width = 170;
            dataGridView1.Columns[2].Width = dataGridView1.Columns[3].Width = 200;
            ColorFulGrid9();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                genIDNo = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                foreach (Control control in PanelMain.Controls)
                {
                    if ((control is TextBox || control is ComboBox || control is CheckBox) && !control.Name.Contains("Off"))
                    {
                        if (!control.Name.Contains("التاريخ") && !control.Name.Contains("موظف"))
                            try
                            {
                                control.Text = dataGridView1.CurrentRow.Cells[control.Name].Value.ToString();
                            }
                            catch (Exception ex) { }
                    }
                    
                }
                gridFill = false;
                التعليقات_السابقة_Off.Text = dataGridView1.CurrentRow.Cells["تعليق"].Value.ToString();
                if (dataGridView1.CurrentRow.Cells["اسم_الزوج"].Value.ToString() == "")
                {
                    newData = true;
                    FillDatafromGenArch("data1", genIDNo.ToString(), "TableMerrageDoc");
                }
                
                AddEdit = false;
                labDescribed.Visible = dataGridView1.Visible = false;
                PanelMain.Visible = true;
                backgroundWorker1.RunWorkerAsync();
            }
            //if (رقم_الوثيقة.Text != "" && رقم_الوثيقة.Text != "بدون") 
            //    رقم_الوثيقة.Enabled = false;
            //else رقم_الوثيقة.Enabled = true;
            
        }

        private bool ready()
        {
            for (int i = 0; i < allList.Length; i++)
            {
                foreach (Control control in PanelMain.Controls)
                {
                    if (control.Name == allList[i])
                    {
                        if (control.Visible && (control.Text == "" || (control is ComboBox && ((ComboBox)control).SelectedIndex == -1)))
                        {
                            MessageBox.Show("يرجى إضافة بيانات " + control.Name); return false;
                        }

                        if (control.Visible && (control.Name.Contains("ميلاد_") && control.Text.Length != 10))
                        {
                            MessageBox.Show("يرجى إضافة عام الميلاد لخانة " + control.Name); return false;
                        }

                        if (control.Visible && (control.Name.Contains("جواز") && control.Text.Length != 9))
                        {
                            MessageBox.Show("يرجى إضافة رقم الجواز بصورة صحيحة لخانة " + control.Name); return false;
                        }
                        
                        if (عمر_الزوج_الحرج)
                        {
                            MessageBox.Show("عمر أحد الزوجين أقل من العمر الذي نص عليه القانون " + control.Name);
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        private void ColorFulGrid9()
        {

            int arch = 0;
            int inComb = 0;
            int i = 0;
            for (; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[2].Value.ToString() == "")
                {
                    inComb++;
                }
                if (dataGridView1.Rows[i].Cells["حالة_الارشفة"].Value.ToString() == "مؤرشف نهائي")
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    
                    arch++;
                }
            }
            labDescribed.Text = "عدد (" + i.ToString() + ") معاملة .. عدد (" + inComb.ToString() + ") غير مكتمل.. والمؤرشف منها عدد (" + arch.ToString() + ")...";

        }
        void FillDatafromGenArch(string doc, string id, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource.Replace("AhwalDataBase", "ArchFilesDB"));
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableGeneralArch where  رقم_المرجع='" + id + "' and نوع_المستند='" + doc + "' and docTable = '" + table+"'", sqlCon);
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

        private void button1_Click(object sender, EventArgs e)
        {            
            string part1to3 = رقم_المعاملة.Text.Split('/')[0] + "/" + رقم_المعاملة.Text.Split('/')[1] + "/" + رقم_المعاملة.Text.Split('/')[2] + "/" + رقم_المعاملة.Text.Split('/')[3] + "/";
            رقم_المعاملة.Text = part1to3 + رقم_الوثيقة.Text;
            addNewAppNameInfo1(اسم_الزوج);
            addNewAppNameInfo2(اسم_الزوجة);
            addNewAppNameInfo3(الشاهد_الاول, وثيقة_الشاهد_الاول);
            addNewAppNameInfo3(الشاهد_الثاني, وثيقة_الشاهد_الثاني);
            if (!ready()) return;
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
            updateGenName(رقم_المعاملة.Text, genIDNo.ToString());
            if (!checkSentSMS(genIDNo, "TableMerrageDoc")) 
                SMS(genIDNo, "TableMerrageDoc");
            
            if (newData) {
                colIDs[0] = رقم_المعاملة.Text; 
                colIDs[1] = genIDNo.ToString();
                colIDs[2] = GregorianDate;
                colIDs[3] = اسم_الزوج.Text;
                colIDs[4] = الموظف.Text;
                colIDs[5] = "حضور مباشرة إلى القنصلية";
                colIDs[6] = "";
                colIDs[7] = "new";
                addarchives(colIDs);
            }
            fillPreDoc();
            fillDocFileAppInfo();
            fillPrintDocx();
            
            this.Close();
        }
        private void addNewAppNameInfo1(TextBox textName)
        {

            string query = "insert into TableGenNames ([الاسم], رقم_الهوية,تاريخ_الميلاد,المهنة) values (@col1,@col2,@col3,@col4) ;SELECT @@IDENTITY as lastid";
            string id = checkExist(textName.Text);
            if (id != "0")
            {
                query = "update TableGenNames set [الاسم] =  @col1,[رقم_الهوية] = @col2,[تاريخ_الميلاد] = @col3,[المهنة] = @col4 where ID = " + id;
                //MessageBox.Show(query);
            }
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();

            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@col1", اسم_الزوج.Text);
            sqlCommand.Parameters.AddWithValue("@col2", جواز_الزوج.Text);
            sqlCommand.Parameters.AddWithValue("@col3", تاريخ_الميلاد.Text);
            sqlCommand.Parameters.AddWithValue("@col4", المهنة.Text);

            var reader = sqlCommand.ExecuteReader();
            if (reader.Read())
            {
                MessageBox.Show(reader["lastid"].ToString());
            }
            try
            {


            }
            catch (Exception ex)
            {
                MessageBox.Show("addNewAppNameInfo");
            }
        }

        private void addNewAppNameInfo2(TextBox textName)
        {

            string query = "insert into TableGenNames ([الاسم], رقم_الهوية,تاريخ_الميلاد) values (@col1,@col2,@col3) ;SELECT @@IDENTITY as lastid";
            string id = checkExist(textName.Text);
            if (id != "0")
            {
                query = "update TableGenNames set [الاسم] =  @col1,[رقم_الهوية] = @col2,[تاريخ_الميلاد] = @col3 where ID = " + id;
                //MessageBox.Show(query);
            }
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();

            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@col1", اسم_الزوجة.Text);
            sqlCommand.Parameters.AddWithValue("@col2", جواز_الزوجة.Text);
            sqlCommand.Parameters.AddWithValue("@col3", ميلاد_الزوجة.Text);

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

        private void addNewAppNameInfo3(TextBox textName, TextBox textDoc)
        {

            string query = "insert into TableGenNames ([الاسم], رقم_الهوية) values (@col1,@col2) ;SELECT @@IDENTITY as lastid";
            string id = checkExist(textName.Text);
            if (id != "0")
            {
                query = "update TableGenNames set [الاسم] =  @col1,[رقم_الهوية] = @col2 where ID = " + id;
                //MessageBox.Show(query);
            }
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();

            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@col1", textName.Text);
            sqlCommand.Parameters.AddWithValue("@col2", textDoc.Text);

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
        private void fillPrintDocx( )
        {
            string pdfFile = localCopy_off.Replace("docx", "pdf");
            oBDoc.SaveAs2(localCopy_off);
            oBDoc.ExportAsFixedFormat(pdfFile, Word.WdExportFormat.wdExportFormatPDF);
            oBDoc.Close(false, oBMiss);
            oBMicroWord.Quit(false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBMicroWord);
            System.Diagnostics.Process.Start(pdfFile);
            File.Delete(localCopy_off);            
            object doNotSaveChanges = Word.WdSaveOptions.wdSaveChanges;

        }

        private void chooseDocxFile(string docId, string name)
        {
            string RouteFile = FilespathIn + "TableMerrageDoc.docx";
            if (name != "")
                localCopy_off = FilespathOut + name + DateTime.Now.ToString("ddmmss") + ".docx";
            else localCopy_off = FilespathOut + docId.Replace("/", "_") + DateTime.Now.ToString("ddmmss") + ".docx";
            while (File.Exists(localCopy_off))
            {
                if (name != "")
                    localCopy_off = FilespathOut + name + DateTime.Now.ToString("ddmmss") + ".docx";
                else localCopy_off = FilespathOut + docId.Replace("/", "_") + DateTime.Now.ToString("ddmmss") + ".docx";
            }
            //
            System.IO.File.Copy(RouteFile, localCopy_off);
            FileInfo fileInfo = new FileInfo(localCopy_off);
            if (fileInfo.IsReadOnly) fileInfo.IsReadOnly = false;

            //MessageBox.Show(localCopy_off );
        }

        private void updateGenName(string name, string idDoc)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource.Replace("AhwalDataBase", "ArchFilesDB"));
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    string query = "update TableGeneralArch set رقم_معاملة_القسم=N'" + name + "' where رقم_المرجع = '" + idDoc + "' and docTable=N'TableMerrageDoc'";
                    SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
                    if (sqlCon.State == ConnectionState.Closed)
                        sqlCon.Open();
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.ExecuteNonQuery();
                    sqlCon.Close();
                }
                catch (Exception ex) { }
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
            string[] phoneNo = new string[10] { "", "", "", "", "", "", "", "", "", "" };
            int i = 0;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["الصفة"].ToString().Contains("قسم الأحوال الشخصية"))
                {
                    string smsText = "تم إنهاء معاملة قسيمة زواج بالرقم  " + رقم_المعاملة.Text.Split('/')[4] + " للمواطن/ " + اسم_الزوج.Text + " بتاريخ:" + GregorianDate;
                    SendSms(dataRow["MandoubPhones"].ToString(), smsText);                    
                    UpdateState(id, "sms", "sent", table);
                }
            }

        }
        private bool checkSentSMS(int id, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select sms from " + table+" where ID ='" + id.ToString() +"'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string[] phoneNo = new string[10] { "", "", "", "", "", "", "", "", "", "" };
            int i = 0;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["sms"].ToString() == "sent")
                {
                    return true;
                }
            }
            return false;
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
        private void MerriageDoc_Load(object sender, EventArgs e)
        {
            autoCompleteTextBox1(اسم_الزوج, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox1(اسم_الزوجة, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox1(الشاهد_الاول, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox1(الشاهد_الثاني, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox1(وكيل_الزوج, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox1(وكيل_الزوجة, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox(المهنة, DataSource, "jobs", "TableListCombo"); autoCompleteTextBox(المهنة, DataSource, "jobs", "TableListCombo");
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
        private string commentInfo()
        {
            string comment = "";
            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text == "")
                comment = "قام  " + الموظف.Text + " بإدخال البيانات " + Environment.NewLine + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine;

            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text != "")
                comment = "قام  " + الموظف.Text + " ببعض التعديلات " + Environment.NewLine + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine + التعليقات_السابقة_Off.Text;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text == "")
                comment = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + "قام  " + الموظف.Text + " ببعض التعديلات " + Environment.NewLine + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text != "")
                comment = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + "قام  " + الموظف.Text + " ببعض التعديلات " + Environment.NewLine + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine + "*" + التعليقات_السابقة_Off.Text.Trim();

            return comment;
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

        private void btnFile1_Click(object sender, EventArgs e)
        {
            FillDatafromGenArch("data1", genIDNo.ToString(), "TableMerrageDoc");
        }

        private void btnFile2_Click(object sender, EventArgs e)
        {
            FillDatafromGenArch("data2", genIDNo.ToString(), "TableMerrageDoc");
        }

        private void حالة_الزوجة_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (حالة_الزوجة.SelectedIndex == 1)
                MessageBox.Show("يجب التأكد من وجود وثيقة طلاق موثقة من الجهات المعنية");
            else if (حالة_الزوجة.SelectedIndex == 2)
                MessageBox.Show("يجب التأكد من وجود شهادة وفاة للمتوفى موثقة من الجهات المعنية");
        }

        private void صلة_الوكيل_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(صلة_الوكيل.SelectedIndex != 0) 
                MessageBox.Show("يجب التأكد من وجود شهادة وفاة للاب وإقرار من ولي الزوجة بأهلية الولاية");
        }

        private void deleteRow_Click(object sender, EventArgs e)
        {
            var selectedOption = MessageBox.Show("", "سيتم حذف بيانات الوثيقة وجميع الملفات المتعلقة بها؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                deleteRowsData(رقم_المعاملة.Text, "TableMerrageDoc");
                this.Close();
                //deleteRowsData(رقم_الوثيقة.Text, "TableGeneralArch");
                //deleteRowsData(رقم_الوثيقة.Text, "archives");
            }
        }

        private void deleteRowsData(string v1, string table)
        {
            string query;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "DELETE FROM " + table + " where رقم_المعاملة = @رقم_المعاملة";
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_المعاملة", v1);
            sqlCmd.ExecuteNonQuery();
            Con.Close();
        }

        private void btnListView_Click(object sender, EventArgs e)
        {
            if (PanelMain.Visible)
            {
                labDescribed.Visible = dataGridView1.Visible = true;
                PanelMain.Visible = false;
                dataGridView1.BringToFront();
            }
            else
            {
                labDescribed.Visible = dataGridView1.Visible = false;
                PanelMain.Visible = true;
                dataGridView1.SendToBack();
            }
        }

        
        string lastInput2 = "";
        bool عمر_الزوج_الحرج = true;
        private void تاريخ_الميلاد_TextChanged(object sender, EventArgs e)
        {
            if (تاريخ_الميلاد.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(تاريخ_الميلاد.Text, 1, 2));
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //VitxtDate1.Text = "";
                    تاريخ_الميلاد.Text = SpecificDigit(تاريخ_الميلاد.Text, 3, 10);
                    return;
                }
                عمر_الزوج_الحرج = getDate( تاريخ_الميلاد);
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
        private void MerriageDoc_FormClosed(object sender, FormClosedEventArgs e)
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

        private void مؤخر_الصداق_كتابة_TextChanged(object sender, EventArgs e)
        {

        }
        string lastInput1 = "";
        bool عمر_الزوجة_الحرج = true;
        private void ميلاد_الزوجة_TextChanged(object sender, EventArgs e)
        {
            if (ميلاد_الزوجة.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(ميلاد_الزوجة.Text, 1, 2));
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //VitxtDate1.Text = "";
                    ميلاد_الزوجة.Text = SpecificDigit(ميلاد_الزوجة.Text, 3, 10);
                    return;
                }
                عمر_الزوجة_الحرج = getDate(ميلاد_الزوجة);
            }

            if (ميلاد_الزوجة.Text.Length == 11)
            {
                ميلاد_الزوجة.Text = lastInput1; return;
            }
            if (ميلاد_الزوجة.Text.Length == 10) return;
            if (ميلاد_الزوجة.Text.Length == 4) ميلاد_الزوجة.Text = "-" + ميلاد_الزوجة.Text;
            else if (ميلاد_الزوجة.Text.Length == 7) ميلاد_الزوجة.Text = "-" + ميلاد_الزوجة.Text;
            lastInput1 = ميلاد_الزوجة.Text;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(عمر_الزوج_الحرج || عمر_الزوجة_الحرج)
                button1Enabel = false;
            else
                button1Enabel = true;
            ColorFulGrid9();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

            string docType = "";
            if (button1.InvokeRequired)
            {
                button1.Invoke(new MethodInvoker(delegate { button1.Enabled = false; }));
            }
            if (اسم_الزوج.InvokeRequired)
            {
                اسم_الزوج.Invoke(new MethodInvoker(delegate { docType = اسم_الزوج.Text; }));
            }
            chooseDocxFile(رقم_المعاملة.Text, docType); 
            prepareDocxfile();
            if (button1.InvokeRequired)
            {
                button1.Invoke(new MethodInvoker(delegate { button1.Enabled = true; }));
            }
        }
        private void prepareDocxfile()
        {

            oBMiss = System.Reflection.Missing.Value;
            oBMicroWord = new Word.Application();

            object objCurrentCopy = localCopy_off;

            oBDoc = oBMicroWord.Documents.Open(objCurrentCopy, oBMiss);
            oBMicroWord.Selection.Find.ClearFormatting();
            oBMicroWord.Selection.Find.Replacement.ClearFormatting();

        }
        private void fillPreDoc()
        {
            اسم_الزوج1_off.Text =اسم_الزوج2_off.Text =اسم_الزوج3_off.Text = اسم_الزوج.Text;
            اسم_الزوجة1_off.Text = اسم_الزوجة2_off.Text = اسم_الزوجة.Text;
            وكيل_الزوجة1_off.Text = وكيل_الزوجة.Text;
                
            if (حالة_الزوجة.SelectedIndex == 0) {
                صفة_الزوجة2_off.Text =  صفة_الزوجة1_off.Text = "البكر البالغ الرشيد";
            }
            else صفة_الزوجة2_off.Text = صفة_الزوجة1_off.Text = "البالغ الرشيد";

            if (طريقة_الإجراء.SelectedIndex == 0)
            {
                ضمير_الزوج1_off.Text = "ك";
                ضمير_الزوج2_off.Text = "";
                اسم_الزوج3_off.Text = اسم_الزوج1_off.Text = "";
                وكيل_الزوج1_off.Text = اسم_الزوج.Text;
            }
            else {
                ضمير_الزوج1_off.Text = "";
                ضمير_الزوج2_off.Text = " إلى موكلك";
                ضمير_الزوج3_off.Text = " إلى موكلي";
                وكيل_الزوج1_off.Text = وكيل_الزوج.Text;
            }
        }
            private void fillDocFileAppInfo()
        {
            foreach (Control control in PanelMain.Controls)
            {
                if (control is TextBox || control is ComboBox)
                {
                    try
                    {
                        //if (control.Name == "التوقيع") 
                        //    MessageBox.Show(panel.Name + control.Text);
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

        private void ListSearch_TextChanged(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '" + ListSearch.Text + "%'";
            dataGridView1.DataSource = bs; 
            ColorFulGrid9();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            ColorFulGrid9();
        }

        private void اسم_الزوج_TextChanged(object sender, EventArgs e)
        {
            getID(جواز_الزوج, تاريخ_الميلاد, المهنة, اسم_الزوج.Text);
        }
        bool gridFill = true;
        public void getID(TextBox رقم_الهوية_1,  TextBox تاريخ_الميلاد_1, TextBox المهنة_1, string name)
        {
            if (gridFill) return;
            DataTable dtbl = new DataTable();
            string query = "SELECT * FROM TableGenNames where الاسم = N'" + name + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            
           
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            
            sqlDa.Fill(dtbl);
            }

            catch (Exception ex)
            {
                رقم_الهوية_1.Text = "P0";
                المهنة_1.Text = "";
                تاريخ_الميلاد_1.Text = "";
                return;
            }
            رقم_الهوية_1.Text = "P0";
            المهنة_1.Text = "";
            تاريخ_الميلاد_1.Text = "";
            foreach (DataRow row in dtbl.Rows)
            {
                رقم_الهوية_1.Text = row["رقم_الهوية"].ToString();
                المهنة_1.Text = row["المهنة"].ToString();
                تاريخ_الميلاد_1.Text = row["تاريخ_الميلاد"].ToString();
                return;
            }
            //MessageBox.Show(رقم_الهوية_1.Text);
        }
        public void getID(TextBox رقم_الهوية_1,  TextBox تاريخ_الميلاد_1, string name)
        {
            if (gridFill) return;
            DataTable dtbl = new DataTable();
            string query = "SELECT * FROM TableGenNames where الاسم = N'" + name + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            
            sqlDa.Fill(dtbl);
            }
            catch (Exception ex)
            {
                رقم_الهوية_1.Text = "P0";
                تاريخ_الميلاد_1.Text = "";
                return;
            }
            رقم_الهوية_1.Text = "P0";
            تاريخ_الميلاد_1.Text = "";
            foreach (DataRow row in dtbl.Rows)
            {
                رقم_الهوية_1.Text = row["رقم_الهوية"].ToString();
                تاريخ_الميلاد_1.Text = row["تاريخ_الميلاد"].ToString();
                return;
            }
            //MessageBox.Show(رقم_الهوية_1.Text);
        }
        public void getID(TextBox رقم_الهوية_1 , string name)
        {
            if (gridFill) return;
            DataTable dtbl = new DataTable();
            string query = "SELECT * FROM TableGenNames where الاسم = N'" + name + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            
            sqlDa.Fill(dtbl);
            }
            catch (Exception ex)
            {
                رقم_الهوية_1.Text = "P0";
                return;
            }
            رقم_الهوية_1.Text = "P0";
            foreach (DataRow row in dtbl.Rows)
            {
                رقم_الهوية_1.Text = row["رقم_الهوية"].ToString();
                return;
            }
            //MessageBox.Show(رقم_الهوية_1.Text);
        }

        private void اسم_الزوجة_TextChanged(object sender, EventArgs e)
        {
            getID(جواز_الزوجة, ميلاد_الزوجة, اسم_الزوجة.Text);
        }

        private void وكيل_الزوجة_TextChanged(object sender, EventArgs e)
        {
            //getID(جواز_وكيل_الزوجة, وكيل_الزوجة.Text, "رقم_الهوية", "P0");
        }

        private void الشاهد_الاول_TextChanged(object sender, EventArgs e)
        {
            //getID(وثيقة_الشاهد_الاول, الشاهد_الاول.Text, "رقم_الهوية", "P0");
        }

        private void الشاهد_الثاني_TextChanged(object sender, EventArgs e)
        {
            //getID(وثيقة_الشاهد_الثاني, الشاهد_الثاني.Text, "رقم_الهوية", "P0");
        }

        private void وكيل_الزوج_TextChanged(object sender, EventArgs e)
        {
            //getID(جواز_وكيل_الزوج, وكيل_الزوج.Text, "رقم_الهوية", "P0");             
        }
        public void getID(TextBox textTo, string name, string controlType, string def)
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
            int index = 0;
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

        string lastInput3 = "";
        private void تاريخ_الايصال_TextChanged(object sender, EventArgs e)
        {
            if (تاريخ_الايصال.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(تاريخ_الايصال.Text, 1, 2));
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //VitxtDate1.Text = "";
                    تاريخ_الايصال.Text = SpecificDigit(تاريخ_الايصال.Text, 3, 10);
                    return;
                }
            }

            if (تاريخ_الايصال.Text.Length == 11)
            {
                تاريخ_الايصال.Text = lastInput3; return;
            }
            if (تاريخ_الايصال.Text.Length == 10) return;
            if (تاريخ_الايصال.Text.Length == 4) تاريخ_الايصال.Text = "-" + تاريخ_الايصال.Text;
            else if (تاريخ_الايصال.Text.Length == 7) تاريخ_الايصال.Text = "-" + تاريخ_الايصال.Text;
            lastInput3 = تاريخ_الايصال.Text;
        }
    }
}
