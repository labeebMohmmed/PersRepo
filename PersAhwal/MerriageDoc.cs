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
        public MerriageDoc(string dataSource, bool addEdit, string empName, int atVCIndex, string gregorianDate, string hijriDate)
        {
            InitializeComponent();
            DataSource = dataSource;
            AddEdit = addEdit;
            AtVCIndex = atVCIndex;
            allList = getColList("TableMerrageDoc");

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
            طريقة_الطلب.SelectedIndex = 0;
            اسم_المندوب.Text = "";
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
                
                if (row["name"].ToString() != "ID" && row["name"].ToString() != "حالة_الارشفة"&& row["name"].ToString() != "sms")
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

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (طريقة_الإجراء.SelectedIndex == 1) {
                labhusSideName.Visible = وكيل_الزوج.Visible = labhusSidePass.Visible = جواز_وكيل_الزوج.Visible = labhusSideIqama.Visible = إقامة_وكيل_الزوج.Visible = label22.Visible = هاتف_وكيل_الزوج.Visible = true;
            }
            else 
                labhusSideName.Visible = وكيل_الزوج.Visible = labhusSidePass.Visible = جواز_وكيل_الزوج.Visible = labhusSideIqama.Visible = إقامة_وكيل_الزوج.Visible = label22.Visible = هاتف_وكيل_الزوج.Visible = false;
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
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableMerrageDoc", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtblMain = new DataTable();
            sqlDa.Fill(dtblMain);
            dataGridView1.DataSource = dtblMain;
            sqlCon.Close();


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
                        if(!control.Name.Contains("التاريخ") && !control.Name.Contains("موظف"))
                        control.Text = dataGridView1.CurrentRow.Cells[control.Name].Value.ToString();
                    }
                    
                }
                if (dataGridView1.CurrentRow.Cells["اسم_الزوج"].Value.ToString() == "")
                {
                    newData = true;
                    FillDatafromGenArch("data1", genIDNo.ToString(), "TableMerrageDoc");
                }
                    AddEdit = false;
                dataGridView1.Visible = false;
                PanelMain.Visible = true;
            }
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

                        if (control.Visible && (control.Name.Contains("ميلاد_") && control.Text.Length != 4))
                        {
                            MessageBox.Show("يرجى إضافة عام الميلاد لخانة " + control.Name); return false;
                        }

                        if (control.Visible && (control.Name.Contains("هاتف") && control.Text.Length != 12))
                        {
                            MessageBox.Show("يرجى إضافة رقم الهاتف بخانة " + control.Name); return false;
                        }
                        if (control.Visible && (control.Name.Contains("قامة") && control.Text.Length != 10))
                        {
                            MessageBox.Show("يرجى إضافة رقم الإقامة بصورة صحيحة لخانة " + control.Name); return false;
                        }
                        if (control.Visible && (control.Name.Contains("جواز") && control.Text.Length != 9))
                        {
                            MessageBox.Show("يرجى إضافة رقم الجواز بصورة صحيحة لخانة " + control.Name); return false;
                        }
                    }
                }
            }
            return true;
        }

        void FillDatafromGenArch(string doc, string id, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
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
            
            if(!checkSentSMS(genIDNo, "TableMerrageDoc")) 
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
            this.Close();
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
                    string smsText = "تم إنهاء معاملة قسيمة زواج بالرقم  " + رقم_الوثيقة.Text + " للمواطن/ " + اسم_الزوج.Text + " بتاريخ:" + GregorianDate;
                    SendSms(dataRow["MandoubPhones"].ToString(), smsText);
                    SendSms(هاتف_الزوج.Text, smsText);
                    UpdateState(id, "sms", "sent", table);
                }
            }

        }
        private bool checkSentSMS(int id, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select sms,الصفة from " + table+" where ID ='" + id.ToString() +"'", sqlCon);
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
                deleteRowsData(رقم_الوثيقة.Text, "TableMerrageDoc");
                //deleteRowsData(رقم_الوثيقة.Text, "TableGeneralArch");
                //deleteRowsData(رقم_الوثيقة.Text, "archives");
            }
        }

        private void deleteRowsData(string v1, string table)
        {
            string query;
            SqlConnection Con = new SqlConnection(DataSource);
            query = "DELETE FROM " + table + " where رقم_الوثيقة = @رقم_الوثيقة";
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_الوثيقة", v1);
            sqlCmd.ExecuteNonQuery();
            Con.Close();
        }

        private void btnListView_Click(object sender, EventArgs e)
        {
            if (PanelMain.Visible)
            {
                dataGridView1.Visible = true;
                PanelMain.Visible = false;
                dataGridView1.BringToFront();
            }
            else
            {
                dataGridView1.Visible = false;
                PanelMain.Visible = true;
                dataGridView1.SendToBack();
            }
        }

        private void طريقة_الطلب_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (طريقة_الطلب.SelectedIndex != 0)
                اسم_المندوب.Visible = true;
            else
            {
                اسم_المندوب.Text = "";
                اسم_المندوب.Visible = false;
            }
        }
    }
}
