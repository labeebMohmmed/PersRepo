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
    
    public partial class FormDivorce : Form
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

        public FormDivorce(string dataSource, bool addEdit, string empName, int atVCIndex, string gregorianDate, string hijriDate)
        {
            InitializeComponent();
            definColumn(dataSource);
            DataSource = dataSource;
            AddEdit = addEdit;
            AtVCIndex = atVCIndex;
            allList = getColList("TableDivorce");

            التاريخ_الهجري.Text = HijriDate = hijriDate;
            التاريخ_الميلادي.Text = GregorianDate = gregorianDate;
            fillFileBox(DataSource);
            if (AddEdit)
            {
                dataGridView1.Visible = false;
                PanelMain.Visible = true;
            }
            else
            {
                dataGridView1.Visible = true;
                PanelMain.Visible = false;
            }

            colIDs[4] = موظف_الإدخال.Text = الموظف.Text = empName;
            موظف_الإدخال.Text = الموظف.Text = empName;
            fileComboBox(المأذون, DataSource, "ArabicAttendVC", "TableListCombo");
            if (المأذون.Items.Count > AtVCIndex)
                المأذون.SelectedIndex = AtVCIndex;
            else المأذون.SelectedIndex = 0;
            طريقة_الطلب.SelectedIndex = 0;
            
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
            SqlCommand sqlCmd = new SqlCommand("alter table TableDivorce add " + Columnname.Replace(" ", "_") + " nvarchar(150)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private bool checkColumnName(string colNo, string dataSource)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableDivorce", sqlCon);
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
            insertAll = "INSERT INTO " + table + "(" + insertItems + ") values (" + insertValues + ")";

            return allList;

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
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableDivorce", sqlCon);
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
                        if (!control.Name.Contains("التاريخ") && !control.Name.Contains("موظف"))
                            control.Text = dataGridView1.CurrentRow.Cells[control.Name].Value.ToString();
                    }

                }
                if (dataGridView1.CurrentRow.Cells["اسم_الزوج"].Value.ToString() == "")
                {
                    newData = true;
                    FillDatafromGenArch("data1", genIDNo.ToString(), "TableDivorce");
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

                        if (control.Visible && (control.Name.Contains("ميلاد_") && control.Text.Length != 10))
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

        private void button1_Click(object sender, EventArgs e)
        {
            string part4 = رقم_المعاملة.Text.Split('/')[4];
            string part1to3 = رقم_المعاملة.Text.Split('/')[0] + "/" + رقم_المعاملة.Text.Split('/')[1] + "/" + رقم_المعاملة.Text.Split('/')[2] + "/" + رقم_المعاملة.Text.Split('/')[3] + "/";
            //MessageBox.Show(part4);
            //MessageBox.Show(part1to3);
            //MessageBox.Show(part1to3+ رقم_الوثيقة.Text);
            if(رقم_الوثيقة.Text!=""&&رقم_الوثيقة.Text!="بدون")
                رقم_المعاملة.Text = part1to3 + رقم_الوثيقة.Text;
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
            if (newData)
            {
                colIDs[0] = رقم_المعاملة.Text;
                colIDs[1] = genIDNo.ToString();
                colIDs[2] = GregorianDate;
                colIDs[3] = اسم_الزوج.Text;
                colIDs[4] = الموظف.Text;
                colIDs[5] = "";
                colIDs[6] = "";
                colIDs[7] = "new";
                addarchives(colIDs);
            }
            this.Close();
        }
        private void updateGenName(string name, string idDoc)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string query = "update TableGeneralArch set رقم_معاملة_القسم=N'" + name + "' where رقم_المرجع = '" + idDoc + "' and docTable=N'TableDivorce'";
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
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

        private void button6_Click(object sender, EventArgs e)
        {
            FillDatafromGenArch("data1", genIDNo.ToString(), "TableDivorce");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FillDatafromGenArch("data2", genIDNo.ToString(), "TableDivorce");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var selectedOption = MessageBox.Show("", "سيتم حذف بيانات الوثيقة وجميع الملفات المتعلقة بها؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                deleteRowsData(رقم_الوثيقة.Text, "TableDivorce");                
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

        private void button5_Click(object sender, EventArgs e)
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
            
        }

        private void FormDivorce_FormClosed(object sender, FormClosedEventArgs e)
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

        private void طريقة_الإجراء_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (طريقة_الإجراء.SelectedIndex == 1)
            {
                labhusSideName.Visible = وكيل_الزوج.Visible = labhusSidePass.Visible = جواز_وكيل_الزوج.Visible = labhusSideIqama.Visible = إقامة_وكيل_الزوج.Visible = true;
            }
            else
                labhusSideName.Visible = وكيل_الزوج.Visible = labhusSidePass.Visible = جواز_وكيل_الزوج.Visible = labhusSideIqama.Visible = إقامة_وكيل_الزوج.Visible = false;

        }
        string lastInput2 = "";
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
            }

            if (ميلاد_الزوجة.Text.Length == 11)
            {
                ميلاد_الزوجة.Text = lastInput2; return;
            }
            if (ميلاد_الزوجة.Text.Length == 10) return;
            if (ميلاد_الزوجة.Text.Length == 4) ميلاد_الزوجة.Text = "-" + ميلاد_الزوجة.Text;
            else if (ميلاد_الزوجة.Text.Length == 7) ميلاد_الزوجة.Text = "-" + ميلاد_الزوجة.Text;
            lastInput2 = ميلاد_الزوجة.Text;
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
        string lastInput1 = "";
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
            }

            if (تاريخ_الميلاد.Text.Length == 11)
            {
                تاريخ_الميلاد.Text = lastInput1; return;
            }
            if (تاريخ_الميلاد.Text.Length == 10) return;
            if (تاريخ_الميلاد.Text.Length == 4) تاريخ_الميلاد.Text = "-" + تاريخ_الميلاد.Text;
            else if (تاريخ_الميلاد.Text.Length == 7) تاريخ_الميلاد.Text = "-" + تاريخ_الميلاد.Text;
            lastInput1 = تاريخ_الميلاد.Text;
        }

        private void رقم_الوثيقة_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(رقم_الوثيقة.Text);

        }
    }
}
