using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Data.SqlTypes;
using DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;
using static Azure.Core.HttpHeader;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
namespace PersAhwal
{
    public partial class Settings : Form
    {
        private string DataSource56, DataSource57, FilepathIn, FilepathOut, ArchFile, FormDataFile;
        private static bool NewSettings = false;
        string comboBoxOptions1 = "", comboBoxOptions2 = "";
        string[] txtComboOptions = new string[5] { "", "", "", "", "" };
        string[] DPTitle = new string[5];
        int pTextHieght = 42;
        int pComboHieght = 42;
        int pCheckHieght = 42;
        int pDateHieght = 42;
        int pbuttonHieght = 42;
        string ColumnName = "";
        bool NewColumn = false;
        string AuthBody1 = "لينوب عني ويقوم مقامي في ";
        int combo1index = 0, combo2index = 0, combo3index = 0, combo4index = 0, combo5index = 0;
        int id = 0;
        int idIndex = 1;
        bool review1 = false;
        string RightColumnName = "";
        int Nobox = 0;
        int listchecked = 0;
        DataTable checkboxdt;
        int LastID = 0, LastTabIndex = 0;
        static string[,] preffix = new string[10, 20];
        static string[] Text_statis = new string[5];
        static int[] statistic = new int[100];
        static int[] staticIndex = new int[100];
        static int[] times = new int[100];
        string[] allList = new string[100];
        static string[] Empty = new string[1] { "" };
        string Server = "M";
        string DataSource;
        int CombAuthTypeIndex = 1;
        string updateAll = "";
        string insertAll = "";
        string[] errors;
        string editRights;
        string errorList;
        string CurrentFile = "";
        bool readyToUpload = false;
        int ProcReqID = 0;
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;
        int rCnt;
        int cCnt;
        int rw = 0;
        int cl = 0;
        string revised = "";
        bool checkIndex = false;
        string formNo = "";
        string[] itemsicheck1 = new string[5];
        bool AuthType = true;
        bool reqGrid = false;
        
        public Settings(string server, bool newSettings, string dataSource56, string dataSource57, bool setDataBase, string filepathIn, string filepathOut, string archFile, string formDataFile, string colName)
        {
            InitializeComponent();
            Server = server;
            DataSource56 = dataSource56;
            DataSource57 = dataSource57;

            if (Server == "57")
                DataSource = DataSource57;
            else if (Server == "56")
                DataSource = DataSource56;
            allList = getColList("TableAddContext");

            NewSettings = newSettings;

            FilepathIn = filepathIn + @"\";
            FilepathOut = filepathOut;
            ArchFile = archFile;
            FormDataFile = formDataFile;
            if (!setDataBase)
            {
                if (!newSettings) loadSettings();
                else
                {
                    SaveSettings.Text = "أدخل بيانات قاعدة بيانات صحيحة";
                    MessageBox.Show("لا توجد قاعدة بيانات مسجلة");
                }
            }
            fillSubComboBox(mainTypeAuth, DataSource, "AuthTypes", "TableListCombo", false);
            //autoCompleteTextBox(newCombAuthType, DataSource, "AuthTypes", "TableListCombo");
            fillSubComboBox(mainTypeIqrar, DataSource, "ArabicGenIgrar", "TableListCombo", false);
            newFillComboBox2(subTypeIqrar, DataSource, mainTypeIqrar.SelectedIndex.ToString(), langIqrar.Text);
            AddRightColumn();
            if (colName != "")
            {
                //MessageBox.Show(colName);
                mainTypeAuth.SelectedIndex = Convert.ToInt32(colName.Split('-')[1]);
                subTypeAuth.SelectedIndex = Convert.ToInt32(colName.Split('-')[0]);
                //FillDataGridView("");
                //flllPanelItemsboxes("ColName", colName);
            }
            Suffex_preffixList();
            //for (int x = 0; x < mainTypeAuth.Items.Count; x++)
            //{
            //    sunInfo(mainTypeAuth.Items[x].ToString().Replace(" ", "_"), x);
            //}
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
            txtSearch.Select();
        }

        private void insertRow(string ColName)
        {
            //MessageBox.Show(data[1]);
            SqlConnection sqlCon = new SqlConnection(DataSource);


            string query = "INSERT INTO TableAddContext (ColName) values (@ColName)";

            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { }
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ColName", ColName);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }


        private bool PreReqFound(string proName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query = "SELECT رقم_المعاملة FROM TableProcReq where المعاملة=N'" + proName + "'";
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count > 0)
                return true;
            else return false;
        }

        private void sunInfo(string colName, int index)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query = "select [" + colName + "] from TableListCombo where [" + colName + "] is not null and [" + colName + "] <> ''";
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
                sqlCon.Close();
                foreach (DataRow row in dtbl.Rows)
                {
                    string formNo2 = row[colName].ToString().Trim() + "-" + index.ToString();

                    if (!TableAddContext(formNo2))
                    {
                        insertRow(formNo2);

                    }
                }
            }
            catch (Exception ex) { }
        }

        private bool TableAddContext(string proName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query = "SELECT ColName FROM TableAddContext where ColName=N'" + proName + "'";
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count > 0)
                return true;
            else return false;
        }

        private string[] getColList(string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource57);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return allList; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('" + table + "')", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            int i = 0;
            string insertItems = "";
            string insertValues = "";
            string updateValues = "";
            int colCount = 0;
            foreach (DataRow row in dtbl.Rows)
            {

                if (row["name"].ToString() != "ID" && row["name"].ToString() != "حالة_الارشفة" && row["name"].ToString() != "sms")
                {
                    colCount++;
                }
            }

            allList = new string[colCount];

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
        private bool checkColExist(string table, string colName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource57);
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

        private void button17_Click(object sender, EventArgs e)
        {

        }

        private void txtModel_TextChanged(object sender, EventArgs e)
        {

        }

        private void button20_Click(object sender, EventArgs e)
        {

        }

        private void txtOutput_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void ArchiveFile_TextChanged(object sender, EventArgs e)
        {

        }

        private void button21_Click(object sender, EventArgs e)
        {

        }

        private void txtServerIP_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void txtDatabase_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void txtLogin_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void txtPass_TextChanged(object sender, EventArgs e)
        {

        }

        private void CreateColumn(string Columnname, string tableName)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("alter table " + tableName + " add " + Columnname + " nvarchar(150)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            try
            {
                sqlCmd.ExecuteNonQuery();
            }catch (Exception ex) { return; }   
            sqlCon.Close();
        }



        private string getAppFolder()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return ""; }
            string settingData = "select FolderApp from TableSettings where ID='1'";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            string ver = "";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                ver = dataRow["FolderApp"].ToString();

            }
            return ver;
        }


        private void button3_Click(object sender, EventArgs e)
        {
            txtApp.Text = getAppFolder();
            if (SettingsPanel.Visible)
            {
                flowLayoutPanel9.Visible = SettingsPanel.Visible = false;

            }
            else
            {
                txtSearch.Visible = button32.Visible = txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = false;
                flowLayoutPanel9.Visible = SettingsPanel.Visible = true;
                ContextPanel.Visible = false;
            }
        }



        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (dateType1.CheckState == CheckState.Checked)
            {
                dateType1.Text = "ميلادي";

            }
            else
            {
                dateType1.Text = "هجري";

            }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            reqGrid = false;
            AuthType = true;
                if(dataGridView1.Visible)
            {
                txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = false;
                flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
                dataGridView1.SendToBack();
                panelLowButtons.Visible = ContextPanel.Visible = true;
                ContextPanel.BringToFront();
            }
            else
            {
                
                repReqPanel.Visible = false;
                dataGridView1.BringToFront();
                FillDataGridView("TableAddContext", AuthType);
                panelLowButtons.Visible = panelIqrar.Visible = panelAuthInfo.Visible = false;
                txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = true;
                flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
                panelLowButtons.Visible = ContextPanel.Visible = true;
                formsBtn.Visible = proFileBtn.Visible = btnRevised.Visible = true;
            }
        }


        void FillDataGridView(string table, bool auth)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }

            //SqlDataAdapter sqlDa = new SqlDataAdapter("select " + getColumnNames(table) + " from " + table + " order by ID desc", sqlCon);
            SqlDataAdapter sqlDa = new SqlDataAdapter("select " + getColumnNames(table) + " from " + table + " where ColRight <> '' order by ID desc", sqlCon);
            if (!auth)
                sqlDa = new SqlDataAdapter("select " + getColumnNames(table) + " from " + table + " where ColRight = '' order by ID desc", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            //dataGridView1.Columns["ID"].Visible = false;
            try
            {
                dataGridView1.Columns[1].Width = 100;
                dataGridView1.Columns[2].Width = 350;
                dataGridView1.Columns[3].Width = 200;
                dataGridView1.Columns[4].Width = 200;
                dataGridView1.Columns[6].Width = 200;
            }
            catch (Exception ex) { }
            sqlCon.Close();
        }
        
        void FillDataGridViewReq(string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }

            SqlDataAdapter sqlDa = new SqlDataAdapter("select " + getColumnNames(table) + " from " + table + " order by  المعاملة asc", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            //dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            //dataGridView1.Columns["ID"].Visible = false;
            try
            {
                dataGridView1.Columns[1].Width = 100;
                dataGridView1.Columns[2].Width = 350;
                dataGridView1.Columns[3].Width = 200;
                dataGridView1.Columns[4].Width = 200;
                dataGridView1.Columns[6].Width = 200;
            }
            catch (Exception ex) { }
            sqlCon.Close();
        }

        private string getColumnNames(string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return ""; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS " + table, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string cols = "";
            int index = 0;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(dataRow["COLUMN_NAME"].ToString()))
                {
                    if (!dataRow["COLUMN_NAME"].ToString().Contains("Data1"))
                    {
                        if (index == 0)
                            cols = dataRow["COLUMN_NAME"].ToString();
                        else cols = cols + "," + dataRow["COLUMN_NAME"].ToString();
                        index++;
                    }
                }
            }
            return cols;
        }

        
        private void Settings_Load(object sender, EventArgs e)
        {

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

        private int getLastID(string source, string comlumnName, string tableName)
        {
            int x = 1;
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select max(ID) as maxID from " + tableName + " where " + comlumnName + " is not null";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;

                //MessageBox.Show(query);
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                foreach (DataRow dataRow in table.Rows)
                {
                    try
                    {
                        return Convert.ToInt32(dataRow["maxID"].ToString()) + 1;
                    }
                    catch (Exception) {
                        return 1;
                    }
                }
                saConn.Close();
            }
            return 1;
        }


        private int getCurrentID(string source, string comlumnName, string tableName, string text)
        {
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select ID from "+ tableName+"  where " + comlumnName+" = N'" + text+"'";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;

                //MessageBox.Show(query);
                cmd.ExecuteNonQuery();
                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);
                
                foreach (DataRow dataRow in table.Rows)
                {
                    try
                    {
                        return Convert.ToInt32(dataRow["ID"].ToString());
                    }
                    catch (Exception) {
                        return 0;
                    }
                }
                saConn.Close();
            }
            return 0;
        }

        private void fillSubComboBox(ComboBox combbox, string source, string comlumnName, string tableName, bool select)
        {
            //MessageBox.Show("source += "+source);
            combbox.Visible = true;
            //MessageBox.Show(source);
            //MessageBox.Show(Server);
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
        
        private bool checkRequInfo(string proID)
        {
            
            using (SqlConnection saConn = new SqlConnection(DataSource))
            {
                saConn.Open();

                string query = "select ID from TableProcReq where المعاملة =N'" + proID + "'";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();

                DataTable table = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(table);

                saConn.Close();
                if (table.Rows.Count > 0)
                {
                    Console.WriteLine("المعاملة موجودة " + proID);
                    return true;
                }
                else return false;
            }
            
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

                    if (dataRow["Lang"].ToString() == Language && dataRow["ColRight"].ToString() != "" && dataRow["ColName"].ToString().Contains("-"))
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
            if (combbox.Items.Count > 0) combbox.SelectedIndex = 0;
        }

        private void newFillComboBox2(ComboBox combbox, string source, string id, string Language)
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

                    if (dataRow["Lang"].ToString() == Language && dataRow["ColRight"].ToString() == "" && !String.IsNullOrEmpty(dataRow["ColName"].ToString()) && dataRow["ColName"].ToString().Contains("-"))
                    {

                        if (dataRow["ColName"].ToString().Split('-')[1].All(char.IsDigit))
                        {
                            try
                            {
                                if (id == dataRow["ColName"].ToString().Split('-')[1])
                                {
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
            if (combbox.Items.Count > 0) combbox.SelectedIndex = 0;
        }

        private bool ComboTest(ComboBox combbox, string source, string comlumnName, string tableName)
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

                    if (dataRow[comlumnName].ToString() == combbox.Text)
                    {
                        return true;
                    }
                }
                saConn.Close();
            }
            return false;
        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        
        private void restShowingItems()
        {
            foreach (Control control in PanelItemsboxes.Controls)
            {
                if (control is TextBox)
                {
                    ((TextBox)control).Text = "";
                    ((TextBox)control).Visible = false;
                    ((TextBox)control).Size = new System.Drawing.Size(200, 35);
                }
                if (control is Label)
                {
                    ((Label)control).Text = "";
                    ((Label)control).Visible = false;
                }
                if (control is ComboBox)
                {
                    ((ComboBox)control).Text = "";
                    ((ComboBox)control).Visible = false;
                }
                if (control is CheckBox)
                {
                    ((CheckBox)control).Text = "";
                    ((CheckBox)control).CheckState = CheckState.Unchecked;
                    ((CheckBox)control).Visible = false;
                }

            }


        }


        private void newCombAuthType_TextChanged(object sender, EventArgs e)
        {

        }

        private void CombAuthType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ComboProcedure_SelectedIndexChanged(object sender, EventArgs e)
        {


        }

        private void button37_Click(object sender, EventArgs e)
        {

        }

        private void addMainAuth(string colText)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableListCombo (AuthTypes) values (@AuthTypes)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@AuthTypes", colText);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private bool checkColumnName(string colNo)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return false; }
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

        private void getColumnName(string colNo)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableListCombo", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(dataRow["COLUMN_NAME"].ToString()))
                {
                    if (dataRow["COLUMN_NAME"].ToString().Contains(colNo))
                    {
                        ColumnName = dataRow["COLUMN_NAME"].ToString();
                        return;
                    }
                    else ColumnName = "";
                }
            }
        }

        private void AddRightColumn()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableAuthRights", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(dataRow["COLUMN_NAME"].ToString()))
                {
                    ColRight.Items.Add(dataRow["COLUMN_NAME"].ToString());
                }
            }
        }


        private void addSubAuth(int id, string colText, string ColName)
        {

            string str = "@" + ColName;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("update TableListCombo set " + ColName + "=" + str + " where ID=@ID", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue(str, colText);
            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void FolderAppUpdate(string folderApp)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("update TableSettings set FolderApp=@FolderApp where ID='1'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@FolderApp", folderApp);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }



        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows.Count > 1 && repReqPanel.Visible)
            {
                المعاملة.Text = dataGridView1.CurrentRow.Cells["المعاملة"].Value.ToString();
                OpenFile(المعاملة.Text, true, reviewForms);
                //MessageBox.Show(CurrentFile);
                string[] colList = new string[11];
                colList[0] = "المعاملة";
                colList[1] = "رقم_المعاملة";
                colList[2] = "المطلوب_رقم1";
                colList[3] = "المطلوب_رقم2";
                colList[4] = "المطلوب_رقم3";
                colList[5] = "المطلوب_رقم4";
                colList[6] = "المطلوب_رقم5";
                colList[7] = "المطلوب_رقم6";
                colList[8] = "المطلوب_رقم7";
                colList[9] = "المطلوب_رقم8";
                colList[10] = "المطلوب_رقم9";
                SqlConnection sqlCon = new SqlConnection(DataSource);
                try
                {
                    if (sqlCon.State == ConnectionState.Closed)
                        sqlCon.Open();
                }
                catch (Exception ex) { return; }
                SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * FROM TableProcReq where المعاملة=N'" + المعاملة.Text + "'", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);
                sqlCon.Close();
                if (dtbl.Rows.Count > 0)
                {
                    foreach (DataRow row in dtbl.Rows)
                    {
                        ProcReqID = Convert.ToInt32(row["ID"].ToString());
                        for (int index = 1; index < 11; index++)
                        {
                            foreach (Control control in repReqPanel.Controls)
                            {
                                if (control.Name == colList[index])
                                {
                                    control.Text = row[colList[index]].ToString();
                                }
                            }
                        }
                    }
                }
                txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = false;
                repReqPanel.Visible = true;
                repReqPanel.BringToFront();
            }
            else if (dataGridView1.Rows.Count > 1 && ContextPanel.Visible)
            {
                string colname = txtProName.Text = dataGridView1.CurrentRow.Cells["ColName"].Value.ToString();
                revised = dataGridView1.CurrentRow.Cells["revised"].Value.ToString();
                labID.Text = "rowID:" + dataGridView1.CurrentRow.Cells["ID"].Value.ToString();
                idIndex = Convert.ToInt32(dataGridView1.CurrentRow.Cells["ID"].Value.ToString());

                langAuth.Text = dataGridView1.CurrentRow.Cells["Lang"].Value.ToString();
                ColRight.Text = dataGridView1.CurrentRow.Cells["ColRight"].Value.ToString().Replace(" ", "_");
                TextModel.Text = dataGridView1.CurrentRow.Cells["TextModel"].Value.ToString();
                editRights = dataGridView1.CurrentRow.Cells["editRights"].Value.ToString();
                errorList = dataGridView1.CurrentRow.Cells["errorList"].Value.ToString();
                
                panelJob(dataGridView1.CurrentRow.Index);
                //ColRight.Text = dataGridView1.CurrentRow.Cells["ColName"].Value.ToString();
                //MessageBox.Show(dataGridView1.CurrentRow.Index.ToString());

                if (langAuth.Text == "" || langAuth.Text == "العربية")
                {
                    langAuth.CheckState = CheckState.Unchecked;
                    langAuth.Text = "العربية";
                    langIqrar.CheckState = CheckState.Unchecked;
                    langIqrar.Text = "العربية";
                }
                else
                {
                    langAuth.CheckState = CheckState.Checked;
                    langAuth.Text = "الانجليزية";
                    langIqrar.CheckState = CheckState.Checked;
                    langIqrar.Text = "الانجليزية";
                }

                if (ColRight.Text != "")
                {
                    panelAuthInfo.Visible = true;
                    panelIqrar.Visible = false;
                    try
                    {
                        mainTypeAuth.SelectedIndex = Convert.ToInt32(colname.Split('-')[1]);
                        subTypeAuth.Text = colname.Split('-')[0];
                        formNo = mainTypeAuth.Text + "-" + subTypeAuth.Text.Trim();

                    }
                    catch (Exception ex) { }
                }
                else
                {
                    panelIqrar.Visible = true;
                    panelAuthInfo.Visible = false;
                    try
                    {
                        mainTypeIqrar.SelectedIndex = Convert.ToInt32(colname.Split('-')[1]);
                        subTypeIqrar.Text = colname.Split('-')[0];
                        formNo = mainTypeAuth.Text + "-" + subTypeAuth.Text.Trim();
                    }
                    catch (Exception ex) { }
                }



                txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = false;
                ContextPanel.BringToFront();
                labID.BringToFront();
                if (checkRequInfo(formNo))
                    OpenFile(formNo, true, formsBtn);
                else
                {
                    repReqPanel.BringToFront();
                    repReqPanel.Visible = true;
                    المعاملة.Text = formNo;
                    revisedRow("");
                    MessageBox.Show("لا يوجد قائمة بالمطلوبات الأولية للمعاملة، يرجى إضافتها");
                }
                if (!formsBtn.Enabled)
                {
                    revisedRow("");
                    formsBtn.Enabled = true;
                    formsBtn.Text = "رفع الاستمارة الأولية";
                }
                panelLowButtons.Visible = ContextPanel.Visible = true;
                button117.Enabled = true;
                //review1 = true;
                review1 = false;
                btnDelete.Visible = btnClear.Visible = true;
                DPTitle[0] = icheckoption11.Text + "_" + icheckoption12.Text;
                dataGridView1.Visible = false;
                //MessageBox.Show(ColRight.Text);
                if(ColRight.Text != "") 
                    btnRightsShow.PerformClick();
            }
        }

        private void dataGridView1_RowIndex(string colName)
        {
            for (int indEX = 0; indEX < dataGridView1.Rows.Count - 1; indEX++)
                if (colName == dataGridView1.Rows[indEX].Cells["ColName"].Value.ToString())
                {
                    string colname = txtProName.Text = dataGridView1.Rows[indEX].Cells["ColName"].Value.ToString();
                    revised = dataGridView1.Rows[indEX].Cells["revised"].Value.ToString();
                    labID.Text = "rowID:" + dataGridView1.Rows[indEX].Cells["ID"].Value.ToString();
                    idIndex = Convert.ToInt32(dataGridView1.Rows[indEX].Cells["ID"].Value.ToString());

                    langAuth.Text = dataGridView1.Rows[indEX].Cells["Lang"].Value.ToString();
                    ColRight.Text = dataGridView1.Rows[indEX].Cells["ColRight"].Value.ToString().Replace(" ", "_");
                    TextModel.Text = dataGridView1.Rows[indEX].Cells["TextModel"].Value.ToString();
                    editRights = dataGridView1.Rows[indEX].Cells["editRights"].Value.ToString();
                    errorList = dataGridView1.Rows[indEX].Cells["errorList"].Value.ToString();

                    panelJob(dataGridView1.Rows[indEX].Index);
                    //ColRight.Text = dataGridView1.Rows[indEX].Cells["ColName"].Value.ToString();
                    //MessageBox.Show(dataGridView1.Rows[indEX].Index.ToString());

                    if (langAuth.Text == "" || langAuth.Text == "العربية")
                    {
                        langAuth.CheckState = CheckState.Unchecked;
                        langAuth.Text = "العربية";
                        langIqrar.CheckState = CheckState.Unchecked;
                        langIqrar.Text = "العربية";
                    }
                    else
                    {
                        langAuth.CheckState = CheckState.Checked;
                        langAuth.Text = "الانجليزية";
                        langIqrar.CheckState = CheckState.Checked;
                        langIqrar.Text = "الانجليزية";
                    }

                    if (ColRight.Text != "")
                    {
                        panelAuthInfo.Visible = true;
                        panelIqrar.Visible = false;
                        try
                        {
                            mainTypeAuth.SelectedIndex = Convert.ToInt32(colname.Split('-')[1]);
                            subTypeAuth.Text = colname.Split('-')[0];
                            formNo = mainTypeAuth.Text + "-" + subTypeAuth.Text.Trim();

                        }
                        catch (Exception ex) { }
                    }
                    else
                    {
                        panelIqrar.Visible = true;
                        panelAuthInfo.Visible = false;
                        try
                        {
                            mainTypeIqrar.SelectedIndex = Convert.ToInt32(colname.Split('-')[1]);
                            subTypeIqrar.Text = colname.Split('-')[0];
                            formNo = mainTypeAuth.Text + "-" + subTypeAuth.Text.Trim();
                        }
                        catch (Exception ex) { }
                    }



                    txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = false;
                    ContextPanel.BringToFront();
                    labID.BringToFront();
                    if (checkRequInfo(formNo))
                        OpenFile(formNo, true, formsBtn);
                    else
                    {
                        repReqPanel.BringToFront();
                        repReqPanel.Visible = true;
                        المعاملة.Text = formNo;
                        revisedRow("");
                        MessageBox.Show("لا يوجد قائمة بالمطلوبات الأولية للمعاملة، يرجى إضافتها");
                    }
                    if (!formsBtn.Enabled)
                    {
                        revisedRow("");
                        formsBtn.Enabled = true;
                        formsBtn.Text = "رفع الاستمارة الأولية";
                    }
                    panelLowButtons.Visible = ContextPanel.Visible = true;
                    button117.Enabled = true;
                    //review1 = true;
                    review1 = false;
                    btnDelete.Visible = btnClear.Visible = true;
                    DPTitle[0] = icheckoption11.Text + "_" + icheckoption12.Text;
                    dataGridView1.Visible = false;
                    //MessageBox.Show(ColRight.Text);
                    if (ColRight.Text != "")
                        btnRightsShow.PerformClick();
                }
        }

        private void checkSexType_CheckedChanged_2(object sender, EventArgs e)
        {

        }

        private void panelJob(int indEX)
        {
            foreach (Control control in panelText.Controls)
            {
                for (int index = 0; index < allList.Length; index++)
                {
                    if (allList[index] == control.Name)
                    {
                        control.Text = dataGridView1.Rows[indEX].Cells[allList[index]].Value.ToString();
                    }
                }
            }
            foreach (Control control in panelDate.Controls)
            {
                for (int index = 0; index < allList.Length; index++)
                {

                    if (allList[index] == control.Name)
                    {
                        control.Text = dataGridView1.Rows[indEX].Cells[allList[index]].Value.ToString();

                    }
                }
            }
            foreach (Control control in panelCombo.Controls)
            {
                for (int index = 0; index < allList.Length; index++)
                {
                    if (allList[index] == control.Name)
                    {
                        control.Text = dataGridView1.Rows[indEX].Cells[allList[index]].Value.ToString();
                    }
                }
            }
            foreach (Control control in panelButton.Controls)
            {
                for (int index = 0; index < allList.Length; index++)
                {
                    if (allList[index] == control.Name)
                    {
                        control.Text = dataGridView1.Rows[indEX].Cells[allList[index]].Value.ToString();
                    }
                }
            }
            foreach (Control control in panelCheck.Controls)
            {
                for (int index = 0; index < allList.Length; index++)
                {
                    if (allList[index] == control.Name)
                    {
                        if (allList[index] == "optionscheck1")
                        {
                            if (dataGridView1.Rows[indEX].Cells["optionscheck1"].Value.ToString().Trim() == "_")
                            {
                                icheckoption11.Text = icheckoption12.Text = "";
                            }
                            else if (dataGridView1.Rows[indEX].Cells["optionscheck1"].Value.ToString().Contains("_"))
                            {
                                icheckoption11.Text = dataGridView1.Rows[indEX].Cells["optionscheck1"].Value.ToString().Split('_')[0];
                                icheckoption12.Text = dataGridView1.Rows[indEX].Cells["optionscheck1"].Value.ToString().Split('_')[1];
                            }
                        }
                        else control.Text = dataGridView1.Rows[indEX].Cells[allList[index]].Value.ToString();
                    }
                }
            }
            for (int index = 0; index < allList.Length; index++)
                for (int x = 1; x < 6; x++)
                {
                    if (allList[index] == "icomboOption" + x.ToString())
                    {
                        txtComboOptions[x - 1] = dataGridView1.Rows[indEX].Cells[allList[index]].Value.ToString();
                    }
                }

            if (txtComboOptions[0] != "")
            {
                Vicombo1.Items.Clear();
                for (int x = 0; x < txtComboOptions[0].Split('_').Length; x++)
                    Vicombo1.Items.Add(txtComboOptions[0].Split('_')[x]);
                Vicombo1.SelectedIndex = 0;
            }
            if (txtComboOptions[1] != "")
            {
                Vicombo2.Items.Clear();
                for (int x = 0; x < txtComboOptions[1].Split('_').Length; x++)
                    Vicombo2.Items.Add(txtComboOptions[1].Split('_')[x]);
                Vicombo2.SelectedIndex = 0;
            }
            if (txtComboOptions[2] != "")
            {
                Vicombo3.Items.Clear();
                for (int x = 0; x < txtComboOptions[2].Split('_').Length; x++)
                    Vicombo3.Items.Add(txtComboOptions[2].Split('_')[x]);
                Vicombo3.SelectedIndex = 0;
            }
            if (txtComboOptions[3] != "")
            {
                Vicombo4.Items.Clear();
                for (int x = 0; x < txtComboOptions[3].Split('_').Length; x++)
                    Vicombo4.Items.Add(txtComboOptions[3].Split('_')[x]);
                Vicombo4.SelectedIndex = 0;
            }
            if (txtComboOptions[4] != "")
            {
                Vicombo5.Items.Clear();
                for (int x = 0; x < txtComboOptions[4].Split('_').Length; x++)
                    Vicombo5.Items.Add(txtComboOptions[4].Split('_')[x]);
                Vicombo5.SelectedIndex = 0;
            }
        }

        private void comboRights_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void deleteItemsAO()
        {
            checkboxdt = new DataTable();
            checkboxdt.Clear();
            Nobox = 0;
            foreach (Control control in panelAuthOptions.Controls)
            {
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


        }

        public void PopulateCheckBoxes(string col, string table, string dataSource)
        {
            if (col == "" || table == "" || dataSource == "") return;
            col = col.Replace("-", "_");
            col = col.Replace(" ", "_");
            string query = "SELECT ID," + col + " FROM " + table;
            //MessageBox.Show(query);
            using (SqlConnection con = new SqlConnection(dataSource))
            {

                using (SqlDataAdapter sda = new SqlDataAdapter(query, con))
                {
                    try
                    {
                        sda.Fill(checkboxdt);
                        listchecked = checkboxdt.Rows.Count;
                        Nobox = 0;
                        foreach (DataRow row in checkboxdt.Rows)
                        {
                            if (checkboxdt.Rows[Nobox][col].ToString() == "" || checkboxdt.Rows[Nobox][col].ToString() == "null") return;
                            //{
                            Text_statis = checkboxdt.Rows[Nobox][col].ToString().Split('_');
                            //MessageBox.Show(checkboxdt.Rows[Nobox][col].ToString() + " ### "+ Text_statis.Length.ToString());
                            if (Text_statis.Length == 5 && checkboxdt.Rows[Nobox]["ID"].ToString() != "1")
                            {
                                CheckBox chk = new CheckBox();
                                chk.TabIndex = Nobox;
                                chk.Width = 80;
                                chk.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                                if (Nobox == 0) chk.Width = panelAuthOptions.Width - 100;
                                else chk.Width = panelAuthOptions.Width - 130;
                                chk.Height = 33;
                                chk.CheckState = CheckState.Unchecked;
                                chk.Location = new System.Drawing.Point(70, 3 + Nobox * 37);
                                chk.Name = "checkBox" + Nobox.ToString();

                                statistic[Nobox] = Convert.ToInt32(Text_statis[1]);
                                times[Nobox] = Convert.ToInt32(Text_statis[2]);
                                staticIndex[Nobox] = Convert.ToInt32(Text_statis[3]);

                                //string text = SuffPrefReplacements(Text_statis[0]);
                                //text = SuffPrefReplacements(text);
                                chk.Text = checkboxdt.Rows[Nobox][col].ToString();
                                chk.Tag = "valid";
                                if (Text_statis[4] == "Star")
                                    chk.CheckState = CheckState.Checked;

                                panelAuthOptions.Controls.Add(chk);
                                PictureBox picboxedit = new PictureBox();
                                picboxedit.Image = global::PersAhwal.Properties.Resources.edit;
                                picboxedit.Location = new System.Drawing.Point(55, Nobox * 37);
                                picboxedit.Name = Nobox.ToString();
                                picboxedit.Size = new System.Drawing.Size(24, 26);
                                picboxedit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                                picboxedit.TabIndex = 175 + Nobox;
                                picboxedit.TabStop = false;
                                picboxedit.Click += new System.EventHandler(this.pictureBoxedit_Click);
                                panelAuthOptions.Controls.Add(picboxedit);

                                PictureBox picboxup = new PictureBox();
                                picboxup.Image = global::PersAhwal.Properties.Resources.arrowup;
                                picboxup.Location = new System.Drawing.Point(86, Nobox * 37);
                                picboxup.Name = "Up";
                                picboxup.Size = new System.Drawing.Size(24, 26);
                                picboxup.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                                picboxup.TabIndex = 176 + Nobox;
                                picboxup.TabStop = false;
                                picboxup.Click += new System.EventHandler(this.pictureBoxup_Click);
                                if (Nobox == 0)
                                {
                                    picboxup.Visible = false;
                                }
                                if (chk.Text.Contains("لمن يشهد والله خير الشاهدين")||chk.Text.Contains("ويعتبر التوكيل")|| chk.Text.Contains("الحق في توكيل الغير") ) picboxup.Visible = false;
                                panelAuthOptions.Controls.Add(picboxup);

                                PictureBox picboxdown = new PictureBox();
                                picboxdown.Image = global::PersAhwal.Properties.Resources.arrowdown;
                                picboxdown.Location = new System.Drawing.Point(55, Nobox * 37);
                                picboxdown.Name = "Down";
                                picboxdown.Size = new System.Drawing.Size(24, 26);
                                picboxdown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                                picboxdown.TabIndex = 177 + Nobox;
                                picboxdown.TabStop = false;
                                picboxdown.Click += new System.EventHandler(this.pictureBoxdown_Click);
                                if (chk.Text.Contains("الحق في توكيل الغير") || chk.Text.Contains("لمن يشهد والله خير الشاهدين") || chk.Text.Contains("ويعتبر التوكيل")) picboxdown.Visible = false;

                                panelAuthOptions.Controls.Add(picboxdown);
                                LastID = Convert.ToInt32(checkboxdt.Rows[Nobox]["ID"].ToString());

                            }
                            Nobox++;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("لا توجد قائمة حقوق بالاسم " + col);
                    }
                }
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

        private void button37_Click_1(object sender, EventArgs e)
        {
            panelAuthOptions.Visible = false;
            //authType.Text = "نموذج صيغة التفويض:";
            //TextModel.Text = "لينوب عني ويقوم مقامي في ";

        }

        private void button35_Click(object sender, EventArgs e)
        {
            if (button109.Text == "إضافة")
            {
                CheckBox chk = new CheckBox();
                chk.TabIndex = Nobox;
                chk.Width = 80;
                chk.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                chk.Width = panelAuthOptions.Width - 130;
                chk.Height = 33;
                chk.Tag = "valid";
                chk.CheckState = CheckState.Checked;
                chk.Location = new System.Drawing.Point(60, 3 + Nobox * 37);
                chk.Name = "checkBox" + Nobox.ToString();
                chk.Text = TextModel.Text;
                TextModel.Clear();
                statistic[Nobox] = 1;
                times[Nobox] = 1;
                panelAuthOptions.Controls.Add(chk);

                PictureBox picboxedit = new PictureBox();
                picboxedit.Image = global::PersAhwal.Properties.Resources.edit;
                picboxedit.Location = new System.Drawing.Point(55, Nobox * 37);
                picboxedit.Name = "Edit";
                picboxedit.Size = new System.Drawing.Size(24, 26);
                picboxedit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                picboxedit.TabIndex = 175 + Nobox;
                picboxedit.TabStop = false;
                picboxedit.Click += new System.EventHandler(this.pictureBoxedit_Click);
                panelAuthOptions.Controls.Add(picboxedit);

                PictureBox picboxup = new PictureBox();
                picboxup.Image = global::PersAhwal.Properties.Resources.arrowup;
                picboxup.Location = new System.Drawing.Point(76, Nobox * 37);
                picboxup.Name = "Up";
                picboxup.Size = new System.Drawing.Size(24, 26);
                picboxup.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                picboxup.TabIndex = 176 + Nobox;
                picboxup.TabStop = false;
                picboxup.Visible = false;
                picboxup.Click += new System.EventHandler(this.pictureBoxup_Click);
                panelAuthOptions.Controls.Add(picboxup);

                PictureBox picboxdown = new PictureBox();
                picboxdown.Image = global::PersAhwal.Properties.Resources.arrowdown;
                picboxdown.Location = new System.Drawing.Point(45, Nobox * 37);
                picboxdown.Size = new System.Drawing.Size(24, 26);
                picboxdown.Name = "Down";
                picboxdown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                picboxdown.TabIndex = 177 + Nobox; ;
                picboxdown.TabStop = false;
                picboxdown.Visible = false;
                picboxdown.Click += new System.EventHandler(this.pictureBoxdown_Click);

                panelAuthOptions.Controls.Add(picboxdown);

                //UpdateColumn(DataSource, RightColumnName, LastID + 1, chk.Text + "_" + statistic[Nobox].ToString() + "_" + times[Nobox].ToString() + "_" + staticIndex[Nobox].ToString() + "_Off", true);
                Nobox++;
                for (int swap = 0; swap < 2; swap++)

                {
                    SwapText(Nobox - swap);
                    ShowArrows(Nobox, swap);
                }

            }
            else if (button109.Text == "تعديل")
            {
                foreach (Control control in panelAuthOptions.Controls)
                {
                    if (control is CheckBox)
                    {
                        if (((CheckBox)control).TabIndex == LastTabIndex)
                        {
                            ((CheckBox)control).Text = TextModel.Text;
                            button109.Text = "لينوب عني ويقوم مقامي ";
                            TextModel.Text = "";
                        }
                    }
                }
            }
        }

        private void UpdateColumn(string source, string comlumnName, int id, string data, bool datatype)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string column = "@" + comlumnName;
            string qurey;
            if (datatype) qurey = "INSERT INTO TableAuthRights (" + comlumnName + ") values(" + column + ")";
            else qurey = "UPDATE TableAuthRights SET " + comlumnName + " = " + column + " WHERE ID = @ID";

            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;

            if (datatype)
            {
                sqlCmd.Parameters.AddWithValue(column, data.Trim());
                sqlCmd.ExecuteNonQuery();
            }
            else
            {

                sqlCmd.Parameters.AddWithValue("@ID", id);
                sqlCmd.Parameters.AddWithValue(column, data.Trim());
                sqlCmd.ExecuteNonQuery();
            }
            sqlCon.Close();
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

        private void button31_Click(object sender, EventArgs e)
        {

        }

        private void ComboProcedure_TextChanged(object sender, EventArgs e)
        {

        }


        public void pictureBoxedit_Click(object sender, EventArgs e)
        {
            PictureBox picbox = (PictureBox)sender;
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (((CheckBox)control).TabIndex == Convert.ToInt32(picbox.Name))
                    {
                        txtRightEditAdd.Text = ((CheckBox)control).Text;
                        btnRightEditAdd.Text = "تعديل";
                        remove.Visible = true;
                        LastTabIndex = ((CheckBox)control).TabIndex;

                    }
                }
            }
        }

        private void button39_Click(object sender, EventArgs e)
        {

        }



        private void btnDelete_Click(object sender, EventArgs e)
        {
            deleteRowsData(idIndex, "TableAddContext", DataSource);
            dataGridView1.BringToFront();
            FillDataGridView("TableAddContext", AuthType);
            panelLowButtons.Visible = panelIqrar.Visible = panelAuthInfo.Visible = false;
            txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = true;
            repReqPanel.Visible = flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
            panelLowButtons.Visible = ContextPanel.Visible = true; 
        }



        private void button40_Click(object sender, EventArgs e)
        {

            if (ContextPanel.Visible)
            {
                ContextPanel.Visible = false;

            }
            else
            {
                txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = false;
                ContextPanel.Visible = false;
                panelLowButtons.Visible = ContextPanel.Visible = true;
                flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
            }
        }

        private void button127_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            txtApp.Text = dlg.FileName;
        }

        private void button121_Click(object sender, EventArgs e)
        {

        }

        private void button121_Click_1(object sender, EventArgs e)
        {
            FolderAppUpdate(txtApp.Text);
            txtApp.Text = "";
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            pTextHieght += 42;
            panelText.Height = pTextHieght;


        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            pTextHieght -= 42;
            panelText.Height = pTextHieght;
        }

        private void IqrarPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void flowLayoutPanel4_MouseHover(object sender, EventArgs e)
        {
            //if(flowLayoutPanel4.Height == 42) 
            //    flowLayoutPanel4.Height = pTextHieght;
        }

        private void IqrarPanel_MouseClick(object sender, MouseEventArgs e)
        {
            panelButton.Height = panelCombo.Height = panelDate.Height = panelCheck.Height = panelText.Height = 42;
        }

        private void flowLayoutPanel4_MouseEnter(object sender, EventArgs e)
        {
            panelButton.Height = panelCombo.Height = panelDate.Height = panelCheck.Height = 42;
            panelText.Height = pTextHieght;
        }

        private void pictureBox23_Click(object sender, EventArgs e)
        {
            pCheckHieght += 42;
            panelCheck.Height = pCheckHieght;
        }

        private void pictureBox24_Click(object sender, EventArgs e)
        {
            pCheckHieght -= 42;
            panelCheck.Height = pCheckHieght;
        }

        private void panelCombo_MouseEnter(object sender, EventArgs e)
        {
            panelButton.Height = panelCombo.Height = panelDate.Height = panelText.Height = 42;
            panelCheck.Height = pComboHieght;
        }

        private void pictureBox34_Click(object sender, EventArgs e)
        {
            pComboHieght += 42;
            panelCombo.Height = pComboHieght;
        }

        private void pictureBox33_Click(object sender, EventArgs e)
        {
            pComboHieght -= 42;
            panelCombo.Height = pComboHieght;
        }

        private void pictureBox47_Click(object sender, EventArgs e)
        {
            pDateHieght += 42;
            panelDate.Height = pDateHieght;
        }

        private void pictureBox48_Click(object sender, EventArgs e)
        {
            pDateHieght -= 42;
            panelDate.Height = pDateHieght;
        }

        private void pictureBox57_Click(object sender, EventArgs e)
        {
            pbuttonHieght += 42;
            panelButton.Height = pbuttonHieght;
        }

        private void pictureBox58_Click(object sender, EventArgs e)
        {
            pbuttonHieght -= 42;
            panelButton.Height = pbuttonHieght;
        }

        private void panelCombo_MouseEnter_1(object sender, EventArgs e)
        {
            panelButton.Height = panelDate.Height = panelCheck.Height = panelText.Height = 42;
            panelCombo.Height = pComboHieght;
        }

        private void button81_Click(object sender, EventArgs e)
        {
            if (combo2index == 0)
                txtComboOptions[1] = iOptions2.Text;
            else txtComboOptions[1] = txtComboOptions[1] + "_" + iOptions2.Text;
            combo2index++;
            iOptions2.Text = "";
        }

        private void check1_CheckedChanged(object sender, EventArgs e)
        {
            if (Vicheck1.CheckState == CheckState.Unchecked) { Vicheck1.Text = DPTitle[0].Split('_')[0]; }
            else
            {
                Vicheck1.Text = DPTitle[0].Split('_')[1];
            }

        }

        private void CombAuthType_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (checkColumnName(mainTypeAuth.Text.Replace(" ", "_")))
            {
                //MessageBox.Show(mainTypeAuth.Text.Replace(" ", "_"));
                subTypeAuth.Items.Clear();
                //newFillComboBox1(subTypeAuth, DataSource, mainTypeAuth.SelectedIndex.ToString(), langAuth.Text);
                fillSubComboBox(subTypeAuth, DataSource, mainTypeAuth.Text.Replace(" ", "_"), "TableListCombo",false);
                NewColumn = false;
                checkIndex = true;
                return;
            }
            NewColumn = true;
            //if (CombAuthType.SelectedIndex >= 2 && CombAuthType.SelectedIndex <= 5)
            //{
            //    ComboProcedure.Items.Clear();
            //    ComboProcedure.Items.Add("إجراء جديد");
            //    ColumnName = "Row1Attach";
            //    fileComboBox(ComboProcedure, DataSource, ColumnName, "TableListCombo");
            //}
            //if (CombAuthType.Text.Contains("زواج"))
            //{
            //    ComboProcedure.Items.Clear();
            //    ComboProcedure.Items.Add("إجراء جديد");
            //    ColumnName = "RowMerrageAttach";

            //    fileComboBox(ComboProcedure, DataSource, "RowMerrageAttach", "TableListCombo");
            //}
            //if (CombAuthType.Text.Contains("ورثة"))
            //{
            //    ComboProcedure.Items.Clear();
            //    ComboProcedure.Items.Add("إجراء جديد");
            //    ColumnName = "RowLegacyAttach";
            //    fileComboBox(ComboProcedure, DataSource, "RowLegacyAttach", "TableListCombo");
            //}
            //if (CombAuthType.Text.Contains("سيارة"))
            //{
            //    ComboProcedure.Items.Clear();
            //    ComboProcedure.Items.Add("إجراء جديد");
            //    ColumnName = "RowCarAttach";
            //    fileComboBox(ComboProcedure, DataSource, "RowCarAttach", "TableListCombo");
            //}
            //if (CombAuthType.Text.Contains("طلاق"))
            //{
            //    ComboProcedure.Items.Clear();
            //    ComboProcedure.Items.Add("إجراء جديد");
            //    ColumnName = "RowDeforceAttach";
            //    fileComboBox(ComboProcedure, DataSource, "RowDeforceAttach", "TableListCombo");
            //}
            //if (CombAuthType.Text.Contains("جامعية"))
            //{
            //    ComboProcedure.Items.Clear();
            //    ComboProcedure.Items.Add("إجراء جديد");
            //    ColumnName = "RowUniversityAttach";
            //    fileComboBox(ComboProcedure, DataSource, "RowUniversityAttach", "TableListCombo");
            //}
            //if (CombAuthType.Text.Contains("ميلاد"))
            //{
            //    ComboProcedure.Items.Clear();
            //    ComboProcedure.Items.Add("إجراء جديد");
            //    ComboProcedure.Items.Add("استخراج وتوثيق");
            //    ColumnName = "";
            //    ComboProcedure.Items.Add(" استخراج وتوثيق بدل فاقد");
            //}
            //if (CombAuthType.Text.Contains("بالتنازل"))
            //{
            //    ComboProcedure.Items.Clear();
            //    ComboProcedure.Items.Add("إجراء جديد");
            //    ColumnName = "GiveAway";
            //    fileComboBox(ComboProcedure, DataSource, "GiveAway", "TableListCombo");
            //}

            //if (CombAuthType.Text.Contains("تأمين"))
            //{
            //    ComboProcedure.Items.Clear();
            //    ComboProcedure.Items.Add("إجراء جديد");
            //    ComboProcedure.Items.Add("استلام تأمين");
            //    ColumnName = "";
            //}
            //if (ComboProcedure.Items.Count > 0) ComboProcedure.SelectedIndex = 0;
        }

        private void ColorFulGrid9()
        {

            int errorList = 0;
            int authrevices = 0;
            int allauth = 0;
            for (int i=0; i < dataGridView1.Rows.Count - 1; i++)
            {
                //

                //if (dataGridView1.Rows[i].Cells["errorList"].Value.ToString() != "")
                //{
                //    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightCyan;
                //    errorList++;
                //}
                if (!reqGrid)
                {
                    try
                    {
                        if (dataGridView1.Rows[i].Cells["revised"].Value.ToString() == "")
                        {
                            dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;
                            if (dataGridView1.Rows[i].Cells["ColRight"].Value.ToString() != "") authrevices++;
                        }
                        else dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                        if (dataGridView1.Rows[i].Cells["ColRight"].Value.ToString() != "") allauth++;
                    }
                    catch (Exception ex) { }
                }
                else
                {
                    allauth++;
                    if (dataGridView1.Rows[i].Cells["revised"].Value.ToString() == "")
                    {
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;
                    }
                    if (dataGridView1.Rows[i].Cells["proForm1"].Value.ToString() == "")
                    {
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightCyan;
                    }
                    else if (dataGridView1.Rows[i].Cells["revised"].Value.ToString() != "")
                    {
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                        authrevices++;
                    }
                    }
             }
            if (errorList > 0) labelarch.Visible = true;
            else labelarch.Visible = false;
            //labelarch.Text = "عدد (" + errorList.ToString() + ") معاملة تحتاج إلى تصحيح ";
            if (!reqGrid)
            {
                labelarch.Text = "عدد (" + authrevices.ToString() + ") تم معاينتها من أصل ( " + allauth + ") معاملة ";
                label1.Text = "عدد (" + authrevices.ToString() + ") تم معاينتها من أصل ( " + allauth + ") معاملة ";
            }
            else {
                labelarch.Text = "عدد (" + authrevices.ToString() + ") تم معاينتها من أصل ( " + allauth + ") معاملة ";
                label1.Text = "عدد (" + authrevices.ToString() + ") تم معاينتها من أصل ( " + allauth + ") معاملة ";
            }

        }

        private void ComboProcedure_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            //if (!NewColumn) return;
            //FillDataGridView("TableAddContext");
            //MessageBox.Show(subTypeAuth.Text + "-" + mainTypeAuth.SelectedIndex.ToString());
            dataGridView1_RowIndex(subTypeAuth.Text + "-" + mainTypeAuth.SelectedIndex.ToString());
                string formNo = mainTypeAuth.Text + "-" + subTypeAuth.Text.Trim();
            //checkIndex = false;
            OpenFile(formNo, false, formsBtn);
            //flllPanelItemsboxes("ColName", subTypeAuth.Text + "-" + mainTypeAuth.SelectedIndex.ToString());
            //for (int id = 0; id < dataGridView1.Rows.Count - 1; id++)
            //{
            //    if (dataGridView1.Rows[id].Cells[25].Value.ToString() == subTypeAuth.Text + "-" + mainTypeAuth.SelectedIndex.ToString())
            //    {
            //        ShowRowNo(id);
            //        review1 = true;
            //        DPTitle[0] = icheckoption11.Text + "_" + icheckoption12.Text;
            //        ExtendedFillBox(itext1.Text, Convert.ToInt32(itext1Length.Text), itext2.Text, Convert.ToInt32(itext2Length.Text), itext3.Text, Convert.ToInt32(itext3Length.Text), itext4.Text, Convert.ToInt32(itext4Length.Text), itext5.Text, Convert.ToInt32(itext5Length.Text), "", 50, "", 50, "", 50, "", 50, "", 50, icheck1.Text, "", "", "", "", itxtDate1.Text, "", "", "", "", icombo1.Text, txtComboOptions[0].Split('_'), icombo2.Text, txtComboOptions[1].Split('_'), "", Empty, "", Empty, "", Empty, ibtnAdd1.Text, "", "", "", "");
            //        idIndex = Convert.ToInt32(dataGridView1.Rows[id].Cells[0].Value.ToString());
            //        button117.Enabled = button109.Enabled = true;
            //    }
            //}
        }


        private string OpenFile(string documenNo, bool printOut, Button button)
        {
            string query = "SELECT ID, proForm1,Data1, Extension1 from TableProcReq where المعاملة=@المعاملة";
            
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@المعاملة", SqlDbType.NVarChar).Value = documenNo;
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            button.Enabled = false;
           
            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                string str = reader["proForm1"].ToString();
                Console.WriteLine(str);
                try
                {
                    var Data = (byte[])reader["Data1"];
                
                CurrentFile = ArchFile + @"\formUpdated\" + str +".docx";
                string filePath = ArchFile + @"\" + str +".docx";
                if (File.Exists(CurrentFile) && !fileIsOpen(CurrentFile)) {
                    File.Delete(CurrentFile);
                }
                
                if (!File.Exists(CurrentFile))
                {
                    try
                    {
                        //File.Delete(CurrentFile);

                        button.Enabled = true;

                        if (printOut)
                        {
                            File.WriteAllBytes(filePath, Data);

                            System.IO.File.Copy(filePath, CurrentFile);

                            FileInfo fileInfo = new FileInfo(CurrentFile);
                            if (fileInfo.IsReadOnly) fileInfo.IsReadOnly = false;
                                Console.WriteLine("CurrentFile " + CurrentFile);
                                System.Diagnostics.Process.Start(CurrentFile);
                        }
                        return CurrentFile;
                    }
                    catch (Exception ex)
                    {
                            Console.WriteLine("fail " +str);
                            button.Enabled = false;
                        return "";
                    }

                }
                else if (File.Exists(CurrentFile) && fileIsOpen(CurrentFile))
                {
                    button.Enabled = false;
                    MessageBox.Show("يرجى إغلاق الملف " + str + " أولا");
                }
                }
                catch (Exception ex)
                {
                    button.Enabled = false;
                    return "";
                }
            }
            else button.Enabled = false;
            Con.Close();
            return "";
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            ColorFulGrid9();
            if (CurrentFile == "" || !readyToUpload) return;
            else if (File.Exists(CurrentFile) && !fileIsOpen(CurrentFile) && readyToUpload)
            {
                //Console.WriteLine("upload file " +CurrentFile);
                uploadForms(CurrentFile);
                readyToUpload = false;
            }
        }

        private void uploadForms(string location)
        {
            if (location != "" && File.Exists(location) && !fileIsOpen(location))
            {
                using (Stream stream = File.OpenRead(location))
                {
                    byte[] buffer1 = new byte[stream.Length];
                    stream.Read(buffer1, 0, buffer1.Length);
                    var fileinfo1 = new FileInfo(location);
                    string formNo = fileinfo1.Name.Replace(".docx","");
                    string query = "UPDATE TableProcReq SET Data1=@Data1 WHERE المعاملة=N'" + formNo + "'";
                    //MessageBox.Show(query);
                    SqlConnection sqlCon = new SqlConnection(DataSource);
                    try
                    {
                        if (sqlCon.State == ConnectionState.Closed)
                            sqlCon.Open();
                    }
                    catch (Exception ex) { return; }
                    SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                    sqlCmd.ExecuteNonQuery();
                    sqlCon.Close();
                    try
                    {
                        //File.Delete(CurrentFile);
                    }
                    catch (Exception ex) { CurrentFile = ""; }
                    label1.Visible = true;
                    return;
                }
            }
        }
        
        private void uploadFormsReq(string location)
        {
            if (location != "" && File.Exists(location) && !fileIsOpen(location))
            {
                using (Stream stream = File.OpenRead(location))
                {
                    byte[] buffer1 = new byte[stream.Length];
                    stream.Read(buffer1, 0, buffer1.Length);
                    var fileinfo1 = new FileInfo(location);
                    string query = "UPDATE TableProcReq SET Data1=@Data1,proForm1=@proForm1 WHERE المعاملة=N'" + المعاملة.Text + "'";
                    //MessageBox.Show(query);
                    SqlConnection sqlCon = new SqlConnection(DataSource);
                    try
                    {
                        if (sqlCon.State == ConnectionState.Closed)
                            sqlCon.Open();
                    }
                    catch (Exception ex) { }
                    SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                    sqlCmd.Parameters.Add("@proForm1", SqlDbType.NVarChar).Value = المعاملة.Text;
                    //MessageBox.Show(المعاملة.Text);
                    sqlCmd.ExecuteNonQuery();
                    sqlCon.Close();
                    
                    label1.Visible = true;
                    return;
                }
            }
        }
        public bool fileIsOpen(string path)
        {
            System.IO.FileStream a = null;

            try
            {
                a = System.IO.File.Open(path,
                System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None);
                return false;
            }
            catch (System.IO.IOException ex)
            {
                return true;
            }

            finally
            {
                if (a != null)
                {
                    a.Close();
                    a.Dispose();
                }
            }
        }
        private void button117_Click(object sender, EventArgs e)
        {
            ColumnJobs();
            createNewColContext();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand(updateAll, sqlCon);
            //MessageBox.Show(idIndex.ToString());
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", idIndex);
            sqlCmd.Parameters.AddWithValue("@errorList", "");
            sqlCmd.Parameters.AddWithValue("@revised", revised);
            addParameters(sqlCmd);
            sqlCmd.Parameters.AddWithValue("@ColRight", ColRight.Text); 
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
            var selectedOption = MessageBox.Show("", "إنهاء المراجعة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                btnRevised.PerformClick();

            }
            else {
                dataGridView1.BringToFront();
                FillDataGridView("TableAddContext", AuthType);
                panelLowButtons.Visible = panelIqrar.Visible = panelAuthInfo.Visible = false;
                txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = true;
                flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
                panelLowButtons.Visible = ContextPanel.Visible = true;
            }
            //this.Close();
            //ClearFileds();
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            mainTypeIqrar.Items.Clear();
            if (langIqrar.CheckState == CheckState.Checked)

            {
                langIqrar.Text = "الانجليزية";
                fileComboBox(mainTypeIqrar, DataSource, "EnglishGenIgrar", "TableListCombo",false);
                newFillComboBox2(subTypeIqrar, DataSource, mainTypeIqrar.SelectedIndex.ToString(), langIqrar.Text);
                PanelItemsboxes.RightToLeft = TextModel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            }
            else
            {
                langIqrar.Text = "العربية";
                fileComboBox(mainTypeIqrar, DataSource, "ArabicGenIgrar", "TableListCombo",false);
                newFillComboBox2(subTypeIqrar, DataSource, mainTypeIqrar.SelectedIndex.ToString(), langIqrar.Text);
                PanelItemsboxes.RightToLeft = TextModel.RightToLeft = System.Windows.Forms.RightToLeft.No;
            }
        }

        private void checkBox2_CheckedChanged_1(object sender, EventArgs e)
        {

        }



        private void language_CheckedChanged(object sender, EventArgs e)
        {
            if (langAuth.CheckState == CheckState.Checked)

            {
                langAuth.Text = "الانجليزية";
            }
            else if (langAuth.CheckState == CheckState.Unchecked)
            {
                langAuth.Text = "العربية";
            }
        }


        private void button12_Click(object sender, EventArgs e)
        {

            panelAuthOptions.Visible = true;
            panelAuthOptions.BringToFront();
            //TextModel.Text = "";
            deleteItemsAO(); 
            RightColumnName = ColRight.Text;
            PopulateCheckBoxes(ColRight.Text.Trim(), "TableAuthRights", DataSource);
            //ContextPanel.Visible = false;
        }

        private void btnRights_Click(object sender, EventArgs e)
        {
            button109.Text = "أنا المواطن/";
            panelAuthInfo.Visible = false;
            panelIqrar.Visible = true;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            button109.Text = "لينوب عني ويقوم مقامي في ";
            panelAuthInfo.Visible = true;
            panelIqrar.Visible = false;
        }

        //private void checkBox1_CheckedChanged_2(object sender, EventArgs e)
        //{
        //    if (newMainType.CheckState == CheckState.Checked)
        //    {
        //        newMainType.Text = "نوع رئيس جديد";
        //        mainTypeAuth.Visible = false;
        //        newCombAuthType.Visible = true;
        //        newCombAuthType.Size = new System.Drawing.Size(200, 35);
        //    }
        //    else
        //    {
        //        newCombAuthType.Size = new System.Drawing.Size(18, 35);
        //        newMainType.Text = "إضافة إلى نوع موجود";
        //        mainTypeAuth.Visible = true;
        //        newCombAuthType.Visible = false;
        //    }
        //}

        private void ColumnJobs() {
            //CombAuthType.SelectedIndex.ToString()
            if (panelAuthInfo.Visible)
            {
                if (!checkColExist("TableListCombo", mainTypeAuth.Text.Replace(" ", "_")))
                {
                    addMainAuth(mainTypeAuth.Text);
                    CreateColumn(mainTypeAuth.Text.Replace(" ", "_"), "TableListCombo");
                    mainTypeAuth.Items.Add(mainTypeAuth.Text);
                }
                CombAuthTypeIndex = mainTypeAuth.Items.Count;
                ColumnName = subTypeAuth.Text.Replace(" ", "_") + "-" + (mainTypeAuth.Items.Count).ToString();
                int curenntID = getCurrentID(DataSource, mainTypeAuth.Text.Replace(" ", "_"), "TableListCombo", subTypeAuth.Text);
                if (curenntID == 0)
                    curenntID = getLastID(DataSource, mainTypeAuth.Text.Replace(" ", "_"), "TableListCombo");
                    //MessageBox.Show(curenntID.ToString());
                    addSubAuth(curenntID, subTypeAuth.Text, mainTypeAuth.Text.Replace(" ", "_"));
                
            }
            else
            {                
                if (!checkColExist("TableListCombo", mainTypeIqrar.Text.Replace(" ", "_")))
                {
                    CreateColumn(mainTypeIqrar.Text.Replace(" ", "_"), "TableListCombo");
                    mainTypeIqrar.Items.Add(mainTypeIqrar.Text);
                }
                ColumnName = subTypeIqrar.Text.Replace(" ", "_") + "-" + mainTypeIqrar.SelectedIndex.ToString();

                int curenntID = getCurrentID(DataSource, mainTypeIqrar.Text.Replace(" ", "_"), "TableListCombo", subTypeIqrar.Text);
                if (curenntID == 0)
                    curenntID = getLastID(DataSource, mainTypeAuth.Text.Replace(" ", "_"), "TableListCombo");
                addSubAuth(curenntID, subTypeIqrar.Text, mainTypeIqrar.Text.Replace(" ", "_"));
            }
            
        }

        private void btnRighrsEditAdd_Click(object sender, EventArgs e)
        {
            if (txtRightEditAdd.Text != "" && btnRightEditAdd.Text == "إضافة")
            {
                CheckBox chk = new CheckBox();
                chk.TabIndex = Nobox;
                chk.Width = 80;
                chk.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                chk.Width = panelAuthOptions.Width - 130;
                chk.Height = 33;
                chk.Tag = "valid";
                chk.CheckState = CheckState.Checked;
                chk.Location = new System.Drawing.Point(60, 3 + Nobox * 37);
                chk.Name = "checkBox" + Nobox.ToString();
                chk.Text = txtRightEditAdd.Text;
                txtRightEditAdd.Clear();
                statistic[Nobox] = 1;
                times[Nobox] = 1;
                panelAuthOptions.Controls.Add(chk);

                PictureBox picboxedit = new PictureBox();
                picboxedit.Image = global::PersAhwal.Properties.Resources.edit;
                picboxedit.Location = new System.Drawing.Point(55, Nobox * 37);
                picboxedit.Name = "Edit";
                picboxedit.Size = new System.Drawing.Size(24, 26);
                picboxedit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                picboxedit.TabIndex = 175 + Nobox;
                picboxedit.TabStop = false;
                picboxedit.Click += new System.EventHandler(this.pictureBoxedit_Click);
                panelAuthOptions.Controls.Add(picboxedit);

                PictureBox picboxup = new PictureBox();
                picboxup.Image = global::PersAhwal.Properties.Resources.arrowup;
                picboxup.Location = new System.Drawing.Point(76, Nobox * 37);
                picboxup.Name = "Up";
                picboxup.Size = new System.Drawing.Size(24, 26);
                picboxup.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                picboxup.TabIndex = 176 + Nobox;
                picboxup.TabStop = false;
                picboxup.Visible = false;
                picboxup.Click += new System.EventHandler(this.pictureBoxup_Click);
                panelAuthOptions.Controls.Add(picboxup);

                PictureBox picboxdown = new PictureBox();
                picboxdown.Image = global::PersAhwal.Properties.Resources.arrowdown;
                picboxdown.Location = new System.Drawing.Point(45, Nobox * 37);
                picboxdown.Size = new System.Drawing.Size(24, 26);
                picboxdown.Name = "Down";
                picboxdown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
                picboxdown.TabIndex = 177 + Nobox; ;
                picboxdown.TabStop = false;
                picboxdown.Visible = false;
                picboxdown.Click += new System.EventHandler(this.pictureBoxdown_Click);

                panelAuthOptions.Controls.Add(picboxdown);

                //UpdateColumn(ParentForm.PublicDataSource, LastCol, LastID + 1, chk.Text + "_" + statistic[Nobox].ToString() + "_" + times[Nobox].ToString() + "_" + staticIndex[Nobox].ToString() + "_Off", true);
                Nobox++;
                for (int swap = 0; swap < 2; swap++)

                {
                    SwapText(Nobox - swap);
                    ShowArrows(Nobox, swap);
                }

            }
            else if (txtRightEditAdd.Text != "" && btnRightEditAdd.Text == "تعديل")
            {
                foreach (Control control in panelAuthOptions.Controls)
                {
                    if (control is CheckBox)
                    {
                        if (((CheckBox)control).TabIndex == LastTabIndex)
                        {
                            ((CheckBox)control).Text = txtRightEditAdd.Text;
                            btnRightEditAdd.Text = "إضافة";
                            txtRightEditAdd.Text = "";
                            remove.Visible = false;
                        }
                    }
                }
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button15_Click(object sender, EventArgs e)
        {
            string[] rights = new string[100];
            int rightIndex = 0;
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox && !control.Text.Contains("(محذوف)") && control.Visible)
                {
                    
                    rights[rightIndex] = ((CheckBox)control).Text;
                    //MessageBox.Show(rights[rightIndex]); 
                    rightIndex++;
                }

            }

            for (int x = 0; x < rightIndex; x++)
            {

                if (rights[x] == "" || rights[x] == "Null") break;
                UpdateColumn(DataSource, ColRight.Text.Trim().Replace("-","_"), x + 2, rights[x]);
            }
            //MessageBox.Show("تم تعديل القائمة");
            button37.PerformClick();
        }
        private void UpdateColumn(string source, string comlumnName, int id, string data)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string column = "@" + comlumnName;
            string qurey = "UPDATE TableAuthRights SET " + comlumnName + " = " + column + " WHERE ID=@ID";
            //MessageBox.Show(qurey);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;

            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.Parameters.AddWithValue(column, data.Trim());
            sqlCmd.ExecuteNonQuery();

            sqlCon.Close();
        }

        private void comboRights_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void FormType_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkIndex = true; 
            switch (mainTypeIqrar.SelectedIndex) {
                case 0:
                    button109.Text = "أنا المواطن/";
                    break;
                case 1:
                    button109.Text = "تُفيد القنصـلية العـامة لجمهـورية الســودان بجـدة";
                    break;
                case 2:
                    button109.Text = "بهذا تشهد القنصـلية العـامة لجمهـورية الســودان بجـدة ";
                    break;
                case 3:
                    button109.Text = "بهذا تشهد القنصـلية العـامة لجمهـورية الســودان بجـدة ";
                    break;
                case 4:
                    button109.Text = "نص المذكرة";
                    break;
                case 5:
                    button109.Text = "نص البرقية";
                    break;
            }

            newFillComboBox1(subTypeIqrar, DataSource, mainTypeIqrar.Text.Replace(" ", "_"));
            //fileComboBox(ProFormType, DataSource, FormType.Text.Replace(" ", "_"), "TableListCombo");
            //newFillComboBox2(subTypeIqrar, DataSource, mainTypeIqrar.SelectedIndex.ToString(), langIqrar.Text);
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
        private void Settings_FormClosed(object sender, FormClosedEventArgs e)
        {
            string primeryLink = @"D:\PrimariFiles\";
            if (!Directory.Exists(@"D:\"))
            {
                ///وله بموجب هذا التوكيل الحق في استخراج الشهادة، ومقابلة كافة الجهات المختصة في كل من مصلحة الأراضي والتسجيلات والمساحة، والوقوف والمقاضاة نيابة عني أمام كافة المحاكم والنيابات بمختلف أنواعها ودرجاتها، والقيام بكافة الإجراءات التي تتطلب حضوري، والتوقيع نيابة عني على كافة الأوراق والمستندات اللازمة لذلك، وله الحق في توكيل الغير في بعض أو كل مما أوكل فيه، وحظر قطعة الارض وأذنت لمن يشهد والله خير الشاهدين
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


        private void icheckoption11_TextChanged(object sender, EventArgs e)
        {
            optionscheck1.Text = icheckoption11.Text + "_" + icheckoption12.Text;
        }

        private void icheckoption12_TextChanged(object sender, EventArgs e)
        {
            optionscheck1.Text = icheckoption11.Text + "_" + icheckoption12.Text;
        }

        private void ProFormType_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillDataGridView("TableAddContext", AuthType);
            dataGridView1_RowIndex(subTypeIqrar.Text + "-" + mainTypeIqrar.SelectedIndex.ToString());
            checkIndex = false;
            //for (int id = 0; id < dataGridView1.Rows.Count - 1; id++)
            //{
            //    if (dataGridView1.Rows[id].Cells[25].Value.ToString() == subTypeIqrar.Text + "-" + mainTypeIqrar.SelectedIndex.ToString())
            //    {
            //        ShowRowNo(id);
            //        review1 = true;
            //        DPTitle[0] = icheckoption11.Text + "_" + icheckoption12.Text;                    
            //        idIndex = Convert.ToInt32(dataGridView1.Rows[id].Cells[0].Value.ToString());
            //        button117.Enabled = button109.Enabled = true;
            //    }
            //}
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            reqGrid = false; 
            AuthType = false; 
            if (dataGridView1.Visible)
            {
                txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = false;
                flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
                dataGridView1.SendToBack();
                panelLowButtons.Visible = ContextPanel.Visible = true;
                ContextPanel.BringToFront();
            }
            else
            {

                repReqPanel.Visible = false;
                dataGridView1.BringToFront();
                FillDataGridView("TableAddContext", false);
                panelLowButtons.Visible = panelIqrar.Visible = panelAuthInfo.Visible = false;
                txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = true;
                flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
                panelLowButtons.Visible = ContextPanel.Visible = true;
                formsBtn.Visible = proFileBtn.Visible = btnRevised.Visible = true;
            }
            //if (dataGridView1.Visible)
            //{
            //    txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = false;
            //    flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
            //}
            //else
            //{
            //    panelIqrar.Visible = panelAuthInfo.Visible = panelLowButtons.Visible = false;
            //    txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = true;
            //    flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
            //    ContextPanel.Visible = false;
            //}
        }

        private void addParameters( SqlCommand sqlCmd)
        {
            foreach (Control control in panelText.Controls)
            {
                for (int index = 0; index < allList.Length; index++)
                {
                    if (allList[index] == control.Name)
                    {
                        sqlCmd.Parameters.AddWithValue("@" + control.Name, control.Text);
                    }
                }
            }
            foreach (Control control in panelDate.Controls)
            {
                for (int index = 0; index < allList.Length; index++)
                {
                    if (allList[index] == control.Name)
                    {
                        sqlCmd.Parameters.AddWithValue("@" + control.Name, control.Text);
                    }
                }
            }
            
            foreach (Control control in panelCombo.Controls)
            {
                for (int index = 0; index < allList.Length; index++)
                {
                    if (allList[index] == control.Name)
                    {
                        sqlCmd.Parameters.AddWithValue("@" + control.Name, control.Text);
                    }                    
                }
            }
            for (int index = 0; index < allList.Length; index++) 
                for (int x = 1; x < 6; x++)
                    {
                        if (allList[index] == "icombo" + x.ToString() + "Option")
                        {
                            sqlCmd.Parameters.AddWithValue("@" + allList[index], txtComboOptions[x - 1]);
                        }
                    }
            
            for (int index = 0; index < allList.Length; index++) 
                for (int x = 1; x < 6; x++)
                    {
                        if (allList[index] == "icheck" + x.ToString() + "Option")
                        {
                            sqlCmd.Parameters.AddWithValue("@" + allList[index], txtComboOptions[x - 1]);
                        }
                    }

            foreach (Control control in panelButton.Controls)
            {
                for (int index = 0; index < allList.Length; index++)
                {
                    if (allList[index] == control.Name)
                    {
                        sqlCmd.Parameters.AddWithValue("@" + control.Name, control.Text);
                    }
                }
            }
            foreach (Control control in panelCheck.Controls)
            {
                for (int index = 0; index < allList.Length; index++)
                {
                    if (allList[index] == control.Name)
                    {
                        sqlCmd.Parameters.AddWithValue("@" + control.Name, control.Text);
                    }
                }
            }
            sqlCmd.Parameters.AddWithValue("@TextModel", TextModel.Text);
            
            sqlCmd.Parameters.AddWithValue("@ColName", txtProName.Text);
            if (panelAuthInfo.Visible)
            {
                sqlCmd.Parameters.AddWithValue("@authIqrar", "auth");
                sqlCmd.Parameters.AddWithValue("@Lang", langAuth.Text);
                sqlCmd.Parameters.AddWithValue("@editRights", editRights);

                //sqlCmd.Parameters.AddWithValue("@ColName", (subTypeAuth.Text + "-" + CombAuthTypeIndex.ToString()).Replace("--", "-"));
            }
            else
            {
                sqlCmd.Parameters.AddWithValue("@authIqrar", "iqrar");
                sqlCmd.Parameters.AddWithValue("@Lang", langIqrar.Text);
                sqlCmd.Parameters.AddWithValue("@editRights", "");
                //sqlCmd.Parameters.AddWithValue("@ColName", (subTypeIqrar.Text + "-" + mainTypeIqrar.SelectedIndex.ToString()).Replace("--", "-"));
            }
        }
        private void createNewColContext()
        {
            foreach (Control control in panelText.Controls)
            {
                if (control is TextBox)
                    if (!checkColExist("TableAddContext", control.Name.Replace(" ", "_")) && control.Name.Contains("itext") && control.Text != "")
                    {
                        CreateColumn(control.Name.Replace(" ", "_"), "TableAddContext");
                        if (control.Name.Contains("1"))
                            CreateColumn("itext1Length", "TableAddContext");
                        else if (control.Name.Contains("2"))
                            CreateColumn("itext2Length", "TableAddContext");
                        else if (control.Name.Contains("3"))
                            CreateColumn("itext3Length", "TableAddContext");
                        else if (control.Name.Contains("4"))
                            CreateColumn("itext4Length", "TableAddContext");
                        else if (control.Name.Contains("5"))
                            CreateColumn("itext5Length", "TableAddContext");
                        else if (control.Name.Contains("6"))
                            CreateColumn("itext6Length", "TableAddContext");
                        else if (control.Name.Contains("7"))
                            CreateColumn("itext7Length", "TableAddContext");
                        else if (control.Name.Contains("8"))
                            CreateColumn("itext8Length", "TableAddContext");
                        else if (control.Name.Contains("9"))
                            CreateColumn("itext9Length", "TableAddContext");
                    }
            }
            foreach (Control control in panelCheck.Controls)
            {
                if (control is TextBox)
                    if (!checkColExist("TableAddContext", control.Name.Replace(" ", "_")) && control is TextBox && control.Name.Contains("icheck") && control.Text != "")
                    {
                        CreateColumn(control.Name.Replace(" ", "_"), "TableAddContext");
                        if (control.Name.Contains("1"))
                        {
                            CreateColumn("optionscheck1", "TableAddContext");
                        }
                        else if (control.Name.Contains("2"))
                        {
                            CreateColumn("optionscheck2", "TableAddContext");
                        }
                        else if (control.Name.Contains("3"))
                        {
                            CreateColumn("optionscheck3", "TableAddContext");
                        }
                        else if (control.Name.Contains("4"))
                        {
                            CreateColumn("optionscheck4", "TableAddContext");
                        }
                        else if (control.Name.Contains("5"))
                        {
                            CreateColumn("optionscheck5", "TableAddContext");
                        }
                    }
            }

            foreach (Control control in panelCombo.Controls)
            {
                if (control is TextBox)
                    if (!checkColExist("TableAddContext", control.Name.Replace(" ", "_")) && control is TextBox && control.Name.Contains("icombo")&& !control.Name.Contains("Length") && control.Text != "")
                    {
                        CreateColumn(control.Name.Replace(" ", "_"), "TableAddContext");
                        if (control.Name.Contains("1"))
                        {
                            CreateColumn("icombo1Length", "TableAddContext");
                            CreateColumn("icombo1Option", "TableAddContext");
                        }
                        else if (control.Name.Contains("2"))
                        {
                            CreateColumn("icombo2Length", "TableAddContext");
                            CreateColumn("icombo2Option", "TableAddContext");
                        }
                        else if (control.Name.Contains("3"))
                        {
                            CreateColumn("icombo3Length", "TableAddContext");
                            CreateColumn("icombo3Option", "TableAddContext");
                        }
                        else if (control.Name.Contains("4"))
                        {
                            CreateColumn("icombo4Length", "TableAddContext");
                            CreateColumn("icombo4Option", "TableAddContext");
                        }
                        else if (control.Name.Contains("5"))
                        {
                            CreateColumn("icombo5Length", "TableAddContext");
                            CreateColumn("icombo5Option", "TableAddContext");
                        }
                    }
            }
            foreach (Control control in panelDate.Controls)
            {
                if (control is TextBox)
                    if (!checkColExist("TableAddContext", control.Name.Replace(" ", "_")) && control is TextBox && control.Text != "")
                    {
                        CreateColumn(control.Name.Replace(" ", "_"), "TableAddContext");
                        if (control.Name.Contains("1"))
                            CreateColumn("dateType1", "TableAddContext");
                        else if (control.Name.Contains("2"))
                            CreateColumn("dateType2", "TableAddContext");
                        else if (control.Name.Contains("3"))
                            CreateColumn("dateType3", "TableAddContext");
                        else if (control.Name.Contains("4"))
                            CreateColumn("dateType4", "TableAddContext");
                        else if (control.Name.Contains("5"))
                            CreateColumn("dateType5", "TableAddContext");
                    }
            }
            foreach (Control control in panelButton.Controls)
            {
                if (control is TextBox)
                    if (!checkColExist("TableAddContext", control.Name.Replace(" ", "_")) && control is TextBox && control.Name.Contains("ibtnAdd") && control.Text != "")
                    {
                        CreateColumn(control.Name.Replace(" ", "_"), "TableAddContext");
                        if (control.Name.Contains("1"))
                            CreateColumn("ibtnAdd1Length", "TableAddContext");
                        else if (control.Name.Contains("2"))
                            CreateColumn("ibtnAdd2Length", "TableAddContext");
                        else if (control.Name.Contains("3"))
                            CreateColumn("ibtnAdd3Length", "TableAddContext");
                        else if (control.Name.Contains("4"))
                            CreateColumn("ibtnAdd4Length", "TableAddContext");
                        else if (control.Name.Contains("5"))
                            CreateColumn("ibtnAdd5Length", "TableAddContext");
                    }
            }
        }

        private string SuffPrefReplacements(string text, int appCaseIndex, int intAuthcases)
        {
            appCaseIndex = 0;
            intAuthcases = 0;

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

            if (text.Contains("m1"))
                return text.Replace("m1", Vicombo1.Text);
            if (text.Contains("m2"))
                return text.Replace("m2", Vicombo2.Text);

            if (text.Contains("a1"))
                return text.Replace("a1", LibtnAdd1.Text);

            if (text.Contains("n1"))
                return text.Replace("n1", " " + VitxtDate1.Text + " ");
            if (text.Contains("#*#"))
                return text.Replace("#*#", preffix[appCaseIndex, 10]);

            if (text.Contains("#1"))
                return text.Replace("#1", preffix[appCaseIndex, 11]);

            if (text.Contains("#2"))
                return text.Replace("#2", preffix[appCaseIndex, 12]);
            //if (text.Contains("@*@"))
            //{
            //    spacialCharacter = "@*@";
            //    return text.Replace("@*@", "لدى  برقم الايبان ()");
            //}

            //if (text.Contains("#8"))
            //    return text.Replace("#8", removedDocNo.Text);
            //if (text.Contains("#6"))
            //    return text.Replace("#6", removedDocSource.Text);
            //if (text.Contains("#7"))
            //    return text.Replace("#7", removedDocDate.Text);



            if (text.Contains("#3"))
                return text.Replace("#3", preffix[0, 7]);
            if (text.Contains("#4"))
                return text.Replace("#4", preffix[0, 8]);
            if (text.Contains("#5"))
                return text.Replace("#5", preffix[0, 9]);



            if (text.Contains("$$$"))
                return text.Replace("$$$", preffix[appCaseIndex, 0]);
            if (text.Contains("&&&"))
                return text.Replace("&&&", preffix[appCaseIndex, 1]);
            if (text.Contains("^^^"))
                return text.Replace("^^^", preffix[appCaseIndex, 2]);
            if (text.Contains("###"))
                return text.Replace("###", preffix[intAuthcases, 4]);
            if (text.Contains("***"))
                return text.Replace("***", preffix[intAuthcases, 3]);
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

            else return text;
        }
        private void flllPanelItemsboxes(string rowID, string cellValue)
        {
            //MessageBox.Show("rowID = " + rowID + " - cellValue=" + cellValue);
            int checkIndex = 0;
            if (dataGridView1.Rows.Count > 1)
            {
                for (int index = 0; index < dataGridView1.Rows.Count - 1; index++)
                    if (cellValue == dataGridView1.Rows[index].Cells[rowID].Value.ToString())
                    {
                        TextModel.Text = dataGridView1.Rows[index].Cells["TextModel"].Value.ToString();
                        
                        foreach (Control Lcontrol in PanelItemsboxes.Controls)
                            try
                            {
                                if (Lcontrol is CheckBox)
                                {                                    
                                    itemsicheck1[checkIndex] = dataGridView1.Rows[index].Cells[Lcontrol.Name.Replace("V", "") + "Option"].Value.ToString();
                                    Lcontrol.Text = itemsicheck1[checkIndex].Split('_')[0]; 
                                    checkIndex++;


                                }

                                if (Lcontrol.Name.StartsWith("L"))
                                {
                                    Lcontrol.Text = dataGridView1.Rows[index].Cells[Lcontrol.Name.Replace("L", "")].Value.ToString();
                                    if (Lcontrol.Text != "")
                                    {
                                        Lcontrol.Visible = true;
                                        foreach (Control Vcontrol in PanelItemsboxes.Controls)
                                        {
                                            //MessageBox.Show("View = " + Vcontrol.Name + " - Label=" + Lcontrol.Name.Replace("L", "V"));
                                            if (Vcontrol.Name.Trim() == Lcontrol.Name.Replace("L", "V").Trim())
                                            {
                                                
                                                Vcontrol.Visible = true;
                                                string size = dataGridView1.Rows[index].Cells[Lcontrol.Name.Replace("L", "") + "Length"].Value.ToString();
                                                //MessageBox.Show(Lcontrol.Name +"-"+size); 
                                                Vcontrol.Width = Convert.ToInt32(size);
                                                if (Convert.ToInt32(size) >= 700)
                                                {
                                                    if (Vcontrol is TextBox) ((TextBox)Vcontrol).Multiline = true;
                                                    Vcontrol.Height = 150;
                                                }

                                                if (Vcontrol is ComboBox)
                                                {
                                                    ((ComboBox)Vcontrol).Items.Clear();
                                                    string[] items = dataGridView1.Rows[index].Cells[Lcontrol.Name.Replace("L", "") + "Option"].Value.ToString().Split('_');

                                                    for (int x = 0; x < items.Length; x++)
                                                        ((ComboBox)Vcontrol).Items.Add(items[x]);
                                                }

                                                

                                                //MessageBox.Show(Lcontrol.Name + "Length");
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


        }

        private void Vicheck_CheckedChanged(object sender, EventArgs e)
        {
            //
            if (dataGridView1.Rows.Count > 1)
                for (int index = 0; index < dataGridView1.Rows.Count - 1; index++)
                    if (idIndex == Convert.ToInt32(dataGridView1.Rows[index].Cells[0].Value.ToString()))
                    {
                        
                        CheckBox checkBox = (CheckBox)sender;
                        string optionscheck = dataGridView1.Rows[index].Cells[checkBox.Name.Replace("Vi", "options")].Value.ToString();
                        //MessageBox.Show(optionscheck +"-"+checkBox.Name.Replace("Vi", "options"));
                        if (optionscheck.Contains("_"))
                        {
                            if (checkBox.Checked)
                                checkBox.Text = optionscheck.Split('_')[0];
                            else checkBox.Text = optionscheck.Split('_')[1];
                        }
                    }
        }

        private void button92_Click(object sender, EventArgs e)
        {
            if (combo4index == 0)
                txtComboOptions[3] = iOptions4.Text;
            else txtComboOptions[3] = txtComboOptions[3] + "_" + iOptions3.Text;
            combo4index++;
            iOptions4.Text = "";
        }

        private void button96_Click(object sender, EventArgs e)
        {
            if (combo5index == 0)
                txtComboOptions[4] = iOptions5.Text;
            else txtComboOptions[4] = txtComboOptions[4] + "_" + iOptions5.Text;
            combo5index++;
            iOptions5.Text = "";
        }

        

        private void button22_Click(object sender, EventArgs e)
        {
            panellError.Visible = false;
        }

        private void formsBtn_Click(object sender, EventArgs e)
        {
            if (formsBtn.Text != "رفع الاستمارة الأولية")
            {
                OpenFile(formNo, true, formsBtn);
                readyToUpload = true;
            }
            else {
                OpenFileDialog dlg = new OpenFileDialog();
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    CurrentFile = @dlg.FileName;
                    //MessageBox.Show(CurrentFile);
                    uploadFormsReq(CurrentFile);
                    formsBtn.Text = "الاستمارة الأولية";
                }
            }
        }

        private void revisedRow(string text)
        {
            string colName = subTypeAuth.Text.Trim() + "-" + mainTypeAuth.SelectedIndex.ToString();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("update TableAddContext set revised=N'"+ text+"' where ColName=N'" + colName + "'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
            private void btnRevised_Click(object sender, EventArgs e)
        {
            revisedRow("revised");
            dataGridView1.BringToFront();
            FillDataGridView("TableAddContext", AuthType);
            panelLowButtons.Visible = panelIqrar.Visible = panelAuthInfo.Visible = false;
            txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = true;
            flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
            panelLowButtons.Visible = ContextPanel.Visible = true;
        }

        private void mainTypeAuth_TextChanged(object sender, EventArgs e)
        {
            //if (!checkIndex) return;
            //if (mainTypeAuth.Text != "")
            //{
            //    for (int item = 0; item < mainTypeAuth.Items.Count; item++)
            //    {
            //        if (mainTypeAuth.Items[item].ToString() == mainTypeAuth.Text)
            //        {
            //            mainTypeAuth.SelectedIndex = item;
            //            return;
            //        }
            //    }
            //    //MessageBox.Show(نوع_التوكيل.SelectedIndex.ToString());
            //}
        }

        private void subTypeAuth_TextChanged(object sender, EventArgs e)
        {
            //if (!checkIndex) return; 
            //if (subTypeAuth.Text != "")
            //{
            //    for (int item = 0; item < subTypeAuth.Items.Count; item++)
            //    {
            //        if (subTypeAuth.Items[item].ToString() == subTypeAuth.Text)
            //        {
            //            subTypeAuth.SelectedIndex = item;
            //            return;
            //        }
            //    }
            //    //MessageBox.Show(نوع_التوكيل.SelectedIndex.ToString());
            //}
        }

        private void mainTypeIqrar_TextChanged(object sender, EventArgs e)
        {
            //if (!checkIndex) return; 
            //if (mainTypeIqrar.Text != "")
            //{
            //    for (int item = 0; item < mainTypeIqrar.Items.Count; item++)
            //    {
            //        if (mainTypeIqrar.Items[item].ToString() == mainTypeIqrar.Text)
            //        {
            //            mainTypeIqrar.SelectedIndex = item;
            //            return;
            //        }
            //    }
            //    //MessageBox.Show(نوع_التوكيل.SelectedIndex.ToString());
            //}
        }

        private void subTypeIqrar_TextChanged(object sender, EventArgs e)
        {
            //if (!checkIndex) return; 
            //if (subTypeIqrar.Text != "")
            //{
            //    for (int item = 0; item < subTypeIqrar.Items.Count; item++)
            //    {
            //        if (subTypeIqrar.Items[item].ToString() == subTypeIqrar.Text)
            //        {
            //            subTypeIqrar.SelectedIndex = item;
            //            return;
            //        }
            //    }
            //    //MessageBox.Show(نوع_التوكيل.SelectedIndex.ToString());
            //}
        }

        private void remove_Click(object sender, EventArgs e)
        {
            foreach (Control control in panelAuthOptions.Controls)
            {
                if (control is CheckBox)
                {
                    if (!((CheckBox)control).Text.Contains("(محذوف)") && ((CheckBox)control).TabIndex == LastTabIndex)
                    {
                        ((CheckBox)control).Text = ((CheckBox)control).Text + " (محذوف)";
                        remove.Visible = false;
                    }
                }
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            reqGrid = true; 
            if (dataGridView1.Visible)
            {
                txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = false;
                repReqPanel.Visible = flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
                dataGridView1.SendToBack();
                repReqPanel.BringToFront();
                
            }
            else
            {
                panelLowButtons.Visible = panelIqrar.Visible = panelAuthInfo.Visible = false;
                txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = true;
                
                repReqPanel.Visible = true;
                ContextPanel.Visible = false;
                dataGridView1.BringToFront();
                FillDataGridViewReq("TableProcReq");
                formsBtn.Visible = proFileBtn.Visible = btnRevised.Visible = false;
                fileComboBox(المعاملة, DataSource, "المعاملة", "TableProcReq", true);
            }
        }

        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName, bool clear)
        {

            if (clear) combbox.Items.Clear();
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
                        if (dataRow[comlumnName].ToString() != "")
                        {
                            bool found = false;
                            for (int x = 0; x < combbox.Items.Count; x++)
                            {
                                if (combbox.Items[x].ToString() == dataRow[comlumnName].ToString()) found = true;
                            }
                            if (!found) combbox.Items.Add(dataRow[comlumnName].ToString());
                        }
                    }
                }
                catch (Exception ex) { }
                saConn.Close();
            }
        }

        private void uploadReqFiles_Click(object sender, EventArgs e)
        {
            uploadReqFiles.Enabled = false;
            OpenFileDialog dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(@dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;
                button23.Enabled = false;
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


                string[] strData = new string[11];
                SqlConnection sqlCon = new SqlConnection(DataSource);
                try
                {
                    if (sqlCon.State == ConnectionState.Closed)
                        sqlCon.Open();
                }
                catch (Exception ex) { return; }
                rCnt = cCnt = 1;
                for (; rCnt <= 113; rCnt++)
                {
                    //if (string.IsNullOrEmpty((string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2)) break;
                    strData[6] = "غير مدرج";
                    strData[7] = "غير مدرج";
                    strData[8] = "غير مدرج";
                    strData[9] = "غير مدرج";
                    strData[10] = "غير مدرج";

                    for (cCnt = 1; cCnt <= 6; cCnt++)
                    {
                        //if (string.IsNullOrEmpty((string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2)) break;
                        //MessageBox.Show(rCnt.ToString() + ","+ cCnt.ToString() + ","+ Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2));
                        strData[cCnt - 1] = Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    }
                    if (strData[0].Length == 1) strData[0] = "0" + strData[0];
                    insertRow(DataSource, strData);
                    uploadReqFiles.Enabled = true;
                }

                sqlCon.Close();
                xlWorkBook.Close(true, null, null);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
            button23.Enabled = true;
        }

        private void insertRow(string source, string[] data)
        {
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

        private void button25_Click(object sender, EventArgs e)
        {
            string[] data = new string[11];
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
            for (int index = 0; index < 11; index++)
            {
                foreach (Control control in repReqPanel.Controls)
                {
                    if (control.Name == colList[index])
                    {
                        data[index] = control.Text;
                    }
                }
            }
            insertRow(DataSource, data);
            foreach (Control control in repReqPanel.Controls)
            {
                if (control.Name.Contains("المطلوب_رقم") || control.Name.Contains("btnReq"))
                {
                    control.Text = "";
                }
            }
            repReqPanel.SendToBack();
            repReqPanel.Visible = false;
            MessageBox.Show("تمت إضافة المعاملة بنجاح");
            

        }

        private void button26_Click(object sender, EventArgs e)
        {
            string[] data = new string[11];
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
            for (int index = 0; index < 11; index++)
            {
                foreach (Control control in repReqPanel.Controls)
                {
                    if (control.Name == colList[index])
                    {
                        data[index] = control.Text;
                    }
                }
            }
            updatetRow(ProcReqID, DataSource, data);
            foreach (Control control in repReqPanel.Controls)
            {
                if (control.Name.Contains("المطلوب_رقم") || control.Name.Contains("btnReq"))
                {
                    control.Text = "";
                }
            }
        }
        private void updatetRow(int id, string source, string[] data)
        {
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
            string item = "رقم_المعاملة=@رقم_المعاملة";
            for (int col = 1; col < 11; col++)
            {
                item = item + "," + colList[col] + "=@" + colList[col];

            }

            string qurey = "UPDATE TableProcReq SET " + item + " WHERE ID=@ID";

            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", id);
            for (int col = 0; col < 11; col++)
            {
                sqlCmd.Parameters.AddWithValue(colList[col], data[col]);
            }
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string qurey = "delete from TableProcReq WHERE ID=@ID";
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", ProcReqID);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
            foreach (Control control in repReqPanel.Controls)
            {
                if (control.Name.Contains("المطلوب_رقم") || control.Name.Contains("btnReq"))
                {
                    control.Text = "";
                }
            }
        }

        private void button32_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                CurrentFile = @dlg.FileName;
                //MessageBox.Show(CurrentFile);
                uploadFormsReq(CurrentFile);
                panelLowButtons.Visible = panelIqrar.Visible = panelAuthInfo.Visible = false;
                txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = true;

                repReqPanel.Visible = true;
                ContextPanel.Visible = false;
                dataGridView1.BringToFront();
                FillDataGridViewReq("TableProcReq");
                formsBtn.Visible = proFileBtn.Visible = btnRevised.Visible = false;
                fileComboBox(المعاملة, DataSource, "المعاملة", "TableProcReq", true);
            }
        }

        private void reviewForms_Click(object sender, EventArgs e)
        {
            OpenFile(المعاملة.Text, true, reviewForms); 
            reviewForms.Enabled = false;
            readyToUpload = true;
        }

        private void المعاملة_SelectedIndexChanged(object sender, EventArgs e)
        {
            المعاملة.Text = dataGridView1.CurrentRow.Cells["المعاملة"].Value.ToString();
            
            OpenFile(المعاملة.Text, true, reviewForms);
            string[] colList = new string[11];
            colList[0] = "المعاملة";
            colList[1] = "رقم_المعاملة";
            colList[2] = "المطلوب_رقم1";
            colList[3] = "المطلوب_رقم2";
            colList[4] = "المطلوب_رقم3";
            colList[5] = "المطلوب_رقم4";
            colList[6] = "المطلوب_رقم5";
            colList[7] = "المطلوب_رقم6";
            colList[8] = "المطلوب_رقم7";
            colList[9] = "المطلوب_رقم8";
            colList[10] = "المطلوب_رقم9";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * FROM TableProcReq where المعاملة=N'" + المعاملة.Text + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count > 0)
            {
                foreach (DataRow row in dtbl.Rows)
                {
                    ProcReqID = Convert.ToInt32(row["ID"].ToString());
                    for (int index = 1; index < 11; index++)
                    {
                        foreach (Control control in repReqPanel.Controls)
                        {
                            if (control.Name == colList[index])
                            {
                                control.Text = row[colList[index]].ToString();
                            }
                        }
                    }
                }
            }
            txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = false;
            repReqPanel.Visible = true;
            repReqPanel.BringToFront();
        }

        private void button28_Click(object sender, EventArgs e)
        {
            string colName = subTypeAuth.Text.Trim() + "-" + mainTypeAuth.SelectedIndex.ToString();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("update TableAddContext set ColName=N'" + txtProName.Text + "' where ID=N'" + idIndex.ToString() + "'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();

            dataGridView1.BringToFront();
            FillDataGridView("TableAddContext",AuthType);
            panelLowButtons.Visible = panelIqrar.Visible = panelAuthInfo.Visible = false;
            txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = true;
            flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
            ContextPanel.Visible = false;
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            if (txtSearch.Text.Length != 0 && txtSearch.Text.All(char.IsLetterOrDigit))
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '%" + txtSearch.Text + "%'";
                dataGridView1.DataSource = bs;
            }else FillDataGridView("TableAddContext", AuthType);
            ColorFulGrid9();
        }

        private void button33_Click(object sender, EventArgs e)
        {
            repReqPanel.Visible = false;
            repReqPanel.SendToBack();
        }

        private void Vicheck1_CheckedChanged(object sender, EventArgs e)
        {
            if (Vicheck1.Checked)
                Vicheck1.Text = itemsicheck1[0].Split('_')[0];
            else Vicheck1.Text = itemsicheck1[0].Split('_')[1];
        }

        private void TextModel_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                //MessageBox.Show("enter");
                button117.PerformClick();
            }
        }

        private void txtSearch_MouseClick(object sender, MouseEventArgs e)
        {
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void searchReq_TextChanged(object sender, EventArgs e)
        {
            if (searchReq.Text.Length != 0 && searchReq.Text.All(char.IsLetterOrDigit))
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '%" + searchReq.Text + "%'";
                dataGridView1.DataSource = bs;
            }
            else FillDataGridViewReq("TableProcReq");
            ColorFulGrid9();
        }

        private void button116_Click(object sender, EventArgs e)
        {
            ColumnJobs();
            createNewColContext();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand(insertAll, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            MessageBox.Show(revised);
            sqlCmd.Parameters.AddWithValue("@revised", "");
            MessageBox.Show(ColRight.Text);
            sqlCmd.Parameters.AddWithValue("@ColRight", "");            
            sqlCmd.Parameters.AddWithValue("@errorList", "");
            addParameters(sqlCmd);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
            //this.Close();
            var selectedOption = MessageBox.Show("", "إنهاء المراجعة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                btnRevised.PerformClick();

            }
            else
            {
                dataGridView1.BringToFront();
                FillDataGridView("TableAddContext", AuthType);
                panelLowButtons.Visible = panelIqrar.Visible = panelAuthInfo.Visible = false;
                txtSearch.Visible = button32.Visible = label1.Visible = dataGridView1.Visible = true;
                flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
                panelLowButtons.Visible = ContextPanel.Visible = true;
            }
        }

        private void panelDate_MouseEnter(object sender, EventArgs e)
        {
            panelButton.Height = panelCombo.Height = panelCheck.Height = panelText.Height = 42;
            panelDate.Height = pDateHieght;
        }

        private void button88_Click(object sender, EventArgs e)
        {
            if (combo3index == 0)
                txtComboOptions[2] = iOptions3.Text;
            else txtComboOptions[2] = txtComboOptions[2] + "_" + iOptions3.Text;
            combo3index++;
            iOptions3.Text = "";
        }

        private void panelButton_MouseEnter(object sender, EventArgs e)
        {
            panelCombo.Height = panelDate.Height = panelCheck.Height = panelText.Height = 42;
            panelButton.Height = pbuttonHieght;
        }

        private void button77_Click(object sender, EventArgs e)
        {
            if (combo1index == 0)
                txtComboOptions[0] = iOptions1.Text;
            else txtComboOptions[0] = txtComboOptions[0] + "_" + iOptions1.Text;
            combo1index++;
            iOptions1.Text = "";
        }

        private void button118_Click(object sender, EventArgs e)
        {
            if (PanelItemsboxes.Visible)
            {
                PanelItemsboxes.SendToBack();
                PanelItemsboxes.Visible = false;
                btnreviewPanel.Text = "معاينة الحقول";
            }
            else
            {
                btnreviewPanel.Text = "إخفاء القائمة";
                flllPanelItemsboxes("ID", idIndex.ToString());
                PanelItemsboxes.BringToFront();
                PanelItemsboxes.Visible = true;
                PanelItemsboxes.BringToFront();
            }
            //else
            //{
            //    PanelItemsboxes.Visible = false;
            //    PanelItemsboxes.SendToBack();
            //}

            //if (!review1)
            //{
            //    DPTitle[0] = icheckoption11.Text + "_" + icheckoption12.Text;
            //    ExtendedFillBox(itext1.Text, Convert.ToInt32(itxtlenght1.Text), itext2.Text, Convert.ToInt32(itxtlenght2.Text), itext3.Text, Convert.ToInt32(itxtlenght3.Text), itext4.Text, Convert.ToInt32(itxtlenght4.Text), itext5.Text, Convert.ToInt32(itxtlenght5.Text), "", 50, "", 50, "", 50, "", 50, "", 50, icheck1.Text, "", "", "", "", itxtDate1.Text, "", "", "", "", icombo1.Text, txtComboOptions[0].Split('_'), icombo2.Text, txtComboOptions[1].Split('_'), "", Empty, "", Empty, "", Empty, ibtnAdd1.Text, "", "", "", "");
            //}
            //else if (review1)
            //{
            //    for (int x = 0; x < 30; x++)
            //        TextModel.Text = SuffPrefReplacements(TextModel.Text);
            //}
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
            FillDataGridView("TableAddContext", AuthType);
            btnDelete.Visible = btnClear.Visible = false;
        }


        private void ClearFileds()
        {
            restShowingItems();
            mainTypeAuth.Items.Clear();
            subTypeAuth.Items.Clear();
            NewColumn = false;
        }


        private void loadSettings()
        {
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select Modelfilespath,TempOutput,ServerName,Serverlogin,ServerPass,serverDatabase,FileArchive  from TableSettings where ID=@id", Con);
            sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
            if (Con.State == ConnectionState.Closed)
                try
                {
                    Con.Open();

                    var reader = sqlCmd1.ExecuteReader();

                    if (reader.Read())
                    {
                        NewSettings = true;
                        txtModel.Text = reader["Modelfilespath"].ToString();
                        txtOutput.Text = reader["TempOutput"].ToString();
                        txtServerIP.Text = reader["ServerName"].ToString();
                        txtLogin.Text = reader["Serverlogin"].ToString();
                        txtPass.Text = reader["ServerPass"].ToString();
                        txtDatabase.Text = reader["serverDatabase"].ToString();
                        ArchiveFile.Text = reader["FileArchive"].ToString();
                        if (NewSettings)
                        {
                            SaveSettings.Text = "تعديل";
                            NewSettings = false;
                        }
                    }
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    Con.Close();
                }

        }


        private void SaveSettings_Click_1(object sender, EventArgs e)
        {
            if (NewSettings)
            {
                DataSource = "Data Source=" + txtServerIP.Text + ";Network Library=DBMSSOCN;Initial Catalog=" + txtDatabase.Text + ";User ID=" + txtLogin.Text + ";Password=" + txtPass.Text;
                FilepathIn = txtModel.Text;
                FilepathOut = txtOutput.Text;
            }
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    SqlCommand sqlCmd = new SqlCommand("SettingsAddorEdit", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    if (SaveSettings.Text == "حفظ")
                    {
                        sqlCmd.Parameters.AddWithValue("@ID", 1);
                        sqlCmd.Parameters.AddWithValue("@mode", "Add");
                        sqlCmd.Parameters.AddWithValue("@Modelfilespath", txtModel.Text);
                        sqlCmd.Parameters.AddWithValue("@TempOutput", txtOutput.Text);
                        sqlCmd.Parameters.AddWithValue("@ServerName", txtServerIP.Text);
                        sqlCmd.Parameters.AddWithValue("@Serverlogin", txtLogin.Text);
                        sqlCmd.Parameters.AddWithValue("@ServerPass", txtPass.Text);
                        sqlCmd.Parameters.AddWithValue("@serverDatabase", txtDatabase.Text);
                        sqlCmd.Parameters.AddWithValue("@FileArchive", ArchiveFile.Text);
                        sqlCmd.ExecuteNonQuery();
                    }
                    else
                    {
                        sqlCmd.Parameters.AddWithValue("@ID", 1);
                        sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                        sqlCmd.Parameters.AddWithValue("@Modelfilespath", txtModel.Text);
                        sqlCmd.Parameters.AddWithValue("@TempOutput", txtOutput.Text);
                        sqlCmd.Parameters.AddWithValue("@ServerName", txtServerIP.Text);
                        sqlCmd.Parameters.AddWithValue("@Serverlogin", txtLogin.Text);
                        sqlCmd.Parameters.AddWithValue("@ServerPass", txtPass.Text);
                        sqlCmd.Parameters.AddWithValue("@serverDatabase", txtDatabase.Text);
                        sqlCmd.Parameters.AddWithValue("@FileArchive", ArchiveFile.Text);
                        sqlCmd.ExecuteNonQuery();
                    }
                    sqlCon.Close();
                    this.Hide();
                    var formDataBase = new FormDataBase(Server, DataSource56, DataSource57, FilepathIn, FilepathOut, ArchFile, FormDataFile, "");
                    formDataBase.Closed += (s, args) => this.Close();
                    formDataBase.Show();
                }

                catch (Exception ex)
                {
                    MessageBox.Show("الوصول لقاعدة البيانات غير متاح");
                }
                finally
                {

                    clear_fields();

                }
        }

        private void clear_fields()
        {
            txtDatabase.Text = txtLogin.Text = txtModel.Text = txtOutput.Text = txtPass.Text = txtServerIP.Text = ArchiveFile.Text = "";
        }
    }
}
