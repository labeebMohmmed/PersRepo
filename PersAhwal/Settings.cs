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

namespace PersAhwal
{
    public partial class Settings : Form
    {
        private string DataSource56, DataSource57, FilepathIn, FilepathOut, ArchFile, FormDataFile;
        private static bool NewSettings = false;
        string comboBoxOptions1 = "", comboBoxOptions2 = "";
        string[] txtComboOptions = new string[5] { "","","","",""};
        string[] DPTitle = new string[5];
        int pTextHieght = 42;
        int pComboHieght = 42;
        int pCheckHieght = 42;
        int pDateHieght = 42;
        int pbuttonHieght = 42;
        string ColumnName = "";
        bool NewColumn = false;
        string AuthBody1 = "لينوب عني ويقوم مقامي في ";
        int combo1index = 0, combo2index = 0,combo3index = 0, combo4index = 0,combo5index = 0;
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
        string[] errors ;
        string[] editRights ;
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
            fileComboBox(mainTypeAuth, DataSource, "AuthTypes", "TableListCombo");
            //autoCompleteTextBox(newCombAuthType, DataSource, "AuthTypes", "TableListCombo");
            fileComboBox(mainTypeIqrar, DataSource, "ArabicGenIgrar", "TableListCombo");
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
        }

        private string[] getColList(string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
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
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
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
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("alter table "+tableName+" add " + Columnname + " nvarchar(150)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }



        private string getAppFolder()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
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
                dataGridView1.Visible = false;
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
           
            if (dataGridView1.Visible)
            {
                dataGridView1.Visible = false;
                flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
            }
            else
            {
                FillDataGridView("");
                dataGridView1.Visible = true;
                flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
                ContextPanel.Visible = false;
            }
        }


        void FillDataGridView(string text)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableAddContext order by ID desc", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["ColName"].Width = 250;
            sqlCon.Close();
        }

        private void ShowRowNo(int id)
        {
            if (dataGridView1.Rows.Count > 1 && id < dataGridView1.Rows.Count && id >= 0)
            {
                //label1,lenght1,label2,lenght2,label3,lenght3,label4,lenght4,label5,lenght5,1
                //labelcheck,optionscheck,11
                //labelcomb1,optionscombo1,lenghtscombo1,labelcomb2,optionscombo2,lenghtscombo2,13
                //labelbtn,lenghtsbtn,19
                //dateYN,dateType,TextModel,ColRight,ColName 21
                itext1.Text = dataGridView1.Rows[id].Cells[1].Value.ToString();
                itext1Length.Text = dataGridView1.Rows[id].Cells[2].Value.ToString();
                itext2.Text = dataGridView1.Rows[id].Cells[3].Value.ToString();
                itext2Length.Text = dataGridView1.Rows[id].Cells[4].Value.ToString();
                itext3.Text = dataGridView1.Rows[id].Cells[5].Value.ToString();
                itext3Length.Text = dataGridView1.Rows[id].Cells[6].Value.ToString();
                itext4.Text = dataGridView1.Rows[id].Cells[7].Value.ToString();
                itext4Length.Text = dataGridView1.Rows[id].Cells[8].Value.ToString();
                itext5.Text = dataGridView1.Rows[id].Cells[9].Value.ToString();
                itext5Length.Text = dataGridView1.Rows[id].Cells[10].Value.ToString();
                icheck1.Text = dataGridView1.Rows[id].Cells[11].Value.ToString();
                if (dataGridView1.Rows[id].Cells[12].Value.ToString().Trim() == "_") { icheckoption11.Text = icheckoption12.Text = ""; }
                else
                {
                    icheckoption11.Text = dataGridView1.Rows[id].Cells[12].Value.ToString().Split('_')[0];
                    icheckoption12.Text = dataGridView1.Rows[id].Cells[12].Value.ToString().Split('_')[1];
                }
                icombo1.Text = dataGridView1.Rows[id].Cells[13].Value.ToString();
                txtComboOptions[0] = dataGridView1.Rows[id].Cells[14].Value.ToString();
                icombo1Length.Text = dataGridView1.Rows[id].Cells[15].Value.ToString();
                icombo2.Text = dataGridView1.Rows[id].Cells[16].Value.ToString();
                txtComboOptions[1] = dataGridView1.Rows[id].Cells[17].Value.ToString();
                icombo2Length.Text = dataGridView1.Rows[id].Cells[18].Value.ToString();
                ibtnAdd1.Text = dataGridView1.Rows[id].Cells[19].Value.ToString();
                ibtnAdd1Length.Text = dataGridView1.Rows[id].Cells[20].Value.ToString();
                itxtDate1.Text = dataGridView1.Rows[id].Cells[21].Value.ToString();
                dateType1.Text = dataGridView1.Rows[id].Cells[22].Value.ToString();
                TextModel.Text = dataGridView1.Rows[id].Cells[23].Value.ToString();
                ColRight.Text = dataGridView1.Rows[id].Cells[24].Value.ToString();
                ColumnName = dataGridView1.Rows[id].Cells[25].Value.ToString();
                langAuth.Text = dataGridView1.Rows[id].Cells[26].Value.ToString();
                if (langAuth.Text == "" || langAuth.Text == "العربية")
                {
                    langAuth.CheckState = CheckState.Unchecked;
                    langAuth.Text = "العربية";
                    langIqrar.CheckState = CheckState.Unchecked;
                    langIqrar.Text = "العربية";
                }
                else {
                    langAuth.CheckState = CheckState.Checked;
                    langAuth.Text = "الانجليزية";
                    langIqrar.CheckState = CheckState.Checked;
                    langIqrar.Text = "الانجليزية";
                }
                if (ColRight.Text != "")
                {
                    if (ColumnName.Contains("-"))
                    {
                        if (dataGridView1.Rows[id].Cells[25].Value.ToString().All(char.IsDigit))

                            try
                            {
                                mainTypeAuth.SelectedIndex = Convert.ToInt32(dataGridView1.Rows[id].Cells[25].Value.ToString().Split('-')[1]);
                            }
                            catch (Exception exp)
                            {
                            }


                        subTypeAuth.Text = dataGridView1.Rows[id].Cells[25].Value.ToString().Split('-')[0].Replace("_", " ");
                    }
                    panelAuthInfo.Visible = true;
                    panelIqrar.Visible = false;
                }
                else
                {
                    if (ColumnName.Contains("-"))
                    {
                        if (dataGridView1.Rows[id].Cells[25].Value.ToString().All(char.IsDigit))
                            try
                            {
                                mainTypeAuth.SelectedIndex = Convert.ToInt32(dataGridView1.Rows[id].Cells[25].Value.ToString().Split('-')[1]);
                            }
                            catch (Exception exp)
                            {
                            }

                        subTypeAuth.Text = dataGridView1.Rows[id].Cells[25].Value.ToString().Split('-')[0].Replace("_", " ");
                    }
                    panelAuthInfo.Visible = false;
                    panelIqrar.Visible = true;
                }

                dataGridView1.Visible = false;
                ContextPanel.Visible = true;

                review1 = false;
            }
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

                string query = "select max(ID) from "+ tableName+" as maxID where "+ comlumnName+" is not null";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;


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

        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string tableName)
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

        private void ExtendedFillBox(string v1, int s1, string v2, int s2, string v3, int s3, string v4, int s4, string v5, int s5, string vE1, int sE1, string vE2, int sE2, string vE3, int sE3, string vE4, int sE4, string vE5, int sE5, string vE61, string vE62, string vE63, string vE64, string vE65, string vE71, string vE72, string vE73, string vE74, string vE75, string v81, string[] vE81, string v82, string[] vE82, string v83, string[] vE83, string v84, string[] vE84, string v85, string[] vE85, string button1, string button2, string button3, string button4, string button5)
        {
            restShowingItems();
            if (v1 != "")
            {
                Litext1.Text = v1;
                Litext1.Visible = true;

                if (s1 < 700)
                    Vitext1.Width = s1;
                else
                {
                    Vitext1.Multiline = true;
                    Vitext1.Size = new System.Drawing.Size(s1, 146);
                }
                Vitext1.Visible = true;
            }
            if (v2 != "")
            {
                Litext2.Text = v2;
                Litext2.Visible = true;
                if (s2 < 700)
                    Vitext2.Width = s2;
                else
                {
                    Vitext2.Multiline = true;
                    Vitext2.Size = new System.Drawing.Size(s2, 146);
                }
                Vitext2.Visible = true;
            }
            if (v3 != "")
            {
                Litext3.Text = v3;
                Litext3.Visible = true;
                if (s3 < 700)
                    Vitext3.Width = s3;
                else
                {
                    Vitext3.Multiline = true;
                    Vitext3.Size = new System.Drawing.Size(s3, 146);
                }
                Vitext3.Visible = true;
            }
            if (v4 != "")
            {
                Litext4.Text = v4;
                Litext4.Visible = true;
                if (s4 < 700)
                    Vitext4.Width = s4;
                else
                {
                    Vitext4.Multiline = true;
                    Vitext4.Size = new System.Drawing.Size(s4, 146);
                }
                Vitext4.Visible = true;
            }
            if (v5 != "")
            {
                Litext5.Text = v5;
                Litext5.Visible = true;
                if (s5 < 700)
                    Vitext5.Width = s5;
                else
                {
                    Vitext5.Multiline = true;
                    Vitext5.Size = new System.Drawing.Size(s5, 146);
                }
                Vitext5.Visible = true;
            }



            if (vE1 != "")
            {
                Litext6.Text = vE1;
                Litext6.Visible = true;
                Vitext6.Width = sE1;
                Vitext6.Visible = true;
            }
            if (vE2 != "")
            {
                Litext7.Text = vE2;
                Litext7.Visible = true;
                Vitext7.Width = sE2;
                Vitext7.Visible = true;
            }
            if (vE3 != "")
            {
                Litext8.Text = vE3;
                Litext8.Visible = true;
                Vitext8.Width = sE3;
                Vitext8.Visible = true;
            }
            if (vE4 != "")
            {
                Litext9.Text = vE4;
                Litext9.Visible = true;
                Vitext9.Width = sE4;
                Vitext9.Visible = true;
            }
            //if (vE5 != "")
            //{
            //    labeltxt10.Text = vE5;
            //    labeltxt10.Visible = true;
            //    txt10.Width = sE5;
            //    txt10.Visible = true;
            //}


            if (vE61 != "")
            {
                Licheck1.Text = vE61;
                Licheck1.Visible = true;
                if (DPTitle[0].Contains("_")) Vicheck1.Text = DPTitle[0].Split('_')[0];
                else Vicheck1.Text = DPTitle[0];
                Vicheck1.Visible = true;
            }
            if (vE62 != "")
            {
                Licheck2.Text = vE62;
                Licheck2.Visible = true;
                if (DPTitle[1].Contains("_")) Vicheck2.Text = DPTitle[1].Split('_')[0];
                else Vicheck2.Text = DPTitle[1];
                Vicheck3.Visible = true;
            }
            if (vE63 != "")
            {
                Licheck3.Text = vE63;
                Licheck3.Visible = true;
                if (DPTitle[2].Contains("_")) Vicheck3.Text = DPTitle[2].Split('_')[0];
                else Vicheck3.Text = DPTitle[2];
                Vicheck3.Visible = true;
            }
            if (vE64 != "")
            {
                Licheck4.Text = vE64;
                Licheck4.Visible = true;
                if (DPTitle[3].Contains("_")) Vicheck4.Text = DPTitle[3].Split('_')[0];
                else Vicheck4.Text = DPTitle[3];
                Vicheck4.Visible = true;
            }
            if (vE65 != "")
            {
                Licheck5.Text = vE65;
                Licheck5.Visible = true;
                if (DPTitle[4].Contains("_")) Vicheck5.Text = DPTitle[4].Split('_')[0];
                else Vicheck5.Text = DPTitle[4];
                Vicheck5.Visible = true;
            }

            if (vE71 != "")
            {
                LitxtDate1.Text = vE71;
                LitxtDate1.Visible = true;
                VitxtDate1LD.Visible = true;
                VitxtDate1VD.Visible = true;
                VitxtDate1LM.Visible = true;
                VitxtDate1VM.Visible = true;
                VitxtDate1LY.Visible = true;
                VitxtDate1VY.Visible = true;
            }

            if (vE72 != "")
            {
                LitxtDate2.Text = vE71;
                LitxtDate2.Visible = true;
                VitxtDate2LD.Visible = true;
                VitxtDate2VD.Visible = true;
                VitxtDate2LM.Visible = true;
                VitxtDate2VM.Visible = true;
                VitxtDate2LY.Visible = true;
                VitxtDate2VY.Visible = true;
            }
            if (vE73 != "")
            {
                LitxtDate3.Text = vE71;
                LitxtDate3.Visible = true;
                VitxtDate3LD.Visible = true;
                VitxtDate3VD.Visible = true;
                VitxtDate3LM.Visible = true;
                VitxtDate3VM.Visible = true;
                VitxtDate3LY.Visible = true;
                VitxtDate3VY.Visible = true;
            }
            if (vE74 != "")
            {
                LitxtDate4.Text = vE71;
                LitxtDate4.Visible = true;
                VitxtDate4LD.Visible = true;
                VitxtDate4VD.Visible = true;
                VitxtDate4LM.Visible = true;
                VitxtDate4VM.Visible = true;
                VitxtDate4LY.Visible = true;
                VitxtDate4VY.Visible = true;
            }
            if (vE75 != "")
            {
                LitxtDate5.Text = vE71;
                LitxtDate5.Visible = true;
                VitxtDate5LD.Visible = true;
                VitxtDate5VD.Visible = true;
                lalM5.Visible = true;
                VitxtDate5VM.Visible = true;
                VitxtDate5LY.Visible = true;
                VitxtDate5VY.Visible = true;
            }

            if (v81 != "")
            {
                Licombo1.Visible = true;
                Vicombo1.Visible = true;
                Licombo1.Text = v81;

                Vicombo1.Items.Clear();
                for (int x = 0; x < vE81.Length; x++)
                    Vicombo1.Items.Add(vE81[x]);
                Vicombo1.SelectedIndex = 0;
            }

            if (v82 != "")
            {
                Licombo2.Visible = true;
                Vicombo1.Visible = true;
                Licombo2.Text = v81;

                Vicombo1.Items.Clear();
                for (int x = 0; x < vE82.Length; x++)
                    Vicombo1.Items.Add(vE82[x]);
                Vicombo1.SelectedIndex = 0;
            }

            if (v83 != "")
            {
                Licombo3.Visible = true;
                Vicombo1.Visible = true;
                Licombo3.Text = v81;

                Vicombo1.Items.Clear();
                for (int x = 0; x < vE83.Length; x++)
                    Vicombo1.Items.Add(vE83[x]);
                Vicombo1.SelectedIndex = 0;
            }

            if (v84 != "")
            {
                Licombo4.Visible = true;
                Vicombo1.Visible = true;
                Licombo4.Text = v81;

                Vicombo1.Items.Clear();
                for (int x = 0; x < vE84.Length; x++)
                    Vicombo1.Items.Add(vE84[x]);
                Vicombo1.SelectedIndex = 0;
            }

            if (v85 != "")
            {
                Licombo5.Visible = true;
                Vicombo1.Visible = true;
                Licombo5.Text = v81;

                Vicombo1.Items.Clear();
                for (int x = 0; x < vE85.Length; x++)
                    Vicombo1.Items.Add(vE85[x]);
                Vicombo1.SelectedIndex = 0;
            }
            if (button1 != "")
            {
                LibtnAdd1.Text = button1;
                LibtnAdd1.Visible = true;
            }
            if (button2 != "")
            {
                LibtnAdd2.Text = button2;
                LibtnAdd2.Visible = true;
            }
            if (button3 != "")
            {
                LibtnAdd3.Text = button3;
                LibtnAdd3.Visible = true;
            }
            if (button4 != "")
            {
                LibtnAdd4.Text = button4;
                LibtnAdd4.Visible = true;
            }
            if (button5 != "")
            {
                LibtnAdd5.Text = button5;
                LibtnAdd5.Visible = true;
            }
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





            if (checkColumnName(mainTypeAuth.Text.Replace(" ", "_")))
            {
                subTypeAuth.Items.Clear();
                fileComboBox(subTypeAuth, DataSource, mainTypeAuth.Text.Replace(" ", "_"), "TableListCombo");
                NewColumn = false;
                return;
            }
            NewColumn = true;
            if (mainTypeAuth.SelectedIndex >= 2 && mainTypeAuth.SelectedIndex <= 5)
            {
                subTypeAuth.Items.Clear();
                subTypeAuth.Items.Add("إجراء جديد");
                ColumnName = "Row1Attach";
                fileComboBox(subTypeAuth, DataSource, ColumnName, "TableListCombo");
            }
            if (mainTypeAuth.Text.Contains("زواج"))
            {
                subTypeAuth.Items.Clear();
                subTypeAuth.Items.Add("إجراء جديد");
                ColumnName = "RowMerrageAttach";

                fileComboBox(subTypeAuth, DataSource, "RowMerrageAttach", "TableListCombo");
            }
            if (mainTypeAuth.Text.Contains("ورثة"))
            {
                subTypeAuth.Items.Clear();
                subTypeAuth.Items.Add("إجراء جديد");
                ColumnName = "RowLegacyAttach";
                fileComboBox(subTypeAuth, DataSource, "RowLegacyAttach", "TableListCombo");
            }
            if (mainTypeAuth.Text.Contains("سيارة"))
            {
                subTypeAuth.Items.Clear();
                subTypeAuth.Items.Add("إجراء جديد");
                ColumnName = "RowCarAttach";
                fileComboBox(subTypeAuth, DataSource, "RowCarAttach", "TableListCombo");
            }
            if (mainTypeAuth.Text.Contains("طلاق"))
            {
                subTypeAuth.Items.Clear();
                subTypeAuth.Items.Add("إجراء جديد");
                ColumnName = "RowDeforceAttach";
                fileComboBox(subTypeAuth, DataSource, "RowDeforceAttach", "TableListCombo");
            }
            if (mainTypeAuth.Text.Contains("جامعية"))
            {
                subTypeAuth.Items.Clear();
                subTypeAuth.Items.Add("إجراء جديد");
                ColumnName = "RowUniversityAttach";
                fileComboBox(subTypeAuth, DataSource, "RowUniversityAttach", "TableListCombo");
            }
            if (mainTypeAuth.Text.Contains("ميلاد"))
            {
                subTypeAuth.Items.Clear();
                subTypeAuth.Items.Add("إجراء جديد");
                subTypeAuth.Items.Add("استخراج وتوثيق");
                ColumnName = "";
                subTypeAuth.Items.Add(" استخراج وتوثيق بدل فاقد");
            }
            if (mainTypeAuth.Text.Contains("بالتنازل"))
            {
                subTypeAuth.Items.Clear();
                subTypeAuth.Items.Add("إجراء جديد");
                ColumnName = "GiveAway";
                fileComboBox(subTypeAuth, DataSource, "GiveAway", "TableListCombo");
            }

            if (mainTypeAuth.Text.Contains("تأمين"))
            {
                subTypeAuth.Items.Clear();
                subTypeAuth.Items.Add("إجراء جديد");
                subTypeAuth.Items.Add("استلام تأمين");
                ColumnName = "";
            }
            //if (ComboProcedure.Items.Count > 0) ComboProcedure.SelectedIndex = 0;
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
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableListCombo (AuthTypes) values (@AuthTypes)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@AuthTypes", colText);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
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

        private void getColumnName(string colNo)
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
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableAuthRight", sqlCon);
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
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
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
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("update TableSettings set FolderApp=@FolderApp where ID='1'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@FolderApp", folderApp);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }


        
                        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows.Count > 1)
            {
                idIndex = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                langAuth.Text = dataGridView1.CurrentRow.Cells["Lang"].Value.ToString();
                ColRight.Text = dataGridView1.CurrentRow.Cells["ColRight"].Value.ToString();
                TextModel.Text = dataGridView1.CurrentRow.Cells["TextModel"].Value.ToString();
                editRights = dataGridView1.CurrentRow.Cells["editRights"].Value.ToString().Split('،');
                string error = dataGridView1.CurrentRow.Cells["errorList"].Value.ToString();
                
                if (error != "")
                {
                    panellError.Visible = true;
                    panellError.BringToFront();
                    errors = error.Split('_');
                    //MessageBox.Show(error);MessageBox.Show(errors[0]);
                    error1.Checked = Convert.ToBoolean(errors[0]);
                    error2.Checked = Convert.ToBoolean(errors[1]);
                    error3.Checked = Convert.ToBoolean(errors[2]);
                    error4.Checked = Convert.ToBoolean(errors[3]);
                    error5.Checked = Convert.ToBoolean(errors[4]);
                    if (errors[5] == "False")
                        labelError.Visible = otherError.Visible = error6.Checked = false;
                    else
                    {
                        otherError.Visible = labelError.Visible = true;
                        error6.Checked = true;
                        otherError.Text = errors[5];
                    }
                    button22.Visible = true; 
                    Nobox = 0;
                    foreach (string str in editRights)
                    {
                        if (str != "")
                        {
                            //{
                            CheckBox chk = new CheckBox();
                            chk.TabIndex = Nobox;
                            chk.Width = 80;
                            chk.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                            if (Nobox == 0) chk.Width = panelAuthOptions.Width - 100;
                            else chk.Width = panelAuthOptions.Width - 130;
                            chk.Height = 33;
                            chk.Location = new System.Drawing.Point(70, 3 + Nobox * 37);
                            chk.Name = "checkBox" + Nobox.ToString();
                            try
                            {
                                chk.Text = str.Split('_')[1] + "،";
                                chk.Tag = "valid";                                
                                if (str.Split('_')[0] == "1")
                                    chk.CheckState = CheckState.Checked;
                                else chk.CheckState = CheckState.Unchecked;
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show(str);
                            }
                            Nobox++;
                            panellError.Controls.Add(chk);
                        }
                    }
                }

                foreach (Control control in panelText.Controls)
                {
                    for (int index = 0; index < allList.Length; index++)
                    {
                        if (allList[index] == control.Name)
                        {
                            control.Text = dataGridView1.CurrentRow.Cells[allList[index]].Value.ToString();
                        }
                    }
                }
                foreach (Control control in panelDate.Controls)
                {
                    for (int index = 0; index < allList.Length; index++)
                    {
                        
                        if (allList[index] == control.Name)
                        {
                            control.Text = dataGridView1.CurrentRow.Cells[allList[index]].Value.ToString();
                            
                        }
                    }
                }
                foreach (Control control in panelCombo.Controls)
                {
                    for (int index = 0; index < allList.Length; index++)
                    {
                        if (allList[index] == control.Name)
                        {
                            control.Text = dataGridView1.CurrentRow.Cells[allList[index]].Value.ToString();
                        }
                    }
                }
                foreach (Control control in panelButton.Controls)
                {
                    for (int index = 0; index < allList.Length; index++)
                    {
                        if (allList[index] == control.Name)
                        {
                            control.Text = dataGridView1.CurrentRow.Cells[allList[index]].Value.ToString();
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
                                if (dataGridView1.CurrentRow.Cells["optionscheck1"].Value.ToString().Trim() == "_")
                                {
                                    icheckoption11.Text = icheckoption12.Text = "";
                                }
                                else if(dataGridView1.CurrentRow.Cells["optionscheck1"].Value.ToString().Contains("_"))
                                {
                                    icheckoption11.Text = dataGridView1.CurrentRow.Cells["optionscheck1"].Value.ToString().Split('_')[0];
                                    icheckoption12.Text = dataGridView1.CurrentRow.Cells["optionscheck1"].Value.ToString().Split('_')[1];
                                }
                            }
                            else control.Text = dataGridView1.CurrentRow.Cells[allList[index]].Value.ToString();
                        }
                    }
                }
                for (int index = 0; index < allList.Length; index++)
                    for (int x = 1; x < 6; x++)
                    {
                        if (allList[index] == "icomboOption" + x.ToString())
                        {
                            txtComboOptions[x - 1] = dataGridView1.CurrentRow.Cells[allList[index]].Value.ToString();
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


                //label1,lenght1,label2,lenght2,label3,lenght3,label4,lenght4,label5,lenght5,labelcheck,optionscheck,labelcomb1,optionscombo1,lenghtscombo1     ,labelcomb2,optionscombo2,lenghtscombo2,labelbtn,lenghtsbtn,dateYN,dateType,TextModel,ColRight,ColName }

                //itext1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();

                //itxtlenght1.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                //itext2.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                //itxtlenght2.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                //itext3.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                //itxtlenght3.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                //itext4.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                //itxtlenght4.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                //itext5.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                //itxtlenght5.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                //icheck1.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                //if (dataGridView1.CurrentRow.Cells[12].Value.ToString().Trim() == "_")
                //{
                //    icheckoption11.Text = icheckoption12.Text = "";
                //}
                //else
                //{
                //    icheckoption11.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString().Split('_')[0];
                //    icheckoption12.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString().Split('_')[1];
                //}
                //icombo1.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                //txtComboOptions1 = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                //icomboLength1.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                //icombo2.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();
                //txtComboOptions2 = dataGridView1.CurrentRow.Cells[17].Value.ToString();
                //icomboLength2.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
                //ibtnAdd1.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
                //buttonLength1.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();
                //itxtDate1.Text = dataGridView1.CurrentRow.Cells[21].Value.ToString();
                //dateType1.Text = dataGridView1.CurrentRow.Cells[22].Value.ToString();
                //txtModelData.Text = dataGridView1.CurrentRow.Cells[23].Value.ToString();
                //comboRights.Text = dataGridView1.CurrentRow.Cells[24].Value.ToString();

                //language.Text = dataGridView1.CurrentRow.Cells[26].Value.ToString();
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
                    string strColName = ColumnName = dataGridView1.CurrentRow.Cells["ColName"].Value.ToString();


                    if (ColumnName.Contains("-"))
                    {
                        if (ColumnName.Split('-')[1].All(char.IsDigit))
                            try
                            {
                                mainTypeAuth.SelectedIndex = Convert.ToInt32(ColumnName.Split('-')[1]);
                            }
                            catch (Exception exp)
                            {
                            }
                        //MessageBox.Show(strColName);
                        int x = 0;
                        //MessageBox.Show(ComboProcedure.Items.Count.ToString());
                        for (; x < subTypeAuth.Items.Count; x++)
                        {
                            //MessageBox.Show(x.ToString());
                            if (subTypeAuth.Items[x].ToString().Trim() == strColName.Split('-')[0].Trim())
                            {
                                //MessageBox.Show(ComboProcedure.Items[x].ToString().Trim() +" -- "+ strColName.Split('-')[0].Trim());
                                subTypeAuth.SelectedIndex = x;
                                break;
                            }
                        }

                        //ComboProcedure.Text = ColumnName.Split('-')[0];
                    }
                    panelAuthInfo.Visible = true;
                    panelIqrar.Visible = false;
                }
                else
                {
                    ColumnName = dataGridView1.CurrentRow.Cells[25].Value.ToString();
                    if (ColumnName.Contains("-"))
                    {
                        if (ColumnName.Split('-')[1].All(char.IsDigit))
                            try
                            {
                                mainTypeIqrar.SelectedIndex = Convert.ToInt32(ColumnName.Split('-')[1]);
                            }
                            catch (Exception exp)
                            {
                            }

                        subTypeIqrar.Text = ColumnName.Split('-')[0];
                    }
                    panelAuthInfo.Visible = false;
                    panelIqrar.Visible = true;
                }
                dataGridView1.Visible = false;
                ContextPanel.Visible = true;
                button117.Enabled = true;
                //review1 = true;
                review1 = false;
                btnDelete.Visible = btnClear.Visible = true;
                DPTitle[0] = icheckoption11.Text + "_" + icheckoption12.Text;
                //ExtendedFillBox(itext1.Text, Convert.ToInt32(itxtlenght1.Text), itext2.Text, Convert.ToInt32(itxtlenght2.Text), itext3.Text, Convert.ToInt32(itxtlenght3.Text), itext4.Text, Convert.ToInt32(itxtlenght4.Text), itext5.Text, Convert.ToInt32(itxtlenght5.Text), "", 50, "", 50, "", 50, "", 50, "", 50, icheck1.Text, "", "", "", "", itxtDate1.Text, "", "", "", "", icombo1.Text, txtComboOptions[0].Split('_'), icombo2.Text, txtComboOptions[1].Split('_'), "", Empty, "", Empty, "", Empty, ibtnAdd1.Text, "", "", "", "");
            }
        }

        private void checkSexType_CheckedChanged_2(object sender, EventArgs e)
        {

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

        private string SuffPrefReplacements(string text)
        {
            Suffex_preffixList();
            //if (text.Contains("tN"))
            //    return text.Replace("tN", ApplicantName.Text);
            //if (text.Contains("tI"))
            //    return text.Replace("tP", DocNo.Text);
            //if (text.Contains("tS"))
            //    return text.Replace("tS", DocSource.Text);
            //if (text.Contains("tSS"))
            //{
            //    if (ApplicantSex.Text == "ذكر") return "";
            //    return "ة";
            //}
            //if (text.Contains("tT"))
            //    return text.Replace("tT", titleEng.Text);
            //if (text.Contains("tD"))
            //    return text.Replace("tD", DocType.Text);

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
                return text.Replace("c1", icheck1.Text);

            if (text.Contains("m1"))
                return text.Replace("m1", Vicombo1.Text);
            if (text.Contains("m2"))
                return text.Replace("m2", Vicombo2.Text);

            if (text.Contains("a1"))
                return text.Replace("a1", ibtnAdd1.Text);

            if (text.Contains("n1"))
                return text.Replace("n1", " " + VitxtDate1VD.Text + "/" + VitxtDate1VM.Text + "/" + VitxtDate1VY.Text + " ");
            if (text.Contains("#*#"))
                return text.Replace("#*#", preffix[0, 10]);

            if (text.Contains("#3"))
                return text.Replace("#3", preffix[0, 7]);
            if (text.Contains("#4"))
                return text.Replace("#4", preffix[0, 8]);
            if (text.Contains("#5"))
                return text.Replace("#5", preffix[0, 9]);

            if (text.Contains("#1"))
                return text.Replace("#1", preffix[0, 11]);
            if (text.Contains("#2"))
                return text.Replace("#2", preffix[0, 12]);

            if (text.Contains("@*@"))
                return text.Replace("@*@", "لدى  برقم الايبان ()");
            if (text.Contains("$$$"))
                return text.Replace("$$$", preffix[0, 0]);
            if (text.Contains("&&&"))
                return text.Replace("&&&", preffix[0, 1]);
            if (text.Contains("^^^"))
                return text.Replace("^^^", preffix[0, 2]);
            if (text.Contains("###"))
                return text.Replace("###", preffix[0, 4]);
            if (text.Contains("***"))
                return text.Replace("***", preffix[0, 3]);
            else return text;
        }




        private void Suffex_preffixList()
        {

            preffix[0, 0] = "ي"; //$$$ "ي/نا";
            preffix[1, 0] = "ي";
            preffix[2, 0] = "نا";
            preffix[3, 0] = "نا";
            preffix[4, 0] = "نا";
            preffix[5, 0] = "نا";

            preffix[0, 1] = "ت";//&&& "ت/نا";
            preffix[1, 1] = "ت";
            preffix[2, 1] = "نا";
            preffix[3, 1] = "نا";
            preffix[4, 1] = "نا";
            preffix[5, 1] = "نا";

            preffix[0, 2] = "ني";//^^^ "ني/نا";
            preffix[1, 2] = "ني";
            preffix[2, 2] = "نا";
            preffix[3, 2] = "نا";
            preffix[4, 2] = "نا";
            preffix[5, 2] = "نا";

            preffix[0, 3] = "";//*** "/ت/ا/تا/ن/وا                               
            preffix[1, 3] = "ت";
            preffix[2, 3] = "ا";
            preffix[3, 3] = "تا";
            preffix[4, 3] = "ن";
            preffix[5, 3] = "وا";

            preffix[0, 4] = "ه";//### "ه/ها/هما/هما/من/هم"
            preffix[1, 4] = "ها";
            preffix[2, 4] = "هما";
            preffix[3, 4] = "هما";
            preffix[4, 4] = "هن";
            preffix[5, 4] = "هم";

            preffix[0, 5] = ""; //
            preffix[1, 5] = "ة";
            preffix[2, 5] = "ان";
            preffix[3, 5] = "تان";
            preffix[4, 5] = "ات";
            preffix[5, 5] = "ون";

            preffix[0, 6] = "";
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

            preffix[0, 9] = "نصيبي";
            preffix[1, 9] = "نصيبي";
            preffix[2, 9] = "نصيبينا";
            preffix[3, 9] = "نصيبينا";
            preffix[4, 9] = "أنصبتنا";
            preffix[5, 9] = "أنصبتنا";

            preffix[0, 10] = "ت";//#*#
            preffix[1, 10] = "";

            preffix[0, 11] = "التي";//#1
            preffix[1, 11] = "الذي";

            preffix[0, 12] = "هو";//#2
            preffix[1, 12] = "هي";
            preffix[2, 12] = "هما";
            preffix[3, 12] = "هما";
            preffix[4, 12] = "هن";
            preffix[5, 12] = "هم";
        }


        public void PopulateCheckBoxes(string col, string table, string dataSource)
        {
            if (col == "" || table == "" || dataSource == "") return;
            string query = "SELECT ID," + col + " FROM " + table;
            using (SqlConnection con = new SqlConnection(dataSource))
            {

                using (SqlDataAdapter sda = new SqlDataAdapter(query, con))
                {

                    sda.Fill(checkboxdt);
                    listchecked = checkboxdt.Rows.Count;
                    Nobox = 0;
                    foreach (DataRow row in checkboxdt.Rows)
                    {
                        if (checkboxdt.Rows[Nobox][col].ToString() == "" || checkboxdt.Rows[Nobox][col].ToString() == "null") return;
                        //{
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
                        Text_statis = checkboxdt.Rows[Nobox][col].ToString().Split('_');
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
                        if (chk.Text.Contains("لمن يشهد والله خير الشاهدين")) picboxup.Visible = false;
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
                        if (chk.Text.Contains("الحق في توكيل الغير") || chk.Text.Contains("لمن يشهد والله خير الشاهدين")) picboxdown.Visible = false;

                        panelAuthOptions.Controls.Add(picboxdown);
                        LastID = Convert.ToInt32(checkboxdt.Rows[Nobox]["ID"].ToString());
                        Nobox++;
                        //}
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
            authType.Text = "نموذج صيغة التفويض:";
            TextModel.Text = "لينوب عني ويقوم مقامي في ";

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
            if (datatype) qurey = "INSERT INTO TableAuthRight (" + comlumnName + ") values(" + column + ")";
            else qurey = "UPDATE TableAuthRight SET " + comlumnName + " = " + column + " WHERE ID = @ID";

            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
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
        }



        private void button40_Click(object sender, EventArgs e)
        {

            if (ContextPanel.Visible)
            {
                ContextPanel.Visible = false;

            }
            else
            {
                dataGridView1.Visible = false;
                ContextPanel.Visible = false;
                ContextPanel.Visible = true;
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
                subTypeAuth.Items.Clear();
                newFillComboBox1(subTypeAuth, DataSource, mainTypeAuth.SelectedIndex.ToString(), langAuth.Text);
                //fileComboBox(subTypeAuth, DataSource, mainTypeAuth.Text.Replace(" ", "_"), "TableListCombo");
                NewColumn = false;
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
            for (int i=0; i < dataGridView1.Rows.Count - 1; i++)
            {
                //

                if (dataGridView1.Rows[i].Cells["errorList"].Value.ToString() != "")
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightCyan;
                    errorList++;
                }                
            }
            if (errorList > 0) labelarch.Visible = true;
            else labelarch.Visible = false;
            labelarch.Text = "عدد (" + errorList.ToString() + ") معاملة تحتاج إلى تصحيح ";

        }

        private void ComboProcedure_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            //if (!NewColumn) return;
            FillDataGridView("");
            flllPanelItemsboxes("ColName", subTypeAuth.Text + "-" + mainTypeAuth.SelectedIndex.ToString());
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

        private void button117_Click(object sender, EventArgs e)
        {
            ColumnJobs();
            createNewColContext();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(updateAll, sqlCon);

            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", idIndex);
            sqlCmd.Parameters.AddWithValue("@errorList", "");
            addParameters(sqlCmd);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
            this.Close();
            //ClearFileds();
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            mainTypeIqrar.Items.Clear();
            if (langIqrar.CheckState == CheckState.Checked)

            {
                langIqrar.Text = "الانجليزية";
                fileComboBox(mainTypeIqrar, DataSource, "EnglishGenIgrar", "TableListCombo");
                newFillComboBox2(subTypeIqrar, DataSource, mainTypeIqrar.SelectedIndex.ToString(), langIqrar.Text);
                PanelItemsboxes.RightToLeft = TextModel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            }
            else
            {
                langIqrar.Text = "العربية";
                fileComboBox(mainTypeIqrar, DataSource, "ArabicGenIgrar", "TableListCombo");
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
            TextModel.Text = "";
            deleteItemsAO(); RightColumnName = ColRight.Text;
            PopulateCheckBoxes(ColRight.Text.Trim(), "TableAuthRight", DataSource);
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
                int id = getLastID(DataSource, mainTypeAuth.Text.Replace(" ", "_"), "TableListCombo");
                addSubAuth(id, subTypeAuth.Text, mainTypeAuth.Text.Replace(" ", "_"));                
            }
            else
            {                
                if (!checkColExist("TableListCombo", mainTypeIqrar.Text.Replace(" ", "_")))
                {
                    CreateColumn(mainTypeIqrar.Text.Replace(" ", "_"), "TableListCombo");
                    mainTypeIqrar.Items.Add(mainTypeIqrar.Text);
                }
                ColumnName = subTypeIqrar.Text.Replace(" ", "_") + "-" + mainTypeIqrar.SelectedIndex.ToString();
                int id = getLastID(DataSource, mainTypeIqrar.Text.Replace(" ", "_"), "TableListCombo");
                addSubAuth(id, subTypeIqrar.Text, mainTypeIqrar.Text.Replace(" ", "_"));
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
                if (control is CheckBox)
                {
                    rights[rightIndex] = ((CheckBox)control).Text;
                    rightIndex++;
                }

            }

            for (int x = 0; x < rightIndex; x++)
            {

                if (rights[x] == "" || rights[x] == "Null") break;
                UpdateColumn(DataSource, ColRight.Text.Trim(), x + 1, rights[x]);
            }
            MessageBox.Show("تم تعديل القائمة");
            button37.PerformClick();
        }
        private void UpdateColumn(string source, string comlumnName, int id, string data)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string column = "@" + comlumnName;
            string qurey = "UPDATE TableAuthRight SET " + comlumnName + " = " + column + " WHERE ID=@ID";

            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
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


            //fileComboBox(ProFormType, DataSource, FormType.Text.Replace(" ", "_"), "TableListCombo");
            newFillComboBox2(subTypeIqrar, DataSource, mainTypeIqrar.SelectedIndex.ToString(), langIqrar.Text);
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
            FillDataGridView("");
            for (int id = 0; id < dataGridView1.Rows.Count - 1; id++)
            {
                if (dataGridView1.Rows[id].Cells[25].Value.ToString() == subTypeIqrar.Text + "-" + mainTypeIqrar.SelectedIndex.ToString())
                {
                    ShowRowNo(id);
                    review1 = true;
                    DPTitle[0] = icheckoption11.Text + "_" + icheckoption12.Text;
                    ExtendedFillBox(itext1.Text, Convert.ToInt32(itext1Length.Text), itext2.Text, Convert.ToInt32(itext2Length.Text), itext3.Text, Convert.ToInt32(itext3Length.Text), itext4.Text, Convert.ToInt32(itext4Length.Text), itext5.Text, Convert.ToInt32(itext5Length.Text), "", 50, "", 50, "", 50, "", 50, "", 50, icheck1.Text, "", "", "", "", itxtDate1.Text, "", "", "", "", icombo1.Text, txtComboOptions[0].Split('_'), icombo2.Text, txtComboOptions[1].Split('_'), "", Empty, "", Empty, "", Empty, ibtnAdd1.Text, "", "", "", "");
                    idIndex = Convert.ToInt32(dataGridView1.Rows[id].Cells[0].Value.ToString());
                    button117.Enabled = button109.Enabled = true;
                }
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Visible)
            {
                dataGridView1.Visible = false;
                flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
            }
            else
            {
                dataGridView1.Visible = true;
                flowLayoutPanel9.Visible = SettingsPanel.Visible = false;
                ContextPanel.Visible = false;
            }
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
                        if (allList[index] == "icomboOption" + x.ToString())
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

            if (panelAuthInfo.Visible)
            {
                sqlCmd.Parameters.AddWithValue("@authIqrar", "auth");
                sqlCmd.Parameters.AddWithValue("@Lang", langAuth.Text);
                sqlCmd.Parameters.AddWithValue("@ColRight", ColRight.Text);
                sqlCmd.Parameters.AddWithValue("@ColName", (subTypeAuth.Text + "-" + CombAuthTypeIndex.ToString()).Replace("--", "-"));
            }
            else
            {
                sqlCmd.Parameters.AddWithValue("@authIqrar", "iqrar");
                sqlCmd.Parameters.AddWithValue("@Lang", langIqrar.Text);
                sqlCmd.Parameters.AddWithValue("@ColRight", "");
                sqlCmd.Parameters.AddWithValue("@ColName", (subTypeIqrar.Text + "-" + mainTypeIqrar.SelectedIndex.ToString()).Replace("--", "-"));
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
                            CreateColumn("icomboOption1", "TableAddContext");
                        }
                        else if (control.Name.Contains("2"))
                        {
                            CreateColumn("icombo2Length", "TableAddContext");
                            CreateColumn("icomboOption2", "TableAddContext");
                        }
                        else if (control.Name.Contains("3"))
                        {
                            CreateColumn("icombo3Length", "TableAddContext");
                            CreateColumn("icomboOption3", "TableAddContext");
                        }
                        else if (control.Name.Contains("4"))
                        {
                            CreateColumn("icombo4Length", "TableAddContext");
                            CreateColumn("icomboOption4", "TableAddContext");
                        }
                        else if (control.Name.Contains("5"))
                        {
                            CreateColumn("icombo5Length", "TableAddContext");
                            CreateColumn("icomboOption5", "TableAddContext");
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
        private void flllPanelItemsboxes(string rowID, string cellValue)
        {
            //MessageBox.Show("rowID = " + rowID + " - cellValue=" + cellValue);
            if (dataGridView1.Rows.Count > 1)
            {
                for (int index = 0; index < dataGridView1.Rows.Count - 1; index++)
                    if (cellValue == dataGridView1.Rows[index].Cells[rowID].Value.ToString())
                    {
                        TextModel.Text = dataGridView1.Rows[index].Cells["TextModel"].Value.ToString();
                        
                        foreach (Control Lcontrol in PanelItemsboxes.Controls)
                            try
                            {
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
                                                //MessageBox.Show(Lcontrol.Name + "Length");
                                                Vcontrol.Visible = true;
                                                string size = dataGridView1.Rows[index].Cells[Lcontrol.Name.Replace("L", "") + "Length"].Value.ToString();
                                                Vcontrol.Width = Convert.ToInt32(size);
                                                if (Convert.ToInt32(size) >= 700)
                                                {
                                                    if (Vcontrol is TextBox) ((TextBox)Vcontrol).Multiline = true;
                                                    Vcontrol.Height = 150;
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            ColorFulGrid9();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            panellError.Visible = false;
        }

        private void button116_Click(object sender, EventArgs e)
        {
            ColumnJobs();
            createNewColContext();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(insertAll, sqlCon);
            sqlCmd.CommandType = CommandType.Text;

            sqlCmd.Parameters.AddWithValue("@errorList", "");
            addParameters(sqlCmd);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
            this.Close();
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
            flllPanelItemsboxes("ID", idIndex.ToString());
            PanelItemsboxes.BringToFront();
            PanelItemsboxes.Visible = true;
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
            FillDataGridView("");
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
