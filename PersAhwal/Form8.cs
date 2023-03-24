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
using System.Security.AccessControl;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Drawing;
using ZXing;

namespace PersAhwal
{
    public partial class Form8 : Form
    {
        string DataSource = "";
        int panelIndex = 0;
        int updateAllIndex = 0;
        string starIndex = "0";
        string starIndexSub = "0";
        string starButton = "";
        bool insert = false;
        string queryAll = "";
        string startingText = "";
        string[] listFiels= new string[100];
        string selectTable = "";
        public Form8(string dataSource)
        {
            InitializeComponent();
            DataSource = dataSource;
            fillSamplesCodes(dataSource);
            getColList("TableAddContext");
            setlistFiels();
            الموضوع.SelectedIndex = 0;
        }

        private void setlistFiels()
        {
            for (int x = 0; x < 100; x++) { listFiels[x] = ""; }
            listFiels[0] = "حقل1";
            listFiels[1] = "حقل2";
            listFiels[2] = "حقل3";
            listFiels[3] = "حقل4";
            listFiels[4] = "حقل5";            
            listFiels[5] = "حقل6";
            listFiels[6] = "تاريخ1";
            listFiels[7] = "تاريخ2";
            listFiels[8] = "تاريخ3";
            listFiels[9] = "تاريخ4";
            listFiels[10] = "تاريخ5";
            listFiels[11] = "خيار متعدد1";
            listFiels[12] = "خيار متعدد2";
            listFiels[13] = "خيار متعدد3";
            listFiels[14] = "خيار متعدد4";
            listFiels[15] = "خيار متعدد5";
            listFiels[16] = "خيار ثنائي1";
            listFiels[17] = "خيار ثنائي2";
            listFiels[18] = "خيار ثنائي3";
            listFiels[19] = "خيار ثنائي4";
            listFiels[20] = "خيار ثنائي5";
            listFiels[21] = "حقل7";
            listFiels[22] = "حقل8";
            listFiels[23] = "حقل9";
            listFiels[24] = "حقل0";
            listFiels[25] = "إضافة";


        }
        private void getColList(string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('" + table + "')", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            allList = new string[dtbl.Rows.Count];
            int i = 0;
            string updateValues = "";
            foreach (DataRow row in dtbl.Rows)
            {
                if (row["name"].ToString() != "ID")
                {

                    allList[i] = row["name"].ToString();
                    i++;
                }
            }
            //updateAll = "UPDATE TableAddContext SET " + updateValues + " where ID = @id";
            //MessageBox.Show(updateAll);
        }

        private bool checkColExistance(string table, string colName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
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


        private bool checkColExist(string dataSource, string table)
        {
            //MessageBox.Show("dataSource " + dataSource);
            //MessageBox.Show("table " + table);
            SqlConnection sqlCon = new SqlConnection(dataSource);
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
            قائمة_النصوص_العامة.Items.Clear();
            foreach (DataRow row in dtbl.Rows)
            {
                
                قائمة_النصوص_العامة.Items.Add(row["name"].ToString().Replace("_", " "));
            }
            return false;

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

        private void Form8_Load(object sender, EventArgs e)
        {
            //checkColExist(DataSource, selectTable);
            fileComboBox(قائمة_النصوص_العامة, DataSource, "ArabicGenIgrar", "TableListCombo", true);
            fileComboBox(قائمة_النصوص_العامة, DataSource, "EnglishGenIgrar", "TableListCombo", false);
            fileComboBox(ColRight, DataSource, "ColRight", "TableAddContext", true);
        }

        private void AppType_CheckedChanged(object sender, EventArgs e)
        {
            if (الموضوع.SelectedIndex == 0)
            {
                
                قائمة_النصوص_العامة.Items.Clear();
                fileComboBox(قائمة_النصوص_العامة, DataSource, "ArabicGenIgrar", "TableListCombo", true);
                fileComboBox(قائمة_النصوص_العامة, DataSource, "EnglishGenIgrar", "TableListCombo", false);
                //checkColExist(DataSource, selectTable);
                الحقوق_lab.Visible = ColRight.Visible = false;
            }
            else if(الموضوع.SelectedIndex == 1)
            {
                قائمة_النصوص_العامة.Items.Clear();
                fileComboBox(قائمة_النصوص_العامة, DataSource, "AuthTypes", "TableListCombo", true);                
                الحقوق_lab.Visible = ColRight.Visible = true;
                //checkColExist(DataSource, selectTable);
            }
            
            else if(الموضوع.SelectedIndex == 2)
            {
                قائمة_النصوص_العامة.Items.Clear();
                fileComboBox(قائمة_النصوص_العامة, DataSource, "AuthTypes", "TableListCombo", true);                
                الحقوق_lab.Visible = ColRight.Visible = true;
                //checkColExist(DataSource, selectTable);
            }
        }
        private void ViewArchShow(string text, string ID)
        {
            //MessageBox.Show(ID);
            Button btnArchieve = new Button();
            btnArchieve.Location = new System.Drawing.Point(12, 1);
            btnArchieve.Name = قائمة_النصوص_العامة.Text.Replace(" ", "_") + "-" + ID;
            btnArchieve.Size = new System.Drawing.Size(667, 146);
            btnArchieve.TabIndex = panelIndex;
            btnArchieve.Text = text;
            btnArchieve.Click += new System.EventHandler(this.button_Click);            
            btnArchieve.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            flowLayoutPanel1.Controls.Add(btnArchieve);
            panelIndex++;
        }

        private string removeSpace(string text)
        {
            string authother = "";
            string removeAuthother = "";
            string lastSentence = "";
            string[] sentences = text.Split('،');
            foreach (string sentence in sentences)
            {
                if (sentence.Contains("الحق في توكيل الغير"))
                    authother = sentence;
                if (sentence.Contains("ويعتبر التوكيل الصادر"))
                    removeAuthother = sentence;
                if (sentence.Contains("لمن يشهد والله"))
                    lastSentence = sentence;
            }
            if (الموضوع.SelectedIndex == 2)
            {
                if (!text.Contains("لمن يشهد والله"))
                    text = text + "، وأذنت لمن يشهد والله خير الشاهدين";
                else
                    text = text.Replace(lastSentence, "، وأذنت لمن يشهد والله خير الشاهدين");
            }
            try
            {
                text = text.Replace(authother, "");
                text = text.Replace(removeAuthother, "");
            }
            catch (Exception ex) { }
            //MessageBox.Show(text);
            for (; text.Contains("،،");)
            {
                text = text.Replace("،،", "، ");
            }
            text = text.Replace("، ،", "، ");
            text = text.Replace("،", "، ");
            text = text.Replace("1_", "");
            text = text.Replace("0_", "");
            text = text.Replace("،،", "،");
            text = text.Replace("..", ".");
            for (; text.Contains("  ");)
            {
                text = text.Replace("  ", " ");
            }
            text = text.Trim();
            

            return text;
        }

        private void button_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            startingText = النص.Text = removeSpace(button.Text);
            starButton = button.Name.Split('-')[1];
            //MessageBox.Show(starButton);
            if (starButton != starIndex)
            {
                picStar.Visible = false;
            }
            else
            {
                picStar.Visible = true;
            }
            النص.Visible = true;
        }

        string[] allList;
        
        private string SuffReplacements(string text, int appCaseIndex, int intAuthcases)
        {

            if (text.Contains("  "))
                text = text.Replace("  ", " ");
            if (text.Contains("t1"))
                text = text.Replace("t1", "حقل1");
            if (text.Contains("t2"))
                text = text.Replace("t2", "حقل2");
            if (text.Contains("t3"))
                text = text.Replace("t3", "حقل3");
            if (text.Contains("t4"))
                text = text.Replace("t4", "حقل4");
            if (text.Contains("t5"))
                text = text.Replace("t5", "حقل5");
            if (text.Contains("t6"))
                text = text.Replace("t6", "حقل6");
            if (text.Contains("t7"))
                text = text.Replace("t7", "حقل7");
            if (text.Contains("t8"))
                text = text.Replace("t8", "حقل8");
            if (text.Contains("t9"))
                text = text.Replace("t9", "حقل9");
            if (text.Contains("t0"))
                text = text.Replace("t0", "حقل0");
            if (text.Contains("c1"))
                text = text.Replace("c1", "خيار ثنائي1");
            if (text.Contains("c2"))
                text = text.Replace("c2", "خيار ثنائي2");
            if (text.Contains("c3"))
                text = text.Replace("c3", "خيار ثنائي3");
            if (text.Contains("c4"))
                text = text.Replace("c4", "خيار ثنائي4");
            if (text.Contains("c5"))
                text = text.Replace("c5", "خيار ثنائي5");
            if (text.Contains("m1"))
                text = text.Replace("m1", "خيار متعدد1");
            if (text.Contains("m2"))
                text = text.Replace("m2", "خيار متعدد2");
            if (text.Contains("m3"))
                text = text.Replace("m3", "خيار متعدد3");
            if (text.Contains("m4"))
                text = text.Replace("m4", "خيار متعدد4");
            if (text.Contains("m5"))
                text = text.Replace("m5", "خيار متعدد5");
            if (text.Contains("n1"))
                text = text.Replace("n1", "تاريخ1");
            if (text.Contains("n2"))
                text = text.Replace("n2", "تاريخ2");
            if (text.Contains("n3"))
                text = text.Replace("n3", "تاريخ3");
            if (text.Contains("n4"))
                text = text.Replace("n4", "تاريخ4");
            if (text.Contains("n5"))
                text = text.Replace("n5", "تاريخ5");
            if (text.Contains("a1"))
                text = text.Replace("a1", "إضافة");
            if (text.Contains("@*@"))
            {
                text = text.Replace("@*@", "لدى  برقم الايبان (حقل3)");
            }
            if (text.Contains("#8"))
                text = text.Replace("#8", "حقل_الحذف1");
            if (text.Contains("#6"))
                text = text.Replace("#6", "حقل_الحذف2");
            if (text.Contains("#7"))
                text = text.Replace("#7", "حقل_الحذف3");

            for (int gridIndex = 0; gridIndex < dataGridView1.Rows.Count - 1; gridIndex++)
            {
                string code = dataGridView1.Rows[gridIndex].Cells["الرموز"].Value.ToString();
                string person = dataGridView1.Rows[gridIndex].Cells["الضمير"].Value.ToString();
                string[] replacemest = new string[6];
                try
                {
                    replacemest[0] = dataGridView1.Rows[gridIndex].Cells["المقابل1"].Value.ToString();
                    replacemest[1] = dataGridView1.Rows[gridIndex].Cells["المقابل2"].Value.ToString();
                    replacemest[2] = dataGridView1.Rows[gridIndex].Cells["المقابل3"].Value.ToString();
                    replacemest[3] = dataGridView1.Rows[gridIndex].Cells["المقابل4"].Value.ToString();
                    replacemest[4] = dataGridView1.Rows[gridIndex].Cells["المقابل5"].Value.ToString();
                    replacemest[5] = dataGridView1.Rows[gridIndex].Cells["المقابل6"].Value.ToString();
                }
                catch (Exception ex) { return text; }
                if (text.Contains(code))
                {
                    if (person == "1")
                        text = text.Replace(code, replacemest[appCaseIndex]);
                    else if (person == "2")
                        text = text.Replace(code, replacemest[intAuthcases]);
                }
            }
            return text;
        }
        
        private string SuffReversReplacements(string text, int appCaseIndex, int intAuthcases)
        {

            if (text.Contains("  "))
                text = text.Replace("  ", " ");
            if (text.Contains("حقل1"))
                text = text.Replace("حقل1", "t1");
            if (text.Contains("حقل2"))
                text = text.Replace("حقل2", "t2");
            if (text.Contains("حقل3"))
                text = text.Replace("حقل3", "t3");
            if (text.Contains("حقل4"))
                text = text.Replace("حقل4", "t4");
            if (text.Contains("حقل5"))
                text = text.Replace("حقل5", "t5");
            if (text.Contains("خيار_ثنائي1"))
                text = text.Replace("خيار_ثنائي1","c1");
            if (text.Contains("خيار_متعدد1"))
                text = text.Replace("خيار_متعدد1", "m1");
            if (text.Contains("خيار_متعدد2"))
                text = text.Replace("خيار_متعدد2", "m2");
            if (text.Contains("تاريخ1"))
                text = text.Replace("تاريخ1", "n1");
            if (text.Contains("لدى  برقم الايبان (حقل3)"))
            {
                text = text.Replace("لدى  برقم الايبان (حقل3", "لدى  برقم الايبان (@*@)");
            }
            if (text.Contains("#حقل_الحذف1"))
                text = text.Replace("حقل_الحذف1", "#8");
            if (text.Contains("#حقل_الحذف2"))
                text = text.Replace("حقل_الحذف2", "#6");
            if (text.Contains("#حقل_الحذف3"))
                text = text.Replace("حقل_الحذف3", "#7");
            text = SuffConvertments(text, 0, 0);
            return text;
        }

        private string SuffConvertments(string text, int person1, int person2)
        {
            string[] words = text.Split(' ');

            foreach (string word in words)
            {
                if (word == "" || word == " ") continue;
                for (int gridIndex = 0; gridIndex < dataGridView1.Rows.Count - 1; gridIndex++)
                {
                    string code = dataGridView1.Rows[gridIndex].Cells["الرموز"].Value.ToString();
                    string person = dataGridView1.Rows[gridIndex].Cells["الضمير"].Value.ToString();

                    string replacemest1 = dataGridView1.Rows[gridIndex].Cells["المقابل" + (person1 + 1).ToString()].Value.ToString();
                    string replacemest2 = dataGridView1.Rows[gridIndex].Cells["المقابل" + (person2 + 1).ToString()].Value.ToString();

                    string[] replacemests = new string[6];
                    replacemests[0] = dataGridView1.Rows[gridIndex].Cells["المقابل1"].Value.ToString();
                    replacemests[1] = dataGridView1.Rows[gridIndex].Cells["المقابل2"].Value.ToString();
                    replacemests[2] = dataGridView1.Rows[gridIndex].Cells["المقابل3"].Value.ToString();
                    replacemests[3] = dataGridView1.Rows[gridIndex].Cells["المقابل4"].Value.ToString();
                    replacemests[4] = dataGridView1.Rows[gridIndex].Cells["المقابل5"].Value.ToString();
                    replacemests[5] = dataGridView1.Rows[gridIndex].Cells["المقابل6"].Value.ToString();

                    for (int cellIndex = 0; cellIndex < 6; cellIndex++)
                    {
                        if (word == replacemests[cellIndex] || word == replacemests[cellIndex] + "،")
                        {
                            Console.WriteLine(word);
                            if (person == "1")
                            {
                                if (word != replacemests[person1])
                                {
                                    var selectedOption = MessageBox.Show("هل تود إجراء التصحيح التلقائي (" + replacemests[person2] + ")", "تم رصد خطاء في الصياغة (" + word + ")", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (selectedOption == DialogResult.Yes)
                                    {
                                        text = text.Replace(word, replacemests[person1]);
                                        break;
                                    }
                                    //    MessageBox.Show(word); 
                                }
                            }
                            if (person == "2")
                            {
                                if (word != replacemests[person2])
                                {
                                    var selectedOption = MessageBox.Show("هل تود إجراء التصحيح التلقائي (" + replacemests[person2] + ")", "تم رصد خطاء في الصياغة (" + word + ")", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (selectedOption == DialogResult.Yes)
                                    {
                                        text = text.Replace(word, replacemests[person2]);
                                        break;
                                    }
                                    //    MessageBox.Show(word); 
                                }
                            }
                            if (person == "3")
                            {
                                if (word != replacemests[person1])
                                {
                                    var selectedOption = MessageBox.Show("هل تود إجراء التصحيح التلقائي (" + replacemests[person1] + ")", "تم رصد خطاء في الصياغة (" + word + ")", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                    if (selectedOption == DialogResult.Yes)
                                    {
                                        text = text.Replace(word, replacemests[person1]);
                                        break;
                                    }
                                    //    MessageBox.Show(word); 
                                }
                            }
                            //text = text.Replace(replacemest[cellIndex], code);
                            //break;
                        }
                    }

                }
            }
            return text;
        }
        private void fillSamplesCodes(string source)
        {
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select * from Tablechar";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.CommandType = CommandType.Text;

                try
                {
                    cmd.ExecuteNonQuery();

                    DataTable table = new DataTable();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                    dataAdapter.Fill(table);
                    dataGridView1.DataSource = table;
                }
                catch (Exception ex) { }
                saConn.Close();
            }
        }

        private void قائمة_النصوص_SelectedIndexChanged(object sender, EventArgs e)
        {            
            newFillComboBox1(قائمة_النصوص_الفرعية, DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_"));

            if (قائمة_النصوص_العامة.Text != "" && قائمة_النصوص_الفرعية.Text != "")
            {
                if (الموضوع.SelectedIndex == 0)
                {
                    checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                    getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight = ''");
                }
                else if (الموضوع.SelectedIndex == 1)
                {
                    checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                    getstarTextSub("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
                }
                else if (الموضوع.SelectedIndex == 2)
                {
                    checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                    getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
                }
            }
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
        //private void flllPanelItemsboxes(string rowID, string cellValue)
        //{
        //    //MessageBox.Show("rowID = " + rowID + " - cellValue=" + cellValue);
        //    string query = "select * from TableAddContext where " + rowID + "=N'" + cellValue + "' and ColRight = ''";
        //    SqlConnection sqlCon = new SqlConnection(DataSource);
        //    if (sqlCon.State == ConnectionState.Closed)
        //        sqlCon.Open();
        //    SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
        //    sqlDa.SelectCommand.CommandType = CommandType.Text;
        //    DataTable dtbl = new DataTable();
        //    sqlDa.Fill(dtbl);
        //    //MessageBox.Show(query);
        //    Console.WriteLine(query + " - " + dtbl.Rows.Count.ToString());
        //    if (dtbl.Rows.Count > 0)

        //        foreach (DataRow dr in dtbl.Rows)
        //        //if (cellValue == dataGridView1.Rows[index].Cells[rowID].Value.ToString())
        //        {
        //            ColName = dr["ColName"].ToString();
        //            ColRight = dr["ColRight"].ToString();
        //            startID = dr["starText"].ToString();
        //            if (startID == "")
        //            {
        //                picStar.Visible = false; btnPrevious.Visible = true;
        //                StrSpecPur = dr["TextModel"].ToString();
        //            }
        //        }
        //}
        private void checkStarTextExist(string dataSource, string col, string genTable)
        {
            string query = "select ID," + col + " from " + genTable;
            Console.WriteLine("checkStarTextExist " + query);
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            int count = 0;
            panelIndex = 0;
            
            try
            {
                sqlDa.Fill(dtbl);
            }
            catch (Exception ex) {
                عدد_النماذج.Text = "عدد النماذج " + count.ToString(); 
                return; 
            }
            
            foreach (Control control in flowLayoutPanel1.Controls)
            {
                control.Visible = false;
                control.Name = "unvalid";
            }

            foreach (DataRow row in dtbl.Rows)
            {

                try
                {
                    if (row[col].ToString() != "")
                    {
                        ViewArchShow(SuffReplacements(row[col].ToString(), 0, 0), row["ID"].ToString());
                        count++;
                    }
                }
                catch (Exception ex) {
                    عدد_النماذج.Text = "عدد النماذج " + count.ToString(); 
                    return; 
                }
            }
            عدد_النماذج.Text = "عدد النماذج " + count.ToString();
            sqlCon.Close();
        }

        private void قائمة_النصوص_الفرعية_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (الموضوع.SelectedIndex == 0)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight = ''");
            }
            else if (الموضوع.SelectedIndex == 1)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                getstarTextSub("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
            }
            else if (الموضوع.SelectedIndex == 2)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
            }



            }
        private void getstarText(string rowID, string cellValue, string colright)
        {
            string query = "select * from TableAddContext where " + rowID + "=N'" + cellValue + "' and "+ colright;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            if (dtbl.Rows.Count > 0)
                foreach (DataRow dr in dtbl.Rows)                
                {
                    starIndex = dr["starText"].ToString();                    
                }
        }
        
        private void getstarTextSub(string rowID, string cellValue, string colright)
        {
            string query = "select * from TableAddContext where " + rowID + "=N'" + cellValue + "' and "+ colright;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            if (dtbl.Rows.Count > 0)
                foreach (DataRow dr in dtbl.Rows)                
                {
                    starIndexSub = dr["starTextSub"].ToString();                    
                }
        }

        private void نص_مرجعي_Click(object sender, EventArgs e)
        {
            updateText();
            //MessageBox.Show(updateText());
            النص.Text = ""; 
            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAddContext SET starText=@starText WHERE ColName = N'" + قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString()+"'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            //MessageBox.Show(starButton);
            sqlCmd.Parameters.AddWithValue("@starText", starButton);
            sqlCmd.ExecuteNonQuery();
            النص.Text = "";
            if (الموضوع.SelectedIndex == 0)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight = ''");
            }
            else if (الموضوع.SelectedIndex == 1)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                getstarTextSub("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
            }
            else if (الموضوع.SelectedIndex == 2)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
            }
        }

        private void تعيين_كخيار_Click(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAddContext SET starText=@starText WHERE ColName = N'" + قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString() + "'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;            
            sqlCmd.Parameters.AddWithValue("@starText", "");
            sqlCmd.ExecuteNonQuery();
            النص.Text = "";
            if (الموضوع.SelectedIndex == 0)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight = ''");
            }
            else if (الموضوع.SelectedIndex == 1)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                getstarTextSub("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
            }
            else if (الموضوع.SelectedIndex == 2)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string query = "delete from TableCollectStarText where ID = '" + starButton + "'";
            if (الموضوع.SelectedIndex == 1)
                query = "delete from TableAuthStarText where ID = '" + starButton + "'";
            if (الموضوع.SelectedIndex == 2)
                query = "delete from TableAuthRightStarText where ID = '" + starButton + "'";
            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            النص.Text = "";
            if (الموضوع.SelectedIndex == 0)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight = ''");
            }
            else if (الموضوع.SelectedIndex == 1)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                getstarTextSub("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
            }
            else if (الموضوع.SelectedIndex == 2)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            insert = false;
            if(الموضوع.SelectedIndex < 2)
                updateAllFields();

            updateText();
            النص.Text = "";
            if (الموضوع.SelectedIndex == 0)
            {
                getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight = ''");
            }
            else if (الموضوع.SelectedIndex == 1)
            {
                getstarTextSub("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
            }
            else if (الموضوع.SelectedIndex == 2)
            {
                getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
            }
        }

        private string updateText() {            
            if (!checkColExistance(selectTable, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_")))
                CreateColumn(قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable, "max");
            string ID = checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), النص.Text, selectTable);

            //MessageBox.Show(starButton);
            النص.Text = SuffReversReplacements(النص.Text, 0, 0);
            string query = "UPDATE "+selectTable+" SET " + قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + "=N'" + النص.Text + "' where ID = " + starButton;
            if(starButton == "")
                query = "insert into " + selectTable + " (" + قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + ") value (N'" + النص.Text + "');SELECT @@IDENTITY as lastid";

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;

            if (starButton == "")
                sqlCmd.ExecuteNonQuery();
            else
            {
                var reader = sqlCmd.ExecuteReader();
                if (reader.Read())
                {
                    return reader["lastid"].ToString();
                }
                sqlCon.Close();
            }
            return starButton;
        }

        private void updateAllFields()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('TableAddContext')", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            allList = new string[dtbl.Rows.Count];
            updateAllIndex = 0;
            string updateValues = "";
            foreach (DataRow row in dtbl.Rows)
            {
                foreach(Control control in PanelItemsboxes.Controls) {
                    if ((row["name"].ToString() == control.Name||row["name"].ToString() == control.Name+"Option") && control.Visible)
                    {
                        allList[updateAllIndex] = row["name"].ToString();
                        if (updateAllIndex == 0)
                        {
                            updateValues = row["name"].ToString() + "=@" + row["name"].ToString();
                        }
                        else
                        {
                            updateValues = updateValues + "," + row["name"].ToString() + "=@" + row["name"].ToString();
                        }
                        updateAllIndex++;
                    }
                }
            }
            queryAll = "UPDATE TableAddContext SET " + updateValues + " WHERE ColName = N'" + قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString() + "'";
            save2DataBase(PanelItemsboxes);
        }
        
        private void insertAllFields()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('TableAddContext')", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            allList = new string[dtbl.Rows.Count];
            updateAllIndex = 0;
            string insertItems = "";
            string insertValues = "";
            foreach (DataRow row in dtbl.Rows)
            {
                foreach(Control control in PanelItemsboxes.Controls) {
                    if ((row["name"].ToString() == control.Name||row["name"].ToString() == control.Name+"Option") && control.Visible)
                    {
                        allList[updateAllIndex] = row["name"].ToString();
                        insertItems = insertItems +","+ row["name"].ToString();
                        insertValues = insertValues + ",@" + row["name"].ToString();
                        updateAllIndex++;
                    }
                }
            }
            queryAll = "insert into TableAddContext (ColRight,ColName,TextModel,starText" + insertItems + ") values (@ColRight,@ColName,@TextModel,@starText" + insertValues + ")";
            save2DataBase(PanelItemsboxes);
        }

        private bool save2DataBase(FlowLayoutPanel panel)
        {
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(queryAll, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            
            bool cont = true;
            
            for (int i = 0; i < updateAllIndex; i++)
            {                
                foreach (Control control in panel.Controls)
                {
                    if (control.Visible)
                    {
                        //MessageBox.Show(control.Name + " - " + control.Text);
                        if (control.Name == allList[i] || (allList[i].Contains("Option") && control.Name == "نص_"+allList[i].Replace( "Option","")))
                        {
                           // MessageBox.Show(control.Name);
                            sqlCommand.Parameters.AddWithValue("@" + allList[i], control.Text);
                            //MessageBox.Show(allList[i] + " - " + control.Text);
                        }
                    }
                }
            }
            try
            {
                if (insert)
                {
                    if (الموضوع.SelectedIndex == 0)
                        sqlCommand.Parameters.AddWithValue("@ColRight", "");
                    else
                        sqlCommand.Parameters.AddWithValue("@ColRight", ColRight.Text);
                    sqlCommand.Parameters.AddWithValue("@ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString());
                    sqlCommand.Parameters.AddWithValue("@TextModel", النص.Text);
                    var selectedOption = MessageBox.Show("تعين النص كمرجع", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (selectedOption == DialogResult.Yes)
                    {
                        sqlCommand.Parameters.AddWithValue("@starText", "1");
                    }
                    else if (selectedOption == DialogResult.No)
                    {
                        sqlCommand.Parameters.AddWithValue("@starText", "");
                    }                     
                }
                sqlCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(queryAll);
            }
            return true;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            insert = true;
            insertAllFields();
            
                النص.Text = SuffReversReplacements(النص.Text,0,0);
            if (الموضوع.SelectedIndex == 0)
            {
                if (!checkColExistance(selectTable, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_")))
                    CreateColumn(قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable, "max");
                
                if (checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), النص.Text, selectTable) == "") return;
            }
            else if (الموضوع.SelectedIndex == 1)
            {
                if (!checkColExistance(selectTable, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_")))
                    CreateColumn(قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable, "max");

                if (checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), النص.Text, selectTable) == "") return;
            }
            else if (الموضوع.SelectedIndex == 2)
            {
                if (!checkColExistance(selectTable, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_")))
                    CreateColumn(قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable, "max");

                if (checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), النص.Text, selectTable) == "") return;
            }
            string query = "insert TableCollectStarText into (" + قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + ") value (N'" + النص.Text + "')";
            if (الموضوع.SelectedIndex == 1)
                query = "insert TableAuthStarText into (" + قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + ") value (N'" + النص.Text + "')";
            if (الموضوع.SelectedIndex == 2)
                query = "insert TableAuthRightStarText into (" + قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + ") value (N'" + النص.Text + "')";

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            النص.Text = "";
            if (الموضوع.SelectedIndex == 0)
            {
                getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight = ''");
            }
            else if (الموضوع.SelectedIndex == 1)
            {                
                getstarTextSub("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
            }
            else if (الموضوع.SelectedIndex == 2)
            {                
                getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
            }
        }

        private string checkStarTextExist(string dataSource, string col, string text, string genTable)
        {
            string query = "select * from " + genTable + " where " + col + "=N'" + text + "'";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            foreach (DataRow row in dtbl.Rows) {
                if (dtbl.Rows.Count > 0) 
                    return row["ID"].ToString();                
            }
            return "";
        }

        private string getTableName(string text)
        {
            if (text.Contains("حقل1"))
                text = text.Replace("حقل1", "itext1");
            if (text.Contains("حقل2"))
                text = text.Replace("حقل2", "itext2");
            if (text.Contains("حقل3"))
                text = text.Replace("حقل3", "itext3");
            if (text.Contains("حقل4"))
                text = text.Replace("حقل4", "itext4");
            if (text.Contains("حقل5"))
                text = text.Replace("حقل5", "itext5");
            if (text.Contains("حقل6"))
                text = text.Replace("حقل6", "itext6");
            if (text.Contains("حقل7"))
                text = text.Replace("حقل7", "itext7");
            if (text.Contains("حق87"))
                text = text.Replace("حقل8", "itext9");
            if (text.Contains("حق9"))
                text = text.Replace("حقل9", "itext9");
            if (text.Contains("حقل0"))
                text = text.Replace("حقل0", "itext0");

            if (text.Contains("خيار ثنائي1"))
                text = text.Replace("خيار ثنائي1", "icheck1");
            if (text.Contains("خيار ثنائي2"))
                text = text.Replace("خيار ثنائي2", "icheck2");
            if (text.Contains("خيار ثنائي3"))
                text = text.Replace("خيار ثنائي3", "icheck3");
            if (text.Contains("خيار ثنائي4"))
                text = text.Replace("خيار ثنائي4", "icheck4");
            if (text.Contains("خيار ثنائي5"))
                text = text.Replace("خيار ثنائي5", "icheck5");

            if (text.Contains("خيار متعدد1"))
                text = text.Replace("خيار متعدد1", "icombo1");
            if (text.Contains("خيار متعدد2"))
                text = text.Replace("خيار متعدد2", "icombo2");
            if (text.Contains("خيار متعدد3"))
                text = text.Replace("خيار متعدد3", "icombo3");
            if (text.Contains("خيار متعدد4"))
                text = text.Replace("خيار متعدد4", "icombo4");
            if (text.Contains("خيار متعدد5"))
                text = text.Replace("خيار متعدد5", "icombo5");

            if (text.Contains("تاريخ1"))
                text = text.Replace("تاريخ1", "itxtDate1");
            if (text.Contains("تاريخ2"))
                text = text.Replace("تاريخ2", "itxtDate2");
            if (text.Contains("تاريخ3"))
                text = text.Replace("تاريخ3", "itxtDate3");
            if (text.Contains("تاريخ4"))
                text = text.Replace("تاريخ4", "itxtDate4");
            if (text.Contains("تاريخ5"))
                text = text.Replace("تاريخ5", "itxtDate5");
            if (text.Contains("إضافة"))
                text = text.Replace("إضافة", "ibtnAdd1");
            if (text != "")
            {
                
                if (!checkColExistance("TableAddContext", text))
                {
                    //Console.WriteLine("checkColExistance " + text);
                    CreateColumn(text, "TableAddContext","max");
                    
                    if(!text.Contains("btnAdd") && !text.Contains("icheck"))
                        CreateColumn(text + "Length", "TableAddContext","5");
                    if(text.Contains("icombo") || !text.Contains("icheck"))
                        CreateColumn(text + "Option", "TableAddContext","max");
                }
            }
            //Console.WriteLine (text);
            return text;    
        }
        private void CreateColumn(string Columnname, string tableName, string size)
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("alter table " + tableName + " add " + Columnname + " nvarchar("+ size+")", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            try
            {
                sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex) { return; }
            sqlCon.Close();
        }
        private void النص_TextChanged(object sender, EventArgs e)
        {
            foreach (Control control in PanelItemsboxes.Controls)
            {
                control.Visible = false;
            }
            for (int index = 0; listFiels[index] != ""; index++)
                {
                    if (النص.Text.Contains(listFiels[index]))
                {
                    panelFill(DataSource, getTableName(listFiels[index]));
                }
            }
        }

        public void panelFill(string dataSource, string field)
        {
            string query = "select * from TableAddContext where ColName = N'" + قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString() + "'";

            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            Console.WriteLine("panelFill " + query);

            foreach (DataRow row in dtbl.Rows)
            {
                foreach (Control control in PanelItemsboxes.Controls)
                {
                    if (control.Name == field)
                    {
                        control.Text = row[field].ToString();

                        control.Visible = true;
                        control.Width = (control.Text.Length * 8) + 1;
                        if (control.Width < 100)
                            control.Width = 100;
                        foreach (Control controlText in PanelItemsboxes.Controls)
                        {
                            if (controlText.Name == "نص_" + field)
                            {
                                controlText.Visible = true;
                                try
                                {
                                    controlText.Text = row[control.Name + "Option"].ToString();

                                }
                                catch (Exception ex) { }
                                controlText.Width = (controlText.Text.Length * 8) + 1;
                                if (controlText.Width < 100)
                                    controlText.Width = 100;
                            }
                        }
                    }

                }
            }
        }

        private void الموضوع_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (الموضوع.SelectedIndex == 0)
            {
                selectTable = "TableCollectStarText"; 
            }
            else if (الموضوع.SelectedIndex == 1)
            {
                selectTable = "TableAuthStarText"; 
            }
            else if (الموضوع.SelectedIndex == 2)
            {
                selectTable = "TableAuthRightStarText";
            }


            if (الموضوع.SelectedIndex == 0)
            {
                fileComboBox(قائمة_النصوص_العامة, DataSource, "ArabicGenIgrar", "TableListCombo", true);
                fileComboBox(قائمة_النصوص_العامة, DataSource, "EnglishGenIgrar", "TableListCombo", false);
            }
            else
            {
                fileComboBox(قائمة_النصوص_العامة, DataSource, "AuthTypes", "TableListCombo", true);                
            }

            if (قائمة_النصوص_العامة.Text != "" && قائمة_النصوص_الفرعية.Text != "")
            {
                if (قائمة_النصوص_العامة.Text != "" && قائمة_النصوص_الفرعية.Text != "")
                {
                    checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                    if (الموضوع.SelectedIndex == 0)
                    {                        
                        getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight = ''");
                    }
                    else if (الموضوع.SelectedIndex == 1)
                    {
                        getstarTextSub("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
                    }
                    else if (الموضوع.SelectedIndex == 2)
                    {
                        getstarText("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString(), "ColRight <> ''");
                    }
                }
            }
        }

        private void ColRight_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateCheckBoxes(txtRights.Text.Replace(" ", "_").Replace("-", "_"), "TableAuthRights", DataSource);
        }
        public void PopulateCheckBoxes(string col, string table, string dataSource)
        {

            if (col == "الحقوق" || col == "Col" || col == "" || table == "" || dataSource == "") return;
            string query = "SELECT ID," + col.Replace("-", "_") + " FROM " + table;
            string authother = "";
            string removeAuthother = "";
            string lastSentence = "";
            txtRights.Text = "";
            panelRights.BringToFront();
            panelRights.Visible = true;
            using (SqlConnection con = new SqlConnection(dataSource))
            {
                DataTable checkboxdt = new DataTable();
                using (SqlDataAdapter sda = new SqlDataAdapter(query, con))
                {
                    Console.WriteLine(query);
                    try
                    {
                        sda.Fill(checkboxdt);
                    }
                    catch (Exception ex) { return; }
                    foreach (DataRow row in checkboxdt.Rows)
                    {
                        string [] Text_statis = row[col.Replace("-", "_")].ToString().Split('_');
                        if (row[col.Replace("-", "_")].ToString() == "") continue;
                        
                        string text = SuffReplacements(Text_statis[0], 0,0);
                        if (text.Contains("الحق في توكيل الغير"))
                            authother = text;

                        if (text.Contains("ويعتبر التوكيل الصادر"))
                            removeAuthother = text;
                        if (text.Contains("لمن يشهد والله"))
                            lastSentence = text;                        
                        try
                        {
                            if (!txtRights.Text.Contains(lastSentence))
                                txtRights.Text = txtRights.Text + "، " + lastSentence;
                            if (!txtRights.Text.Contains(text))
                                txtRights.Text = txtRights.Text + text + " ";
                            txtRights.Text = txtRights.Text.Replace(authother, "") + " ";
                            txtRights.Text = txtRights.Text.Replace(removeAuthother, "") + " ";
                        }
                        catch (Exception ex) { }
                    }
                }
            }
            //autoCompleteTextBox(txtAddRight, DataSource, "قائمة_الحقوق_الكاملة", "TableAuthRight");
        }

    }
}
