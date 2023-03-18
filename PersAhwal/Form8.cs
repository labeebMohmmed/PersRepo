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

namespace PersAhwal
{
    public partial class Form8 : Form
    {
        string DataSource = "";
        int panelIndex = 0;
        string starIndex = "0";
        string starButton = "";
        string startingText = "";
        public Form8(string dataSource)
        {
            InitializeComponent();
            DataSource = dataSource;
            fillSamplesCodes(dataSource);
        }

        private bool checkColExist(string dataSource, string table)
        {
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
            //checkColExist(DataSource, "TableCollectStarText");
            fileComboBox(قائمة_النصوص_العامة, DataSource, "ArabicGenIgrar", "TableListCombo", true);
            fileComboBox(قائمة_النصوص_العامة, DataSource, "EnglishGenIgrar", "TableListCombo", false);
        }

        private void AppType_CheckedChanged(object sender, EventArgs e)
        {
            if (AppType.Checked)
            {
                AppType.Text = "مجموع المعاملات";
                قائمة_النصوص_العامة.Items.Clear();
                fileComboBox(قائمة_النصوص_العامة, DataSource, "ArabicGenIgrar", "TableListCombo", true);
                fileComboBox(قائمة_النصوص_العامة, DataSource, "EnglishGenIgrar", "TableListCombo", false);
                checkColExist(DataSource, "TableCollectStarText");
            }
            else
            {
                قائمة_النصوص_العامة.Items.Clear();
                fileComboBox(قائمة_النصوص_العامة, DataSource, "AuthTypes", "TableListCombo", true);
                AppType.Text = "التوكيلات";
                checkColExist(DataSource, "TableAuthStarText");
            }
        }
        private void ViewArchShow(string text, string ID)
        {
            //MessageBox.Show(ID);
            Button btnArchieve = new Button();
            btnArchieve.Location = new System.Drawing.Point(12, 1);
            btnArchieve.Name = قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + ID;
            btnArchieve.Size = new System.Drawing.Size(667, 146);
            btnArchieve.TabIndex = panelIndex;
            btnArchieve.Text = text;
            btnArchieve.Click += new System.EventHandler(this.button_Click);
            btnArchieve.RightToLeft = RightToLeft.Yes;
            btnArchieve.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            flowLayoutPanel1.Controls.Add(btnArchieve);
            panelIndex++;
        }
        private void button_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            startingText = النص.Text = button.Text;
            starButton = button.Name.Split('_')[1];
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

        private string SuffReplacements(string text, int appCaseIndex, int intAuthcases)
        {

            if (text.Contains("  "))
                text = text.Replace("  ", " ");
            if (text.Contains("t1"))
                text = text.Replace("t1", "الحقل1");
            if (text.Contains("t2"))
                text = text.Replace("t2", "الحقل2");
            if (text.Contains("t3"))
                text = text.Replace("t3", "الحقل3");
            if (text.Contains("t4"))
                text = text.Replace("t4", "الحقل4");
            if (text.Contains("t5"))
                text = text.Replace("t5", "الحقل5");
            if (text.Contains("c1"))
                text = text.Replace("c1", "خيار_ثنائي1");
            if (text.Contains("m1"))
                text = text.Replace("m1", "خيار_متعدد1");
            if (text.Contains("m2"))
                text = text.Replace("m2", "خيار_متعدد2");
            if (text.Contains("n1"))
                text = text.Replace("n1", "تاريخ1");
            if (text.Contains("@*@"))
            {
                text = text.Replace("@*@", "لدى  برقم الايبان (الحقل3)");
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
            if (text.Contains("الحقل1"))
                text = text.Replace("الحقل1", "t1");
            if (text.Contains("الحقل2"))
                text = text.Replace("الحقل2", "t2");
            if (text.Contains("الحقل3"))
                text = text.Replace("الحقل3", "t3");
            if (text.Contains("الحقل4"))
                text = text.Replace("الحقل4", "t4");
            if (text.Contains("الحقل5"))
                text = text.Replace("الحقل5", "t5");
            if (text.Contains("خيار_ثنائي1"))
                text = text.Replace("خيار_ثنائي1","c1");
            if (text.Contains("خيار_متعدد1"))
                text = text.Replace("خيار_متعدد1", "m1");
            if (text.Contains("خيار_متعدد2"))
                text = text.Replace("خيار_متعدد2", "m2");
            if (text.Contains("تاريخ1"))
                text = text.Replace("تاريخ1", "n1");
            if (text.Contains("لدى  برقم الايبان (الحقل3)"))
            {
                text = text.Replace("لدى  برقم الايبان (الحقل3", "لدى  برقم الايبان (@*@)");
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
        }

        //if (AppType.Checked)
        //    checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_"), "TableCollectStarText");
        //else checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_"), "TableAuthStarText");



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

            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            int count = 0;
            panelIndex = 0;
            
            foreach (Control control in flowLayoutPanel1.Controls)
            {
                control.Visible = false;
                control.Name = "unvalid";
            }

            foreach (DataRow row in dtbl.Rows)
            {
                if (row[col].ToString() != "")
                    ViewArchShow(SuffReplacements(row[col].ToString(), 0, 0), row["ID"].ToString());
                count++;
            }
            عدد_النماذج.Text = "عدد المكاذج " + count;
            sqlCon.Close();
        }

        private void قائمة_النصوص_الفرعية_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (AppType.Checked)
                checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), "TableCollectStarText");
            else checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), "TableAuthStarText");

            flllPanelItemsboxes("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString());
        }
        private void flllPanelItemsboxes(string rowID, string cellValue)
        {
            string query = "select * from TableAddContext where " + rowID + "=N'" + cellValue + "' and ColRight = ''";
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

        private void نص_مرجعي_Click(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAddContext SET starText=@starText WHERE ColName = N'" + قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString()+"'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            //MessageBox.Show(starButton);
            sqlCmd.Parameters.AddWithValue("@starText", starButton);
            sqlCmd.ExecuteNonQuery();
            النص.Text = "";
            if (AppType.Checked)
                checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), "TableCollectStarText");
            else checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), "TableAuthStarText");

            flllPanelItemsboxes("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString());
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
            if (AppType.Checked)
                checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), "TableCollectStarText");
            else checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), "TableAuthStarText");

            flllPanelItemsboxes("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string query = "delete from TableCollectStarText where ID = '" + starButton + "'";
            if (!AppType.Checked)
                query = "delete from TableAuthStarText where ID = '" + starButton + "'";
            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            النص.Text = "";
            if (AppType.Checked)
                checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), "TableCollectStarText");
            else checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), "TableAuthStarText");

            flllPanelItemsboxes("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            النص.Text = SuffReversReplacements(النص.Text, 0, 0); 
            string query = "UPDATE TableCollectStarText SET "+ قائمة_النصوص_الفرعية.Text.Replace(" ","_") + "=N'" + النص.Text + "'";
            if (!AppType.Checked)
                query = "UPDATE TableAuthStarText SET " + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + "=N'" + النص.Text + "'"; 
            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            النص.Text = "";
            if (AppType.Checked)
                checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), "TableCollectStarText");
            else checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), "TableAuthStarText");

            flllPanelItemsboxes("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            النص.Text = SuffReversReplacements(النص.Text,0,0);
            if (AppType.Checked)
            {
                if (checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), النص.Text, "TableCollectStarText")) return;
            }
            else if (!AppType.Checked)
            {
                if (checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), النص.Text, "TableAuthStarText")) return;
            }
            string query = "insert TableCollectStarText into (" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + ") value (N'" + النص.Text + "')";
            if (!AppType.Checked)
                query = "insert TableAuthStarText into (" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + ") value (N'" + النص.Text + "')";

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            النص.Text = "";
            if (AppType.Checked)
                checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), "TableCollectStarText");
            else checkStarTextExist(DataSource, قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), "TableAuthStarText");

            flllPanelItemsboxes("ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString());
        }

        private bool checkStarTextExist(string dataSource, string col, string text, string genTable)
        {
            string query = "select * from " + genTable + " where " + col + "=N'" + text + "'";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            if (dtbl.Rows.Count > 0) return true;
            else return false;
            sqlCon.Close();
        }
    }
}
