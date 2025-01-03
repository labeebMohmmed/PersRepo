﻿using System;
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
using DocumentFormat.OpenXml.Office2013.Excel;
using System.Data.SqlTypes;
using Color = System.Drawing.Color;

namespace PersAhwal
{
    public partial class Form8 : Form
    {
        
        int idList = 1;
        int FinalProcReqID = 0;
        string DataSource = "";
        string AddContextID = "";
        string TableProcReqID = "";
        int panelIndex = 0;
        int updateAllIndex = 0;
        int insertAllIndex = 0;
        string starIndex = "0";
        string starIndexSub = "0";
        string starButton = "";
        bool insert = false;
        string queryAll = "";
        string startingText = "";
        string[] listFiels = new string[100];
        string[] checlList;
        string selectTable = "";
        string ArchFile = "";
        string CurrentFile = "";
        int[] IDs;
        int IDsMax = 1;
        int IDsMin = 1;
        public Form8(string dataSource, string archFile)
        {
            InitializeComponent();
            ArchFile = archFile;
            DataSource = dataSource;
            fillSamplesCodes(dataSource);
            getColList("TableAddContext");
            setlistFiels();
            الموضوع.SelectedIndex = otherPro.SelectedIndex = 0;
            الموضوع.Select();
            //altColName();
            //updateList();
        }

        private string[] getRightColList()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('TableAuthRights') and name <> 'ID' ", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string[] allList = new string[dtbl.Rows.Count];
            int i = 0;
            string[] rightsList= new string[dtbl.Rows.Count];
            
            foreach (DataRow row in dtbl.Rows)
            {
                rightsList[i] = row["name"].ToString();
                i++;
                
            }
            return rightsList;
        }
        
        private string getRight(string col)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT "+ col + "  FROM TableAuthRights where ID  = 22 ", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string[] allList = new string[dtbl.Rows.Count];
            int i = 0;
            string right= "";
            
            foreach (DataRow row in dtbl.Rows)
            {
                right = row[col].ToString();
                
            }
            return right;
        }
        
        private void updateList()
        {
            string[] RightList = getRightColList();
            for (int x = 0; x < RightList.Length; x++)
            {
                SqlConnection sqlCon = new SqlConnection(DataSource);
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("select ID,replace(altSubColName + ' ' + altColName, ' ', '_') as col from TableAddContext where editRights <> '' ", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);
                sqlCon.Close();
                

                foreach (DataRow row in dtbl.Rows)
                {
                    if (RightList[x] == row["col"].ToString())
                    {
                        updateRights(getRight(RightList[x]), row["ID"].ToString());
                    }

                }
            }
            return ;
        }

        private void updateRights(string right, string id)
        {
            string query = "update TableAddContext set قائمة_الحقوق = N'" + right + "' where ID = " + id;
            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            try
            {
                sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(query);
            }

        }
            

        private void setCheclList()
        {
            checlList = new string[7];
            checlList[0] = "نص موضوع الانابة غير موجود";
            checlList[1] = "نص موضوع المكاتبة غير موجود";
            checlList[2] = "نص الحقوق غير موجود";
            checlList[3] = "المطلوبات الأولية غير محددة";
            checlList[4] = "استمارة الطلب غير موجودة";
            checlList[6] = "المطلوبات النهائية غير محددة";
            checlList[5] = "شرح المعاملة غير موجودة";
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
            listFiels[11] = "خ_متعدد1";
            listFiels[12] = "خ_متعدد2";
            listFiels[13] = "خ_متعدد3";
            listFiels[14] = "خ_متعدد4";
            listFiels[15] = "خ_متعدد5";
            listFiels[16] = "خ_ثنائي1";
            listFiels[17] = "خ_ثنائي2";
            listFiels[18] = "خ_ثنائي3";
            listFiels[19] = "خ_ثنائي4";
            listFiels[20] = "خ_ثنائي5";
            listFiels[21] = "حقل7";
            listFiels[22] = "حقل8";
            listFiels[23] = "حقل9";
            listFiels[24] = "حقل0";
            listFiels[25] = "إضافة";
            listFiels[26] = "اسم_مقدم_الطلب";
            listFiels[27] = "مكان_الاصدار";
            listFiels[28] = "نوع_الوثيقة";
            listFiels[30] = "تاريخ_الميلاد";
            listFiels[29] = "رقم_الوثيقة";
            listFiels[29] = "انتهاء_الصلاحية";


            for (int index = 0; listFiels[index] != ""; index++)
                addFields(index, listFiels[index]);

        }

        private void addFields(int index, string text) 
        {
            Button btn = new Button();
            btn.Font = new System.Drawing.Font("Arabic Typesetting", 15F);
            btn.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            btn.Location = new System.Drawing.Point(576, 3);
            btn.Name = "buttonFields_"+ index.ToString();
            btn.Height =  30;
            int width = text.Length * 8;
            if (width < 55) width = 55;
            btn.Width = width;
            btn.TabIndex = 840;
            btn.Text = text;
            btn.UseVisualStyleBackColor = true;
            btn.Click += new System.EventHandler(this.btn_Click_1);
            panel_fields.Controls.Add(btn);
        }
        private void btn_Click_1(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            if (!النص.Text.Contains(button.Text))
            {
                النص.Text = النص.Text + " " + button.Text;
                //button.Enabled = false; 
            }
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
        
        private void altColName()
        {
            //fileComboBox(قائمة_النصوص_العامة, DataSource, "ArabicGenIgrar", "TableListCombo", true);
            //fileComboBox(قائمة_النصوص_العامة, DataSource, "EnglishGenIgrar", "TableListCombo", false);

            fileComboBox(قائمة_النصوص_العامة, DataSource, "EnglishGenIgrar", "TableListCombo", true);

            string query = "select ID, right(ColName,len(ColName)-len(Substring(ColName,0,CharIndex('-',ColName)))-1) as col ,Substring(ColName,0,CharIndex('-',ColName)) as subCol , ColName from TableAddContext  where ColRight = '' and Lang = N'الانجليزية'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return ; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            foreach (DataRow row in dtbl.Rows)
            {
                try
                {
                    string column = قائمة_النصوص_العامة.Items[Convert.ToInt32(row["col"].ToString())].ToString();
                    //MessageBox.Show(column);
                    updatealtColName(row["ID"].ToString(), column);
                }catch (Exception ex) { }  
            }
            
        }

        
        
        private void updatealtColName(string id,string col)
        {
            string query = "update TableAddContext set altColName = N'"+ col +"' where ID = " + id;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;            
            sqlCmd.ExecuteNonQuery();
        }


        private bool checkColExist(string dataSource, string table, string Subtable)
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
                if(Subtable == row["name"].ToString().Replace("_", " "))
                    return true;
            }
            return false;

        }
        
        private bool checkSubColExist(string dataSource, string table, string subTable)
        {
            string query = "SELECT " + table + " FROM TableListCombo where " + table + "=N'" + subTable + "'";
            SqlConnection sqlCon = new SqlConnection(dataSource);
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
            Console.WriteLine("checkSubColExist " + query);
            //MessageBox.Show(dtbl.Rows.Count.ToString());
            if (dtbl.Rows.Count > 0) 
                return true;

            return false;

        }
        
        
        private bool checkRowExist(string dataSource, string col, string subTable)
        {
            string column = قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_");
            //string query = "select " + column + " from " +selectTable+ " where "+ column+ " = N'" + النص.Text +"'";
            string query = "SELECT ID FROM TableAddContext where altColName = N'" + col + "' and altSubColName=N'" + subTable + "'";
            SqlConnection sqlCon = new SqlConnection(dataSource);
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
            Console.WriteLine( query);
            //MessageBox.Show(dtbl.Rows.Count.ToString());
            if (dtbl.Rows.Count > 0)
            {
                
                return true;
            }

            return false;

        }


        private void fileComboBox(ComboBox combbox, string source, string comlumnName, string right, bool clear)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct "+ comlumnName+" from TableAddContext where "+ comlumnName+" is not null and ColRight "+ right+" order by "+ comlumnName+" asc";
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
        
        private void fileComboBox(ComboBox combbox, string source, string comlumnName)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct "+ comlumnName+ " from TableProcReq where " + comlumnName+" is not null";
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
        
        private void fileComboBoxSub(ComboBox combbox, string source, string comlumnName, string SubComlumnName, string right, bool clear)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct "+ SubComlumnName + " from TableAddContext where "+ SubComlumnName + " is not null and "+ comlumnName + " =N'"+ قائمة_النصوص_العامة .Text+ "'and ColRight " + right+" order by "+ SubComlumnName + " asc";
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

        private void Form8_Load(object sender, EventArgs e)
        {
            fileComboBox(قائمة_النصوص_العامة, DataSource, "altColName", "=''", false);
            
            fileComboBox(المعاملة, DataSource, "المعاملة");
        }

        private void AppType_CheckedChanged(object sender, EventArgs e)
        {
            if (الموضوع.SelectedIndex == 0)
            {

                قائمة_النصوص_العامة.Items.Clear();
                fileComboBox(قائمة_النصوص_العامة, DataSource, "altColName", "=''", true);
                الحقوق_lab.Visible = false;
            }
            else if (الموضوع.SelectedIndex == 1)
            {
                قائمة_النصوص_العامة.Items.Clear();
                fileComboBox(قائمة_النصوص_العامة, DataSource, "altColName", "<>''", true);
                الحقوق_lab.Visible =  true;
                //checkColExist(DataSource, selectTable);
            }

            else if (الموضوع.SelectedIndex == 2)
            {
                قائمة_النصوص_العامة.Items.Clear();
                fileComboBox(قائمة_النصوص_العامة, DataSource, "altColName", "<>''", true);
                الحقوق_lab.Visible = true;
                //checkColExist(DataSource, selectTable);
            }
        }
        private void ViewArchShow(string text, string ID)
        {
            //MessageBox.Show(ID);
            Button btnArchieve = new Button();
            btnArchieve.Location = new System.Drawing.Point(12, 1);
            btnArchieve.Name = قائمة_النصوص_العامة.Text.Replace(" ", "_") + "-" + ID;
            btnArchieve.Size = new System.Drawing.Size(650, 146);
            btnArchieve.TabIndex = panelIndex;
            btnArchieve.Text = SuffReplacements(text,0,0);
            btnArchieve.Click += new System.EventHandler(this.button_Click);
            btnArchieve.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            panel_النص.Controls.Add(btnArchieve);
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
            text = text.Replace("..", ".");
            text = text.Replace("!", "'");
            for (; text.Contains("  ");)
            {
                text = text.Replace("  ", " ");
            }
            text = text.Replace("، ،", "، ");
            text = text.Replace("  ", " ");
            text = text.Trim();
            

            return text;
        }

        private void button_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
           //MessageBox.Show(button.Text);
            startingText = النص.Text = removeSpace(button.Text);
            starButton = button.Name.Split('-')[1];
            //MessageBox.Show(starButton);
            //if (starButton != starIndex)
            //{
            //    picStar.Visible = false;
            //}
            //else
            //{
            //    picStar.Visible = true;
            //}
            النص.Visible = true;
        }

        string[] allList;

        private string SuffReplacements(string text, int appCaseIndex, int intAuthcases)
        {

            if (text.Contains("!"))
                text = text.Replace("!", "'");
            if (text.Contains("  "))
                text = text.Replace("  ", " ");
            if (text.Contains("tN"))
                text = text.Replace("tN", "اسم_مقدم_الطلب");
            if (text.Contains("tP"))
                text = text.Replace("tP", "رقم_الوثيقة");
            if (text.Contains("tS"))
                text = text.Replace("tS", "مكان_الاصدار");
            if (text.Contains("tX"))
                text = text.Replace("tX", "");
            if (text.Contains("tD"))
                text = text.Replace("tD", "نوع_الوثيقة");
            if (text.Contains("tB"))
                text = text.Replace("tB", "تاريخ_الميلاد");
            if (text.Contains("fD"))
                text = text.Replace("fD", "انتهاء_الصلاحية");
            
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
                text = text.Replace("c1", "خ_ثنائي1");
            if (text.Contains("c2"))
                text = text.Replace("c2", "خ_ثنائي2");
            if (text.Contains("c3"))
                text = text.Replace("c3", "خ_ثنائي3");
            if (text.Contains("c4"))
                text = text.Replace("c4", "خ_ثنائي4");
            if (text.Contains("c5"))
                text = text.Replace("c5", "خ_ثنائي5");
            if (text.Contains("m1"))
                text = text.Replace("m1", "خ_متعدد1");
            if (text.Contains("m2"))
                text = text.Replace("m2", "خ_متعدد2");
            if (text.Contains("m3"))
                text = text.Replace("m3", "خ_متعدد3");
            if (text.Contains("m4"))
                text = text.Replace("m4", "خ_متعدد4");
            if (text.Contains("m5"))
                text = text.Replace("m5", "خ_متعدد5");
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

            if (text.Contains("'"))
                text = text.Replace("'", "!");
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
            if (text.Contains("حقل6"))
                text = text.Replace("حقل6", "t6");
            if (text.Contains("حقل7"))
                text = text.Replace("حقل7", "t7");
            if (text.Contains("حقل8"))
                text = text.Replace("حقل8", "t8");
            if (text.Contains("حقل9"))
                text = text.Replace("حقل9", "t9");
            if (text.Contains("حقل0"))
                text = text.Replace("حقل0", "t0");

            if (text.Contains("اسم_مقدم_الطلب"))
                text = text.Replace( "اسم_مقدم_الطلب", "tN");
            if (text.Contains("رقم_الوثيقة"))
                text = text.Replace("رقم_الوثيقة","tP" );
            if (text.Contains("مكان_الاصدار"))
                text = text.Replace("مكان_الاصدار","tS");
            if (text.Contains("نوع_الوثيقة"))
                text = text.Replace( "نوع_الوثيقة","tD");
            if (text.Contains("انتهاء_الصلاحية"))
                text = text.Replace("انتهاء_الصلاحية", "fD");
            if (text.Contains("تاريخ_الميلاد"))
                text = text.Replace( "تاريخ_الميلاد","tB");

            if (text.Contains("خ_ثنائي1"))
                text = text.Replace("خ_ثنائي1", "c1");
            if (text.Contains("خ_ثنائي2"))
                text = text.Replace("خ_ثنائي2", "c2");
            if (text.Contains("خ_ثنائي3"))
                text = text.Replace("خ_ثنائي3", "c3");
            if (text.Contains("خ_ثنائي4"))
                text = text.Replace("خ_ثنائي4", "c4");
            if (text.Contains("خ_ثنائي5"))
                text = text.Replace("خ_ثنائي5", "c5");
            if (text.Contains("خ_متعدد1"))
                text = text.Replace("خ_متعدد1", "m1");
            if (text.Contains("خ_متعدد2"))
                text = text.Replace("خ_متعدد2", "m2");
            if (text.Contains("خ_متعدد3"))
                text = text.Replace("خ_متعدد3", "m3");
            if (text.Contains("خ_متعدد4"))
                text = text.Replace("خ_متعدد4", "m4");
            if (text.Contains("خ_متعدد5"))
                text = text.Replace("خ_متعدد5", "m5");
            
            if (text.Contains("تاريخ1"))
                text = text.Replace("تاريخ1", "n1");
            if (text.Contains("تاريخ2"))
                text = text.Replace("تاريخ2", "n2");
            if (text.Contains("تاريخ3"))
                text = text.Replace("تاريخ3", "n3");
            if (text.Contains("تاريخ4"))
                text = text.Replace("تاريخ4", "n4");
            if (text.Contains("تاريخ5"))
                text = text.Replace("تاريخ5", "n5");

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
            //newFillComboBox1(قائمة_النصوص_الفرعية, DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_"));
            setCheclList();
            string right = "= ''";
            if(الموضوع.SelectedIndex != 0)
                right = "<> ''";
            fileComboBoxSub(قائمة_النصوص_الفرعية, DataSource, "altColName", "altSubColName", right, true);
            if (قائمة_النصوص_العامة.Text != "" && قائمة_النصوص_الفرعية.Text != "")
            {
                if (الموضوع.SelectedIndex == 0)
                {
                    checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);

                }
                else if (الموضوع.SelectedIndex == 1)
                {
                    selectTable = "TableAuthStarText";
                    checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                    selectTable = "TableAuthRightStarText";
                    checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                }
                view_PreReq(false);

                if (الموضوع.SelectedIndex == 1)
                {
                    PopulateCheckBoxes(قائمة_النصوص_الفرعية.Text.Replace(" ", "_").Replace("-", "_") + "_" + قائمة_النصوص_العامة.Text.ToString(), "TableAuthRights", DataSource);
                }
                    flllPanelItemsboxes();
                
                معاملات_السابق.Visible = معاملات_التالي.Visible = جميع_المعاملات.Visible = false;
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
                try
                {cmd.ExecuteNonQuery();

                
                    DataTable table = new DataTable();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                    dataAdapter.Fill(table);
                    foreach (DataRow dataRow in table.Rows)
                    {
                        combbox.Items.Add(dataRow[colName].ToString());
                    }
                    saConn.Close();
                }
                catch (Exception ex) { }
            }
            //if (combbox.Items.Count > 0) combbox.SelectedIndex = 0;
        }
        private void flllPanelItemsboxes()
        {
            //MessageBox.Show("rowID = " + rowID + " - cellValue=" + cellValue);
            string query = "select * from TableAddContext  where altColName = N'" + قائمة_النصوص_العامة.Text + "' and altSubColName=N'" + قائمة_النصوص_الفرعية.Text + "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            //MessageBox.Show(query);
            Console.WriteLine(query + " - " + dtbl.Rows.Count.ToString());
            if (dtbl.Rows.Count > 0)

                foreach (DataRow dr in dtbl.Rows)
                //if (cellValue == dataGridView1.Rows[index].Cells[rowID].Value.ToString())
                {
                    النص.Text = SuffReplacements(dr["TextModel"].ToString(), 0, 0);
                    قائمة_الحقوق.Text = SuffReplacements(dr["قائمة_الحقوق"].ToString(), 0, 0);
                    اللغة.Text = dr["Lang"].ToString();
                    AddContextID = dr["ID"].ToString();
                    payment.Text = dr["payment"].ToString();
                    if (الموضوع.SelectedIndex == 1)
                    {
                        الأهلية.Text = SuffReplacements(dr["الأهلية"].ToString(), 0, 0);
                        الأهلية.Visible = true;
                    }
                    else الأهلية.Visible = false;
                    titleDefault.Text = dr["titleDefault"].ToString();
                    if (اللغة.Text == "الانجليزية")
                        اللغة.Checked = true;
                    else اللغة.Checked = false;

                    if (payment.Text == "معفي")
                        payment.Checked = true;
                    else
                    {
                        payment.Checked = false;
                        payment.Text = "عادي";
                    }


                    foreach (Control control in PanelItemsboxes.Controls) 
                    {
                        try
                        {
                            control.Text = dr[control.Name].ToString();
                        }
                        catch (Exception ex) { 
                        }
                        if (control.Name.Contains("نص") && control.Name.Contains("comb"))
                        {
                            //MessageBox.Show(control.Name.Replace("نص_", "Option"));
                            try
                            {
                                control.Text = dr[control.Name.Replace("نص_", "") + "Option"].ToString();
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                    }
                }
        }
        private void checkStarTextExist(string dataSource, string col, string genTable)
        {
            deletDup(col);
            string query = "select ID," + col + " from " + genTable;
            Console.WriteLine(query);
           // MessageBox.Show("checkStarTextExist " + query);
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
            catch (Exception ex)
            {
//                عدد_النماذج.Text = "عدد النماذج " + count.ToString();
                return;
            }

            foreach (Control control in panel_النص.Controls)
            {
                control.Visible = false;
                control.Name = "unvalid";
            }

            foreach (DataRow row in dtbl.Rows)
            {

                try
                {
                    if (row[col].ToString() != "" && !row[col].ToString().Contains("removed"))
                    {
                        ViewArchShow(SuffReplacements(row[col].ToString(), 0, 0), row["ID"].ToString());
                        count++;
                    }
                }
                catch (Exception ex)
                {
                    //عدد_النماذج.Text = "عدد النماذج " + count.ToString();
                    return;
                }
            }
            //عدد_النماذج.Text = "عدد النماذج " + count.ToString();
            sqlCon.Close();
        }

        private void قائمة_النصوص_الفرعية_SelectedIndexChanged(object sender, EventArgs e)
        {
            setCheclList();
            if (الموضوع.SelectedIndex == 0)
            {
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                
            }
            else if (الموضوع.SelectedIndex == 1)
            {
                selectTable = "TableAuthStarText";
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
                selectTable = "TableAuthRightStarText";
                checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);
            }
            
            view_PreReq(false);

            if (الموضوع.SelectedIndex == 1)
            {
                PopulateCheckBoxes(قائمة_النصوص_الفرعية.Text.Replace(" ", "_").Replace("-", "_") + "_" + قائمة_النصوص_العامة.Text.ToString(), "TableAuthRights", DataSource);
            }
                flllPanelItemsboxes();
            
            معاملات_السابق.Visible = معاملات_التالي.Visible = جميع_المعاملات.Visible = false;
            
        }

        private void OpenFile(string documenNo)
        {
            string query = "SELECT ID, proForm1,توضيح_المعاملة,Data1, Extension1,المعاملات_المرتبطة from TableProcReq where المعاملة=N'" + documenNo + "'";            
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);

            if (Con.State == ConnectionState.Closed)
                Con.Open();
            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                string str = reader["proForm1"].ToString();
                try
                {
                    var Data = (byte[])reader["Data1"];
                }
                catch (Exception ex)
                {
                    return ;
                }
            }
            Con.Close();
            
        }

        private void view_PreReq(bool view)
        {
            OpenFile(قائمة_النصوص_العامة.Text + "-" + قائمة_النصوص_الفرعية.Text, view);
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
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * FROM TableProcReq where  المعاملة=N'" + قائمة_النصوص_العامة.Text + "-" + قائمة_النصوص_الفرعية.Text + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            idList = 0;
            if (dtbl.Rows.Count > 0)
            {


                checlList[3] = "";//c
                
                foreach (DataRow row in dtbl.Rows)
                {
                    idList = Convert.ToInt32(row["ID"].ToString());
                    if (row["proForm1"].ToString() != "")
                        checlList[4] = "";//c

                    if (row["المطلوب_رقم1"].ToString() != "")
                        btnUploadFroms.Visible = true;
                    else
                    {
                        btnUploadFroms.Visible = reviewForms.Visible = false;
                        //MessageBox.Show(row["proForm1"].ToString());
                    }

                    if (CurrentFile != "")
                        reviewForms.Visible = true;
                    else
                        reviewForms.Visible = false;
                    //MessageBox.Show(row["توضيح_المعاملة"].ToString());

                    if (row["توضيح_المعاملة"].ToString() != "")
                    {
                        checlList[5] = "";
                        txtExplain.Text = row["توضيح_المعاملة"].ToString();
                    }

                    for (int index = 2; index < 11; index++)
                    {
                        foreach (Control control in panel_المستندات.Controls)
                        {
                            if (control.Name == colList[index])
                            {
                                control.Text = row[colList[index]].ToString();
                            }
                        }
                    }
                }
            }
            //MessageBox.Show(checlList[3]);
        }
        
        private bool view_PreReqID(int ID)
        {
            
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

            bool rowFound = false;

            

            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * FROM TableProcReq where  ID=N'"+ID+"'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();            
            if (dtbl.Rows.Count > 0)
            {
                rowFound = true;

                checlList[3] = "";//c
                foreach (DataRow row in dtbl.Rows)
                {
                    idList = Convert.ToInt32(row["ID"].ToString());
                    labID.Text = row["ID"].ToString();
                    if (row["proForm1"].ToString() != "")
                        checlList[4] = "";//c
                    
                    if (row["توضيح_المعاملة"].ToString() != "")
                        checlList[5] = "";
                    جميع_المعاملات.Text = row[colList[0]].ToString();
                    for (int index = 2; index < 11; index++)
                    {
                        foreach (Control control in panel_المستندات.Controls)
                        {
                            if (control.Name == colList[index])
                            {
                                control.Text = row[colList[index]].ToString();
                            }
                        }
                    }
                }
            }
            OpenFile(جميع_المعاملات.Text, false);
            return rowFound;
        }
        
        private bool view_ProDocID(int ID)
        {
            //MessageBox.Show(ID.ToString());
            bool rowFound = false; 
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return false; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * FROM TableProcReq where  ID=N'"+ID+"'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();            
            if (dtbl.Rows.Count > 0)
            {
                rowFound = true;
                foreach (DataRow row in dtbl.Rows)
                {
                    TableProcReqID = row["ID"].ToString();
                    labID.Text = row["ID"].ToString();
                    //MessageBox.Show(row["توضيح_المعاملة"].ToString());
                    txtExplain.Text = row["توضيح_المعاملة"].ToString();
                    جميع_المعاملات.Text = row["المعاملة"].ToString();
                }
            }
            return rowFound;
        }

        private int[] getIDs(string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) {  }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT ID FROM " + table, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            IDs = new int[dtbl.Rows.Count];
            int index = 0;
            IDsMax = IDsMin = 1;
            //reviewForms.Visible = btnUploadFroms.Visible = false;
            foreach (DataRow row in dtbl.Rows)
            {
                IDs[index] = Convert.ToInt32(row["ID"].ToString());
                labID.Text = row["ID"].ToString();
                if (IDs[index] > IDsMax)
                    IDsMax = IDs[index];
                if (IDs[index] < IDsMin)
                    IDsMin = IDs[index];
                index++;
            }
            //MessageBox.Show(IDsMax.ToString());
            return IDs;
        }
            private string OpenFile(string documenNo, bool printOut)
        {
            string query = "SELECT ID, proForm1,Data1, Extension1 from TableProcReq where المعاملة = N'"+ documenNo+"'";
            Console.WriteLine(query);
            reviewForms.Enabled = false;
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                string str = reader["proForm1"].ToString();
                //MessageBox.Show(str);
                try
                {
                    var Data = (byte[])reader["Data1"];

                    CurrentFile = ArchFile + @"\formUpdated\" + str + ".docx";
                    string filePath = ArchFile + @"\" + str + ".docx";
                    if (File.Exists(CurrentFile) && !fileIsOpen(CurrentFile))
                    {
                        File.Delete(CurrentFile);
                    }
                    if (!File.Exists(CurrentFile))
                    {
                        try
                        {
                            File.WriteAllBytes(filePath, Data);
                            System.IO.File.Copy(filePath, CurrentFile);
                            FileInfo fileInfo = new FileInfo(CurrentFile);
                            if (fileInfo.IsReadOnly) fileInfo.IsReadOnly = false;
                            Console.WriteLine("CurrentFile " + CurrentFile);

                            if (printOut)
                            {                                
                                System.Diagnostics.Process.Start(CurrentFile);
                            }
                            reviewForms.Enabled = true; 
                            return CurrentFile;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("fail " + str);
                            return "";
                        }

                    }
                    else if (File.Exists(CurrentFile) && fileIsOpen(CurrentFile))
                    {
                        MessageBox.Show("يرجى إغلاق الملف " + str + " أولا");
                    }
                }
                catch (Exception ex)
                {
                    return "";
                }
            }
            
            Con.Close();
            return "";
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
       


        private void تعيين_كخيار_Click(object sender, EventArgs e)
        {
            
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string query = "update "+selectTable+" set "+ قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + " =N'"+ النص .Text+"removed"+ "'where ID = '" + starButton + "'";
            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            النص.Text = "";
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //insert = false;
            //if (الموضوع.SelectedIndex < 2)
            //    //updateAllFields();

            
        }

        private string updateText(bool newData, string table, string text)
        {
            if (!checkColExistance(table, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_")))
                CreateColumn(قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), table, "max");

            text = SuffReversReplacements(text, 0, 0);
            string query = "insert into " + table + " (" + قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + ") values (N'" + text.Replace("'", "!") + "Star');SELECT @@IDENTITY as lastid";
           
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;

            if (!newData)
                sqlCmd.ExecuteNonQuery();
            else
            {
                var reader = sqlCmd.ExecuteReader();
                if (reader.Read())
                {
                    starButton =  reader["lastid"].ToString();
                }
                sqlCon.Close();
            }
            return starButton;
        }

        private void updateAllFields(string text, string lawfull, string right)
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
                foreach (Control control in PanelItemsboxes.Controls)
                {
                    if ((row["name"].ToString() == control.Name || row["name"].ToString() == control.Name + "Option") && control.Visible)
                    {
                        allList[updateAllIndex] = row["name"].ToString();
                         updateValues = updateValues + "," + row["name"].ToString() + "=@" + row["name"].ToString();
                        
                        updateAllIndex++;
                    }
                }
            }
            queryAll = "UPDATE TableAddContext SET payment=@payment,الأهلية=@الأهلية ,قائمة_الحقوق=@قائمة_الحقوق,titleDefault=@titleDefault,Lang=@Lang,TextModel=@TextModel,ColName=@ColName,ColRight=@ColRight,altColName=@altColName,altSubColName=@altSubColName" + updateValues + " WHERE ID = " + AddContextID;
            Console.WriteLine(queryAll);
            save2DataBase(PanelItemsboxes, updateAllIndex, text, lawfull, right);
        }
        
        private void insertAllFields(string text, string lawfull, string right)
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
            insertAllIndex = 0;
            string insertItems = "TextModel,ColName,ColRight,altColName,altSubColName,Lang,titleDefault,الأهلية,قائمة_الحقوق,payment";
            string insertValues = "@TextModel,@ColName,@ColRight,@altColName,@altSubColName,@Lang,@titleDefault,@الأهلية,@قائمة_الحقوق,@payment";
            foreach (DataRow row in dtbl.Rows)
            {
                foreach (Control control in PanelItemsboxes.Controls)
                {
                    if ((row["name"].ToString() == control.Name || row["name"].ToString() == control.Name + "Option") && control.Visible)
                    {
                        allList[insertAllIndex] = row["name"].ToString();
                        insertItems = insertItems + ","+ row["name"].ToString();
                            insertValues = insertValues +",@" + row["name"].ToString();
                        
                        insertAllIndex++;
                    }
                }
            }
            queryAll = "insert into TableAddContext (" + insertItems + ") values (" + insertValues + ");SELECT @@IDENTITY as lastid";
            save2DataBase(PanelItemsboxes, insertAllIndex, text, lawfull, right);
            
        }

        

        private bool save2DataBase(FlowLayoutPanel panel, int index, string text, string lawfull, string right)
        {
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(queryAll, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            for (int i = 0; i < index; i++)
            {
                foreach (Control control in panel.Controls)
                {
                    if (control.Visible)
                    {
                        if (control.Name == allList[i] || (allList[i].Contains("Option") && control.Name == "نص_" + allList[i].Replace("Option", "")))
                        {
                            sqlCommand.Parameters.AddWithValue("@" + allList[i], control.Text);

                        }
                    }
                }
            }
            StaticParameters(sqlCommand, text);
            sqlCommand.Parameters.AddWithValue("@payment", payment.Text);
            sqlCommand.Parameters.AddWithValue("@Lang", اللغة.Text);
            sqlCommand.Parameters.AddWithValue("@الأهلية", lawfull);
            sqlCommand.Parameters.AddWithValue("@قائمة_الحقوق", right);
            sqlCommand.Parameters.AddWithValue("@titleDefault", titleDefault.Text);
            //MessageBox.Show(queryAll);
            
            //sqlCommand.ExecuteNonQuery();
            if (queryAll.Contains("UPDATE"))
                sqlCommand.ExecuteNonQuery();
            else
            {
                var reader = sqlCommand.ExecuteReader();
                if (reader.Read())
                {
                    AddContextID =reader["lastid"].ToString();
                }
                sqlConnection.Close();
            }
            return true;
        }

        private void StaticParameters(SqlCommand sqlCommand, string text)
        {
            if (الموضوع.SelectedIndex == 0)
                sqlCommand.Parameters.AddWithValue("@ColRight", "");
            else
                sqlCommand.Parameters.AddWithValue("@ColRight", قائمة_النصوص_الفرعية.Text +"_"+ قائمة_النصوص_العامة.SelectedIndex.ToString());
            sqlCommand.Parameters.AddWithValue("@ColName", قائمة_النصوص_الفرعية.Text + "-" + قائمة_النصوص_العامة.SelectedIndex.ToString());
            sqlCommand.Parameters.AddWithValue("@TextModel", text);
            sqlCommand.Parameters.AddWithValue("@altSubColName", قائمة_النصوص_الفرعية.Text);
            sqlCommand.Parameters.AddWithValue("@altColName", قائمة_النصوص_العامة.Text);
        }

        private void addMainAuth( string col, string colText)
        {
            string ID = getID(DataSource, col);
            string query = "update TableListCombo set " + col + "=N'" + colText + "' where ID = " + ID;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }
            Console.WriteLine("addMainAuth " + query);
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            insert = true;
            if (النص.Text != "")
                checlList[1] = "";
            string RevText = SuffReversReplacements(النص.Text, 0, 0);
            string lawfull = SuffReversReplacements(الأهلية.Text, 0, 0);
            string rights = SuffReversReplacements(قائمة_الحقوق.Text, 0, 0);
            if (!checkRowExist(DataSource, قائمة_النصوص_العامة.Text, قائمة_النصوص_الفرعية.Text))
            {
                //MessageBox.Show("new");
                idList = 0;
                insertAllFields(RevText, lawfull, rights);
                if (الموضوع.SelectedIndex == 0)
                {
                    if (!checkColExistance(selectTable, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_")))
                        CreateColumn(قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable, "max");
                }
                else if (الموضوع.SelectedIndex == 1)
                {
                    selectTable = "TableAuthStarText";
                    if (!checkColExistance(selectTable, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_")))
                        CreateColumn(قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable, "max");

                    selectTable = "TableAuthRightStarText";
                    if (!checkColExistance(selectTable, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_")))
                        CreateColumn(قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable, "max");
                }


                string query = "";
                if (الموضوع.SelectedIndex == 0)
                {
                    query = "insert into TableCollectStarText (" + قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + ") values (N'" + النص.Text.Replace("'", "!") + "')";
                    insertData(query);
                }
                else if (الموضوع.SelectedIndex == 1)
                {
                    query = "insert into TableAuthStarText (" + قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + ") values (N'" + النص.Text.Replace("'", "!") + "')";
                    insertData(query);
                    query = "insert into TableAuthRightStarText (" + قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_") + ") values (N'" + قائمة_الحقوق.Text.Replace("'", "!") + "')";
                    insertData(query);
                }
               // this.Close();

            }
            else
            {
                updateAllFields(RevText, lawfull, rights);
                if (الموضوع.SelectedIndex == 0)
                {
                    updateText(false, "TableCollectStarText", RevText);
                }
                else if (الموضوع.SelectedIndex == 1)
                {
                    updateText(false, "TableAuthStarText", RevText);
                    updateText(false, "TableAuthsub", rights);
                }
            }
            goPanel();
            //this.Close();
        }

        private void insertData(string query) {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
        }
            private void deletDup(string col)
        {
            string table = "TableCollectStarText";
            if (الموضوع.SelectedIndex == 1)
                table = "TableAuthStarText";
            if (الموضوع.SelectedIndex == 2)
                table = "TableAuthsub";

            
            
            string dupQuery = "update "+ table+ " set " + col + " = null where ID in(SELECT ID FROM " + table +
                       " WHERE  " + col + " is not null " +
                       " and ID NOT IN (SELECT miN(ID) " +
                       " FROM "+ table+
                       " GROUP BY " + col + ")) ";

            
            Console.WriteLine("dupQuery " + dupQuery);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(dupQuery, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            try
            {
                sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex) { }
            
        }

        /*
         النص
المستندات المطلوبة واستمارة الطلب
المستندات النهائية للارشفة
الاستمارات الأولية
توضيح المعاملات
الرموز المرجعية
         */
        private bool  goPanel() {
            for (int i = 0; i < 6; i++)
                if (checlList[i] != "")
                {
                    if (i == 4)
                    {
                        otherPro.SelectedIndex = 1;
                        otherPro.Text = "المستندات المطلوبة للإجراء";
                        MessageBox.Show(checlList[i]);
                        return true;
                    }
                    else if (i == 3 )
                    {
                        btnUploadFroms.Visible = true;
                        otherPro.SelectedIndex = 1;
                        otherPro.Text = "المستندات المطلوبة للإجراء";
                        MessageBox.Show(checlList[i]);
                        return true;
                    }
                    
                    else if (i == 5)
                    {
                        otherPro.SelectedIndex = 4;
                        otherPro.Text = "توضيح المعاملات";
                        dataGridView2.Visible = false;
                        MessageBox.Show(checlList[i]);
                        return true;
                    }                    

                }
            this.Close();
            return false;
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
            foreach (DataRow row in dtbl.Rows)
            {
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
            if (text.Contains("8"))
                text = text.Replace("حقل8", "itext8");
            if (text.Contains("حق9"))
                text = text.Replace("حقل9", "itext9");
            if (text.Contains("حقل0"))
                text = text.Replace("حقل0", "itext0");

            if (text.Contains("خ_ثنائي1"))
                text = text.Replace("خ_ثنائي1", "icheck1");
            if (text.Contains("خ_ثنائي2"))
                text = text.Replace("خ_ثنائي2", "icheck2");
            if (text.Contains("خ_ثنائي3"))
                text = text.Replace("خ_ثنائي3", "icheck3");
            if (text.Contains("خ_ثنائي4"))
                text = text.Replace("خ_ثنائي4", "icheck4");
            if (text.Contains("خ_ثنائي5"))
                text = text.Replace("خ_ثنائي5", "icheck5");

            if (text.Contains("خ_متعدد1"))
                text = text.Replace("خ_متعدد1", "icombo1");
            if (text.Contains("خ_متعدد2"))
                text = text.Replace("خ_متعدد2", "icombo2");
            if (text.Contains("خ_متعدد3"))
                text = text.Replace("خ_متعدد3", "icombo3");
            if (text.Contains("خ_متعدد4"))
                text = text.Replace("خ_متعدد4", "icombo4");
            if (text.Contains("خ_متعدد5"))
                text = text.Replace("خ_متعدد5", "icombo5");

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
                    CreateColumn(text, "TableAddContext", "max");

                    if (!text.Contains("btnAdd") && !text.Contains("icheck"))
                        CreateColumn(text + "Length", "TableAddContext", "5");
                    if (text.Contains("icombo") || !text.Contains("icheck"))
                        CreateColumn(text + "Option", "TableAddContext", "max");
                }
            }
            //Console.WriteLine (text);
            //MessageBox.Show(text);
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
            SqlCommand sqlCmd = new SqlCommand("alter table " + tableName + " add " + Columnname + " nvarchar(" + size + ")", sqlCon);
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
            
            foreach (Control control in panel_fields.Controls)
            {
                control.Enabled = true;
                for (int index = 0; listFiels[index] != ""; index++)
                {
                    if (listFiels[index] == control.Text && النص.Text.Contains(listFiels[index]))
                    {
                        control.Enabled = false;
                        Console.WriteLine(control.Text);
                    }
                    else if (listFiels[index] == control.Text && الأهلية.Text.Contains(listFiels[index]))
                    {
                        control.Enabled = false;
                        Console.WriteLine(control.Text);
                    }
                }
            }
            for (int index = 0; listFiels[index] != ""; index++)
            {
                if (النص.Text.Contains(listFiels[index]))
                {                   
                    panelFill(DataSource, getTableName(listFiels[index]));
                }
                if (الأهلية.Text.Contains(listFiels[index]))
                {                   
                    panelFill(DataSource, getTableName(listFiels[index]));
                }
            }
            
        }

        public string getID(string dataSource, string col)
        {
            string query = "select cast( max(ID) as int) + 1  as idCount from TableListCombo where "+ col + " is not null";
            string id = "1";
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
                id = row["idCount"].ToString();               
            }
            if (id == "") id = "1";
            return id;
        }
            public void panelFill(string dataSource, string field)
        {
            string query = "select * from TableAddContext where altColName = N'" + قائمة_النصوص_الفرعية.Text + "' and altSubColName = N'" + قائمة_النصوص_العامة.Text+ "'";

            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            Console.WriteLine("panelFill " + query);
            foreach (Control control in PanelItemsboxes.Controls)
            {
                if (control.Name == field)
                {
                    control.Visible = true;
                    control.Width = (control.Text.Length * 8) + 1;
                    if (control.Width < 100)
                        control.Width = 100;
                    foreach (Control controlText in PanelItemsboxes.Controls)
                    {
                        if (controlText.Name == "نص_" + field)
                        {
                            controlText.Visible = true;
                            controlText.Width = (controlText.Text.Length * 8) + 1;
                            if (controlText.Width < 100)
                                controlText.Width = 100;
                        }
                    }
                }
            }

            foreach (DataRow row in dtbl.Rows)
            {
                foreach (Control control in PanelItemsboxes.Controls)
                {
                    if (control.Name == field)
                    {
                        control.Text = row[field].ToString();

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
                            }
                        }
                    }

                }
            }
        }

        private void الموضوع_SelectedIndexChanged(object sender, EventArgs e)
        {
            label3.Visible =  false;
            if (الموضوع.SelectedIndex == 0)
            {
                selectTable = "TableCollectStarText";
                setCheclList();
                otherPro.Items.Clear();
                otherPro.Items.Add("النص");
                otherPro.Items.Add("المستندات المطلوبة للإجراء");
                otherPro.Items.Add("المستندات النهائية للارشفة");
                otherPro.Items.Add("استمارة الطلب");
                otherPro.Items.Add("الرموز المرجعية");
                
                
                //checlList[0] = "";
                //checlList[1] = "نص موضوع المكاتبة غير موجود";
                //checlList[2] = "";
                //checlList[4] = "استمارة الطلب غير موجودة";//c
                //checlList[3] = "المطلوبات الأولية غير محددة";//c
                //checlList[6] = "المطلوبات النهائية غير محددة";
                //checlList[5] = "شرح المعاملة غير موجودة";
                قائمة_الحقوق.Text = "";
                الأهلية.Text = "";
                قائمة_الحقوق.Visible = label8.Visible = label3.Visible = label8.Visible = الأهلية.Visible = false;
                النص.Size = new System.Drawing.Size(680, 361);
                
                label7.Text = "النص";
            }
            else if (الموضوع.SelectedIndex == 1)
            {
                setCheclList();
                selectTable = "TableAuthStarText";
                النص.Size = new System.Drawing.Size(680, 117);

                قائمة_الحقوق.Visible = label8.Visible = label3.Visible = label8.Visible = الأهلية.Visible = true;
                label7.Text = "موضوع الانابة:";

            }
            else if (الموضوع.SelectedIndex == 2)
            {
                selectTable = "TableAuthRightStarText";
                label3.Visible =true;

            }

            if (الموضوع.SelectedIndex != 0)
            {
                otherPro.Items.Clear();
                otherPro.Items.Add("النص");
                otherPro.Items.Add("قوائم الحقوق");
                otherPro.Items.Add("المستندات المطلوبة للإجراء");
                otherPro.Items.Add("المستندات النهائية للارشفة");
                otherPro.Items.Add("استمارة الطلب");
                otherPro.Items.Add("الرموز المرجعية");
                setCheclList();
                checlList[1] = "";

            }

            otherPro.Items.Add("توضيح المعاملات");
            if (الموضوع.SelectedIndex == 0)
            {
                fileComboBox(قائمة_النصوص_العامة, DataSource, "altColName", "= '' ", true);                
            }
            else
            {
                fileComboBox(قائمة_النصوص_العامة, DataSource, "altColName", "<> ''", true);
            }

            if (قائمة_النصوص_العامة.Text != "" && قائمة_النصوص_الفرعية.Text != "")
            {
                if (قائمة_النصوص_العامة.Text != "" && قائمة_النصوص_الفرعية.Text != "")
                {
                    checkStarTextExist(DataSource, قائمة_النصوص_العامة.Text.Replace(" ", "_") + "_" + قائمة_النصوص_الفرعية.Text.Replace(" ", "_"), selectTable);                    
                }
                view_PreReq(false);

                if (الموضوع.SelectedIndex == 1)
                {
                    PopulateCheckBoxes(قائمة_النصوص_الفرعية.Text.Replace(" ", "_").Replace("-", "_") + "_" + قائمة_النصوص_العامة.Text.ToString(), "TableAuthRights", DataSource);
                }
                    flllPanelItemsboxes();
                
                معاملات_السابق.Visible = معاملات_التالي.Visible = جميع_المعاملات.Visible = false;
            }
            
        }

        private void ColRight_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateCheckBoxes(txtRights.Text.Replace(" ", "_").Replace("-", "_"), "TableAuthRights", DataSource);
        }
        public void PopulateCheckBoxes(string col, string table, string dataSource)
        {

            if (col == "الحقوق" || col == "Col" || col == "" || table == "" || dataSource == "") return;
            string query = "SELECT ID," + col.Replace(" ", "_").Replace("-", "_") + " FROM " + table + " where ID  = 22";
            Console.WriteLine(query);
            //MessageBox.Show(query);
            string authother = "";
            string removeAuthother = "";
            string lastSentence = "";
            txtRights.Text = "";
            panel_الحقوق.BringToFront();
            panel_الحقوق.Visible = true;

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
                sqlCon.Close();
                int rowIndex = 0;
                foreach (DataRow row in dtbl.Rows)
                {
                    قائمة_الحقوق.Text = removeSpace(row[col.Replace(" ", "_").Replace("-", "_")].ToString());
                    //MessageBox.Show(قائمة_الحقوق.Text);
                    if (rowIndex != 0)
                    {

                        setCheclList();
                        checlList[2] = "";

                        string[] Text_statis = row[col.Replace("-", "_")].ToString().Split('_');
                        
                        if (row[col.Replace("-", "_")].ToString() == "") continue;

                        string text = SuffReplacements(Text_statis[0], 0, 0);
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
                            //MessageBox.Show(txtRights.Text);
                        }
                        catch (Exception ex) { }
                    }
                    rowIndex++;
                }
            }
            catch (Exception ex) { }

            txtRights.Text = removeSpace(txtRights.Text);
        }

        private void picStar_Click(object sender, EventArgs e)
        {

        }

        private void otherPro_SelectedIndexChanged(object sender, EventArgs e)
        {
            معاملات_السابق.Visible = معاملات_التالي.Visible = جميع_المعاملات.Visible = false;
            panel_النص.Visible = false;
            panel_الحقوق.Visible = false;
            panel_المستندات.Visible = false;
            dataGridView2.Visible = false;
            dataGridView4.Visible = false;
            Panel_الرموز.Visible = false;
            if (otherPro.Text == "النص")
            {
                panel_النص.Visible = true;
                panel_النص.Size = new System.Drawing.Size(502, 478);
                panel_النص.BringToFront();
                panel_النص.Location = new System.Drawing.Point(4, 50);
            }
            else if (otherPro.Text == "قوائم الحقوق")
            {
                
                panel_الحقوق.Visible = true;
                panel_الحقوق.Size = new System.Drawing.Size(502, 478);
                panel_الحقوق.BringToFront();
                panel_الحقوق.Location = new System.Drawing.Point(4,  50);
            }
            else if (otherPro.Text == "المستندات المطلوبة للإجراء")
            {
                

                panel_المستندات.Visible = true;
                panel_المستندات.Size = new System.Drawing.Size(502, 478);
                panel_المستندات.BringToFront();
                panel_المستندات.Location = new System.Drawing.Point(4,  50);
                getIDs("TableProcReq");
               
                
            }
            
            else if (otherPro.Text == "توضيح المعاملات")
            {
                //dataGridView2.Visible = true;                
                //dataGridView2.BringToFront();
                //dataGridView2.Location = new System.Drawing.Point(4,  50);
                panel_شرح.Visible = true;
                panel_شرح.BringToFront();
                getIDs("TableProcReq");
                FillDataGridViewReq("TableProcReq", dataGridView2);
            }
            else if (otherPro.Text == "الرموز المرجعية")
            {
                dataGridView4.Visible = true;
                dataGridView4.BringToFront();                
                Panel_الرموز.Visible = true;
                Panel_الرموز.BringToFront();
                getIDs("Tablechar");
                FillDataGridViewCode("Tablechar", dataGridView4);
            }
        }

        private void ColorFulGrid2()
        {
            for (int i=0; i < dataGridView2.Rows.Count - 1; i++)
            {
                dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.White;
                if (dataGridView2.Rows[i].Cells[2].Value.ToString() != "")
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
            }
        }
        void FillDataGridViewReq(string table, DataGridView dataGridView)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }

            SqlDataAdapter sqlDa = new SqlDataAdapter("select ID, المعاملة, توضيح_المعاملة from " + table + " order by  المعاملة asc", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView.DataSource = dtbl;
            //dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            //dataGridView1.Columns["ID"].Visible = false;
            try
            {
                dataGridView.Columns[0].Visible = dataGridView.Columns[2].Visible = false;
                dataGridView.Columns[1].Width = 550;

            }
            catch (Exception ex) { }
            sqlCon.Close();
        }
        
        void FillDataGridViewCode(string table, DataGridView dataGridView)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex) { return; }

            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from " + table + " order by ID asc", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView.DataSource = dtbl;
            //dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            //dataGridView1.Columns["ID"].Visible = false;
            try
            {
                dataGridView.Columns[0].Visible = false;
                dataGridView.Columns[1].Width = 100;
                dataGridView.Columns[2].Width = dataGridView.Columns[3].Width = dataGridView.Columns[4].Width = 70;
                dataGridView.Columns[4].Width = dataGridView.Columns[5].Width = dataGridView.Columns[6].Width = dataGridView.Columns[7].Width = dataGridView.Columns[8].Width = 70;

            }
            catch (Exception ex) { }
            sqlCon.Close();
        }

       
        private void panelPro(string name)
        {
            foreach (Control control in this.Controls)
            {
                if (control.Name == "panel_" + name.Trim())
                {
                    //MessageBox.Show(control.Name);
                    control.Visible = true;
                    control.Size = new System.Drawing.Size(667, 478);
                    control.BringToFront();
                }
                else
                {
                    //control.Visible = false;
                }
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            string[] data = new string[11];
            string[] colList = new string[11];
            if(الموضوع.SelectedIndex == 0)
                data[0] = "10";
            else if(الموضوع.SelectedIndex != 0)
                data[0] = "12";
            data[1] = قائمة_النصوص_العامة.Text + "-" + قائمة_النصوص_الفرعية.Text;

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
            for (int index = 2; index < 11; index++)
            {
                foreach (Control control in panel_المستندات.Controls)
                {
                    if (control.Name == colList[index])
                    {
                        data[index] = control.Text;
                    }
                }
            }
            updatetRow(idList, DataSource, data);
            foreach (Control control in panel_المستندات.Controls)
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

        private void reviewForms_Click(object sender, EventArgs e)
        {
            if (CurrentFile != "")
                try
                {
                    System.Diagnostics.Process.Start(CurrentFile);
                }
                catch (Exception ex) { }
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
            if (الموضوع.SelectedIndex == 0)
                data[0] = "10";
            else if (الموضوع.SelectedIndex != 0)
                data[0] = "12";
            if (جميع_المعاملات.Text == "")
                data[1] = قائمة_النصوص_العامة.Text + "-" + قائمة_النصوص_الفرعية.Text;
            else 
                data[1] = جميع_المعاملات.Text;


            for (int index = 2; index < 11; index++)
            {
                foreach (Control control in panel_المستندات.Controls)
                {
                    if (control.Name == colList[index])
                    {
                        //MessageBox.Show(control.Name +" - "+ colList[index]);
                        
                        data[index] = control.Text;
                    }
                }
            }
            if (idList == 0)
            {
                insertRow(DataSource, data);
                btnUploadFroms.Visible = true;
            }
            else
                updatetRow(idList, DataSource, data);

            TableProcReqID = idList.ToString();
            foreach (Control control in panel_المستندات.Controls)
            {
                if (control is Button && (control.Name.Contains("المطلوب_رقم") || control.Name.Contains("btnReq")))
                {
                    control.Text = "";
                }
            }
            if (معاملات_التالي.Visible)
                معاملات_التالي.PerformClick();
            checlList[3] = "";
            goPanel();
        }

        private void insertRow(string source, string[] data)
        {
            SqlConnection sqlCon = new SqlConnection(source);
            string[] colList = new string[11];
            colList[1] = "المعاملة";
            colList[0] = "رقم_المعاملة";
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

            string query = "INSERT INTO TableProcReq (" + item + ") values (" + value + ");SELECT @@IDENTITY as lastid";

            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            }
            catch (Exception ex)
            {
                Console.WriteLine(query);
                return; }
            sqlCmd.CommandType = CommandType.Text;
            Console.WriteLine(query);
            for (int col = 0; col < 11; col++)
            {
                sqlCmd.Parameters.AddWithValue(colList[col], data[col]);
            }
            try
            {
                var reader = sqlCmd.ExecuteReader();
                if (reader.Read())
                {
                    idList =  Convert.ToInt32(reader["lastid"].ToString());
                }

            }
            catch (Exception ex)
            {
                idList = 0;
                MessageBox.Show(query);
            }
            
            sqlCon.Close();
        }

        

        private void btnUploadFroms_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                CurrentFile = @dlg.FileName;                
                uploadFormsReq(CurrentFile);
                reviewForms.Visible = true;
            }
            checlList[4] = "";
            if (!goPanel())
            النص.Text = "";
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
                    //string query = "UPDATE TableProcReq SET Data1=@Data1,proForm1=@proForm1 WHERE ID=" + AddContextID;
                    string query = "UPDATE TableProcReq SET Data1=@Data1,proForm1=@proForm1 WHERE ID=" + idList.ToString();
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
                    sqlCmd.Parameters.Add("@proForm1", SqlDbType.NVarChar).Value = قائمة_النصوص_العامة.Text + "-" + قائمة_النصوص_الفرعية.Text;
                    sqlCmd.ExecuteNonQuery();
                    sqlCon.Close();
                    checlList[4] = "";//c
                    label1.Visible = true;
                    return;
                }
            }
        }

        private void الموضوع_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                قائمة_النصوص_العامة.Select();
            }
        }

        private void قائمة_النصوص_العامة_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                قائمة_النصوص_الفرعية.Select();
            }
        }

        private void قائمة_النصوص_الفرعية_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                otherPro.Select();
            }
        }

        private void نص_itext1_TextChanged(object sender, EventArgs e)
        {

        }

        

        private void timer1_Tick(object sender, EventArgs e)
        {
            ColorFulGrid2();
        }

        private void dataGridView2_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.Rows.Count > 1)
            {
                try
                {
                    txtExplain.Text = dataGridView2.CurrentRow.Cells["توضيح_المعاملة"].Value.ToString();
                    if (dataGridView2.CurrentRow.Cells["المعاملة"].Value.ToString() == "")
                        btnDeleteReq.Enabled = true;
                    else 
                        btnDeleteReq.Enabled = true;
                    
                    dataGridView2.Visible = false;
                    TableProcReqID = dataGridView2.CurrentRow.Cells["ID"].Value.ToString();
                    panel_شرح.Visible = true;
                    panel_شرح.Size = new System.Drawing.Size(667, 478);
                    panel_شرح.BringToFront();
                    panel_شرح.Location = new System.Drawing.Point(4,  50);

                }
                catch (Exception ex) { }
            }
        }

        private void btnExpl_Click(object sender, EventArgs e)
        {
            if (btnExpl.Text == "")
            {
                
            }
            string query = "UPDATE TableProcReq SET توضيح_المعاملة=N'" + txtExplain.Text + "'  where  ID = " + TableProcReqID + "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            Console.WriteLine(query);
            sqlCmd.ExecuteNonQuery();
            if (معاملات_التالي.Visible)
                معاملات_التالي.PerformClick();
            checlList[5] = "";
            MessageBox.Show("تمت الإضافة بنجاح");
            this.Close();
            //button2.PerformClick();
            //if (!goPanel())
            //    النص.Text = "";
        }

        private void button15_Click(object sender, EventArgs e)
        {
            string query = "delete from TableProcReq where ID = " + TableProcReqID + "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            button2.PerformClick();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            dataGridView2.Visible = true;
            dataGridView2.BringToFront();
            dataGridView2.Location = new System.Drawing.Point(4,50);

            FillDataGridViewReq("TableProcReq", dataGridView2);
        }

        private void المعاملة_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (قائمة_النصوص_العامة.Text == "" || قائمة_النصوص_الفرعية.Text == "") return;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableProcReq SET المعاملات_المرتبطة=N'" + المعاملة.Text + "'  where المعاملة = N'" + قائمة_النصوص_العامة.Text + "_" + قائمة_النصوص_الفرعية.Text + "'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();            
        }

        private void button17_Click(object sender, EventArgs e)
        {
            checlList[5] = "";
            if (!goPanel())
                النص.Text = "";
            this.Close();
        }

        private void معاملات_الجميع_Click(object sender, EventArgs e)
        {
            //معاملات_الجميع.Size = new System.Drawing.Size(90, 35);
            //معاملات_الجميع.Location = new System.Drawing.Point(578, 5);
            معاملات_السابق.Visible = معاملات_التالي.Visible = جميع_المعاملات.Visible = true;
            idList = 1;
            switch (otherPro.Text)
            {
                case "المستندات المطلوبة للإجراء":
                    while (!view_PreReqID(idList)) 
                        idList++;

                    getIDs("TableProcReq");
                    panel_المستندات.Visible = true;
                    panel_المستندات.BringToFront();
                break;
                
                case "توضيح المعاملات":
                    while (!view_ProDocID(idList))
                        idList++;

                    getIDs("TableProcReq");
                    panel_شرح.Visible = true;
                    panel_شرح.BringToFront();
                break;
            }
        }

        private void معاملات_السابق_Click(object sender, EventArgs e)
        {
            if (idList > IDsMin)
                idList--;
            else idList = 1;
            
            


            while (!getOverAllID()) {
                if (idList > IDsMin)
                    idList--;
                else idList = 1;
            }
        }
        private bool getOverAllID() {
            bool IDFound = false;
            switch (otherPro.Text)
            {
                case "المستندات المطلوبة للإجراء":
                    IDFound = view_PreReqID(idList);
                    break;

                case "توضيح المعاملات":
                    IDFound = view_ProDocID(idList);
                    break;
            }
            return IDFound;
        }
        private void معاملات_التالي_Click(object sender, EventArgs e)
        {           
            if (idList < IDsMax)
                idList++;
            else idList = 1;
            
            while (!getOverAllID()) {
                if (idList < IDsMax)
                    idList++;
                else idList = 1;
            }
        }

        private void button15_Click_1(object sender, EventArgs e)
        {
            string query = "delete from TableProcReq where ID = " + idList ;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            معاملات_التالي.PerformClick();
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView4.Visible = false;
            foreach(Control control in Panel_الرموز.Controls)
            if(control is TextBox)
                control.Text = dataGridView4.CurrentRow.Cells[control.Name].Value.ToString();
            
        }

        private void button23_Click(object sender, EventArgs e)
        {
            dataGridView4.Visible = true;
            dataGridView4.BringToFront();
            Panel_الرموز.Visible = true;
            Panel_الرموز.BringToFront();
            getIDs("Tablechar");
            FillDataGridViewCode("Tablechar", dataGridView4);
        }

        private void button28_Click(object sender, EventArgs e)
        {
            string query = "delete from Tablechar where ID = " + ID.Text;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            قائمة_الرموز.PerformClick();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            string colItemValue = "";
            int items = 0;
            foreach (Control control in Panel_الرموز.Controls)
                if (control is TextBox && control.Name != "ID")
                {
                    if (items == 0)
                        colItemValue = control.Name + " = N'" + control.Text + "'";
                    else
                        colItemValue = colItemValue + ", " + control.Name + " = N'" + control.Text + "'";
                    items++;
                }

            string query = "update Tablechar set " + colItemValue + " where ID = " + ID.Text;
            Console.WriteLine(query);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            foreach (Control control in Panel_الرموز.Controls)
                if (control is TextBox)
                    control.Text = "";
            قائمة_الرموز.PerformClick();
        }

        private void button23_Click_1(object sender, EventArgs e)
        {
            if (button23.Text == "إظهار خيارات الادخال") {
                button23.Text = "إخفاء خيارات الادخال";
                panel_fields.Visible = true;
                panel_fields.BringToFront();
            }
            else  {
                button23.Text = "إظهار خيارات الادخال";
                panel_fields.Visible = false;

            }
        }

        private void جميع_المعاملات_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void اللغة_CheckedChanged(object sender, EventArgs e)
        {
            if (!اللغة.Checked)
                اللغة.Text = "العربية";
            else 
                اللغة.Text = "الانجليزية";
        }

        private void الأهلية_TextChanged(object sender, EventArgs e)
        {
            foreach (Control control in PanelItemsboxes.Controls)
            {
                control.Visible = false;
            }

            foreach (Control control in panel_fields.Controls)
            {
                control.Enabled = true;
                for (int index = 0; listFiels[index] != ""; index++)
                {
                    if (listFiels[index] == control.Text && النص.Text.Contains(listFiels[index]))
                    {
                        control.Enabled = false;
                        Console.WriteLine(control.Text);
                    }
                    else if (listFiels[index] == control.Text && الأهلية.Text.Contains(listFiels[index]))
                    {
                        control.Enabled = false;
                        Console.WriteLine(control.Text);
                    }
                }
            }
            for (int index = 0; listFiels[index] != ""; index++)
            {
                if (النص.Text.Contains(listFiels[index]))
                {
                    panelFill(DataSource, getTableName(listFiels[index]));
                }
                if (الأهلية.Text.Contains(listFiels[index]))
                {
                    panelFill(DataSource, getTableName(listFiels[index]));
                }
            }
        }

        private void payment_CheckedChanged(object sender, EventArgs e)
        {
            if (payment.Checked)
                payment.Text = "معفي";
            else
                payment.Text = "عادي";
        }
    }
}
