using IronBarCode;
using SautinSoft.Document;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ZXing;
using Color = System.Drawing.Color;

using static Azure.Core.HttpHeader;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Bibliography;
using Aspose.Words.Settings;
using Control = System.Windows.Forms.Control;
using OfficeOpenXml;
using System.Data.SqlTypes;

namespace PersAhwal
{
    public partial class Accountant : Form
    {
        string DocxOutFile = @"D:\ArchiveFiles\"+DateTime.Now.ToString("dd-hh-mm-ss") +"ايصال.docx";
        string pdfOutFile = @"D:\ArchiveFiles\"+DateTime.Now.ToString("dd-hh-mm-ss") +"ايصال.pdf";        
        string DocxInFile  = @"D:\PrimariFiles\ModelFiles\الايصال.docx";
        string pictureName = @"D:\PrimariFiles\ModelFiles\صقر.png";
        string DataSource = "";
        bool gridFill = true;
        string GreDate = "";
        string SearchDate = "";
        string txtMissionCodeNum = "";
        string txtMissionCode = "";
        string[] foundList;
        string[] allList;
        string[] values = new string[8];
        string[] items = new string[8] { "barcode", "التاريخ_الميلادي", "القيمة", "المتحصل", "رقم_المعاملة", "مقدم_الطلب" , "المعاملة" , "البعثة" };
        int intID = 0;
        int itemID = 0;
        string updateAll, insertAll;
        string AccountantEmp = "";
        string Joposition = "";
        string colName = "";
        public Accountant(string dataSource, string greDate, string accountant, string joposition)
        {
            InitializeComponent();
            DataSource = dataSource;
            Joposition = joposition;
            string missionInfo = missionBasicInfo().Split('*')[3];
            txtMissionCode = missionInfo.Split('/')[1];
            //MessageBox.Show(txtMissionCode);
            txtMissionCodeNum = missionInfo.Split('/')[0];
            //MessageBox.Show(txtMissionCodeNum);
            التاريخ_الميلادي.Text = التاريخ.Text = GreDate = SearchDate= greDate;
            المتحصل.Text = AccountantEmp = accountant;
            البعثة.Text = values[7] = missionBasicInfo().Split('*')[0];
            //MessageBox.Show(values[7]);
            allList = getColList("TableReceipt");
            FillDataGridView(DataSource, greDate.Split('-')[2]);
            FillDataGridViewItems(DataSource);
            values[0] = @"D:\ArchiveFiles\" + DateTime.Now.ToString("dd-hh-mm-ss") + "الباركود.png";
            comboBox1.SelectedIndex = 0;
            آلية_البحث.SelectedIndex = 0;
            if(Joposition != "مدير")
                button13.Enabled = false;
        }
        private string missionBasicInfo()
        {
            string infoDet = "";
            string query = "select بيانات_البعثة from TableSettings";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return ""; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
            }
            catch (Exception ex)
            {
                return "";
            }



            sqlCon.Close();

            foreach (DataRow dataRow in dtbl.Rows)
            {
                try
                {
                    infoDet = dataRow["بيانات_البعثة"].ToString();
                }
                catch (Exception ex)
                {

                }
            }
            return infoDet;
        }
        private string getTables(string id) {
            string table = "";
            switch (id) {
                case "10":
                    table = "TableCollection";
                    colName = "رقم_المعاملة";
                    break;
                case "12":
                    table = "TableAuth";
                    colName = "رقم_التوكيل";
                    break;
                case "15":
                    table = "TableMerrageDoc";
                    colName = "رقم_المعاملة";
                    break;
                case "17":
                    table = "TableDivorce";
                    colName = "رقم_المعاملة";
                    break;
                case "21":
                        table = "TableHandAuth";
                    colName = "رقم_معاملة_القسم";
                    break;
                }
            return table;
        }
        private void paid(string form, string payState) {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("UPDATE "+ getTables(form) + " SET حالة_السداد =N'" + payState + "' where " +colName+ " = N'" +رقم_معاملة_القسم.Text+"'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
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
                if (row["name"].ToString() != "ID")
                {

                    allList[i] = row["name"].ToString();
                    //MessageBox.Show(allList[i]);
                    if (i == 0)
                    {
                        insertItems = row["name"].ToString();
                        insertValues = "@"+row["name"].ToString();
                        updateValues = row["name"].ToString() + "=@" + row["name"].ToString();
                    }
                    else
                    {
                        insertItems = insertValues + ", "+ row["name"].ToString();
                        insertValues = insertValues + ", @"+ row["name"].ToString();
                        updateValues = updateValues + "," + row["name"].ToString() + "=@" + row["name"].ToString();
                    }
                    i++;
                }
            }
            insertAll = "insert into " + table + "(" + insertItems + ") values (" + insertValues +")";
            updateAll = "UPDATE " + table + " SET " + updateValues + " where ID = @id";
            //MessageBox.Show(updateAll);
            return allList;

        }

        private void autoCompleteBulk(TextBox textbox, string source, string col, string table)
        {

            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct " + col + " from " + table + " where " + col + " is not null";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                DataTable Textboxtable = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(Textboxtable);
                AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();
                bool newSrt = true;
                textbox.AutoCompleteCustomSource.Clear();

                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    autoComplete.Add(dataRow[col].ToString());
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }

        private void Accountant_Load(object sender, EventArgs e)
        {
            autoCompleteBulk(مقدم_الطلب, DataSource, "الاسم", "TableGenNames");
            fileColComboBox(المعاملة, DataSource, "البند", "TableReceiptItems");
        }

        private void fileColComboBox(ComboBox combbox, string source, string comlumnName, string column)
        {
            combbox.Items.Clear();
            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select distinct " + comlumnName + " from "+ column + " where " + comlumnName + " is not null order by " + comlumnName + " asc";
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
        public void FillDataGridView(string dataSource, string year)
        {
            string query = "select * from TableReceipt where DATEPART(year,التاريخ_الميلادي) =" + year + " order by ID";
            if (year == "جميع الأعوام")
                query = "select * from TableReceipt order by ID";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            dataGridView1.Columns[0].Visible = false;
            //dataGridView1.Columns["نوع_المعاملة"].Visible = false ;
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[2].Width = 200;
            dataGridView1.Columns[3].Width = 200;
            //dataGridView1.Columns[3].Width = 40;
            sqlCon.Close();
            ColorFulGrid9();
        }
        
        public void FillDataGridViewItems(string dataSource)
        {
            string query = "select * from TableReceiptItems order by ID";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView2.DataSource = dtbl;
            dataGridView2.Columns[0].Visible = false ;
            dataGridView2.Columns[1].Width = 380;
            dataGridView2.Columns[2].Width = 200;
            sqlCon.Close();
            
        }

        private void ColorFulGrid9()
        {

        //    int genAuth = 0;
        //    int arch = 0;
        //    int unDesc = 0;
        //    int inComb = 0;
            int i = 0;
            for (; i < dataGridView1.Rows.Count - 1; i++)
            {
                
                if (dataGridView1.Rows[i].Cells["حالة_الايصال"].Value.ToString() == "ملغي")
                {
                    // timerColor = false;
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightPink;

                }
                else 
                {
                    // timerColor = false;
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;

                }
                


            }
            //labDescribed.Text = "عدد (" + i.ToString() + ") معاملة .. عدد (" + inComb.ToString() + ") غير مكتمل.. والمؤرشف منها عدد (" + arch.ToString() + ")...";

        }

        public bool newReceipt()
        {
            string query = "select * from TableReceipt where رقم_معاملة_القسم = N'" + رقم_معاملة_القسم .Text+ "'";
            Console.WriteLine("newReceipt  " + query);   
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);            
            sqlCon.Close();
            if (dtbl.Rows.Count == 1)
                return true;
            else return false;
        }
        
        public string existed(string table, string colName)
        {
            string query = "select حالة_السداد from " + table + " where "+ colName + " = N'" + رقم_معاملة_القسم .Text+ "'";
            Console.WriteLine(query);
            string state = "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);            
            sqlCon.Close();
            if (dtbl.Rows.Count == 0)
                state = "لا توجد معاملة";
            else
                foreach (DataRow dataRow in dtbl.Rows)
                {
                    state = dataRow["حالة_السداد"].ToString();
                }
            return state;
        }

        private void رقم_المعاملة_TextChanged(object sender, EventArgs e)
        {
            ZXing.BarcodeWriter writer = new ZXing.BarcodeWriter() { Format = BarcodeFormat.QR_CODE };
            pictureBox1.Image = writer.Write(txtMissionCode + رقم_المعاملة.Text);
            pictureBox1.Image.Save(values[0]);
            if (رقم_المعاملة.Text.Length > 4)
            {
                int FormType = Convert.ToInt32(SpecificDigit(رقم_المعاملة.Text, 3, 4));
                string noForm = SpecificDigit(رقم_المعاملة.Text, 3, 4);
                string rowCount = SpecificDigit(رقم_المعاملة.Text, 5, رقم_المعاملة.Text.Length);
                رقم_المعاملة.Text = رقم_المعاملة.Text.TrimStart().TrimEnd();
                string year = SpecificDigit(رقم_المعاملة.Text, 1, 2).Trim();
                رقم_معاملة_القسم.Text = txtMissionCodeNum +"/"+ txtMissionCode +"/"+ year + "/" + noForm + "/" + rowCount;
            }
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

        //private void fillDocFileAppInfo(FlowLayoutPanel panel)
        //{
        //    //MessageBox.Show(panel.Name);
        //    foreach (Control control in panel.Controls)
        //    {
        //        //MessageBox.Show(panel.Name + " - " + control.Name + " - " + control.Text);
        //        if (control is TextBox || control is ComboBox)
        //        {
        //            try
        //            {
        //                object ParaAuthIDNo = control.Name;
        //                Word.Range BookAuthIDNo = oBDoc.Bookmarks.get_Item(ref ParaAuthIDNo).Range;
        //                if (control.Name == "موقع_المعاملة")
        //                    BookAuthIDNo.Text = control.Text + AuthTitleLast;
        //                else BookAuthIDNo.Text = control.Text;
        //                if ((control.Name == "التاريخ_الميلادي" || control.Name == "التاريخ_الهجري") && اللغة.Checked)
        //                    BookAuthIDNo.Text = control.Text.Split('-')[1] + "-" + control.Text.Split('-')[0] + "-" + control.Text.Split('-')[2];

        //                object rangeAuthIDNo = BookAuthIDNo;
        //                oBDoc.Bookmarks.Add(control.Name, ref rangeAuthIDNo);

        //                //MessageBox.Show(panel.Name+ " - "+control.Name+ " - "+control.Text);
        //                Console.WriteLine(panel.Name + " - " + control.Name + " - " + control.Text);
        //            }
        //            catch (Exception ex)
        //            {
        //                //    MessageBox.Show(control.Name); 
        //            }
        //        }
        //    }
        //}
        
        private void print()
        {
            string noForm = SpecificDigit(رقم_المعاملة.Text, 3, 4);
            
            string state = existed(getTables(noForm), colName);
            if (state == "لا توجد معاملة" && نوع_المعاملة.SelectedIndex != 4)
            {
                
                    var selectedOption = MessageBox.Show("لا توجد معاملة بالرقم الموضح", "المعاملة لا تتعلق بقسم الأحوال الشخصية؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.No)
                {
                    return;
                }
                
            }
            else if (state == "تم السداد")
            {
                MessageBox.Show("المعاملة الموضح تم سدادها مسبقا");
                //return;
            }
            else if (state == "تم الالغاء")
            {
                var selectedOption = MessageBox.Show("المعاملة الموضح تم إلغاءها", "معاينة تفاصيل الالغاء؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                {

                }
                return;
            }
            else if (القيمة.Text == "بدون")
            {
                MessageBox.Show("المعاملة غير مدرجة بقائمة البنود المعتمدة، لا يمكن المتابعة");
                return;
            }


            if (!checkEmpty(panelMain) )
            {
                return;
            }

             DocxOutFile = @"D:\ArchiveFiles\" + DateTime.Now.ToString("dd-hh-mm-ss") + "ايصال.docx";
             pdfOutFile = @"D:\ArchiveFiles\" + DateTime.Now.ToString("dd-hh-mm-ss") + "ايصال.pdf";
             DocxInFile = @"D:\PrimariFiles\ModelFiles\الايصال.docx";
             pictureName = @"D:\PrimariFiles\ModelFiles\صقر.png";
            values[1] = التاريخ_الميلادي.Text;
            values[2] = القيمة.Text;
            values[3] = المتحصل.Text;
            values[4] = رقم_المعاملة.Text;
            values[5] = مقدم_الطلب.Text;
            values[6] = المعاملة.Text;
            bool insertCase = newReceipt();
            save2DataBase(panelMain, insertCase);
            
            if (نوع_المعاملة.SelectedIndex == 4)
                save2HandAuth();

            PrintDoc();
            paid(noForm, "تم السداد");
        }

        private void PrintDoc() {
            System.IO.File.Copy(DocxInFile, DocxOutFile);
            FileInfo fileInfo = new FileInfo(DocxOutFile);
            if (fileInfo.IsReadOnly) fileInfo.IsReadOnly = false;

            Word._Application wordApp = new Word.Application();
            //wordApp.Visible = true;
            Word._Document wordDoc = wordApp.Documents.Open(DocxOutFile, ReadOnly: false, Visible: true);

            int count = wordDoc.Bookmarks.Count;
            //MessageBox.Show(count.ToString());
            for (int index = 1; index < count + 1; index++)
            {
                //try
                //{
                //MessageBox.Show(wordDoc.Bookmarks[index].Name.ToString());
                if (wordDoc.Bookmarks[index].Name.ToString() == items[0])
                {
                    object oRange = wordDoc.Bookmarks[index].Range;
                    object saveWithDocument = true;
                    object missing = Type.Missing;
                    //wordDoc.InlineShapes.AddPicture(pictureName, ref missing, ref saveWithDocument, ref oRange);
                    wordDoc.InlineShapes.AddPicture(values[0], ref missing, ref saveWithDocument, ref oRange);

                }
                else
                {
                    object ParaAuthIDNo = items[index - 1];
                    Word.Range BookAuthIDNo = wordDoc.Bookmarks.get_Item(ref ParaAuthIDNo).Range;
                    BookAuthIDNo.Text = values[index - 1];
                    object rangeAuthIDNo = BookAuthIDNo;
                    wordDoc.Bookmarks.Add(items[index - 1], ref rangeAuthIDNo);
                    //MessageBox.Show (items[index - 1] + " - " + values[index - 1]);
                }
                //}
                //catch (Exception ex) { }
            }
            wordDoc.Save();
            wordDoc.ExportAsFixedFormat(pdfOutFile, Word.WdExportFormat.wdExportFormatPDF);
            wordDoc.Close();
            wordApp.Quit();
            System.Diagnostics.Process.Start(pdfOutFile);
        }
            private bool checkEmpty(Panel panel)
        {
            foreach (Control control in panel.Controls)
            {
                if (control is TextBox || control is ComboBox)
                {
                    if (control.Text == "" || control.Text.Contains("إختر"))
                    {
                        control.BackColor = System.Drawing.Color.MistyRose;
                        MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل " + control.Name);
                        return false;
                    }
                }
            }
            return true;
        }

        private string save2DataBase(Panel panel, bool insert)
        {
            string query = checkList(panel, allList, "TableReceipt", insert);
            Console.WriteLine(query);
            string id = "";
            //MessageBox.Show(query);
            if (query == "UPDATE TableReceipt SET where ID = @id") return "";
            Console.WriteLine(panel.Name + " - " + query);
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@id", intID);
            bool cont = true;
            for (int i = 0; i < foundList.Length; i++)
            {

                    foreach (Control control in panel.Controls)
                    {
                        string name = control.Name;
                        if (name == foundList[i])
                        {
                            sqlCommand.Parameters.AddWithValue("@" + foundList[i], control.Text);
                        //MessageBox.Show(i.ToString() + " " + foundList[i] + " - " + control.Text);
                        Console.WriteLine(i.ToString() + " " + foundList[i] + " - " + control.Text);
                            break;
                        }
                    }
            }
            if (query.Contains("update"))
                sqlCommand.ExecuteNonQuery();
            else {
                var reader = sqlCommand.ExecuteReader();
                if (reader.Read())
                {
                    id =  reader["lastid"].ToString();
                }
                //sqlCommand.Close();
            }
            return id;
        }
        
        private bool save2HandAuth()
        {
            رقم_المعاملة.Text = DocIDGenerator();
            string query = "insert into TableHandAuth (Viewed,حالة_السداد,رقم_معاملة_القسم) values (N'" + مقدم_الطلب .Text+ "',N'تم السداد',N'"+ رقم_معاملة_القسم .Text+ "')";
            Console.WriteLine(query);
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            try
            {
                sqlCommand.ExecuteNonQuery();
            }catch(Exception ex) {
                return false;
            }
            return true;
        }

        private string checkList(Panel panel, string[] List, string table, bool insert)
        {
            string insertItems= "";
            string insertValues = "";
            string updateValues = "";

            foundList = new string[List.Length];
            for (int f = 0; f < List.Length; f++)
                foundList[f] = "";

            int found = 0;
            foreach (Control control in panel.Controls)
            {
                string name = control.Name;
                //if (panel.Name == "PanelItemsboxes")
                //    name = name.Replace("V", "");
                if (control is TextBox || control is ComboBox )
                    for (int col = 0; col < List.Length; col++)
                        if (name == List[col])
                        {
                            foundList[found] = name;
                            //if (panel.Name == "panelapplicationInfo") MessageBox.Show(foundList[found]);
                            if (found == 0)
                            {
                                insertItems = name;
                                insertValues = "@" + name;
                                updateValues = name + "=@" + name;
                            }
                            else
                            {
                                insertItems = insertItems +", "+ name;
                                insertValues = insertValues +", @" + name;
                                updateValues = updateValues + "," + name + "=@" + name;
                            }
                            found++;
                        }
            }
            //MessageBox.Show(updateValues);
            if(insert)                
            return updateAll = "UPDATE " + table + " SET " + updateValues + " where ID = @id";
            else
                return insertAll = "insert into " + table + "(" + insertItems + ") values (" + insertValues + ");SELECT @@IDENTITY as lastid";
            
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                gridFill = true;
                foreach (Control control in panelMain.Controls)
                {
                    panelFill(control);
                }
            }
        }

       
        public void panelFill(Control control)
        {
            for (int col = 0; col < allList.Length; col++)
            {
                 if (control.Name == allList[col])
                {
                    if (dataGridView1.CurrentRow.Cells[allList[col]].Value.ToString() != "")
                    {
                        control.Text = dataGridView1.CurrentRow.Cells[allList[col]].Value.ToString();
                        intID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                        حالة_الايصال.Visible = button7.Visible = true;
                        Console.WriteLine(control.Text);                       
                    }
                    control.Enabled = false;
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedIndex)
            {
                case 0:
                    print();
                    FillDataGridView(DataSource, GreDate.Split('-')[2]);
                    break;
                case 1:
                    رقم_المعاملة.Enabled = المعاملة.Enabled = مقدم_الطلب.Enabled = true;
                    رقم_المعاملة.Text = المعاملة.Text = رقم_معاملة_القسم.Text = مقدم_الطلب.Text = القيمة.Text = "";
                    التاريخ_الميلادي.Text = التاريخ.Text;
                    حالة_الايصال.Visible = button7.Visible = false;
                    المتحصل.Text = AccountantEmp;
                    comboBox1.SelectedIndex = 0;
                    break;
                case 2:
                    if (التاريخ_الميلادي.Text == التاريخ.Text) {
                        حالة_الايصال.Text = "تم الالغاء";

                        SqlConnection sqlCon = new SqlConnection(DataSource);
                        if (sqlCon.State == ConnectionState.Closed)
                            try
                            {
                                sqlCon.Open();
                            }
                            catch (Exception ex) { return; }
                        SqlCommand sqlCmd = new SqlCommand("UPDATE TableReceipt SET حالة_الايصال =N'تم الالغاء' where ID = " + intID.ToString(), sqlCon);
                        sqlCmd.CommandType = CommandType.Text;
                        
                        sqlCmd.ExecuteNonQuery();

                    }
                    else {
                        MessageBox.Show("غير ممكن إلغاء معاملة بتاريخ مسبق");

                    }
                    string noForm = SpecificDigit(رقم_المعاملة.Text, 3, 4);
                    paid(noForm, "تم الالغاء");
                    FillDataGridView(DataSource, GreDate.Split('-')[2]);
                    break;
            }
            
        }

        private void نوع_المعاملة_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (نوع_المعاملة.SelectedIndex == 4)
                رقم_المعاملة.Text = DocIDGenerator();
        }

        private string DocIDGenerator()
        {
            string formtype = "21";
            string year = DateTime.Now.Year.ToString().Replace("20", "");
            string query = "select max(cast (right(رقم_معاملة_القسم,LEN(رقم_معاملة_القسم) - 15) as int)) as newDocID from TableHandAuth where رقم_معاملة_القسم like N'" + txtMissionCodeNum + "/" + txtMissionCode + "/" + year + "/" + formtype + "%'";
            Console.WriteLine(query);
            return  year + formtype  + getUniqueID(query);
        }

        private string getUniqueID(string query)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string maxID = "1";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                try
                {
                    maxID = (Convert.ToInt32(dataRow["newDocID"].ToString()) + 1).ToString();
                    //MessageBox.Show(dataRow["newDocID"].ToString());
                }
                catch (Exception ex)
                {
                    return maxID;
                }
            }
            return maxID;
        }

        private void المعاملة_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                btnEnd.PerformClick();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[آلية_البحث.Text.Replace(" ", "_")].HeaderText.ToString() + " LIKE '" + ListSearch.Text + "%'";
            dataGridView1.DataSource = bs;
            ColorFulGrid9();
        }

        private void المعاملة_SelectedIndexChanged(object sender, EventArgs e)
        {
            string query = "select القيم from TableReceiptItems where البند = N'" + المعاملة.Text+ "'"; ;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            القيمة.Text = "بدون";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                القيمة.Text = dataRow["القيم"].ToString();
            }
        }
        
        private bool checkStatus(string table, string column)
        {
            bool allowed = true;
            string query = "select endTime from " + table + " where "+ column + " = N'" + رقم_معاملة_القسم.Text+ "'";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            القيمة.Text = "بدون";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if( dataRow["endTime"].ToString() != "")
                    allowed = false;
            }
            return allowed;
        }

        private void button6_Click_2(object sender, EventArgs e)
        {
            filllExcelGrid(SearchDate);
        }

        private void yearReport_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void dateTimeFrom_ValueChanged(object sender, EventArgs e)
        {
            SearchDate = dateTimeFrom.Text.Split('-')[1] +"-"+dateTimeFrom.Text.Split('-')[0]+"-"+dateTimeFrom.Text.Split('-')[2];
            MessageBox.Show(SearchDate);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (button9.Text == "إضافة")
            {
                string query = "insert into TableReceiptItems (القيم,البند) values (N'" + القيم.Text + "',N'" + البند.Text + "')";
                Console.WriteLine(query);
                SqlConnection sqlConnection = new SqlConnection(DataSource);
                if (sqlConnection.State == ConnectionState.Closed)
                    sqlConnection.Open();
                SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
                sqlCommand.CommandType = CommandType.Text;
                try
                {
                    sqlCommand.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    return;
                }
            }
            else {
                SqlConnection sqlCon = new SqlConnection(DataSource);
                if (sqlCon.State == ConnectionState.Closed)
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return; }
                SqlCommand sqlCmd = new SqlCommand("UPDATE TableReceiptItems SET البند =N'" + البند.Text + "',القيم=N'"+ القيم.Text + "'  where ID = " + itemID.ToString(), sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.ExecuteNonQuery();
                button9.Text = "إضافة";
            }
            البند.Text = القيم.Text = "";
            FillDataGridViewItems(DataSource);
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentRow.Index != -1)
            {
                gridFill = true;
                itemID = Convert.ToInt32(dataGridView2.CurrentRow.Cells[0].Value.ToString());
                foreach (Control control in panelMain.Controls)
                {
                    البند.Text = dataGridView2.CurrentRow.Cells["البند"].Value.ToString();
                    القيم.Text = dataGridView2.CurrentRow.Cells["القيم"].Value.ToString();
                    button9.Text = "تعديل";
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string query = "delete from TableReceiptItems where ID = " + itemID.ToString();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            البند.Text = القيم.Text = "";
        }

        private void button12_Click(object sender, EventArgs e)
        {
            panelItems.Visible = false;
            panelItems.SendToBack();    
        }

        private void button13_Click(object sender, EventArgs e)
        {
            panelItems.Visible = true;
            panelItems.BringToFront();
        }

        private void المعاملة_TextChanged(object sender, EventArgs e)
        {
            القيمة.Text = "بدون";
        }

        private void filllExcelGrid(string date)
        {            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string strQuery = "select رقم_معاملة_القسم,مقدم_الطلب,المعاملة,القيمة from TableReceipt where التاريخ_الميلادي = '" + date + "' and حالة_الايصال = N'تم السداد' order by ID desc";
            SqlDataAdapter sqlDa = new SqlDataAdapter(strQuery, sqlCon);

            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();

            sqlDa.Fill(dtbl);
            gridExcel.DataSource = dtbl;
            
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
    }
}
