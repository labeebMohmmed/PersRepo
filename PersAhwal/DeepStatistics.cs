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
using System.Configuration;
using System.Globalization;
using System.Threading;
using System.IO;
using Xceed.Document.NET;
using Xceed.Words.NET;
using System.Diagnostics;
using System.Windows.Forms.DataVisualization.Charting;
using DocumentFormat.OpenXml;
using System.Security.Policy;

namespace PersAhwal
{
    //https://www.stats.gov.sa/ar/saudi-standard-classification-of-occupations?combine=&account=&combine_5=All&education=&Skills=&T-skills=
    public partial class DeepStatistics : Form
    {
        string dataSource = "";
        int gridID = 0;
        string archFile = @"D:\ArchiveFiles\";
        string FilespathIn, FilespathOut;
        string[] quorterS = new string[12];
        string[] quorterE = new string[12];
        string queryInfo = "";
        bool shortRange = false;
        string sqlTables = "";
        string symbol = "";
        int suitID = 0;
        string[] columnList= new string[10];
        string columnDate= "";
        char symbolChar;
        //string mainItems = "";
        int borderDashStyle1 = 0;
        int borderDashStyle2 = 0;
        string[] wrongeItems = new string[2000];
        string[] trueItems = new string[2000];
        string mainGroup1 = "";
        string mainGroup2 = "";
        string mainGroup3 = "";
        string mainGroup4 = "";
        string mainGroup5 = "";

        //int[,] gridItems = new int[10, 10];
        string dateItems = "";
        int comIndex = 0;
        int[,,,] DeepReport_Year;
        int[,,] DeepReportsingle = new int[10, 12, 32];
        int comboIndex = 0;
        static int[] Months = new int[12];
        int[] itemCount;
        static string[] List0 = new string[100];
        static string[] List1 = new string[100];
        static string[] List2 = new string[100];

        static string[] ListGroup; 
        static string[] ListTimeLine; 
        static int[] ListValue;

        string[] Words = new string[30000];
        int proTypeCount = 0;
        string dataSource57;
        string dataSource56;
        int chartAreas1 = 0;
        int chartAreas2 = 0;
        
        int chart2Areas = 0;
        bool subMain = false;
        int start = 0;
        bool holdData1 = false;
        int holdData1count = 0;
        int holdData2count = 0;
        bool holdData2 = false;
        int end = 0;
        bool newSet = true;
        string nameSeries = "series";
        int ItemsCount = 3;
        DataTable dataRowTable;
        string[] WorkOffices = new string[15];
        string[] Spect = new string[15];
        string dateLike = "";
        string name = "";
        string DateSearch = "";
        string dateType = "";
        string dateValue = "";
        string datePArt = "";
        public DeepStatistics(string Source57, string Source56, string filespathIn, string filespathOut)
        {
            InitializeComponent();
            dataSource = dataSource57 = Source57;
            dataSource56 = Source56;
            dataGridView5.DataSource = TrendingUpdate();
            FilespathIn = filespathIn;
            FilespathOut = filespathOut;
            AllStatistData();
            AllColumn();
            intiData();
            updateInfo();
            GenPreparations();
            WorkOffices[0] = "مكة المكرمة";
            WorkOffices[1] = "الطائف";
            WorkOffices[2] = "المدينة المنورة";
            WorkOffices[3] = "تبوك";
            WorkOffices[4] = "محايل عسير";
            WorkOffices[5] = "الباحة";
            WorkOffices[6] = "نجران";
            WorkOffices[7] = "جازان";
            WorkOffices[8] = "ابها";
            WorkOffices[9] = "عسير";
            WorkOffices[10] = "خميس مشيط";
            WorkOffices[11] = "بيشة";
            WorkOffices[12] = "ينبع";
            WorkOffices[13] = "القنفذة";

        }

        private void GenPreparations()
        {
            comPeriode.Location = new System.Drawing.Point(1204, 18);//1320, 18
            comPeriode.Size = new System.Drawing.Size(235, 35);
        }
        private void correctDataSuit(string col, string table, bool contain)
        {

            Console.WriteLine(table);
            SqlConnection sqlCon = new SqlConnection(dataSource56);
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select ID,GenWork,المهن_المعدلة,المهنة,نوع_المعاملة,جهة_العمل,رقم_الهوية,التاريخ_الميلادي,الحالة,تصنيف_عام from TableCollective", sqlCon);

            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl1 = new DataTable();
            sqlDa1.Fill(dtbl1);
            dataGridView6.DataSource = dtbl1;
            int count = fillSatatInfoGrid("4331", "4491");
            //MessageBox.Show(wrongeItems[idindexS].ToString() + " - " + idindexS.ToString() + " - " + count.ToString());
            //MessageBox.Show(wrongeItems[idindexE].ToString() + " - " + idindexE.ToString() + " - " + count.ToString());
            //MessageBox.Show(dataGridView6.RowCount.ToString());
            //for (int i = 0; i < dataGridView6.RowCount - 1; i++)
            //{
            //    int id = Convert.ToInt32(dataGridView6.Rows[i].Cells[0].Value.ToString());
            //    //if(id >= 7045) { 
            //    //MessageBox.Show(dataGridView6.Rows[i].Cells[6].Value.ToString());
            //    sqlCon = new SqlConnection(dataSource56);
            //    sqlDa1 = new SqlDataAdapter("select ID,رقم_الهوية,المهنة,الحالة,جهة_العمل,نوع_المعاملة from " + table, sqlCon);

            //    sqlDa1.SelectCommand.CommandType = CommandType.Text;
            //    DataTable dtbl2 = new DataTable();
            //    sqlDa1.Fill(dtbl2);
            //        foreach (DataRow dataRow in dtbl2.Rows)
            //        {
            //        Console.WriteLine(table + " - "+dataRow["ID"].ToString() +" - "+dataRow["رقم_الهوية"].ToString());
            //        int cells0 = Convert.ToInt32(dataRow["ID"].ToString());
            //            string cells1 = dataRow["نوع_المعاملة"].ToString();
            //            string cells2 = dataRow["جهة_العمل"].ToString();
            //            string cells3 = dataRow["الحالة"].ToString();
            //        //MessageBox.Show(cells1);
            //        if (dataRow["رقم_الهوية"].ToString().Trim() == dataGridView6.Rows[i].Cells[6].Value.ToString().Trim())
            //        {

            //            //UpdateState(id, cells3, "الحالة", "TableCollective");

            //            //UpdateState(id, cells2, "جهة_العمل", "TableCollective");
            //            int workID = -1;
            //            if (cells2 != "" && cells2.All(char.IsDigit))
            //            {
            //                //MessageBox.Show(cells2);
            //                workID = Convert.ToInt32(cells2);
            //                if (workID != -1 && WorkOffices[workID] != "")
            //                    UpdateState(id, WorkOffices[workID], "جهة_العمل", "TableCollective");


            //            }
            //            else if (cells2 != "" && !cells2.All(char.IsDigit))
            //            {
            //                //MessageBox.Show(cells2);
            //                if (cells2 != "-1")
            //                    UpdateState(id, cells2, "جهة_العمل", "TableCollective");
            //            }
            //            //else UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), "غير محدد", "جهة_العمل", table);

            //            //if (cells1 == "3" || cells3 == "تغيب عن العمل" || cells3 == "مجهول" || cells3 == "عمرة" || cells3 == "هروب")
            //            //{
            //            //   // MessageBox.Show(id.ToString() + " - " + cells1 + " -- " + cells3);
            //            //    UpdateState(id, "خروج نهائي بالترحيل", "نوع_المعاملة", "TableCollective");
            //            //}
            //            //else if (cells1 == "5")
            //            //    UpdateState(id, "خروج نهائي بالمحكمة العمالية", "نوع_المعاملة", "TableCollective");
            //            //else
            //            //    UpdateState(id, "خروج نهائي نظامي", "نوع_المعاملة", "TableCollective");



            //            //}
            //        }
            //        }
            //}
            //foreach (DataRow dataRow in dtbl1.Rows)
            //{
            //    //Convert.ToInt32(dataRow["جهة_العمل"].ToString()) < 0 && 
            //    if (dataRow["المهنة"].ToString() == "" && dataRow["المهن_المعدلة"].ToString() == "")
            //    {
            //        deleteRowsData(Convert.ToInt32(dataRow["ID"].ToString()), "TableCollective");
            //    }
            //}
            //MessageBox.Show(" count " + count.ToString());
            foreach (DataRow dataRow in dtbl1.Rows)
            {
                if (Convert.ToInt32(dataRow["ID"].ToString()) >= 3 && Convert.ToInt32(dataRow["ID"].ToString()) >= 9)
                {
                    //int workID = -1;
                    //if (dataRow["جهة_العمل"].ToString() != "" && dataRow["جهة_العمل"].ToString().All(char.IsDigit))
                    //    workID = Convert.ToInt32(dataRow["جهة_العمل"].ToString().Trim());

                    //////MessageBox.Show(dataRow["ID"].ToString() + " - " + trueItems[x]);
                    //if (workID != -1)
                    //    UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), WorkOffices[workID], "جهة_العمل", "TableCollective");
                    //else UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), "غير محدد", "جهة_العمل", "TableCollective");


                    //if (dataRow["نوع_المعاملة"].ToString() == "3" || dataRow["الحالة"].ToString() == "تغيب عن العمل" || dataRow["الحالة"].ToString() == "مجهول" || dataRow["الحالة"].ToString() == "عمرة" || dataRow["الحالة"].ToString() == "هروب")
                    //    UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), "خروج نهائي بالترحيل", "نوع_المعاملة", table);
                    //else if (dataRow["نوع_المعاملة"].ToString() == "5")
                    //    UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), "خروج نهائي بالمحكمة العمالية", "نوع_المعاملة", table);
                    //else
                    //    UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), "خروج نهائي نظامي", "نوع_المعاملة", table);


                    //UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), dataRow["المهنة"].ToString().Trim(), "المهن_المعدلة", table);
                    //UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), dataRow["المهنة"].ToString().Trim(), "المهنة", table);
                    //addCollectiveData(dataRow["المهنة"].ToString(), dataRow["نوع_المعاملة"].ToString(), dataRow["جهة_العمل"].ToString(), dataRow["رقم_الهوية"].ToString(), dataRow["التاريخ_الميلادي"].ToString(), dataRow["الحالة"].ToString());
                    //if (dataRow["المهن_المعدلة"].ToString() == "" && (dataRow["الحالة"].ToString().Contains("عمر") ||dataRow["الحالة"].ToString().Contains("متوف")||dataRow["الحالة"].ToString().Contains("مجهول")))
                    //UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), dataRow["المهنة"].ToString(), "تصنيف_عام", table);
                    //if (Convert.ToInt32(dataRow["ID"].ToString()) > 602 && Convert.ToInt32(dataRow["ID"].ToString()) < 610)
                    //{
                    //if (dataRow["رقم_الهوية"].ToString() )
                    //{
                    //    UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), dataRow["المهن_المعدلة"].ToString(), "تصنيف_عام", table);
                    //}
                    //if (dataRow["المهن_المعدلة"].ToString() == "فني اتصالات")
                    //{
                    //    MessageBox.Show(dataRow["تصنيف_عام"].ToString());
                    //}

                    for (int x = idindexS; x < count && x < idindexE; x++)
                    {
                        if (dataRow["تصنيف_عام"].ToString().Trim() == wrongeItems[x].Trim())
                        {
                            UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), trueItems[x], "تصنيف_عام", table);
                        }
                    }

                    //suitID = Convert.ToInt32(dataRow["ID"].ToString());
                    //if (dataRow["GenWork"].ToString().Contains(" و") && suitID > Convert.ToInt32(txtIDCase.Text))
                    //{
                    //    txtIDCase.Text = dataRow["ID"].ToString();
                    //    wrongName.Text = dataRow["القضية"].ToString();

                    //    return;
                    //}

                    //if (dataRow["القضية"].ToString().Contains(" و") && suitID > Convert.ToInt32(txtIDCase.Text))
                    //{
                    //    txtIDCase.Text = dataRow["ID"].ToString();
                    //    wrongName.Text = dataRow["القضية"].ToString();

                    //    return;
                    //}
                    //else {

                    //}
                    //}
                }
            }
        }

        private DataTable TrendingUpdate()
        {
            for (int x = 0; x < 30000; x++)
                Words[x] = "";
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select ID,Words,WordsCount,WordsRows from TableTrendings", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable rowTable1 = new DataTable();
            sqlDa.Fill(rowTable1);
            sqlCon.Close();
            return rowTable1;
        }
            private void TrendingWords()
        {
            for (int x = 0; x < 30000; x++)
                Words[x] = "";

            dataGridView5.DataSource = TrendingUpdate();
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select SpecType,SpecText from TableFreeForm", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable rowTable = new DataTable();
            sqlDa.Fill(rowTable);
            sqlCon.Close();
            int rowcount = 0;
            foreach (DataRow dataRow in rowTable.Rows)
            {
                if (dataRow["SpecText"].ToString() != "" && rowcount >= Convert.ToInt32(dataGridView5.Rows[0].Cells[3].Value.ToString()))
                {
                    string[] wordsRow = dataRow["SpecText"].ToString().Split(' ');
                    for (int x = 0; x < wordsRow.Length; x++)
                    {
                        bool found = false;
                        for (int row = 0; row < dataGridView5.RowCount - 1; row++)
                        {
                            //MessageBox.Show(row.ToString());
                            int countPlus = Convert.ToInt32(dataGridView5.Rows[row].Cells[2].Value.ToString());
                            if (wordsRow[x] == dataGridView5.Rows[row].Cells[1].Value.ToString())
                            {
                                //
                                   
                                addKeyWord(Convert.ToInt32(dataGridView5.Rows[row].Cells[0].Value.ToString()), wordsRow[x], (countPlus+1).ToString());
                                found = true;
                            }
                        }
                        if(!found) 
                            addKeyWord(0, wordsRow[x], "1");
                    }
                    dataGridView5.DataSource = TrendingUpdate();
                    addKeyRows(rowcount+1);
                    //MessageBox.Show(dataRow["SpecText"].ToString());
                }
                rowcount++;
                
            }
        }
        private void deleteDuplicate()
        {
            DataTable rowTable = TrendingUpdate();
            dataGridView5.DataSource = rowTable;
            foreach (DataRow dataRow in rowTable.Rows)
            {
                int id = Convert.ToInt32(dataRow["ID"].ToString());
                int WordsCount = Convert.ToInt32(dataRow["WordsCount"].ToString());
                string Words = dataRow["Words"].ToString();
                for (int row = 0; row < dataGridView5.RowCount - 1; row++)
                {
                    int gridID = Convert.ToInt32(dataGridView5.Rows[row].Cells[0].Value.ToString());
                    string gridWord = dataGridView5.Rows[row].Cells[1].Value.ToString();
                    int countPlus = Convert.ToInt32(dataGridView5.Rows[row].Cells[2].Value.ToString());
                    if (Words == gridWord && gridID != id)
                    {
                        addKeyWord(id, Words, (WordsCount + countPlus).ToString());
                        deleteRowsData(gridID, "TableTrendings");
                    }
                }
            }
            dataGridView5.DataSource = TrendingUpdate();
        }
        private void deleteRowsData(int v1, string v2)
        {
            //MessageBox.Show(v1.ToString());
            string query;
            SqlConnection Con = new SqlConnection(dataSource56);
            query = "DELETE FROM " + v2 + " where ID = @ID";
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", v1);
            sqlCmd.ExecuteNonQuery();
            Con.Close();
        }
        private void addKeyRows(int row)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("update TableTrendings set WordsRows=@WordsRows where ID=@id", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", 1);
            sqlCmd.Parameters.AddWithValue("@WordsRows", row.ToString());
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        private void addKeyWord(int id,string word,string wordCount)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableTrendings (Words,WordsCount) values (@Words,@WordsCount)", sqlCon);
            if (id != 0)
                sqlCmd = new SqlCommand("update TableTrendings set Words=@Words,WordsCount=@WordsCount where ID=@id", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@Words", word);
            sqlCmd.Parameters.AddWithValue("@WordsCount", wordCount);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        int idindexS = 1;
        int idindexE = 1;
        private int fillSatatInfoGrid(string idS, string idE)
        {
            wrongeItems = new string[2000];
            trueItems = new string[2000];

            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableStatisInfo", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable rowTable = new DataTable();
            sqlDa.Fill(rowTable);
            sqlCon.Close();
            dataGridView4.DataSource = rowTable;
            int x = 0;
            foreach (DataRow dataRow in rowTable.Rows)
            {
                if (dataRow["wrongeItems"].ToString() != "")
                {
                    wrongeItems[x] = dataRow["wrongeItems"].ToString();
                    trueItems[x] = dataRow["trueItems"].ToString();
                    x++;
                }
                if (idS == dataRow["ID"].ToString())
                    idindexS = x;
                if (idE == dataRow["ID"].ToString())
                    idindexE = x;
            }
            return x;
        }

        private string MergData(int main, int sub, string item)
        {
            string mainGroupMerg = item;
            string merg = "merg_" + main.ToString();// + "_" + sub.ToString();
            if(dataGridView4.Rows.Count > 1)
            {
                for (int row = 0; row < dataGridView4.RowCount - 1; row++) {
                    string mainGroup = dataGridView4.Rows[row].Cells[1].Value.ToString();
                    string mergedItem = dataGridView4.Rows[row].Cells[4].Value.ToString();
                    string col2 = dataGridView4.Rows[row].Cells[5].Value.ToString();
                    string col3 = dataGridView4.Rows[row].Cells[6].Value.ToString();
                    string col4 = dataGridView4.Rows[row].Cells[7].Value.ToString();
                    string col5= dataGridView4.Rows[row].Cells[8].Value.ToString();
                    string col6= dataGridView4.Rows[row].Cells[9].Value.ToString();
                    switch (sub)
                    {
                        case 0:
                            if (mainGroup == merg + "_0")
                            {
                               // MessageBox.Show(col2 + " - " + col3 + " - " + col4 + " - " + col5 + " - " + col6);
                                if (item == col2 || item == col3 || item == col4 || item == col5 || item == col6)
                                    mainGroupMerg = mergedItem;
                                //MessageBox.Show(mergedItem);
                            }
                            break;

                        case 1:
                            if (mainGroup == merg + "_1")
                            {
                                if (item == col2 || item == col3 || item == col4 || item == col5 || item == col6)
                                    mainGroupMerg = mergedItem;
                            }
                            break;

                        case 2:
                            if (mainGroup == merg + "_2")
                            {
                                if (item == col2 || item == col3 || item == col4 || item == col5 || item == col6)
                                    mainGroupMerg = mergedItem;
                            }
                            break;

                        case 3:
                            if (mainGroup == merg + "_3")
                            {
                                if (item == col2 || item == col3 || item == col4 || item == col5 || item == col6)
                                    mainGroupMerg = mergedItem;
                            }
                            break;
                    }
                }
            }
            return mainGroupMerg;
        }

        private void FillInDataGrid(string SQLstring)
        {
            string cn = ConfigurationManager.ConnectionStrings["Scratchpad"].ConnectionString; //hier wordt de databasestring opgehaald
            DataSet ds = new DataSet();
            // dispose objects that implement IDisposable
            using (SqlConnection myConnection = new SqlConnection(cn))
            {
                SqlDataAdapter dataadapter = new SqlDataAdapter(SQLstring, myConnection);

                // set the CommandTimeout
                dataadapter.SelectCommand.CommandTimeout = 60;  // seconds

                myConnection.Open();
                dataadapter.Fill(ds, "Authors_table");
            }
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "Authors_table";
        }
        private void intiData()
        {
            

            symbolChar = '_';
            symbolChar = '*';
            symbolChar = ' ';
            quorterS[0] = "****-01-01";
            quorterE[0] = "****-01-31";
            quorterS[1] = "****-02-01";
            quorterE[1] = "****-02-29";
            quorterS[2] = "****-03-01";
            quorterE[2] = "****-03-31";
            quorterS[3] = "****-04-01";
            quorterE[3] = "****-04-30";
            quorterS[4] = "****-05-01";
            quorterE[4] = "****-05-31";
            quorterS[5] = "****-06-01";
            quorterE[5] = "****-06-30";
            quorterS[6] = "****-07-01";
            quorterE[6] = "****-07-31";
            quorterS[7] = "****-08-01";
            quorterE[7] = "****-08-31";
            quorterS[8] = "****-09-01";
            quorterE[8] = "****-09-30";
            quorterS[9] = "****-10-01";
            quorterE[9] = "****-10-31";
            quorterS[10] = "****-11-01";
            quorterE[10] = "****-11-30";
            quorterS[11] = "****-12-01";
            quorterE[11] = "****-12-31";
        }
        private void BackupDataBase(string source, string dataBase)
        {
            //OpenFileDialog dlg = new OpenFileDialog();
            //dlg.ShowDialog();
            //source = source + "Connection Timeout=3600";
            string file = "D:";
            //dlg.FileName = ;
            string query = "BACKUP DATABASE " + dataBase + " TO  DISK = '" + file + "\\" + dataBase + "-" + DateTime.Now.Ticks.ToString() + ".bak'";
            string query1 = "BACKUP DATABASE [AhwalDataBase] TO  DISK = N'D:\\SudanAffairs452145' WITH NOFORMAT, NOINIT,  NAME = N'AhwalDataBase-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10";
            try
            {
                SqlConnection sqlCon = new SqlConnection(source);
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlCommand cmd = new SqlCommand(query1, sqlCon);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Backup is done !!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void addComboData(DataTable rows,string items)
        {
            foreach (DataRow dataRow in rows.Rows)
            {
                bool found1 = false;

                for (int a = 1; a < subComb0.Items.Count; a++)
                {
                    if (dataRow[items].ToString() == subComb0.Items[a].ToString())
                        found1 = true;
                }
                if (!found1)
                {
                    if (dataRow[items].ToString() != "")
                        subComb0.Items.Add(dataRow[items].ToString());
                }
            }
        }

        private void updateInfo()
        {
            
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select ID,CountryDest from TableVisaApp", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable tablerow = new DataTable();
            sqlDa.Fill(tablerow);
            sqlCon.Close();
            foreach (DataRow dataRow in tablerow.Rows)
            {
                //Console.WriteLine(dataRow["CountryDest"].ToString());
                if (dataRow["CountryDest"].ToString().Contains("الهند"))
                {
                    ulterData(Convert.ToInt32(dataRow["ID"].ToString()), "the Republic of India");
                    //Console.WriteLine(dataRow["CountryDest"].ToString());
                }
            }
        }

        private void ulterData(int v1, string v2)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("update TableVisaApp set CountryDest=@CountryDest where ID=@id", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", v1);
            sqlCmd.Parameters.AddWithValue("@CountryDest", v2);
            sqlCmd.ExecuteNonQuery();
        }

        private void addComboData1(string items2, string items1, string table)
        {
            if (columnList[1] == "") return;



            DataTable dataRowTable = new DataTable();
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select distinct " + items2 + " from " + table +" where " + items1 + " = @" + items1 , sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@"+ items1, subComb0.Text);
            sqlDa.Fill(dataRowTable);
            //MessageBox.Show(items1 + " - " + dataRowTable.Rows.Count.ToString());
            sqlCon.Close();
            



            foreach (DataRow dataRow in dataRowTable.Rows)
            {
                //MessageBox.Show(dataRow[items2].ToString());
                bool add = true;
                try
                {
                    mainGroup2 = dataRow[items2].ToString();

                    //if (gridcel == "شهادة بحث بغرض البيع" || gridcel == "بيع ارض" || gridcel == "إقرار بالتنازل")
                    //    gridcel = "بيع";
                    mainGroup2 = MergData(genTypes.SelectedIndex, 1, mainGroup2);

                    //if (mainGroup2 != "" && dataRow[columnList[1]].ToString() == subComb0.Text)
                    if (mainGroup2 != "" && !mainGroup2.All(char.IsDigit))
                    {
                        //MessageBox.Show(dataRow[items].ToString());
                        for (int a = 0; a < subComb1.Items.Count; a++)
                        {
                            if (mainGroup2 == subComb1.Items[a].ToString())
                            {
                                add = false;
                                break;
                            }
                        }
                        if (add)
                        {
                            // MessageBox.Show(dataRow[subItems[genTypes.SelectedIndex, 0]].ToString());
                            subComb1.Items.Add(mainGroup2);
                        }
                    }
                }
                catch (Exception ex) { }
                // else add = false;




            }
            string[] arrangedArray = new string[500];
            int aa = 0;
            for (; aa < subComb1.Items.Count; aa++)
            {
                arrangedArray[aa] = subComb1.Items[aa].ToString();
                //MessageBox.Show(arrangedArray[aa]);
            }
            var ordered = arrangedArray.OrderBy(item => item, StringComparer.Ordinal);

            subComb1.Items.Clear();
            string[] strArray = string.Join(",", ordered).Split(',');
            for (int item = 0; item < strArray.Length; item++)
                if (!string.IsNullOrEmpty(strArray[item])) subComb1.Items.Add(strArray[item]);

        }

        private void addComboData2(DataTable rows, string items)
        {
            //MessageBox.Show(items);
            foreach (DataRow dataRow in rows.Rows)
            {
                bool add = true;
                mainGroup3 = dataRow[items].ToString();
                //MessageBox.Show(items);
                string gridcel1 = dataRow[columnList[2]].ToString();
                string gridcel0 = dataRow[columnList[1]].ToString();
                mainGroup3 = MergData(genTypes.SelectedIndex, 2, mainGroup3);
                if (mainGroup3 != "" && gridcel1 == subComb1.Text && gridcel0 == subComb0.Text)
                //if (mainGroup3 == "امبده" || mainGroup3 == "ابوسعد" || mainGroup3 == "الثورة")
                {
                    //MessageBox.Show(dataRow["ID"].ToString());
                    //splitCol(Convert.ToInt32(dataRow["ID"].ToString()), dataRow["موضوع_التوكيل"].ToString().Split('_'));
                    for (int a = 0; a < subComb2.Items.Count; a++)
                    {
                        if (mainGroup3 == subComb2.Items[a].ToString())
                        {
                            add = false;
                            break;
                        }
                    }
                    if (add)
                    {
                        subComb2.Items.Add(mainGroup3);
                    }
                }                
            }
            string[] arrangedArray = new string[500];
            int aa = 0;
            for (; aa < subComb2.Items.Count; aa++)
            {
                arrangedArray[aa] = subComb2.Items[aa].ToString();
                //MessageBox.Show(arrangedArray[aa]);
            }
            var ordered = arrangedArray.OrderBy(item => item, StringComparer.Ordinal);

            subComb2.Items.Clear();
            string[] strArray = string.Join(",", ordered).Split(',');
            for (int item = 0; item < strArray.Length; item++)
                if (!string.IsNullOrEmpty(strArray[item])) subComb2.Items.Add(strArray[item]);

        }

        private void addComboData3(DataTable rows, string items)
        {

            foreach (DataRow dataRow in rows.Rows)
            {
                bool add = true;
                mainGroup4 = dataRow[items].ToString();
                string gridcel0 = dataRow[columnList[1]].ToString();
                string gridcel1 = dataRow[columnList[2]].ToString();
                string gridcel2 = dataRow[columnList[3]].ToString();
                string test = mainGroup4;
                mainGroup4 = MergData(genTypes.SelectedIndex, 3, mainGroup4);
                if (mainGroup4 != ""  && gridcel0 == subComb0.Text && gridcel1 == subComb1.Text && gridcel2 == subComb2.Text)
                //if ((mainGroup4 == "امبده" || mainGroup4 == "ابوسعد" || mainGroup4 == "الثورة") && gridcel0 == subComb0.Text && gridcel1 == subComb1.Text && gridcel2.Contains("امدرمان"))
                {
                    //MessageBox.Show(test);
                    //MessageBox.Show(dataRow["ID"].ToString());
                    //splitCol(Convert.ToInt32(dataRow["ID"].ToString()), dataRow["موضوع_التوكيل"].ToString().Split('_'));
                    for (int a = 0; a < subComb3.Items.Count; a++)
                    {
                        if (mainGroup4 == subComb3.Items[a].ToString())
                        {
                            add = false;
                            break;
                        }
                    }
                    if (add)
                    {
                        subComb3.Items.Add(mainGroup4);
                    }
                }                
            }
            string[] arrangedArray = new string[500];
            int aa = 0;
            for (; aa < subComb3.Items.Count; aa++)
            {
                arrangedArray[aa] = subComb3.Items[aa].ToString();
                //MessageBox.Show(arrangedArray[aa]);
            }
            var ordered = arrangedArray.OrderBy(item => item, StringComparer.Ordinal);

            subComb3.Items.Clear();
            string[] strArray = string.Join(",", ordered).Split(',');
            for (int item = 0; item < strArray.Length; item++)
                if (!string.IsNullOrEmpty(strArray[item])) 
                    subComb3.Items.Add(strArray[item]);

        }
        private void addComboData4(DataTable rows, string items)
        {
            foreach (DataRow dataRow in rows.Rows)
            {
                bool add = true;
                mainGroup4 = dataRow[items].ToString();
                string gridcel0 = dataRow[columnList[1]].ToString();
                string gridcel1 = dataRow[columnList[2]].ToString();
                string gridcel2 = dataRow[columnList[3]].ToString();
                string gridcel3 = dataRow[columnList[4]].ToString();
                mainGroup5 = MergData(genTypes.SelectedIndex, 4, mainGroup5);
                //if (mainGroup5 != "" && gridcel0 == subComb0.Text && gridcel1 == subComb1.Text && gridcel2 == subComb2.Text && gridcel3 == subComb3.Text)
                if (mainGroup5 != "" && gridcel0 == subComb0.Text && gridcel1 == subComb1.Text && gridcel2 == subComb2.Text && gridcel3 == subComb3.Text)
                {
                    for (int a = 0; a < subComb4.Items.Count; a++)
                    {
                        if (mainGroup5 == subComb4.Items[a].ToString())
                        {
                            add = false;
                            break;
                        }
                    }
                    if (add)
                    {
                        subComb4.Items.Add(mainGroup5);
                    }
                }
            }
        }
        private void reFillRows()
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            dataRowTable = new DataTable();
            try
            {
                
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter(queryInfo, sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                sqlDa.Fill(dataRowTable);
            }
            catch (Exception ex) { }
            finally
            {
                sqlCon.Close();
            }
        }
            private int itemsType(string table,string items1, string items2, bool date, ComboBox comb)
        {
            DataTable dataRowTable = new DataTable();
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select distinct " + items1 + " from " + table, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.Fill(dataRowTable);
            //MessageBox.Show(items1 + " - " + dataRowTable.Rows.Count.ToString());
            sqlCon.Close();
            labebSum.Text = dataRowTable.Rows.Count.ToString();
            
            int count = fillSatatInfoGrid("1","1");
            string[] arrangedArray = new string[500];
            
            DataTable dataRowTable1 = new DataTable();
            sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlDa = new SqlDataAdapter("select distinct " + items2 + " from " + table, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.Fill(dataRowTable1);
            //MessageBox.Show(items2 +" - "+ dataRowTable1.Rows.Count.ToString());
            sqlCon.Close();

            foreach (DataRow dataRow in dataRowTable1.Rows)
            {
                bool found2 = false;
                string name2 = dataRow[items2].ToString();
                

                if (date)
                {
                    for (int a = 1; a < comYear.Items.Count; a++)
                    {
                        if (name2.Split('-').Length == 4)
                        {
                            found2 = true; break;
                        }
                        if (name2.Split('-').Length == 3)
                        {
                            if (name2.Split('-')[2] == comYear.Items[a].ToString())
                                found2 = true;
                        }
                        else found2 = true;

                    }
                    //if (!found2 && dataRow[items2].ToString().Split('-')[2].Contains("20"))
                    //{
                    //    if (name2.Split('-')[2] != "")
                    //    {

                    //        //comYear.Items.Add(name2.Split('-')[2]);
                    //        comYear.SelectedIndex = 1;

                    //    }

                    //}
                }

            }
            foreach (DataRow dataRow in dataRowTable.Rows)
            {
                bool found1 = false;
                string name1 = dataRow[items1].ToString();
                name1 = MergData(genTypes.SelectedIndex, 0, name1);

                for (int a = 1; a < comb.Items.Count; a++)
                {
                    if (name1 == comb.Items[a].ToString())
                    { found1 = true; break; }
                    else found1 = false;

                }

                if (name1 != "" && !found1)
                {
                    comb.Items.Add(name1);
                    //MessageBox.Show(mainGroup1);
                }

                

            }
            int aa = 0;
            for (; aa < comb.Items.Count; aa++)
            {
                arrangedArray[aa] = comb.Items[aa].ToString();
                //MessageBox.Show(arrangedArray[aa]);
            }
            var ordered = arrangedArray.OrderBy(item => item, StringComparer.Ordinal);

            comb.Items.Clear();
            string[] strArray = string.Join(",", ordered).Split(',');
            for (int item = 0; item < strArray.Length; item++)
                if (!string.IsNullOrEmpty(strArray[item])) comb.Items.Add(strArray[item]);
            //MessageBox.Show();
            //label1.Text = "";
            //string[] overall = new string[comb.Items.Count];
            //for (int x = 0; x < comb.Items.Count; x++)
            //{
            //    int sumItems = 0;
            //    foreach (DataRow dataRow in dataRowTable.Rows)
            //    {
            //        if (comb.Items[x].ToString() == dataRow[items1].ToString())
            //            sumItems++;
            //    }

            //    label1.Text = label1.Text  + comb.Items[x].ToString() + " - (" + sumItems.ToString() + ")"+Environment.NewLine;
            //overall[x] = comb.Items[x].ToString() + " - (" + sumItems.ToString() + ")";
            //}
            //comb.Items.Clear(); comb.Items.AddRange(overall);

            return comb.Items.Count;
        }

        

        private void AddAchartData(string[] valueName, string name, int[] data, bool col_Graph, string yAxisTitle, bool col)
        {
            int sum = 0;
            if (!holdData1) {
                while (chart1.Series.Count> 0) {
                    chart1.Series.RemoveAt(0);
                    comSeriers.Items.RemoveAt(0);
                    
                }
                borderDashStyle1=0;
            }
            else borderDashStyle1++;
            bool found = false;
            for (int x = 0; x < comSeriers.Items.Count; x++)
            {
                if (comSeriers.Items[x].ToString() == name)
                {
                    found = true;
                }
            }
            if (!found)
            {
                chart1.Series.Add(name);
                comSeriers.Items.Add(name);
                //MessageBox.Show(name);
                //chart1.Series[name].IsValueShownAsLabel = بشم;
                
            }
            if (chartAreas1 == 0)
            {
                
                chart1.ChartAreas[chartAreas1].AxisY.Title = yAxisTitle;
                chart1.ChartAreas[chartAreas1].AxisX.LabelStyle.Interval = 1;
                chart1.ChartAreas[chartAreas1].AxisY.LabelStyle.Interval = Convert.ToInt32(combGrad1.Text);
                chart1.ChartAreas[chartAreas1].AxisY.LabelStyle.Font = new System.Drawing.Font("Arabic Typesetting", 16F, System.Drawing.FontStyle.Regular);
                
            }
            if (col_Graph)
            {
                switch (borderDashStyle1)
                {
                    case 0:
                        chart1.Series[name].BorderDashStyle = ChartDashStyle.Solid;
                        break;
                    case 1:
                        chart1.Series[name].BorderDashStyle = ChartDashStyle.Dot;
                        break;
                    case 2:
                        chart1.Series[name].BorderDashStyle = ChartDashStyle.Dash;
                        break;
                    case 3:
                        chart1.Series[name].BorderDashStyle = ChartDashStyle.DashDot;
                        break;
                    case 4:
                        chart1.Series[name].BorderDashStyle = ChartDashStyle.DashDotDot;
                        borderDashStyle1 = 0;
                        break;
                }

                chart1.Series[name].BorderWidth = 3;
                //chart1.Series[name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            }
            if (radioColumns.Checked)
                chart1.Series[name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
            else
                chart1.Series[name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;

            //string[] comboBoxNames = new string[comboBox.Items.Count];
            //int count = fillSatatInfoGrid("1", "1");

            //for (int y = 0; y < comboBox.Items.Count; y++) 
            //{
            //    comboBoxNames[y] = comboBox.Items[y].ToString();
            //    //MessageBox.Show(comboBoxNames[y]);
            //    //for (int x = 0; x < count; x++)
            //    //{
            //    //    if (comboBoxNames[y] == wrongeItems[x])
            //    //    {
            //    //        comboBoxNames[y] = trueItems[x];
            //    //    }
            //    //}
            //}
            
            for (int x = 0; x < valueName.Length; x++)
            {
                chart1.Series[name].Points.AddXY(valueName[x], data[x]);
            }

            
            OverallStatis.Items.Clear();
            for (int x = 0; x < data.Length; x++)
            {
                for (int y = 0; y < data.Length - 1; y++)
                {
                    int temp = data[y];
                    data[y] = data[y + 1];
                    data[y + 1] = temp;
                }                   
            }
            chartAreas1++;
            //btnSumStatis.Text = overSum.ToString() + " معاملة";
        }
        
        private void AddAchartTimeLine(string[] nameValue, string name, int[] data, bool col_Graph, string yAxisTitle)
        {
            int sum = 0;
            if (!holdData2) {
                while (chart2.Series.Count> 0) {
                    chart2.Series.RemoveAt(0);                    
                }
                borderDashStyle2=0;
            }
            else borderDashStyle2++;
            try
            {
                chart2.Series.Add(name);
            }
            catch (Exception ex) { }
            if (chartAreas2 == 0)
            {

                chart2.ChartAreas[chartAreas2].AxisY.Title = yAxisTitle;
                chart2.ChartAreas[chartAreas2].AxisX.LabelStyle.Interval = 1;
                chart2.ChartAreas[chartAreas2].AxisY.LabelStyle.Interval = Convert.ToInt32(combGrad1.Text);
                chart2.ChartAreas[chartAreas2].AxisY.LabelStyle.Font = new System.Drawing.Font("Arabic Typesetting", 16F, System.Drawing.FontStyle.Regular);
                
            }
            if (col_Graph)
            {
                switch (borderDashStyle1)
                {
                    case 0:
                        chart2.Series[name].BorderDashStyle = ChartDashStyle.Solid;
                        break;
                    case 1:
                        chart2.Series[name].BorderDashStyle = ChartDashStyle.Dot;
                        break;
                    case 2:
                        chart2.Series[name].BorderDashStyle = ChartDashStyle.Dash;
                        break;
                    case 3:
                        chart2.Series[name].BorderDashStyle = ChartDashStyle.DashDot;
                        break;
                    case 4:
                        chart2.Series[name].BorderDashStyle = ChartDashStyle.DashDotDot;
                        borderDashStyle1 = 0;
                        break;
                }
                chart2.Series[name].BorderWidth = 3;
            }
            if (radioColumns.Checked)
                chart2.Series[name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
            else
                chart2.Series[name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;

            for (int x = 0; x < nameValue.Length; x++)
            {
                chart2.Series[name].Points.AddXY(Monthorder(nameValue[x]), data[x]);
            }            
            chartAreas2++;
        }
        
        private void ReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            datePArt = " where DATENAME(" + dateType + ", " + columnDate + ") = " + dateValue + " and " + columnDate + " like N'%" + comYear.Text + "'";
            if (dateValue == "ALL")
                datePArt = "";
            else if (dateValue == "YEAR")
                datePArt = " dateColumn like N'%" + comYear.Text + "' ";
            switch (comboIndex)
            {
                case 0:
                    getMainData(columnList[1], columnDate, dateType, dateValue);
                    AddAchartData(ListGroup, name.Replace(" ", "_"), ListValue, false, "عدد معاملات " + genTypes.Text, true);
                    getTimeLine();
                    AddAchartTimeLine(ListTimeLine, "القيد الزمني للعام " + comYear.Text, ListValue, false, "عدد معاملات " + genTypes.Text);
                    break;
                case 1:
                    getSub1Data(columnList[1], columnList[2], columnDate, dateType, dateValue);
                        
                    AddAchartData(ListGroup, name.Replace(" ", "_"), ListValue, false, "عدد معاملات " + genTypes.Text, true);
                    getTimeLine();
                    AddAchartTimeLine(ListTimeLine, "القيد الزمني للعام " + comYear.Text, ListValue, false, "عدد معاملات " + genTypes.Text);
                    break;
            }
                    /*bool column = true;
                    if(ReportType.SelectedIndex == 2) column = false;
                    // itemsOfSubitems(subComb0, false);
                    //MessageBox.Show(comboIndex.ToString());
                    //MessageBox.Show(comboIndex.ToString());
                    switch (comboIndex)
                    {
                        case 0:
                            getGroupCount0(columnList[1], dateLike);
                            AddAchartData(subComb0, name.Replace(" ", "_"), itemCount, false, "عدد معاملات "  + genTypes.Text , column);
                            //MessageBox.Show(DateSearch);

                            //MessageBox.Show(columnList[1]);
                            AddAchart2Data(subComb0, DateSearch.Split('-'), name, itemCount, "عدد المعاملات");
                            //if (subBtn0.Text == "الكل")
                            //{
                            //    //MessageBox.Show(subComb0.SelectedIndex.ToString());
                            //    itemsOfSubitems(subComb0, true);
                            //    itemsOfTime(subComb0, column);
                            //}
                            //else itemsOfSubitems(subComb0, false);
                            break;

                        case 1:
                            //MessageBox.Show(dateLike);
                            getGroupCount1(columnList[2], dateLike, subComb0.Text);
                            AddAchartData(subComb1, name.Replace(" ", "_"), itemCount, false, "عدد معاملات " + subComb0.Text, column);
                            getDayCount1(columnList[1], subComb0.Text, DateSearch);
                            //getDayCount2(columnList[1], subComb0.Text, columnList[2], subComb1.Text, DateSearch);
                            AddAchart2Data(subComb1, DateSearch.Split('-'), name, itemCount, "عدد المعاملات");
                            //if (subBtn1.Text == "الكل")
                            //{
                            //    itemsOfSubitems(subComb1, true);
                            //    itemsOfTime(subComb1, column);
                            //}
                            //else itemsOfSubitems(subComb1, false);
                            break;

                        case 2:
                            if (subBtn2.Text == "الكل")
                            {
                                itemsOfSubitems(subComb2, true);
                                itemsOfTime(subComb2, column);
                            }
                            else itemsOfSubitems(subComb2, false);
                            break;
                        case 3:
                            if (subBtn2.Text == "الكل")
                            {
                                itemsOfSubitems(subComb3, true);
                                itemsOfTime(subComb3, column);
                            }
                            else itemsOfSubitems(subComb3, false);
                            break;
                    }
                    holdData1count = 0;
                    holdData1 = false;
                    */
            }
        private void getMainData(string colGroup, string dateColumn, string dateType, string dateValue)
        {
            
            string query = "select  " + colGroup + " , count ( " + colGroup + " ) as dataCount from TableTempData " + datePArt + " group by  " + colGroup;
            Console.WriteLine(query);
            //MessageBox.Show(query);
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();            
            ListValue = new int[table.Rows.Count];
            ListGroup = new string[table.Rows.Count];
            int i = 0;
            foreach (DataRow dataRow in table.Rows)
            {
                ListValue[i] = Convert.ToInt32(dataRow["dataCount"].ToString());
                ListGroup[i] = dataRow[colGroup].ToString();
                i++;
            }
        }
        
        private void getMainDataGen(string colGroup)
        {
            subComb0.Items.Clear();
            
            string query = "select  " + colGroup + " , count ( " + colGroup + " ) as dataCount from TableTempData group by  " + colGroup;
            Console.WriteLine(query);
            //MessageBox.Show(query);
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();            
            int i = 0;
            foreach (DataRow dataRow in table.Rows)
            {
                subComb0.Enabled = true;
                subComb0.Items.Add(dataRow[colGroup].ToString());
                i++;
            }
        }
        private void getSub1Data(string colMainGroup,string colSub1Group, string dateColumn, string dateType, string dateValue)
        {
            string query = "select "+ colSub1Group+ " , count ( "+ colSub1Group+ " ) as dataCount from TableTempData " + datePArt+ " and "+ colSub1Group+ " <> N'إختر الإجراء' and "+ colMainGroup + " = N'"+ subComb0.Text +"' group by  " + colSub1Group;
            Console.WriteLine(query);
            //MessageBox.Show(query);
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            ListGroup = new string[table.Rows.Count];
            ListValue = new int[table.Rows.Count];
            int i = 0;
            foreach (DataRow dataRow in table.Rows)
            {
                ListValue[i] = Convert.ToInt32(dataRow["dataCount"].ToString());
                ListGroup[i] = dataRow[colSub1Group].ToString();
                i++;
            }
        }
        private void getSub1DataGen(string colMainGroup,string colSub1Group)
        {
            subComb1.Items.Clear();
            
            string query = "select "+ colSub1Group+ " , count ( "+ colSub1Group+" ) from TableTempData where "+ colSub1Group+ " <> N'إختر الإجراء' and "+ colMainGroup + " = N'"+ subComb0.Text +"' group by  " + colSub1Group;
            Console.WriteLine(query);
            //MessageBox.Show(query);
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            ListTimeLine = new string[table.Rows.Count];
            ListValue = new int[table.Rows.Count];
            int i = 0;
            foreach (DataRow dataRow in table.Rows)
            {
                subComb1.Enabled = true;   
                    subComb1.Items.Add(dataRow[colSub1Group].ToString());
                i++;
            }
        }
        private void getSub2DataGen(string colMainGroup,string colSub1Group)
        {
            subComb2.Items.Clear();
            
            string query = "select "+ colSub1Group+ " , count ( "+ colSub1Group+" ) from TableTempData where "+ colSub1Group+ " <> N'إختر الإجراء' and "+ colMainGroup + " = N'"+ subComb0.Text +"' group by  " + colSub1Group;
            Console.WriteLine(query);
            //MessageBox.Show(query);
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            ListTimeLine = new string[table.Rows.Count];
            ListValue = new int[table.Rows.Count];
            int i = 0;
            foreach (DataRow dataRow in table.Rows)
            {
                subComb1.Enabled = true;   
                    subComb1.Items.Add(dataRow[colSub1Group].ToString());
                i++;
            }
        }
        private void getTimeLine()
        {
            string query = "select DATEpart(MONTH, التاريخ_الميلادي) as timeLine ,count(*) as countTime from TableTempData where التاريخ_الميلادي like N'%"+comYear.Text+"' group by DATEpart(MONTH, التاريخ_الميلادي) order by DATEpart(MONTH, التاريخ_الميلادي)";
            //Console.WriteLine(query);
            //MessageBox.Show(query);
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            ListTimeLine = new string[table.Rows.Count];
            ListValue = new int[table.Rows.Count];
            int i = 0;
            foreach (DataRow dataRow in table.Rows)
            {
                ListValue[i] = Convert.ToInt32(dataRow["countTime"].ToString());
                ListTimeLine[i] = dataRow["timeLine"].ToString();                
                i++;
            }
        }


        private void itemsOfSubitems(ComboBox comboBox, bool multi)
        {
            if (shortRange) return;
            Console.WriteLine(comboBox.Items.Count.ToString());
            if (comboBox.Items.Count == 0 || shortRange) return;
            int[,] dataMonth = new int[12 * (comYear.Items.Count - 1), comboBox.Items.Count];
            int[] dataItem = new int[12 * (comYear.Items.Count - 1)];
            int[] sumItem = new int[12 * (comYear.Items.Count - 1)];
            int[] avgItem = new int[12 * (comYear.Items.Count - 1)];
            string[] nameItem = new string[12 * (comYear.Items.Count - 1)];

            int sumDays = 0;

            for (int item = 0; item < comboBox.Items.Count; item++)
            {
                int monthCount = 0;
                for (int year = 1; year < comYear.Items.Count; year++)
                    for (int month = 0; month < 12; month++)
                    {
                        sumDays = 0;
                        for (int day = 0; day < 31; day++)
                        {
                            sumDays = sumDays + DeepReport_Year[year - 1, month, day, item];
                        }
                        dataMonth[monthCount, item] = sumDays;
                        if (item == 0)
                            nameItem[monthCount] = comYear.Items[year].ToString() + "-" + Monthorder(month + 1);
                        Console.WriteLine("monthCount = " + monthCount.ToString() + " nameItem[monthCount] -- " + nameItem[monthCount] + " ---- sumDays = " + dataMonth[monthCount, item].ToString());
                        monthCount++;
                    }
            }
            int activeMonths = 1;
            for (int month = 0; month < 12 * (comYear.Items.Count - 1); month++)
            {
                sumItem[month] = 0;
                for (int item = 0; item < comboBox.Items.Count; item++)
                {

                    sumItem[month] += dataMonth[month, item];
                }
                if (sumItem[month] > 10) activeMonths++;
            }
            for (int month = 0; month < 12 * (comYear.Items.Count - 1); month++)
            {
                avgItem[month] = 1 + (sumItem[month] / comboBox.Items.Count);
            }

            string seriesName = " متوسط المعاملات";
            AddAchart2Data(comboBox, nameItem, seriesName, avgItem, "عدد المعاملات");
            holdData2 = true;
            for (int item = 0; item < comboBox.Items.Count; item++)
            {
                if (multi)
                {

                    for (int month = 0; month < 12 * (comYear.Items.Count - 1); month++)
                    {
                        dataItem[month] = dataMonth[month, item];
                        Console.WriteLine(nameItem[month] + " ---- dataItem = " + dataItem[month].ToString());

                    }
                    seriesName = comboBox.Items[item].ToString().Replace(" ", "_");
                    AddAchart2Data(comboBox, nameItem, seriesName, dataItem, "عدد المعاملات");
                    holdData2 = true;

                }
                else if (!multi && comboBox.SelectedIndex == item)
                {
                    for (int month = 0; month < 12 * (comYear.Items.Count - 1); month++)
                        dataItem[month] = dataMonth[month, item];

                    seriesName = comboBox.Items[item].ToString().Replace(" ", "_");
                    AddAchart2Data(comboBox, nameItem, seriesName, dataItem, "عدد المعاملات");
                }
            }
            holdData2 = false;
        }
        private void AddAchart2Data(ComboBox comboBox, string[] month_Year, string name, int[] data, string YaxisTitle)
        {
            //MessageBox.Show(data.Length.ToString());
            //MessageBox.Show(comboBox.Items.Count.ToString());
            if (!holdData2)
            {
                while (chart2.Series.Count > 0)
                {
                    chart2.Series.RemoveAt(0);
                    comSeriers2.Items.RemoveAt(0);

                }
                borderDashStyle2 = 0;
            }
            else borderDashStyle2++;

            bool found = false;
            for (int x = 0; x < comSeriers2.Items.Count; x++)
            {
                if (comSeriers2.Items[x].ToString() == name)
                {
                    found = true;
                    //comSeriers.Items.RemoveAt(comSeriers.SelectedIndex);
                    //chart2.Series.RemoveAt(comSeriers.SelectedIndex);
                }
            }
            if (chartAreas2 == 0)
            {
                chart2.ChartAreas[chartAreas2].AxisX.Title = "الفترة الزمنية";
                chart2.ChartAreas[chartAreas2].AxisY.Title = YaxisTitle;
                chart2.ChartAreas[chartAreas2].AxisX.LabelStyle.Interval = 1;
                chart2.ChartAreas[chartAreas2].AxisY.LabelStyle.Interval = Convert.ToInt32(combGrad2.Text); ;
                chart2.ChartAreas[chartAreas2].AxisY.LabelStyle.Font = new System.Drawing.Font("Arabic Typesetting", 16F, System.Drawing.FontStyle.Regular);
            }
            if (!found)
            {
                chart2.Series.Add(name);
                comSeriers2.Items.Add(name);
            }

            switch (borderDashStyle2)
            {
                case 0:
                    chart2.Series[name].BorderDashStyle = ChartDashStyle.Solid;
                    break;
                case 1:
                    chart2.Series[name].BorderDashStyle = ChartDashStyle.Dot;
                    break;
                case 2:
                    chart2.Series[name].BorderDashStyle = ChartDashStyle.Dash;
                    break;
                case 3:
                    chart2.Series[name].BorderDashStyle = ChartDashStyle.DashDot;
                    break;
                case 4:
                    chart2.Series[name].BorderDashStyle = ChartDashStyle.DashDotDot;
                    borderDashStyle2 = 0;
                    break;
            }

            chart2.Series[name].BorderWidth = 3;
            chart2.Series[name].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;


            for (int x = 0; x < data.Length; x++)
            {
                //MessageBox.Show(month.ToString());

                try
                {
                    //if(data[x-1] >= 10|| data[x + 1] >= 10)
                    Console.WriteLine(Monthorder(Convert.ToInt32(month_Year[x]))+" - "+ data[x].ToString());
                    chart2.Series[name].Points.AddXY(Monthorder(Convert.ToInt32(month_Year[x])), data[x]);
                }
                catch (Exception e)
                {

                }
            }
            MessageBox.Show("");
            //int overSum = 0;
            //string[] reArrange = new string[comboBox.Items.Count];
            //int[] itemIndex = new int[comboBox.Items.Count];

            //for (int x = 0; x < comboBox.Items.Count; x++)
            //    reArrange[x] = valueName[x].ToString();

            //for (int x = 0; x < comboBox.Items.Count; x++)

            //    itemIndex[x] = x;

            //for (int x = 0; x < data.Length; x++)
            //{
            //    overSum = overSum + data[x];
            //    for (int y = 0; y < data.Length - 1; y++)
            //        if (data[y] < data[y+1])
            //        {
            //            int temp1 = data[y];
            //            data[y] = data[y + 1];
            //            data[y + 1] = temp1;

            //            string temp2 = reArrange[y];
            //            reArrange[y] = reArrange[y + 1];
            //            reArrange[y + 1] = temp2;

            //            int temp3 = itemIndex[y];
            //            itemIndex[y] = itemIndex[y + 1];
            //            itemIndex[y + 1] = temp3;

            //        }
            //}
            //OverallStatisTime.Items.Clear();
            //for (int x = 0; x < data.Length; x++)
            //{
            //    int avg = 1;
            //    if (overSum != 0)
            //        avg = 1 + (100 * data[x]) / overSum;
            //    //OverallStatisTime.Items.Add(reArrange[x] + " % " + avg.ToString());

            //    //drawColumns(itemIndex[x-1], avg, reArrange[x-1]);
            //    //MessageBox.Show(data[x - 1].ToString());
            //}
            chartAreas2++;
        }
        private void itemsOfTime(ComboBox comboBox, bool column)
        {/*
            if (shortRange) return;
            //MessageBox.Show(comboBox.SelectedIndex.ToString());
            int[] dataChart = new int[comboBox.Items.Count ];
            int[] sumChart = new int[comboBox.Items.Count ];
            int[] sumYear = new int[ comboBox.Items.Count  ];
            int[] avgYear = new int[comboBox.Items.Count];
            for (int item = 0; item < comboBox.Items.Count ; item++)
                dataChart[item] = 0;
            //MessageBox.Show(dataChart.Length.ToString() + " -- " + (comboBox.Items.Count - 1).ToString());
            switch (comPeriode.SelectedIndex)
            {
                case 0:

                    for (int item = 0; item < comboBox.Items.Count; item++)
                        for (int day = 0; day < 31; day++)
                            dataChart[item] = DeepReport_Year[comYear.SelectedIndex - 1, comSubPeriode.SelectedIndex, day, item];

                    string name = "تقرير " + ReportType.Text + "لشهر " + comSubPeriode.Text + " للعام " + comYear.Text;

                    
                    
                    AddAchartData(comboBox, name.Replace(" ", "_"), dataChart, false, "عدد المباني والعقارات", column);
                    //holdData1 = true;
                    //AddAchartData(comboBox, name.Replace(" ", "_"), dataChart, false, "عدد المباني والعقارات",false);
                    break;
                case 1:
                    
                    int[] DeepReport_3Month = new int[comboBox.Items.Count];
                    for (int month = 0; month < 3; month++)
                        for (int item = 0; item < comboBox.Items.Count; item++)
                            DeepReport_3Month[item] = 0;

                    int sMonth = 0, eMonth = 2;
                    switch (comSubPeriode.SelectedIndex)
                    {
                        case 0:
                            sMonth = 0; eMonth = 2;
                            break;
                        case 1:
                            sMonth = 3; eMonth = 5;
                            break;
                        case 2:
                            sMonth = 6; eMonth = 8;
                            break;
                        case 3:
                            sMonth = 9; eMonth = 11;
                            break;
                    }

                    int[,] dataMonthChart = new int[3, comboBox.Items.Count  ];
                    for (int month = 0; month < 3; month++)
                        for (int item = 0; item < comboBox.Items.Count; item++)
                            dataMonthChart[month, item] = 0;

                    int m = 0;
                    int Month3Items = 0;
                    for (int month = sMonth; month <= eMonth; month++)
                    {
                        for (int day = 0; day < 31; day++) 
                        {
                            //MessageBox.Show("item " + item.ToString());
                            for (int item = 0; item < comboBox.Items.Count; item++)
                            {
                                DeepReport_3Month[item] = DeepReport_3Month[item] + DeepReport_Year[comYear.SelectedIndex - 1, month, day, item];
                                dataMonthChart[m, item] = DeepReport_3Month[item];
                                if(item == 0) Month3Items+= DeepReport_3Month[item];
                            }
                        }
                        m++;
                    }
                    
                    //MessageBox.Show("Month3Items " + Month3Items.ToString());

                    switch (ReportType.SelectedIndex)
                    {
                        case 2:
                            
                            for (int month = 0; month < 3; month++)
                            {
                                for (int item = 0; item < comboBox.Items.Count; item++)
                                    dataChart[item] = dataMonthChart[month, item] ;
                                
                                nameSeries = "تقرير شهر " + Monthorder(sMonth + month) + " للعام " + comYear.Text;
                                AddAchartData(comboBox, nameSeries.Replace(" ", "_"), dataChart, false, "عدد المباني والعقارات", column);
                                //holdData1 = true; 
                                //AddAchartData(comboBox, nameSeries.Replace(" ", "_"), dataChart, false, "عدد المباني والعقارات",false);
                            }
                            break;
                        case 1:

                            for (int item = 0; item < comboBox.Items.Count; item++)
                                sumChart[item] = 0;

                            for (int month = 0; month < 3; month++)
                            {
                                for (int item = 0; item < comboBox.Items.Count; item++)
                                    sumChart[item] = sumChart[item] + dataMonthChart[month, item];

                            }
                            nameSeries = "المعدل التراكمي " + comSubPeriode.Text + " للعام " + comYear.Text;
                            AddAchartData(comboBox, nameSeries.Replace(" ", "_"), sumChart, true, "عدد المباني والعقارات", column);
                            //holdData1 = true; 
                            //AddAchartData(comboBox, nameSeries.Replace(" ", "_"), sumChart, true, "عدد المباني والعقارات",false);
                            break;
                        case 0:
                            for (int item = 0; item < comboBox.Items.Count; item++)
                                sumChart[item] = 0;
                            int items = 0;
                            int OverAllSum = 0;
                            for (int month = 0; month < 3; month++)
                            {
                                for (int item = 0; item < comboBox.Items.Count; item++)
                                {
                                    sumChart[item] = sumChart[item] + dataMonthChart[month, item];
                                    OverAllSum+= sumChart[item];
                                }
                                items++;
                            }

                            int[] AvaChart = new int[comboBox.Items.Count];
                            for (int item = 0; item < comboBox.Items.Count; item++)
                                AvaChart[item] = 0;
                            for (int month = 0; month < 3; month++)
                            {
                                for (int item = 0; item < comboBox.Items.Count; item++)
                                    AvaChart[item] = sumChart[item] / items;
                            }
                            //btnSumStatis.Text = "إحصائية لعدد " + OverAllSum.ToString() + " معاملة";
                            nameSeries = "متوسط " + comSubPeriode.Text + " للعام " + comYear.Text;
                            string yAxis= "متوسط معاملات " + comSubPeriode.Text + " للعام " + comYear.Text;
                            AddAchartData(comboBox, nameSeries.Replace(" ", "_"), AvaChart, true, "متوسط عدد المباني والعقارات", column);
                            //holdData1 = true; 
                            //AddAchartData(comboBox, nameSeries.Replace(" ", "_"), AvaChart, true, "متوسط عدد المباني والعقارات",false);
                            break;

                    }


                    break;
                case 2:
                    for (int item = 0; item < comboBox.Items.Count; item++)
                        sumChart[item] = 0;


                    for (int item = 0; item < comboBox.Items.Count; item++)
                    {
                        avgYear[item] = sumYear[ item] = 0;
                    }

                    
                    for (int month = 0; month < 12; month++)
                        for (int item = 0; item < comboBox.Items.Count; item++)
                            for (int day = 0; day < 31; day++)
                            {
                                sumYear[ item] = sumYear[item] + DeepReport_Year[comYear.SelectedIndex - 1, month, day, item];
                                Console.WriteLine("sumYear " + sumYear[item]);
                            }
                    for (int item = 0; item < comboBox.Items.Count; item++)
                    {
                        sumChart[item] = sumChart[item] + sumYear[item];
                        Console.WriteLine(sumChart[item]);
                    }

                    switch (ReportType.SelectedIndex)
                    {
                        
                        case 1:
                            //MessageBox.Show(comboBox.SelectedIndex.ToString() + " - "+comboBox.Text);
                            nameSeries = "تراكمي العام " + comYear.Text;
                            AddAchartData(comboBox, subComb0.Text, sumYear, true, nameSeries, column);
                            //holdData1 = true; 
                            //AddAchartData(comboBox, nameSeries.Replace(" ", "_"), sumYear, true, nameSeries,false);
                            break;
                        case 0:
                            //MessageBox.Show(ReportType.Text);
                            for (int item = 0; item < comboBox.Items.Count - 1; item++)
                            {
                                avgYear[item] = sumChart[item] / 365;
                                Console.WriteLine("avgYear " + avgYear[item]);
                            }

                            nameSeries = "متوسط العام " + comYear.Text;
                            AddAchartData(comboBox, nameSeries.Replace(" ", "_"), sumChart, true, nameSeries, column);
                            //holdData1 = true; 
                            //AddAchartData(comboBox, nameSeries.Replace(" ", "_"), sumChart, true, nameSeries,false);
                            break;
                    }
                    break;

                case 3:
                    for (int item = 0; item < comboBox.Items.Count; item++)
                        sumChart[item] = 0;


                    for (int item = 0; item < comboBox.Items.Count; item++)
                    {
                        avgYear[item] = sumYear[item] = 0;
                    }


                    for (int month = 0; month < 12; month++)
                        for (int item = 0; item < comboBox.Items.Count; item++)
                            for (int day = 0; day < 31; day++)
                            {
                                sumYear[item] = sumYear[item] + DeepReport_Year[comYear.SelectedIndex - 1, month, day, item];
                                Console.WriteLine("sumYear " + sumYear[item]);
                            }
                    switch (ReportType.SelectedIndex)
                    {
                        case 1:

                            for (int item = 0; item < comboBox.Items.Count - 1; item++)
                                sumChart[item] = 0;

                            for (int item = 0; item < comboBox.Items.Count - 1; item++)
                                sumChart[item] = sumChart[item] + sumYear[ item];

                            nameSeries = "تراكمي حميع الأعوام " + comYear.Text;
                            AddAchartData(comboBox, nameSeries.Replace(" ", "_"), sumChart, true, nameSeries, column);
                            //holdData1 = true; 
                            //AddAchartData(comboBox, nameSeries.Replace(" ", "_"), sumChart, true, nameSeries,false);
                            break;
                        case 2:
                            for (int item = 1; item < comboBox.Items.Count - 1; item++)
                                avgYear[item] = 0;

                            for (int item = 1; item < comboBox.Items.Count - 1; item++)
                                avgYear[item] = sumChart[item] / (comYear.Items.Count - 1);
                            nameSeries = "متوسط حميع الأعوام " + comYear.Text;
                            AddAchartData(comboBox, nameSeries.Replace(" ", "_"), avgYear, true, nameSeries, column);
                            //holdData1 = true; 
                            //AddAchartData(comboBox, nameSeries.Replace(" ", "_"), avgYear, true, nameSeries,false);
                            break;
                    }
                    break;
            }*/
        }

        private string Monthorder(int month)
        {
            switch (month)
            {
                case 1:
                    return "يناير";


                case 2:
                    return "فبراير";


                case 3:
                    return "مارس";


                case 4:
                    return "ابريل";

                case 5:
                    return "مايو";


                case 6:
                    return "يونيو";


                case 7:
                    return "يوليو";


                case 8:
                    return "أغسطس";

                case 9:
                    return "سبتمبر";


                case 10:
                    return "اكتوبر";


                case 11:
                    return "نوفمبر";


                case 12:
                    return "ديسمبر";
                default:
                    return "";

            }
        }
        
        private string Monthorder(string month)
        {
            switch (month)
            {
                case "1":
                    return "يناير";


                case "2":
                    return "فبراير";


                case "3":
                    return "مارس";


                case "4":
                    return "ابريل";

                case "5":
                    return "مايو";


                case "6":
                    return "يونيو";


                case "7":
                    return "يوليو";


                case "8":
                    return "أغسطس";

                case "9":
                    return "سبتمبر";


                case "10":
                    return "اكتوبر";


                case "11":
                    return "نوفمبر";


                case "12":
                    return "ديسمبر";
                default:
                    return "";

            }
        }

        private string MonthorderInNumber(int month)
        {
            switch (month)
            {
                case 1:
                    return "01";


                case 2:
                    return "02";


                case 3:
                    return "03";


                case 4:
                    return "04";

                case 5:
                    return "05";


                case 6:
                    return "06";


                case 7:
                    return "07";


                case 8:
                    return "08";

                case 9:
                    return "09";


                case 10:
                    return "10";


                case 11:
                    return "11";


                case 12:
                    return "12";
                default:
                    return "";

            }
        }

        private void AuthTypes(int rows, string reportName, int month)
        {
            string route = FilespathIn + "نوع_التواكيل.docx";
            string ActiveCopy = FilespathOut + reportName;
            System.IO.File.Copy(route, ActiveCopy);
            using (var document = DocX.Load(ActiveCopy))
            {
                System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);

                string strHeader = "الرقم : " + "     " + "التاريخ :" + " م" + "     " + "الموافق : " + "هـ" + Environment.NewLine;
                document.InsertParagraph(strHeader)
                .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                .FontSize(16d)
                .Alignment = Alignment.center;
                for (int year = 0; year < comYear.Items.Count - 1; year++)
                {
                    string MessageDir = "التقرير المفصل للعام " + comYear.Items[year].ToString()
                        + Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ";
                    document.InsertParagraph(MessageDir)
                        .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                        .FontSize(18d)
                        .Direction = Direction.RightToLeft;

                    var t = document.AddTable(1, month + 2);
                    t.Design = TableDesign.TableGrid;
                    t.Alignment = Alignment.center;
                    t.SetColumnWidth(month, 140);
                    t.SetColumnWidth(month + 1, 40);
                    for (int m = 1; m <= month; m++)
                    {
                        t.SetColumnWidth(m - 1, 50);
                        t.Rows[0].Cells[12 - m].Paragraphs[0].Append(Monthorder(m)).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    }
                    t.Rows[0].Cells[month].Paragraphs[0].Append("البند").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    t.Rows[0].Cells[month + 1].Paragraphs[0].Append("الرقم").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;
                    int x = 1;
                    //DeepReport[d, x]
                    for (int count = 1; count <= rows; count++)

                    {
                        t.InsertRow();
                        //for (int m = 0; m < month; m++)
                        //    t.Rows[x].Cells[11 - m].Paragraphs[0].Append(DeepReport_Year[year, m, count - 1].ToString()).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Bold().Alignment = Alignment.center;

                        t.Rows[x].Cells[month].Paragraphs[0].Append(subComb0.Items[count - 1].ToString()).Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                        t.Rows[x].Cells[month + 1].Paragraphs[0].Append(x.ToString() + ".").Font(new Xceed.Document.NET.Font("Arabic Typesetting")).FontSize(20d).Direction = Direction.RightToLeft;
                        x++;
                    }



                    var p = document.InsertParagraph(Environment.NewLine);
                    p.InsertTableAfterSelf(t);
                }

                string strAttvCo = Environment.NewLine + "ـــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ" + Environment.NewLine + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + Environment.NewLine + "\t\t\t\t\t\t\t\t\t\t" + "ع/ القنصل العام بالإنابة";
                var AttvCo = document.InsertParagraph(strAttvCo)
                    .Font(new Xceed.Document.NET.Font("Arabic Typesetting"))
                    .FontSize(20d)
                    .Bold()
                    .Alignment = Alignment.center;


                document.Save();
                Process.Start("WINWORD.EXE", ActiveCopy);

            }



        }

        private int daysOfMonth(int month, int year)
        {
            Months[0] = 31;

            if (year % 400 == 0 || (year % 4 == 0 && year % 100 != 0))
                Months[1] = 29;
            else Months[1] = 28;

            Months[2] = 31;

            Months[3] = 30;
            Months[4] = 31;
            Months[5] = 30;

            Months[6] = 31;
            Months[7] = 31;
            Months[8] = 30;

            Months[9] = 31;
            Months[10] = 30;
            Months[11] = 31;

            return Months[month];
        }

        private void comYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            comPeriode.Visible = true;


        }

        private void UpdateCase(int id, string text, string col)
        {

            string qurey = "update TableSuitCase set " + col + "=@" + col + " where ID=@id";
            Console.WriteLine(qurey);

            string colQuery = "@" + col;
            SqlConnection sqlCon = new SqlConnection(dataSource56);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue(colQuery, text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();

        }
        private void UpdateCasetype(int id, string text)
        {

            string qurey = "update TableSuitCase set القضية=@القضية where ID=@id";
            Console.WriteLine(qurey);

            SqlConnection sqlCon = new SqlConnection(dataSource56);
            SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("القضية", text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();

        }

       

        private void UpdateState(int id, string text, string col,string table)
        {
            
            string qurey = "update " + table + " set "+ col + "=@"+ col+" where ID=@id";
            Console.WriteLine(qurey);
            //
            string colQuery = "@" + col;

            try
            {
                SqlConnection sqlCon = new SqlConnection(dataSource56);
                SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@id", id);
                sqlCmd.Parameters.AddWithValue(colQuery, text);
                sqlCmd.ExecuteNonQuery();
                sqlCon.Close();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(text);
            }
            

        }

        private void UpdateSecCase(int id, string text)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource56);
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select * from TableSuitCase" , sqlCon);

            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl1 = new DataTable();
            sqlDa1.Fill(dtbl1);
            foreach (DataRow dataRow in dtbl1.Rows)
            {
                if (dataRow["ID"].ToString() == id.ToString())
                {
                    string Suitcase = text;
                    string finishDate = dataRow["تاريخ_الرفع"].ToString();
                    string receiveDate = dataRow["تاريخ_الاستلام"].ToString();
                    string messDate = dataRow["تاريخ_لبرقية"].ToString();
                    string messID = dataRow["رقم_لبرقية"].ToString();
                    string messName = dataRow["مقدم_الطلب"].ToString();
                    string messComment = "لا تعليق";
                    string GregorianDate = dataRow["التاريخ_الميلادي"].ToString();
                    string HijriDate = dataRow["التاريخ_الهجري"].ToString();
                    string attendedVC = dataRow["مدير_القسم"].ToString();
                    string ConsulateEmployee = dataRow["اسم_الموظف"].ToString();
                    NewMandoubData(Suitcase, finishDate, receiveDate, messDate, messID, messName, messComment, GregorianDate, HijriDate, attendedVC, ConsulateEmployee);
                }
            }
        }


        private void NewMandoubData(string Suitcase, string finishDate, string receiveDate, string messDate, string messID, string messName, string messComment, string GregorianDate, string HijriDate, string attendedVC, string ConsulateEmployee)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource56);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();

            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableSuitCase (رقم_لبرقية, تاريخ_لبرقية, مقدم_الطلب, القضية, تاريخ_الاستلام, تاريخ_الرفع, التاريخ_الميلادي, التاريخ_الهجري, مدير_القسم, اسم_الموظف, تعليق)  values (@رقم_لبرقية, @تاريخ_لبرقية, @مقدم_الطلب, @القضية, @تاريخ_الاستلام, @تاريخ_الرفع, @التاريخ_الميلادي, @التاريخ_الهجري, @مدير_القسم, @اسم_الموظف, @تعليق) ", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_لبرقية", messID);
            sqlCmd.Parameters.AddWithValue("@تاريخ_لبرقية", messDate);
            sqlCmd.Parameters.AddWithValue("@مقدم_الطلب", messName);
            sqlCmd.Parameters.AddWithValue("@القضية", Suitcase);
            sqlCmd.Parameters.AddWithValue("@تاريخ_الاستلام", receiveDate);
            sqlCmd.Parameters.AddWithValue("@تاريخ_الرفع", finishDate);
            sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", GregorianDate);
            sqlCmd.Parameters.AddWithValue("@التاريخ_الهجري", HijriDate);
            sqlCmd.Parameters.AddWithValue("@مدير_القسم", attendedVC);
            sqlCmd.Parameters.AddWithValue("@اسم_الموظف", ConsulateEmployee);
            sqlCmd.Parameters.AddWithValue("@تعليق", messComment);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void addCollectiveData(string text1, string text2, string text3, string text4, string text5, string text6)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource56);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableCollective (جهة_العمل,نوع_المعاملة,المهن_المعدلة,رقم_الهوية,التاريخ_الميلادي,الحالة,تصنيف_عام) values (@جهة_العمل,@نوع_المعاملة,@المهن_المعدلة,@رقم_الهوية,@التاريخ_الميلادي,@الحالة,@تصنيف_عام)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@المهن_المعدلة", text1);
            sqlCmd.Parameters.AddWithValue("@نوع_المعاملة", text2);
            sqlCmd.Parameters.AddWithValue("@جهة_العمل", text3);
            sqlCmd.Parameters.AddWithValue("@رقم_الهوية", text4);
            sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", text5);
            sqlCmd.Parameters.AddWithValue("@الحالة", text6);
            sqlCmd.Parameters.AddWithValue("@تصنيف_عام", text1);
            
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        
        private void correctData(string wronge, string TrueData, string col, string table)
        {
            //MessageBox.Show(" table " + table+ " col " + col);
            SqlConnection sqlCon = new SqlConnection(dataSource);
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select ID,"+ col+" from "+ table, sqlCon);
            
            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl1 = new DataTable();
            sqlDa1.Fill(dtbl1);
            foreach (DataRow dataRow in dtbl1.Rows)
            {

                if (dataRow[col].ToString().Contains (wronge))
                {
                    
                    //MessageBox.Show(wronge + " -- " + dataRow["ID"].ToString());
                    //    string str = dataRow["FullTextData"].ToString().Split('-')[7];
                    //if (str.Contains('*'))  
                    //    str = str.Split('*')[0]; 

                    UpdateState(Convert.ToInt32(dataRow["ID"].ToString()), TrueData, col, table);      
                }

            }
        }

        private void splitData(string wronge, string TrueData, string col)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource56);
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select ID," + col + " from TableSuitCase" , sqlCon);
            //MessageBox.Show("select ID," + col + " from " + columnList[0]);
            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl1 = new DataTable();
            sqlDa1.Fill(dtbl1);
            foreach (DataRow dataRow in dtbl1.Rows)
            {

                if (dataRow[col].ToString().Contains (wronge))
                {
                    //MessageBox.Show(wronge + " -- " +True);
                    //    string str = dataRow["FullTextData"].ToString().Split('-')[7];
                    //if (str.Contains('*'))  
                    //    str = str.Split('*')[0]; 

                    UpdateCase(Convert.ToInt32(dataRow["ID"].ToString()), TrueData, "نوع_عام");
                }

            }
        }

        private void swapData(string text2, string text1, string col1, string col2)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select ID," + col1 +"," + col2+ " from " + columnList[0], sqlCon);
            //MessageBox.Show("select ID," + col1 + "," + col2 + " from " + columnList[0]);
            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl1 = new DataTable();
            sqlDa1.Fill(dtbl1);
            foreach (DataRow dataRow in dtbl1.Rows)
            {
                int id = Convert.ToInt32(dataRow["ID"].ToString());
                //if (dataRow[col1].ToString() == text1 && dataRow[col2].ToString() == text2)
                if (dataRow[col2].ToString() == text2)
                {
                    string qurey = "update " + columnList[0] + " set " + col1 + "=@" + col1 +","+ col2 + "=@" + col2 + " where ID=@id";
                    Console.WriteLine(qurey);
                    //MessageBox.Show(dataRow["ID"].ToString());
                    string colQuery1 = "@" + col1;
                    string colQuery2 = "@" + col2;
                    SqlCommand sqlCmd = new SqlCommand(qurey, sqlCon);
                    if (sqlCon.State == ConnectionState.Closed)
                        sqlCon.Open();
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@id", id);
                    sqlCmd.Parameters.AddWithValue(colQuery1, text2);
                    sqlCmd.Parameters.AddWithValue(colQuery2, text1);
                    sqlCmd.ExecuteNonQuery();
                    sqlCon.Close();
                }
            }
        }

        private void returnID(string text2, string text1, string col1, string col2)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select ID," + col1 + "," + col2 + " from " + columnList[0], sqlCon);
            //MessageBox.Show("select ID," + col1 + "," + col2 + " from " + columnList[0]);
            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl1 = new DataTable();
            sqlDa1.Fill(dtbl1);
            foreach (DataRow dataRow in dtbl1.Rows)
            {
                if (dataRow[col1].ToString() == text1 && dataRow[col2].ToString() == text2)
                {
                    txtGetID.Text = dataRow["ID"].ToString();
                }
            }
        }

        private void DeepStatistics_Load(object sender, EventArgs e)
        {
            fileComboBox(genTypes, dataSource, "iteminfo", "TableStatisInfo");
            combGrad1.SelectedIndex = 2;
            combGrad2.SelectedIndex = 2;

        }

        private string DeepStatistInfo(int id)
        {
            //MessageBox.Show(id.ToString());
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableStatisInfo", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@ID", id);
            DataTable rowTable = new DataTable();
            sqlDa.Fill(rowTable);
            sqlCon.Close();
            string items = "select ID";
            string value = "";
            int x = 1;
            dataSource = dataSource57;
            foreach (DataRow dataRow in rowTable.Rows)
            {
                if (dataRow["iteminfo"].ToString() == genTypes.Text)
                {
                    columnList[0] = dataRow["tablesName"].ToString();
                    columnDate = dataRow["col1"].ToString();
                    for (; x <= 10; x++)
                        if (dataRow["Col" + x.ToString()].ToString() != "")
                        {
                            items = items + ", " + dataRow["Col" + x.ToString()].ToString();
                            if (x >= 3)
                            {
                                columnList[x - 2] = dataRow["Col" + x.ToString()].ToString();
                                //MessageBox.Show("Col" + x.ToString() + " - " + columnList[x - 2]);
                            }
                        }
                    items = items + " from " + columnList[0];
                    if (dataRow["databaseFile"].ToString() == "56") 
                        dataSource = dataSource56;
                    dateItems = dataRow["col1"].ToString();
                    symbol = dataRow["dataSymbol"].ToString();
                    symbolChar = Convert.ToChar(dataRow["dataSymbol"].ToString());
                }                
            }
            //MessageBox.Show(items);
            return items;
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
                    if (!string.IsNullOrEmpty(dataRow[comlumnName].ToString()) && !dataRow[comlumnName].ToString().Contains("merg"))
                    {
                        combbox.Items.Add(dataRow[comlumnName].ToString());
                    }
                }
                saConn.Close();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            chartAreas2 = chartAreas1 = 0;
            //if(subItems[genTypes.SelectedIndex, 0] == "") return; 
            subComb0.Items.Clear();
            
            comYear.Items.Clear();
            comYear.Items.Add("جميع الأعوام");
            
            //MessageBox.Show(DeepStatistInfo(genTypes.SelectedIndex+1));
            queryInfo = DeepStatistInfo(genTypes.SelectedIndex+1);
            
            // subBtn0.Text =
            //itemsType(columnList[0], columnList[1], dateItems, true, subComb0).ToString();
            createTable(genTypes.Text);
            fillYears(comYear, columnDate);
            getMainDataGen(columnList[1]);
            //getGroupCount0(columnList[1], dateLike, subComb0);
            //for (int x = 0; x < itemCount.Length; x++)
            //{
            //    Console.WriteLine(itemCount[x].ToString());
            //}
            //string name = "تقرير " + ReportType.Text + "لشهر " + comSubPeriode.Text + " للعام " + comYear.Text;
            //AddAchartData(subComb0, name.Replace(" ", "_"), itemCount, false, "عدد معاملات " + subComb0.Text, true);

            //if (btnPlotSearch.Text == "فحص")
            //{
            //    DeepStatics(genTypes);
            //    AllStatistData();
            //    prePareToShow(subComb0, genTypes);
            //}
            //else
            //{
            //    prePareToShow(subComb0, genTypes);
            //}
            subComb1.Items.Clear(); subComb1.Text = "المعاملة";
            subComb2.Items.Clear(); subComb2.Text = "المعاملة";
            subComb3.Items.Clear(); subComb3.Text = "المعاملة";
            subComb4.Items.Clear(); subComb4.Text = "المعاملة";

        }

        private void fillYears(ComboBox combo, string column)
        {
            combo.Items.Clear();
            string query = "select distinct DATENAME(YEAR, "+ column+ ")  as years from TableTempData where DATENAME(YEAR, " + column+") like '20%' order by DATENAME(YEAR, " + column+") desc";
            //Console.WriteLine(query);
            //MessageBox.Show(query);
            SqlConnection Con = new SqlConnection(dataSource);
            if (Con.State == ConnectionState.Closed)
                try
                {
                    Con.Open();
                    SqlDataAdapter sqlDa = new SqlDataAdapter(query, Con);
                    sqlDa.SelectCommand.CommandType = CommandType.Text;
                    DataTable dtbl2 = new DataTable();
                    sqlDa.Fill(dtbl2);
                    Con.Close();
                    foreach (DataRow dataRow in dtbl2.Rows)
                    {
                        combo.Items.Add(dataRow["years"].ToString());
                        combo.Visible = true;
                        combo.Text = "إختر العام";
                    }
                }
                catch (Exception ex) { }
        }
        private void createTable(string proName)
        {
            SqlConnection Con = new SqlConnection(dataSource);
            if (Con.State == ConnectionState.Closed)
                try
                {
                    Con.Open();
                    SqlDataAdapter sqlDa = new SqlDataAdapter(proName, Con);
                    sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                    DataTable dtbl2 = new DataTable();
                    sqlDa.Fill(dtbl2);
                    Con.Close();
                }
                catch (Exception ex) { }
        }
        private void prePareToShow(ComboBox comboBox, ComboBox gen)
        {
            shortRange = false;
            int itemOverAll = 0;
            string[] range = getRange(gen).Split('-');
            int start = Convert.ToInt32(range[0]);
            int end = Convert.ToInt32(range[1]);
            Console.WriteLine(gen.Text + " start " + start.ToString() + " end " + end.ToString());
            if (start == 0 && end == 0) { shortRange = true; return; }
            //Console.WriteLine(" --year-- " + year.ToString() + " --month-- " + month.ToString() + " --day-- " + day.ToString() + " --count-- " + count.ToString());
            
            DeepReport_Year = new int[comYear.Items.Count-1, 12, 32, comboBox.Items.Count];

            for (int w = 0;w < comYear.Items.Count - 1; w++)
                for (int x = 0; x < 12; x++)
                for (int y = 0; y < 31; y++)
                    for (int z = 0; z < comboBox.Items.Count; z++)
                        DeepReport_Year[w, x, y, z] = 0;

            
            for (int year = 1; year < comYear.Items.Count; year++)
            {
                int countItem = 0;
                int CurrentYear = Convert.ToInt32(comYear.Items[year].ToString());
                
                for (int row = 0; row < dataGridView1.RowCount - 1; row++)
                {
                    string cellDate = dataGridView1.Rows[row].Cells[1].Value.ToString();
                    for (int month = 0; month < 12; month++)
                    {
                        for (int day = 1; day <= daysOfMonth(month, CurrentYear); day++)
                        {
                            string CurrentDay = "", Currentmonth = "", CurrentDate = "";
                            if ((month + 1) < 10) Currentmonth = "0" + (month + 1).ToString();
                            else Currentmonth = (month + 1).ToString();
                            if (day < 10) CurrentDay = "0" + day.ToString();
                            else CurrentDay = day.ToString();
                            CurrentDate = CurrentDay + "-" + Currentmonth + "-" + CurrentYear.ToString();
                            if (cellDate == CurrentDate)
                            {
                                int countItems = 0;
                                
                                for (int item = start; item <= end; item++)
                                {
                                    
                                    string cellInfo = dataGridView1.Rows[row].Cells[item].Value.ToString();
                                    int count = 0;
                                    if (!string.IsNullOrEmpty(cellInfo))
                                        count = Convert.ToInt32(cellInfo);
                                    if (item == start) { countItem += count; }
                                    DeepReport_Year[year - 1, month, day - 1, countItems] = count;
                                    itemOverAll = itemOverAll + count;
                                    //if (comboIndex == 3)
                                    //{
                                        //MessageBox.Show(comboBox.Items[countItem].ToString());
                                        Console.WriteLine(" --year-- " + year.ToString() + " --month-- " + month.ToString() + " --day-- " + day.ToString() + " --count-- " + count.ToString());
                                    //}
                                    countItems++;
                                }
                            }
                        }
                    }
                }
                //MessageBox.Show("countItem " + countItem.ToString());
            }
        }

        private void btnInvisible_Click(object sender, EventArgs e)
        {
            if (chart1.Series[comSeriers.Text].Enabled)
                chart1.Series[comSeriers.Text].Enabled = false;
            else chart1.Series[comSeriers.Text].Enabled = true;

        }

        private void comSeriers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (chart1.Series[comSeriers.Text].Enabled)
                btnInvisible.Text = "إخفاء";
            else btnInvisible.Text = "إظهار";
            //btnReport.Visible = btnDelete.Visible = btnInvisible.Visible = true;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            
            chart1.Series.RemoveAt(comSeriers.SelectedIndex);
            comSeriers.Items.RemoveAt(comSeriers.SelectedIndex);
            var partAll = MessageBox.Show("", "حذف الجميع؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (partAll == DialogResult.Yes)
            {
                while (chart1.Series.Count > 0)
                {
                    chart1.Series.RemoveAt(0);
                    if (comSeriers.Items.Count > 0) comSeriers.Items.RemoveAt(0);
                }
                
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ReportName = "Report" + DateTime.Now.ToString("mmss") + ".docx";
            AuthTypes(subComb0.Items.Count, ReportName, 12);
        }

        private void comPeriode_SelectedIndexChanged(object sender, EventArgs e)
        {
            comSubPeriode.Items.Clear();
            switch (comPeriode.SelectedIndex)
            {
                case 0:
                    dateType = "MONTH";
                    for (int month = 1; month <= 12; month++)
                        comSubPeriode.Items.Add(MonthorderInNumber(month));
                    comSubPeriode.Text = "إختر الشهر";
                    comPeriode.Location = new System.Drawing.Point(1320, 18);
                    comPeriode.Size = new System.Drawing.Size(119, 35);
                    comSubPeriode.Visible = true;
                    ReportType.Visible = false;
                    
                    //comSubPeriode.Items.Add(Monthorder(month));
                    break;
                case 1:
                    dateType = "QUARTER";
                    comSubPeriode.Items.Add("الربع الأول");
                    comSubPeriode.Items.Add("الربع الثاني");
                    comSubPeriode.Items.Add("الربع الثالث");
                    comSubPeriode.Items.Add("الربع الرابع");
                    comSubPeriode.Text = "إختر ربع العام";
                    comPeriode.Location = new System.Drawing.Point(1320, 18);
                    comPeriode.Size = new System.Drawing.Size(119, 35);
                    comSubPeriode.Visible = true;
                    ReportType.Visible = false;
                    break;
                case 2:
                    dateType = "YEAR";
                    dateValue = comYear.Text;
                    comSubPeriode.Visible = false;
                    ReportType.Visible = true;
                    name = "تقرير " + ReportType.Text + " عام " + comYear.Text;
                    dateLike = "التاريخ_الميلادي like '%" + comYear.Text + "'";
                    DateSearch = "01-02-03-04-05-06-07-08-09-10-11-12";

                    ReportType.Text = "إختر نوع التقرير";
                    ReportType.SelectedIndex = 0;
                    break;
                case 3:
                    dateType = "ALL";
                    dateValue = "";
                    comSubPeriode.Visible = false;
                    ReportType.Visible = true;
                    name = "تقرير " + ReportType.Text;
                    dateLike = "التاريخ_الميلادي like '%'";
                    ReportType.Text = "إختر نوع التقرير";
                    ReportType.SelectedIndex = 0;
                    break;
            }
        }

        private void comDeepPeriode_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comPeriode.SelectedIndex)
            {
                case 0:
                    //DateSearch = "01-02-03-04-05-06-07-08-09-10-11-12-13-14-15-16-17-18-19-20-21-22-23-24-25-26-27-28-29-30-31";

                    name = "تقرير شهر " + Monthorder(Convert.ToInt32(comSubPeriode.Text)) + " للعام " + comYear.Text;
                    //dateLike = "التاريخ_الميلادي like '%-"+ comSubPeriode.Text + "-" + comYear.Text + "'";
                    //ReportType.Text = "إختر نوع التقرير";
                    dateValue = comSubPeriode.Text;
                    ReportType.SelectedIndex = 0;
                    break;
                case 1:
                    name = "تقرير " + ReportType.Text + " " + comSubPeriode.Text + " للعام " + comYear.Text;
                    switch (comSubPeriode.SelectedIndex) {
                        case 0:
                            dateValue = "1";
                            //DateSearch = "01-02-03";
                            //dateLike = "التاريخ_الميلادي like '%-01-" + comYear.Text + "' or التاريخ_الميلادي like '%-02-" + comYear.Text + "' or التاريخ_الميلادي like '%-03-" + comYear.Text + "'";
                            break;
                        case 1:
                            dateValue = "2";
                            //DateSearch = "04-05-06";
                            //dateLike = "التاريخ_الميلادي like '%-04-" + comYear.Text + "' or التاريخ_الميلادي like '%-05-" + comYear.Text + "' or التاريخ_الميلادي like '%-06-" + comYear.Text + "'";
                            break;
                        case 2:
                            dateValue = "3";
                            //DateSearch = "07-08-09";
                            //dateLike = "التاريخ_الميلادي like '%-07-" + comYear.Text + "' or التاريخ_الميلادي like '%-08-" + comYear.Text + "' or التاريخ_الميلادي like '%-09-" + comYear.Text + "'";
                            break;
                        case 3:
                            dateValue = "4";
                            //DateSearch = "10-11-12";
                            //dateLike = "التاريخ_الميلادي like '%-10-" + comYear.Text + "' or التاريخ_الميلادي like '%-11-" + comYear.Text + "' or التاريخ_الميلادي like '%-12-" + comYear.Text + "'"; 
                            break;
                    }
                    ReportType.Text = "إختر نوع التقرير";
                    ReportType.SelectedIndex = 0;
                    break;
                
            }
            ReportType.Visible = true;
        }

        //private void drawColumns(int index, int ratio, string name) {
        //    int color = (255 * ratio) / 100;
        //    //Console.WriteLine(" color " + color.ToString());
            
        //    Button PrintReport = new Button();
        //    PrintReport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(color)))));
        //    PrintReport.FlatAppearance.BorderSize = 0;
        //    PrintReport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
        //    PrintReport.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        //    PrintReport.Location = new System.Drawing.Point(2, 3);
        //    PrintReport.Name = index.ToString();
        //    PrintReport.RightToLeft = System.Windows.Forms.RightToLeft.No;
        //    PrintReport.Size = new System.Drawing.Size(285, 34);
        //    PrintReport.TabIndex = 470;
        //    PrintReport.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
        //    PrintReport.Text = name + " % " + ratio.ToString();
        //    PrintReport.UseVisualStyleBackColor = false;
        //    PrintReport.Visible = true;
        //    PrintReport.Click += new System.EventHandler(PrintReport_Click);

        //    ReportPanel.Controls.Add(PrintReport);
        //}
        private void PrintReport_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            //MessageBox.Show(subAuthTypes.Items[Convert.ToInt32(button.Name)].ToString());
        }

        //private string addcombp(string name) {
        //    ComboBox comboBox = new ComboBox();
        //    comboBox.Font = new System.Drawing.Font("Arabic Typesetting", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        //    comboBox.FormattingEnabled = true;
        //    comboBox.Location = new System.Drawing.Point(3, 3);
        //    comboBox.Name = "newCombo" + comboIndex;
        //    comboBox.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
        //    comboBox.Size = new System.Drawing.Size(296, 35);
        //    comboBox.TabIndex = 843;
        //    comboBox.Text = name;
        //    ReportPanel.Controls.Add(comboBox);
        //    comboIndex++;
        //    return "newCombo" + (comboIndex - 1).ToString();
        //}

        public int filterMainRowWithDate(string date)
        {
            string paraDate = "@" + dateItems;
            DataTable dataRowTable = new DataTable();
            dataRowTable = new DataTable();
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(queryInfo + " where "+ dateItems + "=@"+ dateItems, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;

            sqlDa.SelectCommand.Parameters.AddWithValue(paraDate, date);
            sqlDa.Fill(dataRowTable);
            sqlCon.Close();

            int x = 0;
            dataGridView2.DataSource = dataRowTable;
            //MessageBox.Show(dataRowTable.Rows.Count.ToString());
            //foreach (DataRow dataRow in dataRowTable.Rows)
            //{                
            //    if (dataRow[querySearch[comboType.SelectedIndex]].ToString() != date)
            //    {
            //        dataRowTable.Rows[x].Delete();
            //    }
            //    x++;
            //}
            //dataRowTable.AcceptChanges();
            //switch (comboIndex)
            //{
            //    case 1:                    
            //        foreach (DataRow dataRow in dataRowTable.Rows)
            //        {
            //            if (dataRow[subItems[comboType.SelectedIndex, 1]].ToString() != subAuthTypes.Text)
            //            {
            //                dataRowTable.Rows[x].Delete();
            //            }
            //            x++;
            //        }
            //        break;
            //    case 2:
            //        int y = 0;                    
            //        foreach (Control control in ReportPanel.Controls)
            //        {
            //            if (((ComboBox)control).Name == "newCombo2")
            //                foreach (DataRow dataRow in dataRowTable.Rows)
            //                {
            //                    if (dataRow[subItems[comboType.SelectedIndex, 2]].ToString() != ((ComboBox)control).Text)
            //                    {
            //                        dataRowTable.Rows[y].Delete();
            //                    }
            //                    y++;
            //                }
            //        }
            //        break;
            //}
            //dataGridView2.dataSource57 = dataRowTable;
            //MessageBox.Show(dataGridView2.RowCount.ToString());
            return dataGridView2.RowCount;
        }
        private void AllStatistData()
        {
           SqlConnection sqlCon = new SqlConnection(dataSource57);
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableDeepStatis", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                DataTable table = new DataTable();
                sqlDa.Fill(table);
                sqlCon.Close();
                dataGridView1.DataSource = table;
            }
            catch (Exception ex) { 

            }
        
        }
        private int checklastrange(int id)
        {
            int lastRange = 2;
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select ID,itemID from TableDeepStatis where ID=@id", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@id", id);
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            foreach (DataRow dataRow in table.Rows)
            {
                
                if (dataRow["ID"].ToString() == id.ToString() && dataRow["itemID"].ToString().Contains("-"))
                {
                    string range = dataRow["itemID"].ToString().Split('_')[1];
                    range = range.Split('-')[1];
                    lastRange = Convert.ToInt32(range);
                } 
            }
            return lastRange;
        }
        private void editDate(int id, string dateGre, string table)
        {
            string query = "UPDATE "+ table+ " SET التاريخ_الميلادي=@التاريخ_الميلادي where ID=@id";
            //MessageBox.Show("id " + id.ToString() + " -- " + dateHijri);
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", dateGre);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void addRabge(int id, string item)
        {
            string query = "UPDATE TableDeepStatis SET itemID=@itemID where ID=@id";
            //MessageBox.Show("id " + id.ToString() + " -- " + item);
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@itemID", item);
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        private bool checkRange(string item)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select itemID from TableDeepStatis", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            foreach (DataRow dataRow in table.Rows)
            {
                if (dataRow["itemID"].ToString().Contains(item.Replace(" ","_")))
                {
                    return true;
                }
            }
            return false;
        }
        private int getLastRangeID()
        {
            int itemid = 0;
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select itemID from TableDeepStatis", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            foreach (DataRow dataRow in table.Rows)
            {
                if (dataRow["itemID"].ToString() == "")
                {
                    return itemid;
                }
                itemid++;
            }
            return itemid;
        }

        private string getRange(ComboBox genItems)
        {
            string itemid = "0-0";
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select itemID from TableDeepStatis", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            foreach (DataRow dataRow in table.Rows)
            {
                if (dataRow["itemID"].ToString().Split('_')[0] == genItems.Text)
                {
                    try
                    {
                        itemid = dataRow["itemID"].ToString().Split('_')[1];
                    }
                    catch (Exception ex) { 
                    }
                }
            }
            return itemid;
        }

        private void getGroupCount0(string groupedItem, string text)
        {
            text = "%-01-" + comYear.Text + "' or التاريخ_الميلادي like '%-02-" + comYear.Text + "' or التاريخ_الميلادي like '%-03-" + comYear.Text + "'";

            //MessageBox.Show(text);
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select نوع_التوكيل,count(التاريخ_الميلادي) as dayCount from TableAuth where التاريخ_الميلادي like '%' group by "+ groupedItem, sqlCon);
            //SqlDataAdapter sqlDa = new SqlDataAdapter("select نوع_التوكيل,count(التاريخ_الميلادي) as dayCount from TableAuth group by نوع_التوكيل", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            //sqlDa.SelectCommand.Parameters.AddWithValue("@التاريخ_الميلادي", text);
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            itemCount = new int[table.Rows.Count];
            List0 = new string[table.Rows.Count];
            int i = 0;
            foreach (DataRow dataRow in table.Rows)
            {
                itemCount[i] = Convert.ToInt32(dataRow["dayCount"].ToString());
                List0[i] = dataRow[groupedItem].ToString();
                Console.WriteLine(dataRow[groupedItem].ToString() +" - "+ dataRow["dayCount"].ToString());
                i++;
            }            
        }
        
        private int[] getGroupCount1(string groupedItem, string text, string text1)
        {
            //MessageBox.Show(groupedItem);
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select "+ groupedItem+",count(التاريخ_الميلادي) as dayCount from TableAuth where نوع_التوكيل=@نوع_التوكيل and (" + text + ") group by "+ groupedItem, sqlCon);
            //SqlDataAdapter sqlDa = new SqlDataAdapter("select نوع_التوكيل,count(التاريخ_الميلادي) as dayCount from TableAuth group by نوع_التوكيل", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@نوع_التوكيل", text1);
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            itemCount = new int[table.Rows.Count];
            List1 = new string[table.Rows.Count];
            int i = 0;
            foreach (DataRow dataRow in table.Rows)
            {
                itemCount[i] = Convert.ToInt32(dataRow["dayCount"].ToString());
                List1[i] = dataRow[groupedItem].ToString();
                Console.WriteLine(dataRow[groupedItem].ToString() + " - " + dataRow["dayCount"].ToString());
                i++;
            }
            return itemCount;
        }

        private int[] getDayCount1(string groupedItem0, string text, string text1)
        {
            int i = 0;
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            itemCount = new int[text1.Split('-').Length];
            for (; i < text1.Split('-').Length; i++)
            {
                string date = "%" + text1.Split('-')[i] +"-"+ comYear.Text;
                itemCount[i] = 0;
                //MessageBox.Show(date);
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("select "+ groupedItem0+ ",count(" + groupedItem0 + ") as dayCount from TableAuth where " + groupedItem0 + "=@" + groupedItem0 + " and التاريخ_الميلادي like '"+ date+"' group by " + groupedItem0, sqlCon);
                //SqlDataAdapter sqlDa = new SqlDataAdapter("select count(" + groupedItem0 + ") as dayCount from TableAuth where " + groupedItem0 + "=@" + groupedItem0 + " and ('" + date + "') group by " + groupedItem0, sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                sqlDa.SelectCommand.Parameters.AddWithValue("@" + groupedItem0, text);
                DataTable table = new DataTable();
                sqlDa.Fill(table);
                sqlCon.Close();
                
                List1 = new string[table.Rows.Count];

                foreach (DataRow dataRow in table.Rows)
                {
                    itemCount[i] = Convert.ToInt32(dataRow["dayCount"].ToString());
                    Console.WriteLine("text1.Split('-')["+i.ToString() + "] - " + dataRow["dayCount"].ToString());
                }
            }
            return itemCount;
        }

        private void getDayCount2(string groupedItem0, string text1,string groupedItem1, string text2, string time)
        {
            int i = 0;
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            itemCount = new int[time.Split('-').Length];
            for (; i < time.Split('-').Length; i++)
            {
                string date = "%" + time.Split('-')[i] + "-" + comYear.Text;
                itemCount[i] = 0;
                MessageBox.Show(date);
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("select " + groupedItem0 + ",count(" + groupedItem0 + ") as dayCount from TableAuth where " + groupedItem1 + "=@" + groupedItem1 + " and + " + groupedItem0 + "=@" + groupedItem0 + " and التاريخ_الميلادي like '" + date + "' group by " + groupedItem0, sqlCon);
                //SqlDataAdapter sqlDa = new SqlDataAdapter("select count(" + groupedItem0 + ") as dayCount from TableAuth where " + groupedItem0 + "=@" + groupedItem0 + " and ('" + date + "') group by " + groupedItem0, sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                sqlDa.SelectCommand.Parameters.AddWithValue("@" + groupedItem0, text1);
                sqlDa.SelectCommand.Parameters.AddWithValue("@" + groupedItem1, text2);
                DataTable table = new DataTable();
                sqlDa.Fill(table);
                sqlCon.Close();

                List1 = new string[table.Rows.Count];

                foreach (DataRow dataRow in table.Rows)
                {
                    itemCount[i] = Convert.ToInt32(dataRow["dayCount"].ToString());
                    Console.WriteLine("dayCount - " + dataRow["dayCount"].ToString());
                }
            }
        }


        private int getDataWithDate(string date, string item) 
        {
            item = item.Replace(" ", "_");
            //MessageBox.Show(item);   
            string paraDate = "@" + dateItems;


            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select " + item + ",day from TableDeepStatis where day=@day", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@day", date);
            DataTable dataRowTable = new DataTable();
            sqlDa.Fill(dataRowTable);
            sqlCon.Close();

            int x = 0;
            foreach (DataRow dataRow in dataRowTable.Rows)
            {
                if (dataRow["day"].ToString() == date && !string.IsNullOrEmpty(dataRow[item].ToString()))
                {
                    //MessageBox.Show(dataRow[item].ToString());
                    x = Convert.ToInt32(dataRow[item].ToString());
                }
            }
            return x;
        }
        private void filteringRows(DataTable rows, string items, ComboBox combo)
        {
            int x = 0;
            foreach (DataRow dataRow in rows.Rows)
            {
                if (dataRow[items].ToString() != combo.Text)
                {
                    rows.Rows[x].Delete();
                }
                x++;
            }
        }
        
        private void subAuthTypes_SelectedIndexChanged(object sender, EventArgs e)
        {
            getSub1DataGen(columnList[1], columnList[2]);
            ////chartAreas2 = chartAreas1 = 0;
            ////if (columnList[2] == "") return;
            comboIndex = 1;
            ////subComb1.Items.Clear(); subComb1.Text = "المعاملة";
            ////subComb2.Items.Clear(); subComb2.Text = "المعاملة";
            ////subComb3.Items.Clear(); subComb3.Text = "المعاملة";
            ////subComb4.Items.Clear(); subComb4.Text = "المعاملة";

            //addComboData1(columnList[2],columnList[1],columnList[0]);

            ////if (btnPlotSearch.Text == "فحص")
            ////{
            ////    DeepStatics(subComb0);
            ////    AllStatistData();
            ////    prePareToShow(subComb1, subComb0);
            ////}

            ////else
            ////{
            ////    prePareToShow(subComb1, subComb0);
            ////}
        }

        
        bool DeepStatics(ComboBox comboBox)
        {
            /*
             * drop Table TableDeepStatis
Create Table TableDeepStatis
(
	[ID] int Primary key identity ,
	[day]  nvarchar(150),
	[itemID]  nvarchar(150)
)
          */
            int gridSum = 0;
            int itemSum = 0;
            int row = 0;
            int id = getLastRangeID();
            int last = checklastrange(id);
            int colCount = last + 1;
            int itemOverAll = 0;
            bool foundData = false;
            for (int yearIndex = 1; yearIndex < comYear.Items.Count; yearIndex++)
            {
                int days = 1;
                for (int month = 1; month <= 12; month++)
                {
                    int sum = 0;
                    for (int day = 1; day <= daysOfMonth(month - 1, Convert.ToInt32(comYear.Items[yearIndex].ToString())); day++)
                    {
                        //if (yearIndex == 1 && month <= 9 && day <= 1) { continue; }
                        string CurrentDay = "", Currentmonth = "", CurrentDate = "";
                        if (month < 10) Currentmonth = "0" + month.ToString();
                        else Currentmonth = month.ToString();
                        if (day < 10) CurrentDay = "0" + day.ToString();
                        else CurrentDay = day.ToString();
                        CurrentDate = CurrentDay + "-" + Currentmonth + "-" + comYear.Items[yearIndex].ToString();
                        filterMainRowWithDate(CurrentDate);
                        if (!checkDateExistance(CurrentDate))
                            addBasicData(CurrentDate);
                        gridSum += dataGridView2.RowCount - 1;
                        switch (comboIndex)
                        {
                            case 0:
                                for (int item = 0; item < subComb0.Items.Count; item++)
                                {
                                    int found = 0;
                                    if (!checkColumnName(subComb0.Items[item].ToString()))
                                    {
                                        //Console.WriteLine(" --not founditems-- ");
                                        CreateColumn(subComb0.Items[item].ToString());
                                        //MessageBox.Show(subComb0.Items[item].ToString());
                                        colCount++;
                                    }
                                    //else Console.WriteLine(" --founditems-- ");

                                    for (int grid = 0; grid < dataGridView2.RowCount - 1; grid++)
                                    {
                                        int founditems;
                                        
                                        string nameCell = dataGridView2.Rows[grid].Cells[3].Value.ToString();
                                        string test = nameCell;
                                        //if (nameCell == "قطعة أرض سكنية" || nameCell == "قطعة أرض حيازة" || nameCell == "ساقية")
                                            
                                        //    nameCell = "عقار";
                                        nameCell = MergData(genTypes.SelectedIndex, 0, nameCell);
                                        if (nameCell == subComb0.Items[item].ToString())
                                        {
                                            
                                            founditems = 1;
                                            //if (test == "قطعة أرض سكنية" || test == "قطعة أرض حيازة" || test == "ساقية") MessageBox.Show(test);
                                            if (dataGridView2.Rows[grid].Cells[2].Value.ToString().Contains(symbol))
                                            {
                                                founditems = dataGridView2.Rows[grid].Cells[3].Value.ToString().Split(symbolChar).Length;
                                                

                                                //Console.WriteLine(CurrentDate + " --founditems-- " + founditems.ToString());
                                            }
                                            found += founditems;
                                            if(item == 0) itemSum++;
                                            Console.WriteLine(CurrentDate + "  - item - " + item.ToString() + " --founditems-- " + found.ToString());
                                        }
                                    }
                                    //MessageBox.Show(subComb0.Items[item].ToString() + " - "+found.ToString());
                                    //if (item == 0) MessageBox.Show( itemSum.ToString());
                                    //if (sum == 0) sum = sum + found;
                                    //else if (sum > 0)
                                    //{
                                    //    //Console.WriteLine(" --row-- " + row.ToString() + " --month-- " + month.ToString() + " --day-- " + day.ToString() + " --count-- " + found.ToString());

                                    //}
                                    row++;
                                    Console.WriteLine("CurrentDate " + CurrentDate + "  - item - " + subComb0.Items[item].ToString() + " --found-- " + found.ToString());
                                    addDeepData(CurrentDate, subComb0.Items[item].ToString(), found.ToString());
                                    itemOverAll += found;
                                }
                                break;
                            case 1:
                                for (int item = 0; item < subComb1.Items.Count; item++)
                                {
                                    int found = 0;
                                    string itemSub = subComb1.Items[item].ToString().Replace(" - ", "_");
                                    string colName = subComb0.Text + "_" + itemSub;
                                    //colName = colName.Replace(" ", "_");
                                    if (!checkColumnName(colName))
                                    {

                                        CreateColumn(colName);
                                        colCount++;
                                    }
                                    bool rightCol = false;

                                    for (int grid = 0; grid < dataGridView2.RowCount - 1; grid++)
                                    {

                                        string gridcelSplit = dataGridView2.Rows[grid].Cells[2].Value.ToString();
                                        string gridcel1 = dataGridView2.Rows[grid].Cells[3].Value.ToString();
                                        gridcel1 = MergData(genTypes.SelectedIndex, 0, gridcel1);
                                        string gridcel2 = dataGridView2.Rows[grid].Cells[4].Value.ToString();
                                        gridcel2 = MergData(genTypes.SelectedIndex, 1, gridcel2);                                        
                                        //MessageBox.Show(gridcel1 + " - " + gridcel2 + " - " + gridcelSplit);
                                        if (gridcel1 == subComb0.Text)
                                        {
                                            // MessageBox.Show(gridcel3 + " - " + subComb1.Items[item].ToString());
                                            if (gridcel2 == subComb1.Items[item].ToString())
                                            {
                                                //MessageBox.Show(CurrentDate + " - "+gridcel1 + " - " + gridcel);
                                                rightCol = true;
                                                //MessageBox.Show(gridcel);
                                                if (!checkColumnName(colName))
                                                {
                                                    //MessageBox.Show(CurrentDate);
                                                    CreateColumn(colName);
                                                    colCount++;
                                                }
                                                int founditems = 0;
                                                if (gridcelSplit.Contains(symbol))
                                                {
                                                    founditems = gridcelSplit.Split(symbolChar).Length;
                                                    
                                                }
                                                else founditems = 1;

                                                found += founditems;
                                                Console.WriteLine(comboIndex.ToString() + " - " + CurrentDate + " --founditems-- " + found.ToString());
                                            }
                                        }
                                    }
                                    if (sum == 0) sum = sum + found;
                                    if (sum > 0 && rightCol)
                                    {
                                        addDeepData(CurrentDate, colName, found.ToString());
                                        itemOverAll += found;
                                    }
                                }
                                break;
                            case 2:
                                //MessageBox.Show("2");
                                for (int item = 0; item < subComb2.Items.Count; item++)
                                {
                                    int found = 0;
                                    string itemSub2 = subComb2.Items[item].ToString().Replace(" - ", "_");
                                    string itemSub1 = subComb1.Text.Replace(" - ", "_");
                                    string colName = subComb0.Text + "_" + itemSub1 + "_" + itemSub2;
                                    if (!checkColumnName(colName))
                                    {

                                        CreateColumn(colName);
                                        colCount++;
                                    }
                                    bool rightCol = false;

                                    for (int grid = 0; grid < dataGridView2.RowCount - 1; grid++)
                                    {
                                        string gridcelSplit = dataGridView2.Rows[grid].Cells[2].Value.ToString();

                                        string gridcel1 = dataGridView2.Rows[grid].Cells[3].Value.ToString();
                                        gridcel1 = MergData(genTypes.SelectedIndex, 0, gridcel1);
                                        string gridcel2 = dataGridView2.Rows[grid].Cells[4].Value.ToString();
                                        gridcel2 = MergData(genTypes.SelectedIndex, 1, gridcel2);
                                        string gridcel3 = dataGridView2.Rows[grid].Cells[5].Value.ToString();
                                        gridcel3 = MergData(genTypes.SelectedIndex, 2, gridcel3);
                                        if (gridcel3 != "")
                                            if (gridcel1 == subComb0.Text && gridcel2 == subComb1.Text)
                                        {
                                            // MessageBox.Show(gridcel3 + " - " + subComb1.Items[item].ToString());
                                            if (gridcel3 == subComb2.Items[item].ToString())
                                            {
                                                //MessageBox.Show(gridcel + " - " + gridcel1 + " - " + gridcel4);
                                                rightCol = true;
                                                if (!checkColumnName(colName))
                                                {
                                                    colCount++;
                                                    CreateColumn(colName);
                                                }
                                                int founditems = 0;
                                                if (gridcelSplit.Contains(symbol))
                                                {
                                                    founditems = gridcelSplit.Split(symbolChar).Length;
                                                    //Console.WriteLine(comboIndex.ToString() + " - " + CurrentDate + " --founditems-- " + founditems.ToString());
                                                }
                                                else founditems = 1;
                                                found += founditems;
                                                    Console.WriteLine(item.ToString() + " - " + CurrentDate + " --founditems-- " + found.ToString());
                                                }
                                        }
                                    }

                                    if (sum == 0) sum = sum + found;
                                    if (sum > 0 && rightCol)
                                    {
                                        addDeepData(CurrentDate, colName, found.ToString());
                                        itemOverAll += found;
                                    }
                                }
                                break;
                            case 3:
                                for (int item = 0; item < subComb3.Items.Count; item++)
                                {
                                    int found = 0;
                                    string itemSub3 = subComb3.Items[item].ToString().Replace(" - ", "_");
                                    string itemSub2 = subComb2.Text.Replace(" - ", "_");
                                    string itemSub1 = subComb1.Text.Replace(" - ", "_");
                                    string colName = subComb0.Text + "_" + itemSub1 + "_" + itemSub2+ "_" + itemSub3;
                                    if (!checkColumnName(colName))
                                    {

                                        CreateColumn(colName);
                                        colCount++;
                                    }
                                    bool rightCol = false;

                                    for (int grid = 0; grid < dataGridView2.RowCount - 1; grid++)
                                    {
                                        string gridcelSplit = dataGridView2.Rows[grid].Cells[2].Value.ToString();

                                        string gridcel1 = dataGridView2.Rows[grid].Cells[3].Value.ToString();
                                        gridcel1 = MergData(genTypes.SelectedIndex, 0, gridcel1);
                                        string gridcel2 = dataGridView2.Rows[grid].Cells[4].Value.ToString();
                                        gridcel2 = MergData(genTypes.SelectedIndex, 1, gridcel2);
                                        string gridcel3 = dataGridView2.Rows[grid].Cells[5].Value.ToString();
                                        gridcel3 = MergData(genTypes.SelectedIndex, 2, gridcel3);
                                        string gridcel4 = dataGridView2.Rows[grid].Cells[6].Value.ToString();
                                        gridcel4 = MergData(genTypes.SelectedIndex, 3, gridcel4);
                                        
                                        //if (gridcel4 != "")
                                            if (gridcel4 == subComb3.Items[item].ToString() && gridcel1 == subComb0.Text && gridcel2 == subComb1.Text && gridcel3 == subComb2.Text)
                                            {
                                            //MessageBox.Show(gridcel1 + " - " + gridcel2 + " - " + gridcel3 + " - " + gridcel4);
                                            //if ()
                                                {
                                                    //MessageBox.Show(gridcel + " - " + gridcel1 + " - " + gridcel4);
                                                    rightCol = true;
                                                    if (!checkColumnName(colName))
                                                    {
                                                        colCount++;
                                                        CreateColumn(colName);
                                                    }
                                                    int founditems = 0;
                                                    if (gridcelSplit.Contains(symbol))
                                                    {
                                                        founditems = gridcelSplit.Split(symbolChar).Length;
                                                        //Console.WriteLine(comboIndex.ToString() + " - " + CurrentDate + " --founditems-- " + founditems.ToString());
                                                    }
                                                    else founditems = 1;
                                                    found += founditems;
                                                    Console.WriteLine(item.ToString() + " - " + CurrentDate + " --founditems-- " + found.ToString());
                                                }
                                            }
                                    }

                                    if (sum == 0) sum = sum + found;
                                    if (sum > 0 && rightCol)
                                    {
                                        addDeepData(CurrentDate, colName, found.ToString());
                                        itemOverAll += found;
                                    }
                                }
                                break;
                        }
                        days++;
                        if (CurrentDate == DateTime.Now.ToString("dd-MM-yyyy"))
                        {
                            //btnSumStatis.Text = itemOverAll.ToString() + " معاملة";
                            if (!checkRange(comboBox.Text))
                            {
                                Console.WriteLine(" --ColumnCount-- " + ColumnCount().ToString());
                                addRabge(id + 1, comboBox.Text + "_" + (last + 1).ToString() + "-" + ColumnCount().ToString());
                            }
                            //labebSum.Text = labebSum.Text + " - "+ gridSum.ToString()+ " - "+ itemSum.ToString() ;
                            labebSum.Text =  itemSum.ToString() ;
                            return foundData;
                        }
                    }

                }
            }
            //btnSumStatis.Text = itemOverAll.ToString() + " معاملة";
            if (!checkRange(comboBox.Text))
            {
                Console.WriteLine(" --ColumnCount-- " + ColumnCount().ToString());
                addRabge(id + 1, comboBox.Text + "_" + (last + 1).ToString() + "-" + ColumnCount().ToString());

            }
            //labebSum.Text = labebSum.Text + " - " + gridSum.ToString() + " - " + itemSum.ToString();
            return foundData;
        }
        private void CreateColumn(string Columnname)
        {
            
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("alter table TableDeepStatis add " + Columnname.Replace(" ", "_") + " nvarchar(10)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void fillColName(ComboBox gen)
        {
            
            //dataGridView2.dataSource57 = dataRowTable;
            string[] range = getRange(gen).Split('-');
            int start = Convert.ToInt32(range[0]);
            int end = Convert.ToInt32(range[1]);

            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableDeepStatis", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            int colIndex = 0;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                string col = dataRow["COLUMN_NAME"].ToString();
                
                if (!string.IsNullOrEmpty(col) && colIndex >= start && colIndex <= end)
                {
                    gen.Items.Add(dataRow["COLUMN_NAME"].ToString().Replace("_", " "));                        
                }
                colIndex++;
            }            
        }

        private void AllColumn()
        {

            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableDeepStatis", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            int colIndex = 0;
            foreach (DataRow dataRow in dtbl.Rows)
            {
                string col = dataRow["COLUMN_NAME"].ToString();

                if (!string.IsNullOrEmpty(col) && colIndex>2)
                {
                    //comboBox1.Items.Add(dataRow["COLUMN_NAME"].ToString().Replace("_", " "));
                    //comboBox2.Items.Add(dataRow["COLUMN_NAME"].ToString().Replace("_", " "));
                }
                colIndex++;
            }
        }
        private int ColumnCount()
        {
            int count = 0;
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableDeepStatis", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                count++;
            }
            return count-1;
        }
        private bool checkColumnName(string colNo)
        {
            //MessageBox.Show(dataSource);
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SP_COLUMNS TableDeepStatis", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (!string.IsNullOrEmpty(dataRow["COLUMN_NAME"].ToString()))
                {
                    //MessageBox.Show(dataRow["COLUMN_NAME"].ToString());
                    if (dataRow["COLUMN_NAME"].ToString() == colNo.Replace(" ", "_"))
                    {
                         return true;
                    }
                }
            }
            //MessageBox.Show(colNo + "not found");
            return false;
        }
        private void addDeepData(string date, string items, string data)
        {
            items = items.Replace(" ", "_"); 
            string paraItems = "@" + items;
            string query = "UPDATE TableDeepStatis SET " + items + "=" + paraItems + " where day=@date";
            //MessageBox.Show(query);
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@date", date);
            sqlCmd.Parameters.AddWithValue(paraItems, data);            
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        private void addBasicData(string day)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableDeepStatis (day) values (@day)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;            
            sqlCmd.Parameters.AddWithValue("@day", day);            
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        

        private void subComb1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // chartAreas2 = chartAreas1 = 0; 
            //if (columnList[3] == "") return; 
            comboIndex = 2;
            //subComb2.Items.Clear(); subComb2.Text = "المعاملة";
            //subComb3.Items.Clear(); subComb3.Text = "المعاملة";
            //subComb4.Items.Clear(); subComb4.Text = "المعاملة";

            //addComboData2(dataRowTable, columnList[3]);
            ////MessageBox.Show(gridItems[genTypes.SelectedIndex, 2].ToString());
            //if (btnPlotSearch.Text == "فحص")
            //{
            //    //MessageBox.Show(gridItems[genTypes.SelectedIndex, 2].ToString());
            //    DeepStatics(subComb1);
            //    AllStatistData();
            //    prePareToShow(subComb2, subComb1);
            //}
            //else
            //{
            //    //fillColName(subComb1);
            //    prePareToShow(subComb2, subComb1);
            //}

        }

        
        private void splitCol(int id, string[] text)
        {
            
            string[] str = new string[10] { "غير محددة", "غير محددة", "غير محددة", "غير محددة", "غير محددة", "غير محددة", "غير محددة", "غير محددة", "غير محددة", "غير محددة" };
            for (int iS = 0; (iS < text.Length && iS < 10); iS++)
            {
                if (text[iS] != "")
                    str[iS] = text[iS];
            }

            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET text1=@text1,text2=@text2,text3=@text3,text4=@text4,text5=@text5,check1=@check1,txtD1=@txtD1,combo1=@combo1,combo2=@combo2,addName1=@addName1 WHERE ID = @id", sqlCon);



                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@id", id);
                    sqlCmd.Parameters.AddWithValue("@text1", str[0]);
                    sqlCmd.Parameters.AddWithValue("@text2", str[1]);
                    sqlCmd.Parameters.AddWithValue("@text3", str[2]);
                    sqlCmd.Parameters.AddWithValue("@text4", str[3]);
                    sqlCmd.Parameters.AddWithValue("@text5", str[4]);
                    sqlCmd.Parameters.AddWithValue("@check1", str[5]);
                    sqlCmd.Parameters.AddWithValue("@txtD1", str[6]);
                    sqlCmd.Parameters.AddWithValue("@combo1", str[7]);
                    sqlCmd.Parameters.AddWithValue("@addName1", str[8]);
                    sqlCmd.Parameters.AddWithValue("@combo2", str[9]);
                    sqlCmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    return;
                }

        }
        private void subComb2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //chartAreas2 = chartAreas1 = 0; 
            //if (columnList[4] == "") return;
            ////subBtn2.Text = "محدد"; 
            ////comboIndex = 2;
            //comboIndex = 3;
            //subComb3.Items.Clear(); subComb3.Text = "المعاملة";
            //subComb4.Items.Clear(); subComb4.Text = "المعاملة";

            //addComboData3(dataRowTable, columnList[4]);
            ////MessageBox.Show(subComb3.Items.Count.ToString());
            //if (btnPlotSearch.Text == "فحص")
            //{
            //    //MessageBox.Show(gridItems[genTypes.SelectedIndex, 2].ToString());
            //    DeepStatics(subComb2);
            //    AllStatistData();
            //    prePareToShow(subComb3, subComb2);
            //}
            //else
            //{
            //    //fillColName(subComb1);
            //    prePareToShow(subComb3, subComb2);
            //}
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (subBtn0.Text == "الكل")
            {
                subBtn0.Text = "محدد";
            }
            else
            {
                subBtn0.Text = "الكل";
            }
        }

        private void subBtn1_Click(object sender, EventArgs e)
        {
            if (subBtn1.Text == "الكل")
            {
                subBtn1.Text = "محدد";
            }
            else
            {
                subBtn1.Text = "الكل";
            }
        }

        private void subBtn2_Click(object sender, EventArgs e)
        {
            if (subBtn2.Text == "الكل")
            {
                subBtn2.Text = "محدد";
            }
            else
            {
                subBtn2.Text = "الكل";
            }
        }

        private void subBtn3_Click(object sender, EventArgs e)
        {
            if (subBtn3.Text == "الكل")
            {
                subBtn3.Text = "محدد";
            }
            else
            {
                subBtn3.Text = "الكل";
            }
        }

        private void subBtn4_Click(object sender, EventArgs e)
        {
            if (subBtn4.Text == "الكل")
            {
                subBtn4.Text = "محدد";
            }
            else
            {
                subBtn4.Text = "الكل";
            }
        }

        private void chart1_Click(object sender, EventArgs e)
        {
            holdData1 = true;
            holdData1count++;
            holdData2 = true;
            holdData2count++;
        }

        private void chart2_Click(object sender, EventArgs e)
        {
            holdData1 = true;
            holdData1count++;
            holdData2 = true;
            holdData2count++;
        }

        

        private void fillGrid3(string text, string querytext, string queryName)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(text, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue(queryName, querytext);
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            dataGridView3.DataSource = table;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (chart2.Series[comSeriers2.Text].Enabled)
                chart2.Series[comSeriers2.Text].Enabled = false;
            else chart2.Series[comSeriers2.Text].Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            chart2.Series.RemoveAt(comSeriers2.SelectedIndex);
            comSeriers2.Items.RemoveAt(comSeriers2.SelectedIndex);
            var partAll = MessageBox.Show("", "حذف الجميع؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (partAll == DialogResult.Yes)
            {
                while (chart2.Series.Count > 0)
                {
                    chart2.Series.RemoveAt(0);
                    if (comSeriers2.Items.Count > 0) comSeriers2.Items.RemoveAt(0);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int x = 0; x < comSeriers2.Items.Count; x++)
            {
                if (btnHideAll2.Text == "إخفاء الجميع") chart2.Series[comSeriers2.Items[x].ToString()].Enabled = false;
                else chart2.Series[comSeriers2.Items[x].ToString()].Enabled = true;
            }
            if (btnHideAll2.Text == "إخفاء الجميع")
                btnHideAll2.Text = "إظهار الجميع";
            else btnHideAll2.Text = "إخفاء الجميع";
        }

        private void addTrueWrongData(string wrongData,string trueData)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableStatisInfo (wrongeItems,trueItems) values (@wrongeItems,@trueItems)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@trueItems", trueData);
            sqlCmd.Parameters.AddWithValue("@wrongeItems", wrongData);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            
        }

        private void subComb3_SelectedIndexChanged(object sender, EventArgs e)
        {
            chartAreas2 = chartAreas1 = 0; 
            comboIndex = 4;
            if (columnList[5] == "") return; 
            subComb4.Items.Clear(); subComb4.Text = "المعاملة";
            addComboData4(dataRowTable, columnList[5]);
            
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            //if (wrongName.Text == "" ) return;
            addTrueWrongData(wrongName.Text, trueName.Text);
            if(corrections.SelectedIndex != 3) reFillRows();


            switch (corrections.SelectedIndex) {
                case 0:
                    int count = fillSatatInfoGrid("1", "1");
                    if (comboIndex == 0) return;
                    //for (int x = 0; x < count; x++)
                    //{
                        //correctData(wrongeItems[x], trueItems[x], columnList[comboIndex], columnList[0]);
                        correctDataSuit("المهن_المعدلة", columnList[0],false);
                        //deleteDuplicate();
                    //}
                    
                    break;
                case 1:
                    //MessageBox.Show(comboIndex.ToString() + " - " +columnList[comboIndex] + " -- " + columnList[comboIndex - 1]);
                    swapData(wrongName.Text, trueName.Text, columnList[comboIndex], columnList[comboIndex - 1]);
                    break;
                case 2:
                    //MessageBox.Show(comboIndex.ToString() + " - " +columnList[comboIndex] + " -- " + columnList[comboIndex - 1]);
                    switch (comboIndex)
                    {
                        case 1:
                            returnID(subComb2.Text, subComb3.Text, columnList[comboIndex], columnList[comboIndex - 1]);
                            break;
                        case 2:
                            returnID(subComb2.Text, subComb3.Text, columnList[comboIndex], columnList[comboIndex - 1]);
                            break;
                        case 3:
                            returnID(subComb2.Text, subComb3.Text, columnList[comboIndex], columnList[comboIndex - 1]);
                            break;
                        case 4:
                            returnID(subComb2.Text, subComb3.Text, columnList[comboIndex], columnList[comboIndex - 1]);
                            break;
                    }
                            break;
                case 3:
                    //MessageBox.Show(comboIndex.ToString() + " - " +columnList[comboIndex] + " -- " + columnList[comboIndex - 1]);
                    
                    break;
                case 4:
                    //MessageBox.Show(comboIndex.ToString() + " - " +columnList[comboIndex] + " -- " + columnList[comboIndex - 1]);
                    splitData(wrongName.Text, trueName.Text, "القضية");
                    break;
                case 5:
                    
                    break;

            }
            wrongName.Text = trueName.Text = "";
        }

        private void subComb4_SelectedIndexChanged(object sender, EventArgs e)
        {
            chartAreas2 = chartAreas1 = 0;
        }

        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView5.CurrentRow.Index != -1) {
                gridID = Convert.ToInt32(dataGridView5.CurrentRow.Cells[0].Value.ToString());
                wrongName.Text = dataGridView5.CurrentRow.Cells[1].Value.ToString();
                if (corrections.SelectedIndex == 3)
                {
                    deleteRowsData(gridID, "TableTrendings");
                    addRemovedItem(wrongName.Text);
                    
                    gridID = 0;

                    btnItemsCount.Text = dataGridView5.RowCount.ToString();
                }
            }
        }
        private int lastIdRemoved()
        {
            int id = 1;
            SqlConnection sqlCon = new SqlConnection(dataSource);
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select ID,removedItems from TableTrendings", sqlCon);
            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl1 = new DataTable();
            sqlDa1.Fill(dtbl1);
            foreach (DataRow dataRow in dtbl1.Rows)
            {

                if (dataRow["removedItems"].ToString() == "")
                {
                    return Convert.ToInt32(dataRow["ID"].ToString()); ;
                }
            }
            return id;
        }

        private void addRemovedItem(string word)
        {
            int id = lastIdRemoved();
               SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("update TableTrendings set removedItems=@removedItems where ID=@id", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", id);
            sqlCmd.Parameters.AddWithValue("@removedItems", word);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void btnItemsCount_Click(object sender, EventArgs e)
        {
            dataGridView5.DataSource = TrendingUpdate();
            btnItemsCount.Text = dataGridView5.RowCount.ToString();
        }

        private void btnSumStatis_Click(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource56);
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select ID,القضية from TableSuitCase", sqlCon);

            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl1 = new DataTable();
            sqlDa1.Fill(dtbl1);
            bool once = true;
            int count = 0; textBox4.Text = "1";
            foreach (DataRow dataRow in dtbl1.Rows)
            {
                suitID = Convert.ToInt32(dataRow["ID"].ToString());
                if (dataRow["القضية"].ToString() == wrongName.Text)
                {
                    if (once)
                    {
                        UpdateCasetype(suitID, trueName.Text);

                        if (txtGetID.Text != "")
                            UpdateSecCase(suitID, txtGetID.Text);

                        if (textBox1.Text != "")
                            UpdateSecCase(suitID, textBox1.Text);
                        if (textBox3.Text != "")
                            UpdateSecCase(suitID, textBox3.Text);
                        if (textBox2.Text != "")
                            UpdateSecCase(suitID, textBox2.Text);
                        once = false;
                    }
                    else if(!once && suitID != Convert.ToInt32(txtIDCase.Text)) {
                        UpdateCasetype(suitID, trueName.Text);

                        if (txtGetID.Text != "")
                            UpdateSecCase(suitID, txtGetID.Text);

                        if (textBox1.Text != "")
                            UpdateSecCase(suitID, textBox1.Text);
                        if (textBox3.Text != "")
                            UpdateSecCase(suitID, textBox3.Text);
                        if (textBox2.Text != "")
                            UpdateSecCase(suitID, textBox2.Text);
                    }
                    count++;
                    //MessageBox.Show(suitID.ToString());
                }
            }

            textBox4.Text = count.ToString();   
            txtGetID.Text = textBox1.Text = textBox2.Text = textBox3.Text = "";
        }

        private void button4_Click_2(object sender, EventArgs e)
        {
            correctDataSuit("القضية", "TableSuitCase",false);
        }

        private void trueName_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (e.KeyChar == (char)13)
            //{
            //    addTrueWrongData(wrongName.Text, trueName.Text);
            //    if (corrections.SelectedIndex != 3) reFillRows();
            //    fillSatatInfoGrid("1");
            //    correctDataSuit("تصنيف_عام", columnList[0],false);
            //    wrongName.Text = "";
            //}
        }

        private void trueName_TextChanged(object sender, EventArgs e)
        {

        }

        private void subComb0_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                wrongName.Text = subComb0.Text;
            }
        }

        private void wrongName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                addTrueWrongData(wrongName.Text, trueName.Text);
                if (corrections.SelectedIndex != 3) reFillRows();
                //fillSatatInfoGrid("1");
                correctDataSuit("تصنيف_عام", columnList[0], false);
                wrongName.Text = "";
            }
        }

        private void combGrad1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private bool checkDateExistance(string day)
        {
            bool found = false;
            SqlConnection sqlCon = new SqlConnection(dataSource57);
            SqlDataAdapter sqlDa1 = new SqlDataAdapter("select day from TableDeepStatis", sqlCon);
            sqlDa1.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl1 = new DataTable();
            sqlDa1.Fill(dtbl1);
            foreach (DataRow dataRow in dtbl1.Rows)
            {

                if (dataRow["day"].ToString() == day)
                {
                    found = true;
                }

            }
            return found;
        }
    }
}
