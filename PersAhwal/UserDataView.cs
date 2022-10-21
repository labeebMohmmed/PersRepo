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

namespace PersAhwal
{
    public partial class UserDataView : UserControl
    {
        bool colored = false;
        string AuthNoPart1 = "ق س ج/160/12/";
        string AuthNoPart2 = "";
        public string rowCount = "";
        public bool NewAuth = false;
        int intID = -1;
        string archFile = @"D:\ArchiveFiles\";
        string FilespathIn, FilespathOut;
        bool timerColor = true;
        bool timer = true;
        bool steadyGrid = false;
        public Delegate DataMovePage;
        string[] dataSum = new string[50];
        string dataSource = "Data Source=192.168.100.100,49170;Network Library=DBMSSOCN;Initial Catalog=AhwalDataBase;User ID=ConsJeddahAdmin;Password=DataBC0nsJ49170";
        public Form11 ParentData { get; set; }
        bool fileloaded = false;
        public string[] DatasumValue
        {
            get { return dataSum; }
            set { dataSum = value; }
        }

        public Button deleteRowValue
        {
            get { return deleteRow; }
            set { deleteRow = value; }
        }

        public bool fileloadedValue
        {
            get { return fileloaded; }
            set { fileloaded = value; }
        }
        public CheckBox ArchivedStValue
        {
            get { return ArchivedSt; }
            set { ArchivedSt = value; }
        }

        public bool timer1Value
        {
            get { return timer; }
            set { timer = value; }
        }

        

        public Label labelArchValue
        {
            get { return labelArch; }
            set { labelArch = value; }
        }
        public TextBox ListSearchValue
        {
            get { return ListSearch; }
            set { ListSearch = value; }
        }
        
        public UserDataView()
        {
            InitializeComponent();
            FilespathIn = @"\\192.168.100.56\Users\Public\Documents\ModelFiles\";
            FilespathOut = @"D:\ArchiveFiles\";
            if (!Directory.Exists(@"D:\"))
            {
                string appFileName = Environment.GetCommandLineArgs()[0];
                string directory = Path.GetDirectoryName(appFileName);
                directory = directory + @"\";
                archFile = directory + @"ArchiveFile\";
            }

                
            
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            dataGridView1.BackgroundColor = Color.White;
            string file = archFile + @"\dataSource.txt";
            
            dataSource = File.ReadAllText(file);
            //MessageBox.Show(dataSource);
            FillDataGridView(dataSource);
            //ColorFulGrid9();
            datagrid1();
            //AllProType();

            //proTypeCount = AllProType();

        }

        

       
       


        private void button2_Click(object sender, EventArgs e)
        {
            
        }
        
        
        
        private void button1_Click(object sender, EventArgs e)
        {
            ParentData.txtAuthNoValue.Text = AuthNoPart1 + AuthNoPart2;

            DataMovePage.DynamicInvoke(2);
            NewAuth = true;
        }

        public void FillAuthNo(String DataSource,string  text)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("AuthIDViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@رقم_التوكيل", text);
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);dataGridView1.DataSource = dtbl;
            rowCount = dtbl.Rows.Count.ToString();
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            AuthNoPart2 = dataGridView1.Rows.Count.ToString();
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 150;
            sqlCon.Close();
        }

        public void FillDataGridView(string DataSource)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("AuthViewOrSearch", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            sqlDa.SelectCommand.Parameters.AddWithValue("@مقدم_الطلب", ListSearch.Text.Trim());
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            rowCount = dtbl.Rows.Count.ToString();
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            //dataGridView1.Columns[0].Visible = false ;
            dataGridView1.Columns[1].Width = 180;
            dataGridView1.Columns[3].Width = 50;
            dataGridView1.Columns[8].Width = 50;
            dataGridView1.Columns[9].Width = 170;
            dataGridView1.Columns[7].Width = dataGridView1.Columns[2].Width = 200;
            AuthNoPart2 = dataGridView1.Rows.Count.ToString();
            sqlCon.Close();
            int bre = 0;

            //for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            //{
            //    if (dataGridView1.Rows[i].Cells[2].Value.ToString()!= "")
            //    {
            //        name(Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value.ToString()), dataGridView1.Rows[i].Cells[2].Value.ToString());
            //       Console.WriteLine("id= " + dataGridView1.Rows[i].Cells[0].Value.ToString() );

            //    }
            //    bre++;
            //    //if (bre == 10) return;
            //}

        }

        private void name(int id, string text)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET مقدمي_الطلب=@مقدمي_الطلب WHERE ID = @id", sqlCon);



                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@id", id);
                    sqlCmd.Parameters.AddWithValue("@مقدمي_الطلب", text);
                    sqlCmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    return;
                }

        }
        private void splitCol(int id, string[] text)
        {
            //Console.WriteLine("index= " + i.ToString() + " -- id = " + dataGridView1.Rows[i].Cells[0].Value.ToString() + " -- value= " + dataGridView1.Rows[i].Cells[11].Value.ToString());

            string[] str = new string[10] { "غير محددة" , "غير محددة" , "غير محددة" , "غير محددة" , "غير محددة" , "غير محددة" , "غير محددة" , "غير محددة" , "غير محددة" , "غير محددة" };
            for (int iS = 0; (iS < text.Length && iS<10); iS++)
            {
                if(text[iS] != "")
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
                catch (Exception ex) {
                    return;
                }
            
        }
    
        private void mandoubVisibilty(bool AppType)
        {
            if (AppType)
            {

                ParentData.MandoubNameValue.Visible = false;
                ParentData.MandoubLabelValue.Visible = false;
            }
            else
            {

                ParentData.MandoubNameValue.Visible = true;
                ParentData.MandoubLabelValue.Visible = true;
            }
        }

        public void datagrid1() {

            int currentID = 0;
            if (dataGridView1.Rows.Count == 2)
            {
                ParentData.NewDataValue = true;
                dataSum[9] = dataGridView1.Rows[currentID].Cells[0].Value.ToString();
                
                if (ParentData.JobpositionValue.Contains("قنصل"))
                {
                    deleteRow.Visible = true;
                    timerColor = true;
                    
                }
                foreach (Control control in ParentData.MainPanel.Controls)
                {
                    if (control is TextBox && ((TextBox)control).Name == "txtAuthNo")
                    {
                        
                        ((TextBox)control).Text = dataGridView1.Rows[currentID].Cells[1].Value.ToString();
                    }
                    if (control is TextBox && ((TextBox)control).Name == "txtGreDate")
                    {
                        ((TextBox)control).Text = dataGridView1.Rows[currentID].Cells[14].Value.ToString();
                    }
                    if (control is TextBox && ((TextBox)control).Name == "txtHijDate")
                    {
                        ((TextBox)control).Text = dataGridView1.Rows[currentID].Cells[15].Value.ToString();
                    }
                    if (control is TextBox && ((TextBox)control).Name == "txtAttendVC")
                    {
                        ((TextBox)control).Text = dataGridView1.Rows[currentID].Cells[16].Value.ToString();
                    }
                    if (control is ComboBox && ((ComboBox)control).Name == "ComboAuthDestin")
                    {
                        ((ComboBox)control).Text = dataGridView1.Rows[currentID].Cells[21].Value.ToString();
                    }




                    if (control is CheckBox && ((CheckBox)control).Name == "checkedViewed")
                    {
                        if (dataGridView1.Rows[currentID].Cells[17].Value.ToString() == "غير معالج")
                        {
                            ((CheckBox)control).CheckState = CheckState.Unchecked;
                        }
                        else ((CheckBox)control).CheckState = CheckState.Checked;
                    }

                }
                foreach (Control control in ParentData.subMainPanel.Controls)
                {
                    if (control is CheckBox && ((CheckBox)control).Name == "AppType")
                    {
                        ((CheckBox)control).Text = dataGridView1.Rows[currentID].Cells[18].Value.ToString();
                        if (((CheckBox)control).Text == "حضور مباشرة إلى القنصلية") ((CheckBox)control).CheckState = CheckState.Checked;
                        else ((CheckBox)control).CheckState = CheckState.Unchecked;
                        if (((CheckBox)control).CheckState == CheckState.Unchecked)
                        {
                            if (((CheckBox)control).CheckState == CheckState.Checked)
                            {
                                ((CheckBox)control).Text = "حضور مباشرة إلى القنصلية";
                                ParentData.MandoubNameValue.Visible = false;
                                ParentData.MandoubLabelValue.Visible = false;
                            }
                            else
                            {
                                ((CheckBox)control).Text = "عن طريق أحد مندوبي القنصلية";
                                ParentData.MandoubNameValue.Visible = true;
                                ParentData.MandoubLabelValue.Visible = true;
                            }
                            ParentData.MandoubNameValue.Text = dataGridView1.Rows[currentID].Cells[20].Value.ToString();
                        }
                        if (control is TextBox && ((TextBox)control).Name == "txtComment")
                        {
                            ((TextBox)control).Text = dataGridView1.Rows[currentID].Cells[29].Value.ToString();
                        }

                    }
                }
                if (dataGridView1.Rows[currentID].Cells[2].Value.ToString().Contains("_"))
                {
                    ParentData.strAppnameList = dataGridView1.Rows[currentID].Cells[2].Value.ToString().Split('_');
                    ParentData.intAppCounts = dataGridView1.Rows[currentID].Cells[2].Value.ToString().Split('_').Length;
                }
                else
                {
                    ParentData.strAppnameList[0] = dataGridView1.Rows[currentID].Cells[2].Value.ToString();
                    ParentData.intAppCounts = 1;
                }

                if (dataGridView1.Rows[currentID].Cells[3].Value.ToString().Contains("_"))
                {
                    ParentData.strAppMaleFemaleList = dataGridView1.Rows[currentID].Cells[3].Value.ToString().Split('_');
                }
                else
                {
                    ParentData.strAppMaleFemaleList[0] = dataGridView1.Rows[currentID].Cells[3].Value.ToString();
                }

                if (dataGridView1.Rows[currentID].Cells[4].Value.ToString().Contains("_"))
                {
                    ParentData.strAppDocTypelist = dataGridView1.Rows[currentID].Cells[4].Value.ToString().Split('_');
                }
                else
                {
                    ParentData.strAppDocTypelist[0] = dataGridView1.Rows[currentID].Cells[4].Value.ToString();
                }

                if (dataGridView1.Rows[currentID].Cells[5].Value.ToString().Contains("_"))
                {
                    ParentData.strAppDocNolist = dataGridView1.Rows[currentID].Cells[5].Value.ToString().Split('_');

                }
                else
                {

                    ParentData.strAppDocNolist[0] = dataGridView1.Rows[currentID].Cells[5].Value.ToString();

                }

                if (dataGridView1.Rows[currentID].Cells[6].Value.ToString().Contains("_"))
                {
                    ParentData.strAppissueList = dataGridView1.Rows[currentID].Cells[6].Value.ToString().Split('_');
                }
                else
                {
                    ParentData.strAppissueList[0] = dataGridView1.Rows[currentID].Cells[6].Value.ToString();
                }

                if (dataGridView1.Rows[currentID].Cells[7].Value.ToString().Contains("_"))
                {
                    ParentData.strAuthNames = dataGridView1.Rows[currentID].Cells[7].Value.ToString().Split('_');
                    ParentData.intAuthCount = dataGridView1.Rows[currentID].Cells[7].Value.ToString().Split('_').Length;
                }
                else
                {
                    ParentData.strAuthNames[0] = dataGridView1.Rows[currentID].Cells[7].Value.ToString();
                }

                if (dataGridView1.Rows[currentID].Cells[8].Value.ToString().Contains("_"))
                {
                    ParentData.strAuthMaleFemale = dataGridView1.Rows[currentID].Cells[8].Value.ToString().Split('_');
                }
                else
                {
                    ParentData.strAuthMaleFemale[0] = dataGridView1.Rows[currentID].Cells[8].Value.ToString();
                }
                dataSum[0] = dataGridView1.Rows[currentID].Cells[9].Value.ToString();//نوع_التوكيل
                dataSum[1] = dataGridView1.Rows[currentID].Cells[10].Value.ToString();//رقم_العمود
                dataSum[2] = dataGridView1.Rows[currentID].Cells[11].Value.ToString();//موضوع التوكيل
                dataSum[3] = dataGridView1.Rows[currentID].Cells[12].Value.ToString();//حقوق_التوكيل
                dataSum[4] = dataGridView1.Rows[currentID].Cells[13].Value.ToString();//التاريخ_الميلادي

                dataSum[10] = dataGridView1.Rows[currentID].Cells[23].Value.ToString();//txtWitName1
                if (dataSum[10] == "") dataSum[10] = "P0";
                    dataSum[11] = dataGridView1.Rows[currentID].Cells[24].Value.ToString();//txtWitPass1
                if (dataSum[11] == "") dataSum[11] = "P0";
                dataSum[12] = dataGridView1.Rows[currentID].Cells[25].Value.ToString();//txtWitName2
                if (dataSum[12] == "") dataSum[12] = "P0";
                dataSum[13] = dataGridView1.Rows[currentID].Cells[26].Value.ToString();//txtWitPass2
                if (dataSum[13] == "") dataSum[13] = "P0";
                //
                dataSum[15] = dataGridView1.Rows[currentID].Cells[32].Value.ToString();// "غير مؤرشف"  
                dataSum[16] = dataGridView1.Rows[currentID].Cells[33].Value.ToString();// "إجراء توكيل"  
                ArchivedSt.Visible = true;
                ParentData.NewDataValue = true;
                ParentData.txtAuthNoValue.Text = AuthNoPart1 + AuthNoPart2;
                dataSum[17] = dataGridView1.Rows[currentID].Cells[34].Value.ToString();// "المكاتبات الملغية"  
                dataSum[18] = dataGridView1.Rows[currentID].Cells[27].Value.ToString();// "إجراء توكيل"  
                //text1,text2,text3,text4,text5,check1,txtD1,combo1,combo2,addName1 
                dataSum[35] = dataGridView1.Rows[currentID].Cells[27].Value.ToString();// text1
                dataSum[36] = dataGridView1.Rows[currentID].Cells[27].Value.ToString();// text2
                dataSum[37] = dataGridView1.Rows[currentID].Cells[27].Value.ToString();// text3
                dataSum[38] = dataGridView1.Rows[currentID].Cells[27].Value.ToString();// text4
                dataSum[39] = dataGridView1.Rows[currentID].Cells[27].Value.ToString();// text5
                dataSum[40] = dataGridView1.Rows[currentID].Cells[27].Value.ToString();// check1
                dataSum[41] = dataGridView1.Rows[currentID].Cells[27].Value.ToString();// txtD1  
                dataSum[42] = dataGridView1.Rows[currentID].Cells[27].Value.ToString();// combo1  
                dataSum[43] = dataGridView1.Rows[currentID].Cells[27].Value.ToString();// combo2  
                dataSum[44] = dataGridView1.Rows[currentID].Cells[27].Value.ToString();// addName1 
                
                
                DataMovePage.DynamicInvoke(2);
            }
        }
        //private void dataGridView1_DoubleClick(object sender, EventArgs e)
        private void dataGridView1_Click(object sender, EventArgs e)
        {
            if (steadyGrid)
            {
                intID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                steadyGrid = false;
                
                return;
            }

            if (dataGridView1.CurrentRow.Index != -1)
            {
                
                ParentData.NewDataValue = true;
                dataSum[9] = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                //MessageBox.Show(dataSum[9]);

                foreach (Control control in ParentData.MainPanel.Controls)
                {
                    if (control is TextBox && ((TextBox)control).Name == "txtAuthNo")
                    {
                        ParentData.txtAuthNoValue.Text = dataSum[19] = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                        
                    }
                    if (control is TextBox && ((TextBox)control).Name == "txtGreDate")
                    {
                        ((TextBox)control).Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                    }
                    if (control is TextBox && ((TextBox)control).Name == "txtHijDate")
                    {
                        ((TextBox)control).Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                    }
                    if (control is TextBox && ((TextBox)control).Name == "txtAttendVC")
                    {
                        ((TextBox)control).Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();
                    }
                    if (control is ComboBox && ((ComboBox)control).Name == "ComboAuthDestin")
                    {
                        if(dataGridView1.CurrentRow.Cells[21].Value.ToString() != "") 
                            ((ComboBox)control).Text = dataGridView1.CurrentRow.Cells[21].Value.ToString();
                    }
                    if (control is CheckBox && ((CheckBox)control).Name == "checkedViewed")
                    {
                        if (dataGridView1.CurrentRow.Cells[17].Value.ToString() == "غير معالج")
                        {
                            ((CheckBox)control).CheckState = CheckState.Unchecked;
                        }
                        else ((CheckBox)control).CheckState = CheckState.Checked;
                    }
                }
                foreach (Control control in ParentData.subMainPanel.Controls)
                {
                    if (control is CheckBox && ((CheckBox)control).Name == "AppType")
                    {
                        ((CheckBox)control).Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
                        if (((CheckBox)control).Text == "حضور مباشرة إلى القنصلية") ((CheckBox)control).CheckState = CheckState.Checked;
                        else ((CheckBox)control).CheckState = CheckState.Unchecked;
                        if (((CheckBox)control).CheckState == CheckState.Unchecked)
                        {
                            if (((CheckBox)control).CheckState == CheckState.Checked)
                            {
                                ((CheckBox)control).Text = "حضور مباشرة إلى القنصلية";
                                ParentData.MandoubNameValue.Visible = false;
                                ParentData.MandoubLabelValue.Visible = false;
                            }
                            else
                            {
                                ((CheckBox)control).Text = "عن طريق أحد مندوبي القنصلية";
                                ParentData.MandoubNameValue.Visible = true;
                                ParentData.MandoubLabelValue.Visible = true;
                            }
                            ParentData.MandoubNameValue.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();
                        }
                        if (control is TextBox && ((TextBox)control).Name == "txtComment")
                        {
                            ((TextBox)control).Text = dataGridView1.CurrentRow.Cells[29].Value.ToString();
                        }

                    }
                }
                if (dataGridView1.CurrentRow.Cells[2].Value.ToString().Contains("_"))
                {
                    ParentData.strAppnameList = dataGridView1.CurrentRow.Cells[2].Value.ToString().Split('_');
                    ParentData.intAppCounts = dataGridView1.CurrentRow.Cells[2].Value.ToString().Split('_').Length;
                    //MessageBox.Show("here2");
                    
                }
                else
                {
                    ParentData.strAppnameList[0] = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                    
                    ParentData.intAppCounts = 1;
                }

                if (dataGridView1.CurrentRow.Cells[3].Value.ToString().Contains("_"))
                {
                    ParentData.strAppMaleFemaleList = dataGridView1.CurrentRow.Cells[3].Value.ToString().Split('_');
                }
                else
                {
                    ParentData.strAppMaleFemaleList[0] = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                }

                if (dataGridView1.CurrentRow.Cells[4].Value.ToString().Contains("_"))
                {
                    ParentData.strAppDocTypelist = dataGridView1.CurrentRow.Cells[4].Value.ToString().Split('_');
                }
                else
                {
                    ParentData.strAppDocTypelist[0] = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                }

                if (dataGridView1.CurrentRow.Cells[5].Value.ToString().Contains("_"))
                {
                    ParentData.strAppDocNolist = dataGridView1.CurrentRow.Cells[5].Value.ToString().Split('_');

                }
                else
                {

                    ParentData.strAppDocNolist[0] = dataGridView1.CurrentRow.Cells[5].Value.ToString();

                }

                if (dataGridView1.CurrentRow.Cells[6].Value.ToString().Contains("_"))
                {
                    ParentData.strAppissueList = dataGridView1.CurrentRow.Cells[6].Value.ToString().Split('_');
                }
                else
                {
                    ParentData.strAppissueList[0] = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                }

                if (dataGridView1.CurrentRow.Cells[7].Value.ToString().Contains("_"))
                {
                    ParentData.strAuthNames = dataGridView1.CurrentRow.Cells[7].Value.ToString().Split('_');
                    ParentData.intAuthCount = dataGridView1.CurrentRow.Cells[7].Value.ToString().Split('_').Length;
                }
                else
                {
                    ParentData.strAuthNames[0] = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                }

                if (dataGridView1.CurrentRow.Cells[8].Value.ToString().Contains("_"))
                {
                    ParentData.strAuthMaleFemale = dataGridView1.CurrentRow.Cells[8].Value.ToString().Split('_');
                }
                else
                {
                    ParentData.strAuthMaleFemale[0] = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                }
                dataSum[0] = dataGridView1.CurrentRow.Cells[9].Value.ToString();//نوع_التوكيل
                dataSum[1] = dataGridView1.CurrentRow.Cells[10].Value.ToString();//رقم_العمود
                dataSum[2] = dataGridView1.CurrentRow.Cells[11].Value.ToString();//موضوع التوكيل
                dataSum[3] = dataGridView1.CurrentRow.Cells[12].Value.ToString();//حقوق_التوكيل
                dataSum[4] = dataGridView1.CurrentRow.Cells[13].Value.ToString();//التاريخ_الميلادي

                dataSum[10] = dataGridView1.CurrentRow.Cells[23].Value.ToString();//txtWitName1
                dataSum[11] = dataGridView1.CurrentRow.Cells[24].Value.ToString();//txtWitPass1
                dataSum[12] = dataGridView1.CurrentRow.Cells[25].Value.ToString();//txtWitName2
                dataSum[13] = dataGridView1.CurrentRow.Cells[26].Value.ToString();//txtWitPass2
                                                                                  //
                dataSum[15] = dataGridView1.CurrentRow.Cells[32].Value.ToString();// "غير مؤرشف"  
                dataSum[16] = dataGridView1.CurrentRow.Cells[33].Value.ToString();// "إجراء توكيل"  
                ArchivedSt.Visible = true;
                ParentData.NewDataValue = true;
                ParentData.txtAuthNoValue.Text = AuthNoPart1 + AuthNoPart2;

                dataSum[17] = dataGridView1.CurrentRow.Cells[34].Value.ToString();// "إجراء توكيل"  
                dataSum[18] = dataGridView1.CurrentRow.Cells[27].Value.ToString();// "إجراء توكيل"  
                dataSum[21] = dataGridView1.CurrentRow.Cells[35].Value.ToString();// المكاتبات الملغية
                dataSum[20] = dataGridView1.CurrentRow.Cells[22].Value.ToString();// توكيل مرجعي
                
                DataMovePage.DynamicInvoke(2);
            }
        }
        
        private void ColorFulGrid9()
        {
            
            int genAuth = 0;
            int arch = 0;
            int unDesc = 0;
            int inComb = 0;
            int i = 0;
            for (; i < dataGridView1.Rows.Count - 1; i++)
            {
                //dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;

                if (dataGridView1.Rows[i].Cells[2].Value.ToString() == "")
                {
                    inComb++;
                }
                if (dataGridView1.Rows[i].Cells[18].Value.ToString().Contains("مندوب"))
                {
                    // timerColor = false;
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightPink;
                    
                }
                if (dataGridView1.Rows[i].Cells[32].Value.ToString() == "مؤرشف نهائي")
                {
                   // timerColor = false;
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    arch++;
                }
                if (dataGridView1.Rows[i].Cells[9].Value.ToString() == "توكيل بصيغة غير مدرجة")
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightSkyBlue;
                    unDesc++;
                }
                if (dataGridView1.Rows[i].Cells["نوع_التوكيل"].Value.ToString() == "طلاق" && (dataGridView1.Rows[i].Cells["تاريخ_الميلاد"].Value.ToString() == "" || dataGridView1.Rows[i].Cells["المهنة"].Value.ToString() == ""))
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightCoral;
                    
                }

                //else dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.;
                //
            }
            Console.WriteLine("ColorFulGrid9");
            labDescribed.Text ="عدد ("  + i.ToString() + ") معاملة .. عدد (" + inComb.ToString() + ") غير مكتمل.. والمؤرشف منها عدد (" + arch.ToString()+  ").. توكيل بصيغة غير مدرجة (" + unDesc.ToString() + ".. ) توكيل بصغة عامة (" + genAuth.ToString() + ")...";
            //for (int x = 0; x< dataGridView1.Rows.Count - 1; x++)
            //{
            //    splitCol(x, Convert.ToInt32(dataGridView1.Rows[x].Cells[0].Value.ToString()), dataGridView1.Rows[x].Cells[11].Value.ToString().Split('_'));

            //    //if (string.IsNullOrEmpty(dataGridView1.Rows[x].Cells[38].Value.ToString()))
            //    //{
            //    //    // Console.WriteLine("id= " + dataGridView1.Rows[i].Cells[0].Value.ToString() + " - " + dataGridView1.Rows[i].Cells[11].Value.ToString());
            //    //    break;
            //    //}
            //}


        }

        private void SearchDoc_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            ListSearch.Visible = true;
            ListSearch.Text = dlg.FileName;
            
        }

        private void dataGridView1_Click1(object sender, EventArgs e)
        {

        }

        private void deleteRow_Click_1(object sender, EventArgs e)
        {
            int ApplicantID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            deleteRowsData(ApplicantID, "TableAuth", dataSource);
            deleteRow.Visible = false;
            FillDataGridView(dataSource);
        }


        private void deleteRowsData(int v1, string v2, string source)
        {
            string query;
            SqlConnection Con = new SqlConnection(dataSource);
            query = "DELETE FROM " + v2 + " where ID = @ID";
            if (Con.State == ConnectionState.Closed)
                Con.Open();
            SqlCommand sqlCmd = new SqlCommand(query, Con);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", v1);
            sqlCmd.ExecuteNonQuery();
            Con.Close();
            FillDataGridView(dataSource);
        }

        void FillDatafromGenArch(string doc, string id)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableGeneralArch where  رقم_المرجع='" + id + "' and نوع_المستند='" + doc + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string[] allList = new string[dtbl.Rows.Count];
            int i = 0;
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
        private void button7_Click(object sender, EventArgs e)
        {

            if (intID != -1)
                OpenFile(intID, 1);
        }
        private void OpenFile(int id, int fileNo)
        {
            string query;

            SqlConnection Con = new SqlConnection(dataSource);
            if (fileNo == 1)
            {
                query = "select Data1, Extension1,ارشفة_المستندات from TableAuth where ID=@id";
            }
            else if (fileNo == 2)
            {
                query = "select Data2, Extension2,المكاتبة_النهائية from TableAuth where ID=@id";
            }
            else query = "select Data3, Extension3,DocxData from TableAuth where ID=@id";
            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@Id", SqlDbType.Int).Value = id;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                if (fileNo == 1)
                {
                    var name = reader["ارشفة_المستندات"].ToString();
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ssmmhh")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }
                else if(fileNo == 2)
                {
                    var name = reader["المكاتبة_النهائية"].ToString();
                    var Data = (byte[])reader["Data2"];
                    var ext = reader["Extension2"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ssmmhh")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }
                else
                {
                    var name = reader["DocxData"].ToString();
                    var Data = (byte[])reader["Data3"];
                    var ext = reader["Extension3"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("ssmmhh")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    openFile3(NewFileName, id);
                    System.Diagnostics.Process.Start(NewFileName);
                }

            }
            Con.Close();


        }

        private void openFile3(string filename3, int id)
        {
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET DocxData=@DocxData,fileUpload=@fileUpload WHERE ID = @ID", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@ID", id);
                sqlCmd.Parameters.AddWithValue("@DocxData", filename3);
            sqlCmd.Parameters.AddWithValue("@fileUpload", "No");
            sqlCmd.ExecuteNonQuery();
            

            
        }
        private void btnPrevious_Click(object sender, EventArgs e)
        {
             

            ParentData.txtAuthNoValue.Text = AuthNoPart1 + AuthNoPart2;
            DataMovePage.DynamicInvoke(4);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[1].HeaderText.ToString() + " LIKE '" + ListSearch.Text + "%'";
            dataGridView1.DataSource = bs;
            datagrid1();

            ColorFulGrid9();
            steadyGrid = true;
            deleteRow.Visible = true;

        }

        private void UserDataView_Click(object sender, EventArgs e)
        {
            steadyGrid = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {

            //if (intID != -1)
            
                OpenFile(intID, 2);
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (intID == -1) return;
            OpenFile(intID, 3);
        }

        private void UserDataView_Load(object sender, EventArgs e)
        {

        }

        private void saveToDatabase(string filePath1)
        {            
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableAuth (Data1, Extension1, ارشفة_المستندات ) values (@Data1, @Extension1, @ارشفة_المستندات) ", sqlCon);
            //SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableAuth (Data2, Extension2, المكاتبة_النهائية) values(@Data2, @Extension2, @المكاتبة_النهائية) ", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", 1);

            using (Stream stream = File.OpenRead(filePath1))
            {
                byte[] buffer1 = new byte[stream.Length];
                stream.Read(buffer1, 0, buffer1.Length);
                var fileinfo1 = new FileInfo(filePath1);
                string extn1 = fileinfo1.Extension;
                string DocName1 = fileinfo1.Name;
                sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
                sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
                sqlCmd.Parameters.Add("@ارشفة_المستندات", SqlDbType.NVarChar).Value = DocName1;
            }
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }


        private void button4_Click(object sender, EventArgs e)
        {
            string filePath2 = "";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableAuth SET Data3=@Data3, Extension3=@Extension3,DocxData=@DocxData  from TableAuth where ID=@id", sqlCon);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", intID);
            if (ListSearch.Text != "") filePath2 = ListSearch.Text;
            using (Stream stream = File.OpenRead(filePath2))
            {
                byte[] buffer2 = new byte[stream.Length];
                stream.Read(buffer2, 0, buffer2.Length);
                var fileinfo2 = new FileInfo(filePath2);
                string extn2 = fileinfo2.Extension;
                string DocName2 = fileinfo2.Name;
                sqlCmd.Parameters.Add("@Data3", SqlDbType.VarBinary).Value = buffer2;
                sqlCmd.Parameters.Add("@Extension3", SqlDbType.Char).Value = extn2;
                sqlCmd.Parameters.Add("@DocxData", SqlDbType.NVarChar).Value = DocName2;
                ListSearch.Clear();
            }

                    sqlCmd.ExecuteNonQuery();
                }

        private void SearchFile_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void ListSearch_TextChanged(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '" + ListSearch.Text + "%'";
            dataGridView1.DataSource = bs;

            //ColorFulGrid9();
            //DataMovePage.DynamicInvoke(2);
        }
        
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (colored) return;            
            ColorFulGrid9();
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                //dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;

                if (dataGridView1.Rows[i].DefaultCellStyle.BackColor != Color.White)
                {
                    colored = true;
                    return;
                }


            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {

        }

        
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void comYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void ReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            
        }
        
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
