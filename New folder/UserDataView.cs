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
namespace PersAhwal
{
    public partial class UserDataView : UserControl
    {
        string AuthNoPart1 = "ق س ج/160/12/";
        string AuthNoPart2 = "";
        public Delegate DataMovePage;
        string[] dataSum = new string[50];
        string dataSource = "Data Source=192.168.100.123,49170;Network Library=DBMSSOCN;Initial Catalog=AhwalDataBase;User ID=ConsJeddah;Password=DBaseC@nsJ0d103";
        public Form11 ParentData { get; set; }
        bool fileloaded = false;
        public string[] DatasumValue
        {
            get { return dataSum; }
            set { dataSum = value; }
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
            FillDataGridView(dataSource);
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            dataGridView1.BackgroundColor = Color.White;
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            ParentData.txtAuthNoValue.Text = AuthNoPart1 + AuthNoPart2;
            DataMovePage.DynamicInvoke(4);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ParentData.txtAuthNoValue.Text = AuthNoPart1 + AuthNoPart2;
            
            DataMovePage.DynamicInvoke(2);
        }



    public void FillDataGridView(String DataSource)
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
            dataGridView1.Sort(dataGridView1.Columns["ID"], System.ComponentModel.ListSortDirection.Descending);
            AuthNoPart2 = dataGridView1.Rows.Count.ToString();
            sqlCon.Close();
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
        
        //private void dataGridView1_DoubleClick(object sender, EventArgs e)
        private void dataGridView1_Click(object sender, EventArgs e)
        {
            
            if (dataGridView1.CurrentRow.Index != -1)
            {
                dataSum[9] = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                if (ParentData.JobpositionValue.Contains("قنصل")) deleteRow.Visible = true;
                foreach (Control control in ParentData.MainPanel.Controls)
                {                    
                    if (control is TextBox && ((TextBox)control).Name == "txtAuthNo")
                    {
                        ((TextBox)control).Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
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
                    ParentData.intAppCounts= dataGridView1.CurrentRow.Cells[2].Value.ToString().Split('_').Length;
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
                dataSum[0] = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                dataSum[1] = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                dataSum[2] = dataGridView1.CurrentRow.Cells[11].Value.ToString();//موضوع التوكيل
                dataSum[3] = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                dataSum[4] = dataGridView1.CurrentRow.Cells[13].Value.ToString();

                //dataSum[9] = dataGridView1.CurrentRow.Cells[22].Value.ToString();//RelatedAuthNo

                dataSum[10] = dataGridView1.CurrentRow.Cells[23].Value.ToString();//txtWitName1
                dataSum[11] = dataGridView1.CurrentRow.Cells[24].Value.ToString();//txtWitPass1
                dataSum[12] = dataGridView1.CurrentRow.Cells[25].Value.ToString();//txtWitName2
                dataSum[13] = dataGridView1.CurrentRow.Cells[26].Value.ToString();//txtWitPass2                
                dataSum[15] = dataGridView1.CurrentRow.Cells[30].Value.ToString();// "غير مؤرشف"  
                dataSum[16] = dataGridView1.CurrentRow.Cells[31].Value.ToString();// "إجراء توكيل"  
                ArchivedSt.Visible = true;
                ParentData.NewDataValue = true;
                ParentData.txtAuthNoValue.Text = AuthNoPart1 + AuthNoPart2;
                
                DataMovePage.DynamicInvoke(2);
            }
        }

        private void SearchDoc_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            SearchFile.Visible = true;
            SearchFile.Text = dlg.FileName;
            if (SearchFile.Text != "") fileloaded = true;
        }

        private void dataGridView1_Click1(object sender, EventArgs e)
        {

        }

        private void deleteRow_Click(object sender, EventArgs e)
        {
            int ApplicantID = Convert.ToInt32(dataSum[9]);
            deleteRowsData(ApplicantID, "TableAuth", dataSource);
            deleteRow.Visible = false;
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
    }
}
