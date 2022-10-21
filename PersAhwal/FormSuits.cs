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
namespace PersAhwal
{
    public partial class FormSuits : Form
    {
        string DataSource = "";
        int proID = 0;
        DataTable dtblMain;
        bool grdiFill = false;
        public FormSuits(string dataSource, string empName)
        {
            InitializeComponent();
            DataSource = dataSource;
            labEmp.Text = empName;
            
            fillFileBox();
            comFinalaPro.SelectedIndex = 1;
        }

        private void finalPro_Click(object sender, EventArgs e)
        {
            //checkDataInfo(false);
            //if (picVerify.Visible)
            //{
            //    MessageBox.Show("يوجد إجراء سابق متطابق مع رقم الهوية");
            //    return;
            //}

            switch (comFinalaPro.SelectedIndex)
            {
                case 0:
                    //printSaveDoc();
                    //this.Close();
                    break;
                case 1:
                    Save2DataBase();
                    fillFileBox();
                    
                    break;
            }
            
            comFinalaPro.SelectedIndex = 1;
        }

        private void fillFileBox()
        {
            dtblMain = new DataTable();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            try
            {
                SqlDataAdapter sqlDa = new SqlDataAdapter("TableSuitCaseVoS", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
                sqlDa.Fill(dtblMain);
                dataGridView1.DataSource = dtblMain;
                foreach (DataRow dataRow in dtblMain.Rows)
                {
                    bool found = false;
                    for (int a = 0; a < combSuitsCase.Items.Count; a++)
                    {
                        if (!string.IsNullOrEmpty(dataRow["القضية"].ToString()) && dataRow["القضية"].ToString() == combSuitsCase.Items[a].ToString())
                            found = true;
                    }
                    if (!found)
                    {
                        combSuitsCase.Items.Add(dataRow["القضية"].ToString());
                    }
                }
            }
            catch (Exception ex) { 
            }
            sqlCon.Close();
            dataGridView1.Visible = true;
            PanelMain.Visible = false;
        }
        private void Save2DataBase()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);

            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlCommand sqlCmd = new SqlCommand("TableSuitCaseAoE", sqlCon);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.Parameters.AddWithValue("@ID", proID);
                sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                sqlCmd.Parameters.AddWithValue("@رقم_لبرقية", txtMessageNo.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@تاريخ_لبرقية", txtMesDate.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@مقدم_الطلب", txtPersName.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@النوع", checkSex.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@الميلاد", txtBirth.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@المهنة", txtJob.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@العنوان_بالمملكة", combAddress.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@جهة_العمل", txtWorkPlace.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@تاريخ_الاستقدام", txtArrival.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@نوع_الهوية", combDocType.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@رقم_الهوية", txtDocNo.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@القضية", combSuitsCase.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@تاريخ_الاستلام", txtMesRecive.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@تاريخ_الرفع", txtMesFinishDate.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@التاريخ_الميلادي", GregorianDate.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@التاريخ_الهجري", HijriDate.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@مدير_القسم", combViceConsul.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@اسم_الموظف", labEmp.Text.Trim());
                sqlCmd.Parameters.AddWithValue("@تعليق", txtComment.Text.Trim());
                //MessageBox.Show(txtDocNo.Text);
                sqlCmd.ExecuteNonQuery();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
            }
            finally
            {
                sqlCon.Close();
            }
        }
        private void OpenFileDoc(int id, int fileNo)
        {
            string query, table;

            SqlConnection Con = new SqlConnection(DataSource);
            query = "select Data1, Extension1,ارشفة_المستندات from TableSuitCase  where ID=@id";
            if (fileNo == 2)
            query = "select Data2, Extension2,المكاتبة_النهائية from TableSuitCase  where ID=@id";
            

            SqlCommand sqlCmd1 = new SqlCommand(query, Con);
            sqlCmd1.Parameters.Add("@Id", SqlDbType.Int).Value = id;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();
            if (reader.Read())
            {
                if (fileNo == 1)
                {
                    string name = reader["رد_البرقية"].ToString();
                    if (string.IsNullOrEmpty(name)) return;
                    var Data = (byte[])reader["Data1"];
                    var ext = reader["Extension1"].ToString();
                    if (string.IsNullOrEmpty(ext)) return;
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("hhmmss")) + ext;

                    File.WriteAllBytes(NewFileName, Data);
                    if (string.IsNullOrEmpty(NewFileName))
                        return;
                    try
                    {
                        System.Diagnostics.Process.Start(NewFileName);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("FileName " + NewFileName);
                    }
                }
                else if (fileNo == 2)
                {
                    string name = reader["المكاتبة_النهائية"].ToString();
                    if (string.IsNullOrEmpty(name)) return;
                    var Data = (byte[])reader["Data2"];
                    var ext = reader["Extension2"].ToString();
                    var NewFileName = name.Replace(ext, DateTime.Now.ToString("mmss")) + ext;
                    File.WriteAllBytes(NewFileName, Data);
                    System.Diagnostics.Process.Start(NewFileName);
                }

            }
            Con.Close();


        }
        private void GridViewSeach()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("TableSuitCaseVoS", sqlCon);

            sqlDa.SelectCommand.CommandType = CommandType.StoredProcedure;
            DataTable dtbl = new DataTable();

            sqlDa.Fill(dtbl);
            dataGridView6.DataSource = dtbl;
            sqlCon.Close();

        }

        private int checkDataInfo(bool show)
        {
            grdiFill = true;
            picVerify.Visible = true;
            picVerified.Visible = false;
            GridViewSeach();
            for (int z = 0; z < dataGridView6.RowCount - 1; z++)
            {
                if (dataGridView6.Rows[z].Cells[1].Value.ToString() == txtMessageNo.Text)
                {
                    if (show)
                    {
                        var selectedOption = MessageBox.Show("عرض؟", "رقم البرقية تطابق مع إجرا سابق", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (selectedOption == DialogResult.Yes)
                        {
                            selectedID(z, dataGridView6);
                        }
                    }

                    grdiFill = false;
                    return Convert.ToInt32(dataGridView6.Rows[z].Cells[0].Value.ToString());
                }
            }
            picVerify.Visible = false;
            picVerified.Visible = true;
            return -1;
        }


        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                btnFile1.BackColor = System.Drawing.Color.Gainsboro;
                btnFile2.BackColor = System.Drawing.Color.Gainsboro;
                grdiFill = true;
                dataGridView1.Visible = false;
                PanelMain.Visible = true;
                proID = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());                
                if (dataGridView1.CurrentRow.Cells[3].Value.ToString() == "")
                {
                    OpenFileDoc(proID, 1);
                    return;
                }
                
                txtMessageNo.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                txtMesDate.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                txtPersName.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                checkSex.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                txtBirth.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                txtJob.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                combAddress.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                txtWorkPlace.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                txtArrival.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                combDocType.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                txtDocNo.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                combSuitsCase.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                txtMesRecive.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                txtMesFinishDate.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                GregorianDate.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                HijriDate.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();
                combViceConsul.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
                labEmp.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
                txtComment.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
            }
            comFinalaPro.SelectedIndex = 1;
        }

        private void selectedID(int Rowindex, DataGridView dataGridView)
        {
            if (dataGridView1.RowCount >= Rowindex)
            {
                btnFile1.BackColor = System.Drawing.Color.Gainsboro;
                btnFile2.BackColor = System.Drawing.Color.Gainsboro;
                grdiFill = true;
                dataGridView.Visible = false;
                PanelMain.Visible = true;
                proID = Convert.ToInt32(dataGridView.Rows[Rowindex].Cells[0].Value.ToString());
                
                txtMessageNo.Text = dataGridView.Rows[Rowindex].Cells[1].Value.ToString();
                txtMesDate.Text = dataGridView.Rows[Rowindex].Cells[2].Value.ToString();
                txtPersName.Text = dataGridView.Rows[Rowindex].Cells[3].Value.ToString();
                checkSex.Text = dataGridView.Rows[Rowindex].Cells[4].Value.ToString();
                txtBirth.Text = dataGridView.Rows[Rowindex].Cells[5].Value.ToString();
                txtJob.Text = dataGridView.Rows[Rowindex].Cells[6].Value.ToString();
                combAddress.Text = dataGridView.Rows[Rowindex].Cells[7].Value.ToString();
                txtWorkPlace.Text = dataGridView.Rows[Rowindex].Cells[8].Value.ToString();
                txtArrival.Text = dataGridView.Rows[Rowindex].Cells[9].Value.ToString();
                combDocType.Text = dataGridView.Rows[Rowindex].Cells[10].Value.ToString();
                txtDocNo.Text = dataGridView.Rows[Rowindex].Cells[11].Value.ToString();
                combSuitsCase.Text = dataGridView.Rows[Rowindex].Cells[12].Value.ToString();
                txtMesRecive.Text = dataGridView.Rows[Rowindex].Cells[13].Value.ToString();
                txtMesFinishDate.Text = dataGridView.Rows[Rowindex].Cells[14].Value.ToString();
                GregorianDate.Text = dataGridView.Rows[Rowindex].Cells[15].Value.ToString();
                HijriDate.Text = dataGridView.Rows[Rowindex].Cells[16].Value.ToString();
                combViceConsul.Text = dataGridView.Rows[Rowindex].Cells[17].Value.ToString();
                labEmp.Text = dataGridView.Rows[Rowindex].Cells[18].Value.ToString();
                txtComment.Text = dataGridView.Rows[Rowindex].Cells[19].Value.ToString();
            }
        }
        private void picVerify_Click(object sender, EventArgs e)
        {

        }

        private void comFinalaPro_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void FormSuits_FormClosed(object sender, FormClosedEventArgs e)
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
    }
}
