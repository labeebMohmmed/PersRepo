using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using DocumentFormat.OpenXml.Office2016.Drawing.Charts;
using System.Data.SqlTypes;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace PersAhwal
{
    public partial class SignUp : Form
    {
        string DataSource;
        string Employeename;
        string userpass = "";
        string Jobposition;
        int IDEmp;
        bool resetpassword = false;
        string ServerType = "57";
        string GriDate = "";
        bool grdview = false;
        public SignUp(string employeename, string jobposition, string datasource, string serverType, string griDate)
        {
            InitializeComponent();
            DataSource = datasource;
            Employeename = employeename;
            Jobposition = jobposition;
            ServerType = serverType;
            GriDate = griDate;
            //this.Size = new Size(799, 573);
            //MessageBox.Show(employeename);
            if (jobposition.Contains("قنصل"))
            {
                Console.WriteLine("Cons");
                //Register.Text = "تأكيد مستخدم جديد";
                this.Size = new Size(799, 744);
                fillDatagrid();
            }
            else {
                fillDatagrid();
                dataGridView1.Visible = false;
                this.Size = new Size(799, 427);
                for (int i = 0; i < dataGridView1.RowCount - 1; i++) {
                   // MessageBox.Show(dataGridView1.Rows[i].Cells[1].Value.ToString());
                    Console.WriteLine(dataGridView1.Rows[i].Cells[0].Value.ToString());
                    if (dataGridView1.Rows[i].Cells[1].Value.ToString() == Employeename)
                    {
                        
                        detectName(i);
                        break;
                    }
                }
            }
            //dataGridView1_RowIndex(6024);
        }

        private void fillDatagrid()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string settingData = "select ID,EmployeeName,JobPosition,Gender,UserName,Email,pass,comment from TableUser ";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;            
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            sqlCon.Close();
            dataGridView1.Columns[1].Width = 250;
            dataGridView1.Columns[2].Width = 100;
            dataGridView1.Columns[6].Visible = dataGridView1.Columns[0].Visible = false;
        }

        private bool checkName(string name, string userName)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string settingData = "select EmployeeName,UserName from TableUser ";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (name == dataRow["EmployeeName"].ToString() )
                {
                    Register.Visible = false; 
                    MessageBox.Show("يوجد حساب باسم نفس الموظف، يرجى التواصل مع مدير النظام لإسترداد كلمة المرور الخاصة به");                    
                    return true;
                }
                else if ( userName == dataRow["UserName"].ToString())
                {
                    Register.Visible = false;
                    MessageBox.Show("يوجد حساب بنفس اسم المرور، يرجى التواصل مع مدير النظام لإسترداد كلمة المرور الخاصة به");
                    return true;
                }
            }
            Register.Visible = true;
            return false;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string addInfo = "تم إعادة ضبط كلمة المرور لحساب الموظف/" + ApplicantName.Text + " بتاريخ " + GriDate +Environment.NewLine + "----------------------------------------------";
            if (resetpassword && password1.Text == userpass)
            {

                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlCommand sqlCmd = new SqlCommand("UPDATE TableUser SET Pass = @Pass,RestPAss=@RestPAss,comment=@comment WHERE ID = @ID", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@ID", IDEmp);
                sqlCmd.Parameters.AddWithValue("@Pass", password2.Text);
                sqlCmd.Parameters.AddWithValue("@RestPAss", "done");
                sqlCmd.Parameters.AddWithValue("@comment", addInfo + التعليقات_السابقة_Off.Tag);
                
                sqlCmd.ExecuteNonQuery();
                MessageBox.Show("تم إعادة ضبط كلمة المرور");
                IDEmp = 0;
                this.Close();
                return;
            }
            else if (resetpassword && password1.Text != userpass)
            {
                MessageBox.Show("يرحى إدخال كلمة المرور الحالية بشكل صحيح");
                return;
            }
            else
            {
                try
                {
                    if (password1.Text.Equals(password2.Text) && JobPossition.SelectedIndex != 0)
                    {
                        
                         addInfo = "تم تسجيل حساب الموظف/" + ApplicantName.Text + " بتاريخ " + GriDate + Environment.NewLine + "----------------------------------------------";

                        if (sqlCon.State == ConnectionState.Closed)
                            sqlCon.Open();
                        SqlCommand sqlCmd = new SqlCommand("UserAddorEdit", sqlCon);
                        sqlCmd.CommandType = CommandType.StoredProcedure;                        
                        sqlCmd.Parameters.AddWithValue("@ID", 0);
                        sqlCmd.Parameters.AddWithValue("@mode", "Add");
                        sqlCmd.Parameters.AddWithValue("@EmployeeName", ApplicantName.Text);
                        sqlCmd.Parameters.AddWithValue("@JobPosition", JobPossition.Text);
                        sqlCmd.Parameters.AddWithValue("@Gender", EmpGender.Text);
                        sqlCmd.Parameters.AddWithValue("@UserName", userName.Text);
                        sqlCmd.Parameters.AddWithValue("@Email", "");
                        sqlCmd.Parameters.AddWithValue("@Pass", password1.Text);
                        sqlCmd.Parameters.AddWithValue("@Aproved", "غير مؤكد");
                        sqlCmd.Parameters.AddWithValue("@Purpose", ServerType);
                        sqlCmd.Parameters.AddWithValue("@comment", addInfo + التعليقات_السابقة_Off.Tag);
                        try
                        {
                            sqlCmd.ExecuteNonQuery();
                            MessageBox.Show("تم التسجيل بنجاح");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("خطأ في تسجيل البيانات");
                        }

                    }
                    else
                    {
                        MessageBox.Show("كلمة المرور غير متطابقة");
                    }
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
            this.Close();
        }

        private void EmpGender_CheckedChanged(object sender, EventArgs e)
        {
            if (EmpGender.CheckState == CheckState.Unchecked) EmpGender.Text = "ذكر";
            else EmpGender.Text = "أنثى";
        }

       

        private void detectName(int cell)
        {
            IDEmp = Convert.ToInt32(dataGridView1.Rows[cell].Cells[0].Value.ToString());
            //MessageBox.Show(IDEmp.ToString());
            ApplicantName.Text = dataGridView1.Rows[cell].Cells[1].Value.ToString();
            JobPossition.Text = dataGridView1.Rows[cell].Cells[2].Value.ToString();
            EmpGender.Text = dataGridView1.Rows[cell].Cells[3].Value.ToString();
            if (EmpGender.Text == "ذكر") EmpGender.CheckState = CheckState.Unchecked;
            else EmpGender.CheckState = CheckState.Checked;
            userName.Text = dataGridView1.Rows[cell].Cells[4].Value.ToString();
            //email.Text = dataGridView1.Rows[cell].Cells[5].Value.ToString();
            userpass = dataGridView1.Rows[cell].Cells[6].Value.ToString();
            password1.Visible = false;
            password2.Visible = false;
            labelpass2.Visible = false;
            labelpass1.Visible = false;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            resetpassword = true;
            labelpass1.Text = "كلمة المرور القديمة";
            labelpass2.Text = "كلمة المرور الجديدة";
            password1.Visible = password2.Visible= labelpass1.Visible= labelpass2.Visible= true;
        }

        private void SignUp_FormClosed(object sender, FormClosedEventArgs e)
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                grdview = true;
                //ID,EmployeeName,JobPosition,Gender,UserName,Email 
                IDEmp = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                ApplicantName.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                JobPossition.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                EmpGender.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                if (EmpGender.Text == "ذكر") EmpGender.CheckState = CheckState.Unchecked;
                else EmpGender.CheckState = CheckState.Checked;
                userName.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                التعليقات_السابقة_Off.Text = dataGridView1.CurrentRow.Cells["comment"].Value.ToString();
                dataGridView1.Height = 195;
                //email.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();                
                password1.Visible = false;
                password2.Visible = false;
                labelpass2.Visible = false;
                labelpass1.Visible = false;
                btnActivete.Visible = true;
                btnDeActivete.Visible = true;
                grdview = false;
            }
        }
        
        private void dataGridView1_RowIndex(int ID)
        {
            if (dataGridView1.Rows.Count > 1)
            {
                for (int x = 0; x < dataGridView1.Rows.Count - 1; x++)
                {
                    grdview = true;
                    //ID,EmployeeName,JobPosition,Gender,UserName,Email 
                    IDEmp = Convert.ToInt32(dataGridView1.Rows[x].Cells[0].Value.ToString());
                    if (IDEmp == ID)
                    {
                        ApplicantName.Text = dataGridView1.Rows[x].Cells[1].Value.ToString();
                        JobPossition.Text = dataGridView1.Rows[x].Cells[2].Value.ToString();
                        EmpGender.Text = dataGridView1.Rows[x].Cells[3].Value.ToString();
                        if (EmpGender.Text == "ذكر") EmpGender.CheckState = CheckState.Unchecked;
                        else EmpGender.CheckState = CheckState.Checked;
                        userName.Text = dataGridView1.Rows[x].Cells[4].Value.ToString();
                        التعليقات_السابقة_Off.Text = dataGridView1.Rows[x].Cells["comment"].Value.ToString();
                        dataGridView1.Height = 195;
                        //email.Text = dataGridView1.Rows[x].Cells[5].Value.ToString();                
                        password1.Visible = false;
                        password2.Visible = false;
                        labelpass2.Visible = false;
                        labelpass1.Visible = false;
                        btnActivete.Visible = true;
                        btnDeActivete.Visible = true;
                        grdview = false;
                        dataGridView1.Visible = true;  
                    }
                    else continue;
                }
            }
        }

        private void btnActivete_Click(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UserAddorEdit", sqlCon);
            sqlCmd.CommandType = CommandType.StoredProcedure; 
            string addInfo = "تم تفعيل حساب الموظف/" + ApplicantName.Text + " بتاريخ " + GriDate + Environment.NewLine + "----------------------------------------------";
                sqlCmd.Parameters.AddWithValue("@ID", IDEmp);
                sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                sqlCmd.Parameters.AddWithValue("@EmployeeName", ApplicantName.Text);
                sqlCmd.Parameters.AddWithValue("@JobPosition", JobPossition.Text);
                sqlCmd.Parameters.AddWithValue("@Gender", EmpGender.Text);
                sqlCmd.Parameters.AddWithValue("@UserName", userName.Text);
                sqlCmd.Parameters.AddWithValue("@Email", "");
                sqlCmd.Parameters.AddWithValue("@Pass", password1.Text);
                sqlCmd.Parameters.AddWithValue("@Aproved", "أكده " + Jobposition + " " + Employeename);
                sqlCmd.Parameters.AddWithValue("@Purpose", ServerType);
                sqlCmd.Parameters.AddWithValue("@comment", addInfo + التعليقات_السابقة_Off.Text);
                try
                {
                    sqlCmd.ExecuteNonQuery();
                    MessageBox.Show("تم التأكيد بنجاح");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("خطأ في تأكيد البيانات");
                }


            this.Close();
            
        }

        private void btnDeActivete_Click(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UserAddorEdit", sqlCon);
            sqlCmd.CommandType = CommandType.StoredProcedure;
            string addInfo = "تم تعطيل حساب الموظف/" + ApplicantName.Text + " بتاريخ " + GriDate + Environment.NewLine + "----------------------------------------------";
            sqlCmd.Parameters.AddWithValue("@ID", IDEmp);
            sqlCmd.Parameters.AddWithValue("@mode", "Edit");
            sqlCmd.Parameters.AddWithValue("@EmployeeName", ApplicantName.Text);
            sqlCmd.Parameters.AddWithValue("@JobPosition", JobPossition.Text);
            sqlCmd.Parameters.AddWithValue("@Gender", EmpGender.Text);
            sqlCmd.Parameters.AddWithValue("@UserName", userName.Text);
            sqlCmd.Parameters.AddWithValue("@Email", "");
            sqlCmd.Parameters.AddWithValue("@Pass", password1.Text);
            sqlCmd.Parameters.AddWithValue("@Aproved", "غير مؤكد ");
            sqlCmd.Parameters.AddWithValue("@Purpose", ServerType);
            sqlCmd.Parameters.AddWithValue("@comment", addInfo + التعليقات_السابقة_Off.Text);
            try
            {
                sqlCmd.ExecuteNonQuery();
                MessageBox.Show("تم تعطيل حساب الموظف");
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطأ في تأكيد البيانات");
            }
            this.Close();
        }

        private void ApplicantName_TextChanged(object sender, EventArgs e)
        {
            if(grdview)return;
            checkName(ApplicantName.Text, userName.Text);
        }

        private void userName_TextChanged(object sender, EventArgs e)
        {
            if (grdview) return;
            checkName(ApplicantName.Text, userName.Text);
        }

        private void JobPossition_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (JobPossition.SelectedIndex == 5) {
                comJob2.SelectedIndex = 5;
                panelMandoub.Visible = true;
                panelEmployee.Visible = false;
            }
        }

        private void comJob2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comJob2.SelectedIndex != 5)
            {
                JobPossition.SelectedIndex = comJob2.SelectedIndex ;
                panelMandoub.Visible = false;
                panelEmployee.Visible = true;
            }
        }
    }
}
