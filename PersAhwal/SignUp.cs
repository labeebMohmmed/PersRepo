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
        public SignUp(string employeename, string jobposition, string datasource)
        {
            InitializeComponent();
            DataSource = datasource;
            Employeename = employeename;
            Jobposition = jobposition;
            //this.Size = new Size(799, 573);
            //MessageBox.Show(employeename);
            if (jobposition.Contains("قنصل"))
            {
                Console.WriteLine("Cons");
                Register.Text = "تأكيد مستخدم جديد";
                this.Size = new Size(799, 573);
                fillDatagrid();
            }
            else {
                fillDatagrid();
                dataGridView1.Visible = false;

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

        }

        private void fillDatagrid()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string settingData = "select ID,EmployeeName,JobPosition,Gender,UserName,Email,pass from TableUser ";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;            
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            dataGridView1.DataSource = dtbl;
            sqlCon.Close();
            dataGridView1.Columns[6].Visible = false;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (resetpassword && password1.Text == userpass)
            {

                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlCommand sqlCmd = new SqlCommand("UPDATE TableUser SET Pass = @Pass,RestPAss=@RestPAss WHERE ID = @ID", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@ID", IDEmp);
                sqlCmd.Parameters.AddWithValue("@Pass", password2.Text);
                sqlCmd.Parameters.AddWithValue("@RestPAss", "done");
                
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
                        if (sqlCon.State == ConnectionState.Closed)
                            sqlCon.Open();
                        SqlCommand sqlCmd = new SqlCommand("UserAddorEdit", sqlCon);
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        if (Register.Text == "تسجيل")
                        {
                            sqlCmd.Parameters.AddWithValue("@ID", 0);
                            sqlCmd.Parameters.AddWithValue("@mode", "Add");
                            sqlCmd.Parameters.AddWithValue("@EmployeeName", ApplicantName.Text);
                            sqlCmd.Parameters.AddWithValue("@JobPosition", JobPossition.Text);
                            sqlCmd.Parameters.AddWithValue("@Gender", EmpGender.Text);
                            sqlCmd.Parameters.AddWithValue("@UserName", userName.Text);
                            sqlCmd.Parameters.AddWithValue("@Email", email.Text);
                            sqlCmd.Parameters.AddWithValue("@Pass", password1.Text);
                            sqlCmd.Parameters.AddWithValue("@Aproved", "غير مؤكد");
                            MessageBox.Show("تم التسجيل بنجاح");
                        }
                        else
                        {
                            sqlCmd.Parameters.AddWithValue("@ID", IDEmp);
                            sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                            sqlCmd.Parameters.AddWithValue("@EmployeeName", ApplicantName.Text);
                            sqlCmd.Parameters.AddWithValue("@JobPosition", JobPossition.Text);
                            sqlCmd.Parameters.AddWithValue("@Gender", EmpGender.Text);
                            sqlCmd.Parameters.AddWithValue("@UserName", userName.Text);
                            sqlCmd.Parameters.AddWithValue("@Email", email.Text);
                            sqlCmd.Parameters.AddWithValue("@Pass", password1.Text);
                            sqlCmd.Parameters.AddWithValue("@Aproved", "أكده " + Jobposition + " " + Employeename);
                            MessageBox.Show("تم التأكيد بنجاح");
                        }
                        sqlCmd.ExecuteNonQuery();

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
        }

        private void EmpGender_CheckedChanged(object sender, EventArgs e)
        {
            if (EmpGender.CheckState == CheckState.Unchecked) EmpGender.Text = "ذكر";
            else EmpGender.Text = "أنثى";
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

            if (dataGridView1.CurrentRow.Index != -1)
            {
                //ID,EmployeeName,JobPosition,Gender,UserName,Email 
                IDEmp = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                ApplicantName.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                JobPossition.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                EmpGender.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                if (EmpGender.Text == "ذكر") EmpGender.CheckState = CheckState.Unchecked;
                else EmpGender.CheckState = CheckState.Checked;
                userName.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                email.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                Register.Text = "تأكيد";
                password1.Visible = false;
                password2.Visible = false;
                labelpass2.Visible = false;
                labelpass1.Visible = false;
            }
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
            email.Text = dataGridView1.Rows[cell].Cells[5].Value.ToString();
            userpass = dataGridView1.Rows[cell].Cells[6].Value.ToString();
            Register.Text = "تأكيد";
            password1.Visible = false;
            password2.Visible = false;
            labelpass2.Visible = false;
            labelpass1.Visible = false;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Register.Text = "إعادة تعيين";
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
    }
}
