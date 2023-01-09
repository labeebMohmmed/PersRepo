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
            //MessageBox.Show(DataSource);
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
                sqlCmd.Parameters.AddWithValue("@comment", addInfo + التعليقات_السابقة_Off.Text);
                
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
                        sqlCmd.Parameters.AddWithValue("@comment", addInfo + التعليقات_السابقة_Off.Text);
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
                IDEmp = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                if (panelEmployee.Visible)
                {
                    ApplicantName.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    JobPossition.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                    EmpGender.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                    if (EmpGender.Text == "ذكر") EmpGender.CheckState = CheckState.Unchecked;
                    else EmpGender.CheckState = CheckState.Checked;
                    userName.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();                    
                    Register.Text = "تعديل";
                }
                else {
                    اسم_المندوب.Text = dataGridView1.CurrentRow.Cells["MandoubNames"].Value.ToString();
                    رقم_الهاتف.Text = dataGridView1.CurrentRow.Cells["MandoubPhones"].Value.ToString();
                    اسم_المنطقة.Text = dataGridView1.CurrentRow.Cells["MandoubAreas"].Value.ToString();
                    الصفة.Text = dataGridView1.CurrentRow.Cells["الصفة"].Value.ToString();
                    يوم_المراجعة.Text = dataGridView1.CurrentRow.Cells["مواعيد_الحضور"].Value.ToString();
                    بيانات_المندوب.Text = "تعديل";
                }
                التعليقات_السابقة_Off.Text = dataGridView1.CurrentRow.Cells["comment"].Value.ToString();
                dataGridView1.Height = 195;
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
                btnActiveteM.Visible = btnDeActiveteM.Visible = panelMandoub.Visible = true;
                btnActivete.Visible = btnDeActivete.Visible = panelEmployee.Visible = false;
                fillMandoubGrid();
            }
        }

        private void fillMandoubGrid()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT ID,MandoubNames,MandoubAreas,MandoubPhones,مواعيد_الحضور,الصفة,وضع_المندوب,comment FROM TableMandoudList", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            dataGridView1.DataSource = table;
            if (dataGridView1.Rows.Count > 1)
            {
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns[1].Width = 180;
                dataGridView1.Columns[2].Width = 80;
                dataGridView1.Columns[3].Width = 120;
            }
        }

        private void comJob2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comJob2.SelectedIndex != 5)
            {
                JobPossition.SelectedIndex = comJob2.SelectedIndex ;
                btnActiveteM.Visible = btnDeActiveteM.Visible = panelMandoub.Visible = false;
                btnActivete.Visible = btnDeActivete.Visible = panelEmployee.Visible = true;
                fillDatagrid();
            }
        }

        private void بيانات_المندوب_Click(object sender, EventArgs e)
        {
            string query = "insert into TableMandoudList (MandoubNames,MandoubAreas,MandoubPhones,مواعيد_الحضور,الصفة,وضع_المندوب,comment) values (@MandoubNames,@MandoubAreas,@MandoubPhones,@مواعيد_الحضور,@الصفة,@وضع_المندوب,@comment)";
            string addInfo = "تم تسجيل حساب المندوب/" + اسم_المندوب.Text + " بتاريخ " + GriDate + Environment.NewLine + "----------------------------------------------";
            if (بيانات_المندوب.Text == "تعديل")
                query = "update TableMandoudList set MandoubNames=@MandoubNames,MandoubAreas=@MandoubAreas,MandoubPhones=@MandoubPhones,مواعيد_الحضور=@مواعيد_الحضور,الصفة=@الصفة,وضع_المندوب=@وضع_المندوب,comment=@comment where ID=@id";

                SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            try
            {
                if (بيانات_المندوب.Text == "تعديل")
                {
                    addInfo = "تم تعديل حساب المندوب/" + اسم_المندوب.Text + " بتاريخ " + GriDate + Environment.NewLine + "----------------------------------------------";

                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@id", IDEmp);
                    sqlCmd.Parameters.AddWithValue("@MandoubNames", اسم_المندوب.Text);
                    sqlCmd.Parameters.AddWithValue("@MandoubAreas", اسم_المنطقة.Text);
                    sqlCmd.Parameters.AddWithValue("@MandoubPhones", رقم_الهاتف.Text);
                    sqlCmd.Parameters.AddWithValue("@وضع_المندوب", "في انتظار تفعيل الحساب");
                    sqlCmd.Parameters.AddWithValue("@الصفة", الصفة.Text);
                    sqlCmd.Parameters.AddWithValue("@مواعيد_الحضور", يوم_المراجعة.Text);
                    sqlCmd.Parameters.AddWithValue("@comment", addInfo + التعليقات_السابقة_Off.Text);
                    try
                    {
                        sqlCmd.ExecuteNonQuery();
                        MessageBox.Show("تم التعديل بنجاح");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("خطأ في تعديل البيانات");
                    }
                }
                else
                {

                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@MandoubNames", اسم_المندوب.Text);
                    sqlCmd.Parameters.AddWithValue("@MandoubAreas", اسم_المنطقة.Text);
                    sqlCmd.Parameters.AddWithValue("@MandoubPhones", رقم_الهاتف.Text);
                    sqlCmd.Parameters.AddWithValue("@وضع_المندوب", "في انتظار تفعيل الحساب");
                    sqlCmd.Parameters.AddWithValue("@الصفة", الصفة.Text);
                    sqlCmd.Parameters.AddWithValue("@مواعيد_الحضور", يوم_المراجعة.Text);
                    sqlCmd.Parameters.AddWithValue("@comment", addInfo + التعليقات_السابقة_Off.Text);
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
            }
            finally
            {
                sqlCon.Close();
            }
            this.Close();
        }

        private void btnDeActiveteM_Click(object sender, EventArgs e)
        {
            string query = "update TableMandoudList set وضع_المندوب=@وضع_المندوب,comment=@comment where ID=@id";
            string addInfo = "تم تسجيل تعطيل المندوب/" + اسم_المندوب.Text + " بتاريخ " + GriDate + Environment.NewLine + "----------------------------------------------";

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            try
            {
                if (بيانات_المندوب.Text == "تعديل")
                {
                   

                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@id", IDEmp);
                    sqlCmd.Parameters.AddWithValue("@وضع_المندوب", "الحساب معطل");
                    sqlCmd.Parameters.AddWithValue("@comment", addInfo + التعليقات_السابقة_Off.Text);
                    try
                    {
                        sqlCmd.ExecuteNonQuery();
                        MessageBox.Show("تم تعطيل حساب المندوب بنجاح");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("خطأ في البيانات");
                    }
                }
            }
            catch (Exception ex) { 
            }

            this.Close();
        }

        private void btnActiveteM_Click(object sender, EventArgs e)
        {
            string query = "update TableMandoudList set وضع_المندوب=@وضع_المندوب,comment=@comment where ID=@id";
            string addInfo = "تم تسجيل تفعيل المندوب/" + اسم_المندوب.Text + " بتاريخ " + GriDate + Environment.NewLine + "----------------------------------------------";

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            try
            {
                if (بيانات_المندوب.Text == "تعديل")
                {


                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@id", IDEmp);
                    sqlCmd.Parameters.AddWithValue("@وضع_المندوب", "الحساب مفعل");
                    sqlCmd.Parameters.AddWithValue("@comment", addInfo + التعليقات_السابقة_Off.Text);
                    try
                    {
                        sqlCmd.ExecuteNonQuery();
                        MessageBox.Show("تم تفعيل حساب المندوب بنجاح");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("خطأ في البيانات");
                    }
                }
            }
            catch (Exception ex)
            {
            }
            this.Close();
        }
    }
}
