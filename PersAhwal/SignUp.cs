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
using DocumentFormat.OpenXml.Presentation;

namespace PersAhwal
{
    public partial class SignUp : Form
    {
        string DataSource;
        string Empname;
        string userpass = "";
        string JobPos;
        int IDEmp;
        bool resetpassword = false;
        string ServerType = "57";
        string GriDate = "";
        bool grdview = false;
        string Update = "";
        string[] allList;
        string diplomat = "";

        public SignUp(string employee, string jobposition, string datasource, string serverType, string griDate, string update, string career, int jobIndex)
        {
            InitializeComponent();
            DataSource = datasource;
            allList = getColList("TableUser");
            Empname = employee;
            if (Empname == "جديد")
                olspassword.Visible = label13.Visible = false;
            else olspassword.Visible = label13.Visible = true;
            JobPos = jobposition;
            ServerType = serverType;
            GriDate = griDate;
            Update = update;
            if (update == "yes")
            {
                Register.Text = "تحديث";
                JobPosition.Enabled = false;
                UserName.Enabled = false;
            }

            getInfo(Empname);

            if (career == "مدير نظام" || jobIndex != -1)
            {
                //this.Size = new Size(870, 769);
                btnActivate.Visible = true;
                //fillDatagrid();
            }
            else {
                btnActivate.Visible = true;
                fillDatagrid();
                dataGridView1.Visible = false;
                this.Size = new Size(967, 435);
                //for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                //{
                //    if (dataGridView1.Rows[i].Cells[1].Value.ToString() == Empname)
                //    {
                //        detectName(i);
                //        break;
                //    }
                //}
            }
            if (jobIndex != -1)
            {
                JobPosition.SelectedIndex = jobIndex;
                JobPosition.Enabled = false;
            }
        }

        string updateAll = "";
        string insertAll = "";
        private void getInfo(string name)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * from TableUser WHERE EmployeeName = N'" + name + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow row in dtbl.Rows)
            {
                IDEmp = Convert.ToInt32(row["ID"].ToString());
                JobPosition.Text = row["JobPosition"].ToString();
                EmployeeName.Text = name;
                EngEmployeeName.Text = row["EngEmployeeName"].ToString();
                UserName.Text = row["UserName"].ToString();
                JobPosition.Text = row["JobPosition"].ToString();
                Gender.Text = row["Gender"].ToString();
                if (Gender.Text == "ذكر")
                    Gender.Checked = false;
                else
                    Gender.Checked = true;
                headOfMission.Text = row["headOfMission"].ToString();
                if (headOfMission.Text == "")
                {
                    headOfMission.Checked = false;
                    headOfMission.Text = "عضو بعثة";
                }
                else
                    headOfMission.Checked = true;
                olspassword.Text = userpass = row["Pass"].ToString();
                Register.Text = "تعديل";
                btnActivete.Visible = olspassword.Visible = true;
                btnActiveteM.Visible = false;
                labelpass1.Visible = password1.Visible = password2.Visible = labelpass2.Visible = editPass.Visible = false;
                btnActivete.BringToFront();
            }
        }

        //public void panelFill(Control control)
        //{
        //    for (int col = 0; col < allList.Length; col++)
        //    {
        //        if (control.Name == allList[col])
        //        {
        //            if (dataGridView1.CurrentRow.Cells[allList[col]].Value.ToString() != "")
        //            {
        //                control.Text = dataGridView1.CurrentRow.Cells[allList[col]].Value.ToString();
        //                IDEmp = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());                        
        //            }                    
        //        }
        //    }
        //}
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

                if (row["name"].ToString() != "endTime" && row["name"].ToString() != "ID" && row["name"].ToString() != "تاريخ_الارشفة1" && row["name"].ToString() != "تاريخ_الارشفة2" && row["name"].ToString() != "حالة_الارشفة" && row["name"].ToString() != "sms")
                {
                    allList[i] = row["name"].ToString();
                    //MessageBox.Show(row["name"].ToString());
                    if (i == 0)
                    {
                        insertItems = row["name"].ToString();
                        insertValues = "@" + row["name"].ToString();
                        updateValues = row["name"].ToString() + "=@" + row["name"].ToString();
                    }
                    else
                    {
                        insertItems = insertItems + "," + row["name"].ToString();
                        insertValues = insertValues + "," + "@" + row["name"].ToString();
                        updateValues = updateValues + "," + row["name"].ToString() + "=@" + row["name"].ToString();
                    }
                    i++;
                }
            }
            updateAll = "UPDATE " + table + " SET " + updateValues + " where ID = @id";
            insertAll = "INSERT INTO " + table + "(" + insertItems + ") values (" + insertValues + ")";

            return allList;

        }
        private void fillDatagrid()
        {
            //MessageBox.Show(DataSource);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            string settingData = "select ID,EmployeeName,JobPosition,Gender,UserName,Email,pass,comment,headOfMission,EngEmployeeName from TableUser ";
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
                if (name == dataRow["EmployeeName"].ToString())
                {
                    Register.Visible = false;
                    MessageBox.Show("يوجد حساب باسم نفس الموظف، يرجى التواصل مع مدير النظام لإسترداد كلمة المرور الخاصة به");
                    return true;
                }
                else if (userName == dataRow["UserName"].ToString())
                {
                    Register.Visible = false;
                    MessageBox.Show("يوجد حساب بنفس اسم المرور، يرجى التواصل مع مدير النظام لإسترداد كلمة المرور الخاصة به");
                    return true;
                }
            }
            Register.Visible = true;
            return false;
        }

        private bool passwordCheck()
        {
            string newPass = userpass;
            if (JobPosition.Text == "إختر الوظيفة")
            {
                MessageBox.Show("يرجى اختيار الوظيفة");
                return false;
            }
                if (password1.Text != password2.Text)
            {
                MessageBox.Show("كلمة المرور غير متطابقة");
                return false;
            }
            if (password1.Text == userpass)
            {
                newPass = password1.Text;
                MessageBox.Show("كلمة المرور الجديدة لا يمكن أن تطابق الكلمة السابقة");
                return false;
            }
            if (password1.Text != "" && password1.Text.Length < 6)
            {
                MessageBox.Show("كلمة المرور يجب أن لا تقل عن ستة رموز");
                return false;
            }
            if (password1.Text != "" && password1.Text.All(char.IsDigit))
            {
                MessageBox.Show("كلمة المرور يجب أن تحتوي على أحرف");
                return false;
            }
            return true;
        }
        string[] items = new string[15] { "", "", "","", "","", "", "","", "","", "", "","", ""};
        string[] values = new string[15] { "", "", "","", "","", "", "","", "","", "", "","", ""};
        private bool saveItems(string table, string[] item,string[] value, bool saveType) 
        {
            string query = ""; 
                string itemName = item[0];
            string itemValue = "@" + item[0];
            string itemUpdate = item[0] + "=@" + item[0];
            
            for (int i = 1;  i< item.Length && item[i] != ""; i++)
            {
                if (item[i] != "ID")
                {
                    itemName = itemName + "," + item[i];
                    itemValue = itemValue + ",@" + item[i];
                    itemUpdate = itemUpdate + "," + item[i] + "=@" + item[i];
                }
            }
            if(saveType)
                 query = "insert into " + table + "(" + itemName + ")values ( " + itemValue + ")";
            else 
                 query = "UPDATE " + table + " SET " + itemUpdate + " where ID = " + IDEmp.ToString();
            Console.WriteLine(query);
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;

            for (int i = 0; i<item.Length&& item[i] != ""; i++)
            {
                sqlCmd.Parameters.AddWithValue("@" + item[i], value[i]);
               MessageBox.Show( item[i] + " - "+ value[i]);
            }
            //try
            //{
                sqlCmd.ExecuteNonQuery();
            //}
            //catch (Exception ex) {
            //    return false;
            //}
            return true;
            
        }


        private void button1_Click(object sender, EventArgs e)
        {
            string career = "مدير";
            string newPass = userpass;
            if (JobPosition.Text == "تعين محلي" || JobPosition.Text == "مندوب جالية" || JobPosition.Text == "محاسب")
                career = "موظف";
            
            
            
            if (JobPosition.Text == "محاسب"||JobPosition.Text == "مدير مالي")
                ServerType = "محاسب";            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string addInfo = "قام " + UserName.Text + " بتحديث بيانات حسابه بتاريخ " + GriDate +Environment.NewLine + "----------------------------------------------";
            if (Register.Text == "تسجيل")
            {
                if (!passwordCheck())
                    return;
                addInfo = "تم تسجيل حساب الموظف/" + UserName.Text + " بتاريخ " + GriDate + Environment.NewLine + "----------------------------------------------";
                items = new string[15] { "EmployeeName", "Career", "JobPosition", "Gender", "UserName", "Pass", "EngEmployeeName", "Aproved", "Purpose", "RestPAss", "headOfMission", "comment", "الدبلوماسيون", "", "" };
                values = new string[15] { EmployeeName.Text, career, JobPosition.Text, Gender.Text, UserName.Text, password1.Text, EngEmployeeName.Text, "غير مؤكد", ServerType, "done", headOfMission.Text, addInfo + التعليقات_السابقة_Off.Text, diplomat, "", "" };
                if (saveItems("TableUser", items, values, true))
                    MessageBox.Show("تم التسجيل بنجاح");
                else
                    MessageBox.Show("تعذر التسجيل لاسباب فنية");
            }
           
            

            
            else if (Register.Text == "تعديل")
            {
                if (JobPosition.Text == "إختر الوظيفة")
                {
                    MessageBox.Show("يرجى اختيار الوظيفة");
                    return ;
                }
                addInfo = "تم تعديل بيانات حساب الموظف/" + UserName.Text + " بتاريخ " + GriDate + Environment.NewLine + "----------------------------------------------";
                items = new string[14] { "EmployeeName",     "Career", "JobPosition",    "Gender", "UserName",     "EngEmployeeName",     "Aproved", "Purpose",   "RestPAss", "headOfMission",    "comment", "الدبلوماسيون", "", "" };
                values = new string[14] { EmployeeName.Text, career, JobPosition.Text, Gender.Text, UserName.Text,   EngEmployeeName.Text, "غير مؤكد", ServerType, "done",     headOfMission.Text, addInfo + التعليقات_السابقة_Off.Text, diplomat, "", "" };
                if (saveItems("TableUser", items, values, false))
                    MessageBox.Show("تم التعديل بنجاح");
                else
                    MessageBox.Show("تعذر التعديل لاسباب فنية");
                MessageBox.Show("بتطلب تعديل البيانات إعادة تفعيل الحساب بواسطة مدير القسم الحالي");

            }
            this.Close();
        }

        private void EmpGender_CheckedChanged(object sender, EventArgs e)
        {
            if (Gender.CheckState == CheckState.Unchecked) Gender.Text = "ذكر";
            else Gender.Text = "أنثى";
        }

       

        private void detectName(int cell)
        {
            IDEmp = Convert.ToInt32(dataGridView1.Rows[cell].Cells[0].Value.ToString());
            //MessageBox.Show(IDEmp.ToString());
            UserName.Text = dataGridView1.Rows[cell].Cells[1].Value.ToString();
            JobPosition.Text = dataGridView1.Rows[cell].Cells[2].Value.ToString();
            Gender.Text = dataGridView1.Rows[cell].Cells[3].Value.ToString();
            if (Gender.Text == "ذكر") Gender.CheckState = CheckState.Unchecked;
            else Gender.CheckState = CheckState.Checked;
            EmployeeName.Text = dataGridView1.Rows[cell].Cells[4].Value.ToString();
            //email.Text = dataGridView1.Rows[cell].Cells[5].Value.ToString();
            userpass = dataGridView1.Rows[cell].Cells[6].Value.ToString();
            password1.Visible = false;
            password2.Visible = false;
            labelpass2.Visible = false;
            labelpass1.Visible = false;
            if (Update == "yes")
            {
                password1.Visible = password2.Visible = labelpass1.Visible = labelpass2.Visible = true;
            }
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
                    UserName.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                    JobPosition.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                    Gender.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                    userpass = dataGridView1.CurrentRow.Cells["Pass"].Value.ToString();
                    if (Gender.Text == "ذكر") Gender.CheckState = CheckState.Unchecked;
                    else Gender.CheckState = CheckState.Checked;
                    EmployeeName.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString(); 
                    if (dataGridView1.CurrentRow.Cells["headOfMission"].Value.ToString() == "")
                        headOfMission.Checked = false;
                    else headOfMission.Checked = true;
                    EngEmployeeName.Text = dataGridView1.CurrentRow.Cells["EngEmployeeName"].Value.ToString();                    
                    Register.Text = "تعديل";
                    btnActivete.Visible = true;
                    btnActiveteM.Visible = false;
                    btnActivete.BringToFront();
                    //MessageBox.Show(userpass);
                }
                else {
                    btnActiveteM.Enabled = false;
                    btnActivete.Visible = false;
                    btnActiveteM.Visible = true;
                    btnActiveteM.BringToFront();
                    passport.Text = dataGridView1.CurrentRow.Cells["رقم_الجواز"].Value.ToString();
                    اسم_المندوب.Text = dataGridView1.CurrentRow.Cells["MandoubNames"].Value.ToString();
                    رقم_الهاتف.Text = dataGridView1.CurrentRow.Cells["MandoubPhones"].Value.ToString();
                    اسم_المنطقة.Text = dataGridView1.CurrentRow.Cells["MandoubAreas"].Value.ToString();
                    الصفة.Text = dataGridView1.CurrentRow.Cells["الصفة"].Value.ToString();
                    يوم_المراجعة.Text = dataGridView1.CurrentRow.Cells["مواعيد_الحضور"].Value.ToString();
                    بيانات_المندوب.Text = "تعديل";
                    FillDatafromGenArch("data2", IDEmp.ToString(), "TableMandoudList");
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
                if (Update == "yes")
                {
                    password1.Visible = password2.Visible = labelpass1.Visible = labelpass2.Visible = true;
                }
            }
        }
        void FillDatafromGenArch(string doc, string id, string table)
        {
            string query = "select * from TableGeneralArch where  رقم_المرجع='" + id + "' and نوع_المستند='" + doc + "' and docTable='" + table + "'";
            string database = DataSource.Replace("AhwalDataBase", "ArchFilesDB");
            //database = DataSource.Replace("SudaneseAffairs", "ArchFilesDB");
            
            SqlConnection sqlCon = new SqlConnection(database);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            Console.WriteLine(database);
            Console.WriteLine(query);
            //MessageBox.Show(database);
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if(dtbl.Rows.Count > 0) {
                var selectedOption = MessageBox.Show("عرض الخطاب؟", "للمندوب خطاب تفويض مؤرشف",  MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (selectedOption == DialogResult.Yes)
                {
                    foreach (DataRow reader in dtbl.Rows)
                    {
                        var name = reader["المستند"].ToString();
                        var Data = (byte[])reader["Data1"];
                        var ext = reader["Extension1"].ToString();
                        var NewFileName = name.Replace(ext, DateTime.Now.ToString("ddMMyyyyhhmmss")) + ext;
                        File.WriteAllBytes(NewFileName, Data);
                        System.Diagnostics.Process.Start(NewFileName);
                        
                    }
                }
                btnActiveteM.Enabled = true;
            }
            


            sqlCon.Close();
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
                        UserName.Text = dataGridView1.Rows[x].Cells[1].Value.ToString();
                        JobPosition.Text = dataGridView1.Rows[x].Cells[2].Value.ToString();
                        Gender.Text = dataGridView1.Rows[x].Cells[3].Value.ToString();
                        if (Gender.Text == "ذكر") Gender.CheckState = CheckState.Unchecked;
                        else Gender.CheckState = CheckState.Checked;
                        EmployeeName.Text = dataGridView1.Rows[x].Cells[4].Value.ToString();
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
                        if (Update == "yes")
                        {
                            password1.Visible = password2.Visible = labelpass1.Visible = labelpass2.Visible = true;
                        }
                        dataGridView1.Visible = true;  
                    }
                    else continue;
                }
            }
        }

        private void btnActivete_Click(object sender, EventArgs e)
        {
            string career = "مدير";
            if (JobPosition.Text == "محاسب" || JobPosition.Text == "مدير مالي")
                ServerType = "محاسب";

            if (JobPosition.Text == "تعين محلي" || JobPosition.Text == "مندوب جالية" || JobPosition.Text == "محاسب")
                career = "موظف"; 
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UserAddorEdit", sqlCon);
            sqlCmd.CommandType = CommandType.StoredProcedure;
            string addInfo = "تم تفعيل حساب الموظف/" + UserName.Text + " بتاريخ " + GriDate + Environment.NewLine + "----------------------------------------------";
            sqlCmd.Parameters.AddWithValue("@ID", IDEmp);
            sqlCmd.Parameters.AddWithValue("@mode", "Edit");

            sqlCmd.Parameters.AddWithValue("@EmployeeName", UserName.Text);
            sqlCmd.Parameters.AddWithValue("@Career", career);
            sqlCmd.Parameters.AddWithValue("@JobPosition", JobPosition.Text);
            sqlCmd.Parameters.AddWithValue("@Gender", Gender.Text);
            sqlCmd.Parameters.AddWithValue("@UserName", EmployeeName.Text);
            sqlCmd.Parameters.AddWithValue("@Email", "");
            sqlCmd.Parameters.AddWithValue("@EngEmployeeName", EngEmployeeName.Text);
            sqlCmd.Parameters.AddWithValue("@Pass", userpass);
            sqlCmd.Parameters.AddWithValue("@RestPAss", "done");
            sqlCmd.Parameters.AddWithValue("@Aproved", "أكده " + JobPos + " " + Empname);
            sqlCmd.Parameters.AddWithValue("@Purpose", ServerType);
            if (headOfMission.Checked)
                sqlCmd.Parameters.AddWithValue("@headOfMission", "head");
            else
                sqlCmd.Parameters.AddWithValue("@headOfMission", "");
            sqlCmd.Parameters.AddWithValue("@comment", addInfo + التعليقات_السابقة_Off.Text);


            try
            {
                sqlCmd.ExecuteNonQuery();
                MessageBox.Show("تم تفعيل الحساب بنجاح");
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطأ في تفعيل البيانات");
            }


            this.Close();

        }

        private void btnDeActivete_Click(object sender, EventArgs e)
        {
            string career = "مدير";
            if (JobPosition.Text == "تعين محلي" || JobPosition.Text == "مندوب جالية" || JobPosition.Text == "محاسب")
                career = "موظف";
            if (JobPosition.Text == "محاسب" || JobPosition.Text == "مدير مالي")
                ServerType = "محاسب";

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand("UserAddorEdit", sqlCon);
            sqlCmd.CommandType = CommandType.StoredProcedure;
            string addInfo = "تم تعطيل حساب الموظف/" + UserName.Text + " بتاريخ " + GriDate + Environment.NewLine + "----------------------------------------------";
            sqlCmd.Parameters.AddWithValue("@ID", IDEmp);
            sqlCmd.Parameters.AddWithValue("@mode", "Edit");
            sqlCmd.Parameters.AddWithValue("@EmployeeName", UserName.Text);
            sqlCmd.Parameters.AddWithValue("@JobPosition", JobPosition.Text);
            sqlCmd.Parameters.AddWithValue("@Gender", Gender.Text);
            sqlCmd.Parameters.AddWithValue("@UserName", EmployeeName.Text);
            sqlCmd.Parameters.AddWithValue("@Email", "");
            sqlCmd.Parameters.AddWithValue("@Career", career);
            sqlCmd.Parameters.AddWithValue("@Pass", password1.Text);
            sqlCmd.Parameters.AddWithValue("@Aproved", "غير مؤكد ");
            sqlCmd.Parameters.AddWithValue("@Purpose", ServerType);
            sqlCmd.Parameters.AddWithValue("@RestPAss", "");
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
            if(grdview || Update == "yes") return;
            //MessageBox.Show(Update);
            checkName(UserName.Text, EmployeeName.Text);
        }

        private void userName_TextChanged(object sender, EventArgs e)
        {
            if (grdview || Update == "yes") return;
           // MessageBox.Show(Update);
            checkName(UserName.Text, EmployeeName.Text);
        }

        private void JobPossition_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            if (JobPosition.Text == "القنصل العام" || JobPosition.Text == "القنصل العام بالإنابة" || JobPosition.Text == "نائب قنصل")
                diplomat = "yes"; 
            if (JobPosition.SelectedIndex == 5) {
                comJob2.SelectedIndex = 5;
                btnActiveteM.Visible = btnDeActiveteM.Visible = panelMandoub.Visible = true;
                btnActivete.Visible = btnDeActivete.Visible = panelEmployee.Visible = false;
                fillMandoubGrid();
                //MessageBox.Show("text");
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
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT ID,MandoubNames,MandoubAreas,MandoubPhones,مواعيد_الحضور,الصفة,وضع_المندوب,comment,رقم_الجواز FROM TableMandoudList", sqlCon);
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
                JobPosition.SelectedIndex = comJob2.SelectedIndex ;
                btnActiveteM.Visible = btnDeActiveteM.Visible = panelMandoub.Visible = false;
                btnActivete.Visible = btnDeActivete.Visible = panelEmployee.Visible = true;
                fillDatagrid();
            }
        }

        private void بيانات_المندوب_Click(object sender, EventArgs e)
        {
            string query = "insert into TableMandoudList (MandoubNames,MandoubAreas,MandoubPhones,مواعيد_الحضور,الصفة,وضع_المندوب,comment,رقم_الجواز) values (@MandoubNames,@MandoubAreas,@MandoubPhones,@مواعيد_الحضور,@الصفة,@وضع_المندوب,@comment,@رقم_الجواز)";
            string addInfo = "تم تسجيل حساب المندوب/" + اسم_المندوب.Text + " بتاريخ " + GriDate + Environment.NewLine + "----------------------------------------------";
            if (بيانات_المندوب.Text == "تعديل")
                query = "update TableMandoudList set رقم_الجواز=@رقم_الجواز, MandoubNames=@MandoubNames,MandoubAreas=@MandoubAreas,MandoubPhones=@MandoubPhones,مواعيد_الحضور=@مواعيد_الحضور,الصفة=@الصفة,وضع_المندوب=@وضع_المندوب,comment=@comment where ID=@id";

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
                    sqlCmd.Parameters.AddWithValue("@رقم_الجواز", passport.Text);
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
                    sqlCmd.Parameters.AddWithValue("@رقم_الجواز", passport.Text);
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

        private void button1_Click_2(object sender, EventArgs e)
        {
            string Search = "";
            //MessageBox.Show(Employeename);
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.ShowDialog();
            try
            {
                Search = dlg.FileName;
            } catch (Exception ex) { Search = ""; }
            if (Search != "")
            {
                string query = "update TableUser set Data1=@Data1 where EmployeeName=@EmployeeName";
                
                SqlConnection sqlCon = new SqlConnection(DataSource);
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
                //try
                //{
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@EmployeeName", Empname);
                    using (Stream stream = File.OpenRead(Search))
                    {
                        byte[] buffer1 = new byte[stream.Length];
                        stream.Read(buffer1, 0, buffer1.Length);
                        sqlCmd.Parameters.AddWithValue("@Data1", buffer1);
                    }
                    //try
                    //{
                        sqlCmd.ExecuteNonQuery();
                    //}
                    //catch (Exception ex)
                    //{
                    //    MessageBox.Show("خطأ في البيانات");
                    //}
                //}
                //catch (Exception ex)
                //{
                //}
            }
        }

        private void btnActivate_Click(object sender, EventArgs e)
        {
            this.Size = new Size(967, 769);
            //btnActivate.Visible = true;
            fillDatagrid();
        }

        private void btnActivate_Click_1(object sender, EventArgs e)
        {
            this.Size = new Size(967, 769);
            //btnActivate.Visible = true;
            fillDatagrid();
            dataGridView1.Visible = true;
        }

        private void editPassword_Click(object sender, EventArgs e)
        {
            labelpass1.Visible = password1.Visible = password2.Visible = labelpass2.Visible = editPass.Visible = true;
            olspassword.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (olspassword.Text != userpass)
                    {
                MessageBox.Show("لم يتم إدخال كلمة المرور الحالية بشكل صحيح");
                return;
            }
            if (!passwordCheck())
                return;
            string addInfo = "قام " + UserName.Text + " بإعادة ضبط كلمة مرور حسابه" + GriDate + Environment.NewLine + "----------------------------------------------";
            items = new string[2] { "Pass", "comment" };
            values = new string[2] {  password2.Text, addInfo + التعليقات_السابقة_Off.Text };
            if (saveItems("TableUser", items, values, false))
                MessageBox.Show("تم إعادة الضبط بنجاح");
            else
                MessageBox.Show("تعذر إعادة الضبط لاسباب فنية");
            this.Close();
            return;
        }
    }
}
