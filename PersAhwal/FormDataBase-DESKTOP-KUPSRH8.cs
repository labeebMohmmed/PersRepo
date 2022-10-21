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

namespace PersAhwal
{
    //https://www.youtube.com/watch?v=-2UcDV4uUu8
    public partial class FormDataBase : Form
    {
                    
        static string DataSource, Employee, FilespathIn, FilepathOut;

       // MainForm mainForm;

        public FormDataBase(string dataSource, string filepathIn, string filepathOut)
        {
            InitializeComponent();
            DataSource = dataSource;
            FilespathIn = filepathIn;
            FilepathOut = filepathOut;
            fillDatagrid();
        }

        private void fillDatagrid()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    string settingData = "select * from TableUser";
                    SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
                    sqlDa.SelectCommand.CommandType = CommandType.Text;
                    DataTable dtbl = new DataTable();
                    sqlDa.Fill(dtbl);
                    dataGrid.DataSource = dtbl;
                    greenCircle.Visible = true;
                }
                catch (Exception ex)
                {
                    greenCircle.Visible = false;
                    redCircle.Visible = true;

                    this.Hide();
                    var settings = new Settings(true, DataSource, true, FilespathIn, FilepathOut);
                    settings.Closed += (s, args) => this.Close();
                    settings.Show();                   
                }
                finally
                {
                    sqlCon.Close();
                }            
            
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            bool founded = false ;
            int x;
            if (username.Text == "")
                MessageBox.Show("أدخل اسم مستخدم صحيح أولا");
            else
            {                
                for (int Xindex = 0; Xindex < dataGrid.Rows.Count; Xindex++) 
                {
                    if (dataGrid.Rows[Xindex].Cells[4].Value.ToString() == username.Text.Trim())
                    {
                        founded = true;
                        if (dataGrid.Rows[Xindex].Cells[7].Value.ToString() != "غير مؤكد")
                        {
                            if (Password.Text == dataGrid.Rows[Xindex].Cells[6].Value.ToString())
                            {
                                string joposition = dataGrid.Rows[Xindex].Cells[2].Value.ToString();
                                Employee = username.Text;
                                username.Clear();
                                Password.Clear();
                                this.Hide();
                                var mainForm = new MainForm(dataGrid.Rows[Xindex].Cells[1].Value.ToString(), joposition, DataSource, FilespathIn, FilepathOut);
                                mainForm.Closed += (s, args) => this.Close();
                                mainForm.Show();
                            }
                            else MessageBox.Show("خطأ في اسم الموظف أو كلمة مرور");
                        }
                        else MessageBox.Show("حساب المستخدم غير مفعل");
                        break;
                    }
                }
                if (!founded) MessageBox.Show("خطأ في اسم الموظف أو كلمة مرور");
            }
            
        }

        private void employeeName_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void employeeName_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SignUp signUp = new SignUp("جديد", "غير محدد",DataSource);
            signUp.Show();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.Hide();
            var settings = new Settings(false, DataSource, false, FilespathIn, FilepathOut);
            settings.Closed += (s, args) => this.Close();
            settings.Show();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {

        }

        private void Password_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                btnLog.PerformClick();
        }

        private void SeePass_CheckedChanged(object sender, EventArgs e)
        {
            if (SeePass.CheckState == CheckState.Checked) {
                Password.UseSystemPasswordChar = false;
            }
        }


   
    }
    }
