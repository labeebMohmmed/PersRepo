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

namespace PersAhwal
{
    public partial class Settings : Form
    {
        private string DataSource, FilepathIn, FilepathOut;
        private static bool NewSettings = false;
        public Settings( bool newSettings,string dataSource, bool setDataBase,string filepathIn, string filepathOut)
        {
            InitializeComponent();
            NewSettings = newSettings;
            DataSource = dataSource;
            FilepathIn = filepathIn;
            FilepathOut = filepathOut;
            if (!setDataBase)
            {
                if (!newSettings) loadSettings();
                else
                {
                    SaveSettings.Text = "أدخل بيانات قاعدة بيانات صحيحة";
                    MessageBox.Show("لا توجد قاعدة بيانات مسجلة");
                }
            }
        }

        private void loadSettings()
        {
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select Modelfilespath,TempOutput,ServerName,Serverlogin,ServerPass,serverDatabase  from TableSettings where ID=@id", Con);
            sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
            if (Con.State == ConnectionState.Closed)
                try
                {
                    Con.Open();

                    var reader = sqlCmd1.ExecuteReader();

                    if (reader.Read())
                    {
                        NewSettings = true;
                        txtModel.Text = reader["Modelfilespath"].ToString();
                        txtOutput.Text = reader["TempOutput"].ToString();
                        txtServerIP.Text = reader["ServerName"].ToString();
                        txtLogin.Text = reader["Serverlogin"].ToString();
                        txtPass.Text = reader["ServerPass"].ToString();
                        txtDatabase.Text = reader["serverDatabase"].ToString();
                        if (NewSettings)
                        {
                            SaveSettings.Text = "تعديل";
                            NewSettings = false;
                        }                        
                    }
                }
                catch (Exception ex)
                {
                    
                }
                finally
                {
                    Con.Close();
                }

        }
        

        private void SaveSettings_Click(object sender, EventArgs e)
        {
            if (NewSettings)
            {
                DataSource = "Data Source=" + txtServerIP.Text + ";Network Library=DBMSSOCN;Initial Catalog=" + txtDatabase.Text + ";User ID=" + txtLogin.Text + ";Password=" + txtPass.Text;                
                FilepathIn = txtModel.Text;
                FilepathOut = txtOutput.Text;
            }
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    SqlCommand sqlCmd = new SqlCommand("SettingsAddorEdit", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    if (SaveSettings.Text == "حفظ")
                    {
                        sqlCmd.Parameters.AddWithValue("@ID", 1);
                        sqlCmd.Parameters.AddWithValue("@mode", "Add");
                        sqlCmd.Parameters.AddWithValue("@Modelfilespath", txtModel.Text);
                        sqlCmd.Parameters.AddWithValue("@TempOutput", txtOutput.Text);
                        sqlCmd.Parameters.AddWithValue("@ServerName", txtServerIP.Text);
                        sqlCmd.Parameters.AddWithValue("@Serverlogin", txtLogin.Text);
                        sqlCmd.Parameters.AddWithValue("@ServerPass", txtPass.Text);
                        sqlCmd.Parameters.AddWithValue("@serverDatabase", txtDatabase.Text);
                        sqlCmd.ExecuteNonQuery();
                    }
                    else
                    {
                        sqlCmd.Parameters.AddWithValue("@ID", 1);
                        sqlCmd.Parameters.AddWithValue("@mode", "Edit");
                        sqlCmd.Parameters.AddWithValue("@Modelfilespath", txtModel.Text);
                        sqlCmd.Parameters.AddWithValue("@TempOutput", txtOutput.Text);
                        sqlCmd.Parameters.AddWithValue("@ServerName", txtServerIP.Text);
                        sqlCmd.Parameters.AddWithValue("@Serverlogin", txtLogin.Text);
                        sqlCmd.Parameters.AddWithValue("@ServerPass", txtPass.Text);
                        sqlCmd.Parameters.AddWithValue("@serverDatabase", txtDatabase.Text);
                        sqlCmd.ExecuteNonQuery();
                    }
                    sqlCon.Close();
                    this.Hide();
                    var formDataBase = new FormDataBase(DataSource, FilepathIn, FilepathOut);
                    formDataBase.Closed += (s, args) => this.Close();
                    formDataBase.Show();
                }

                catch (Exception ex)
                {
                    MessageBox.Show("الوصول لقاعدة البيانات غير متاح");                    
                }
                finally
                {
                    
                    clear_fields();
                    
                }
        }

        private void clear_fields()
        {
            txtDatabase.Text = txtLogin.Text = txtModel.Text = txtOutput.Text = txtPass.Text = txtServerIP.Text = "";
        }
    }
}
