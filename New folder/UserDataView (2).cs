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
        public Delegate DataMovePage;
        string dataSource = "Data Source = (LocalDB)\\MSSQLLocalDB;Initial Catalog = myDataBase; Integrated Security = True";
        public UserDataView()
        {
            InitializeComponent();
            FillDataGridView(dataSource);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataMovePage.DynamicInvoke(4);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataMovePage.DynamicInvoke(2);
        }
        private void FillDataGridView(String DataSource)
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
            sqlCon.Close();
        }
    }
}
