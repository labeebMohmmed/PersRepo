using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PersAhwal
{

    static class Program
    {
        static string dataSource;
        static string EmployeeName = "مدير الصالة";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]

        static void Main()
        {
            string filepathIn = @"D:\ModelFiles\";
            string filepathOut = @"D:\AhwalIqrar\";
            //static string DataSource;
            //dataSource = "Data Source = (LocalDB)\\MSSQLLocalDB;Initial Catalog = myDataBase; Integrated Security = True";
            // Hall Computer 
            //dataSource = "Data Source=GOUNSLY\\SQLEXPRESS;Initial Catalog=AhwalDataBase;Integrated Security=True";

            dataSource = "Data Source=192.168.100.123,49170;Network Library=DBMSSOCN;Initial Catalog=AhwalDataBase;User ID=ConsJeddah;Password=DBaseC@nsJ0d103";
            // Office computer 
            //DataSource = "Data Source=DESKTOP-H6E9AI2\\SQLEXPRESS;Initial Catalog=TestProj;Integrated Security=True";
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new FormDataBase(DataSource));

            Application.Run(new Form11());
            
            //Application.Run( new Tawkeel1(EmployeeName, dataSource, filepathIn, filepathOut));
            //Application.Run(new MainForm(EmployeeName,"نائب قنصل", dataSource, filepath));
        }
    }
}
