using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.Threading;
using System.Net.NetworkInformation;

namespace PersAhwal
{
    //https://www.youtube.com/watch?v=MZ1UfBPoxvg&t=527s copy column from different database

    static class Program
    {
        
        static string EmployeeName = "مدير الصالة";
        static string JobPossition = "نائب قنصل";
        static string Modelfilespath, FileArchive, FormDataFile, newFiles;
        //CZssA58@9QdF
        //dataSource = "Data Source = (LocalDB)\\MSSQLLocalDB;Initial Catalog = myDataBase; Integrated Security = True";
        // Hall Computer 
        //dataSource = "Data Source=GOUNSLY\\SQLEXPRESS;Initial Catalog=AhwalDataBase;Integrated Security=True";
        //https://www.datree.io/resources/git-error-fatal-remote-origin-already-exists github
        //https://www.youtube.com/watch?v=dJ6c3OgIVDM git hub
        //static string dataSource56 = "Data Source=192.168.100.56,49170;Network Library=DBMSSOCN;Initial Catalog=AhwalDataBase;User ID=ConsJeddahAdmin;Password=DataBC0nsJ49170";
        //static string dataSource56 = "Data Source=192.168.100.58,49170;Network Library=DBMSSOCN;Initial Catalog=SudaneseAffairs;User ID=SADDB;Password=SADDB96325";
        //static string dataSource57 = "Data Source=192.168.100.57,49170;Network Library=DBMSSOCN;Initial Catalog=AhwalDataBase;User ID=sa;Password=1234";
        static string dataSource100A = "Data Source=192.168.100.100,49170;Network Library=DBMSSOCN;Initial Catalog=AhwalDataBase;User ID=GeneralUSerDB;Password=GeneralUSer_9632";
        static string dataSource100S = "Data Source=192.168.100.100,49170;Network Library=DBMSSOCN;Initial Catalog=SudaneseAffairs;User ID=GeneralUSerDB;Password=GeneralUSer_9632";
        //static string dataSource = "Data Source=192.168.100.100,49170;Initial Catalog=AhwalDataBase;User ID=Admin;Password=admin123";
        //static string dataSource = "Data Source=DESKTOP-KUPSRH8\\SQLEXPRESS;Initial Catalog=AhwalDataBase;Integrated Security=True";
        // Office computer 
        //DataSource = "Data Source=DESKTOP-H6E9AI2\\SQLEXPRESS;Initial Catalog=TestProj;Integrated Security=True";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        
        static void Main()
        {
            string filepathIn = Modelfilespath = @"D:\ModelFiles\";
            string filepathOut = @"D:\ArchiveFiles\";
            string PrimariFiles = @"D:\PrimariFiles";
            string archFile = @"D:\ArchiveFiles\";

            string appFileName = Environment.GetCommandLineArgs()[0];
            string directory = Path.GetDirectoryName(appFileName);
            directory = directory + @"\";
            bool source56 = false;
            bool source57 = false;
            //static string DataSource;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form11(-1, "",dataSource, filepathIn, filepathOut, EmployeeName, "نائب فنصل"));
            SqlConnection Con = new SqlConnection(dataSource100A);
            SqlCommand sqlCmd1 = new SqlCommand("select Modelfilespath,TempOutput,ServerName,Serverlogin,ServerPass,serverDatabase,FileArchive,FormDataFile,newFiles  from TableSettings where ID=@id", Con);
            sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
            if (Con.State == ConnectionState.Closed)
                try
                {
                    Con.Open();

                    var reader = sqlCmd1.ExecuteReader();

                    if (reader.Read())
                    {
                        source57 = true;
                        Modelfilespath = reader["Modelfilespath"].ToString();
                        //txtOutput.Text = reader["TempOutput"].ToString();
                        //txtServerIP.Text = reader["ServerName"].ToString();
                        //txtLogin.Text = reader["Serverlogin"].ToString();
                        //txtPass.Text = reader["ServerPass"].ToString();
                        //txtDatabase.Text = reader["serverDatabase"].ToString();
                        newFiles = FileArchive = reader["newFiles"].ToString();
                        archFile = FileArchive = reader["FileArchive"].ToString();
                        FormDataFile = reader["FormDataFile"].ToString();
                    }
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    Con.Close();
                }
            if (!source57)
            {
                Con = new SqlConnection(dataSource100S);
                sqlCmd1 = new SqlCommand("select Modelfilespath,TempOutput,ServerName,Serverlogin,ServerPass,serverDatabase,FileArchive,FormDataFile,newFiles  from TableSettings where ID=@id", Con);
                sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
                if (Con.State == ConnectionState.Closed)
                    try
                    {
                        Con.Open();

                        var reader = sqlCmd1.ExecuteReader();

                        if (reader.Read())
                        {
                            source56 = true;
                            Modelfilespath = reader["Modelfilespath"].ToString();
                            //txtOutput.Text = reader["TempOutput"].ToString();
                            //txtServerIP.Text = reader["ServerName"].ToString();
                            //txtLogin.Text = reader["Serverlogin"].ToString();
                            //txtPass.Text = reader["ServerPass"].ToString();
                            //txtDatabase.Text = reader["serverDatabase"].ToString();
                            newFiles = FileArchive = reader["newFiles"].ToString();
                            archFile = FileArchive = reader["FileArchive"].ToString();
                            FormDataFile = reader["FormDataFile"].ToString();
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
            if (Directory.Exists(@"D:\"))
            {
                if (!Directory.Exists(archFile))
                {
                    //string pathString = System.IO.Path.Combine(Modelfilespath, "NewFileData");
                    System.IO.Directory.CreateDirectory(archFile);
                }
                if (!Directory.Exists(@"D:\PrimariFiles"))
                {
                    System.IO.Directory.CreateDirectory(@"D:\PrimariFiles");
                }
                if (Directory.Exists(@"D:\PrimariFiles"))
                {
                    if (!Directory.Exists(@"D:\PrimariFiles\Personnel"))
                        System.IO.Directory.CreateDirectory(@"D:\PrimariFiles\Personnel");

                    if (!Directory.Exists(@"D:\PrimariFiles\SuddaneseAffairs"))
                        System.IO.Directory.CreateDirectory(@"D:\PrimariFiles\SuddaneseAffairs");
                }
            }
            else {
                archFile = directory + @"ArchiveFiles\";
                PrimariFiles =  directory+"PrimariFiles";
                if (!Directory.Exists(archFile))
                {                    
                    System.IO.Directory.CreateDirectory(archFile);
                }
                if (!Directory.Exists(directory + @"PrimariFiles\Personnel"))
                    System.IO.Directory.CreateDirectory(directory + @"PrimariFiles\Personnel");

                if (!Directory.Exists(directory + @"PrimariFiles\SuddaneseAffairs"))
                    System.IO.Directory.CreateDirectory(directory + @"PrimariFiles\SuddaneseAffairs");
            }
            //Application.Run(new AllConsArchInfo(dataSource100A, EmployeeName,  "17-01-2023", @"D:\PrimariFiles\ModelFiles\", archFile, "نائب قنصل"));
            //Application.Run(new DeepStatistics(dataSource100A, dataSource100S, Modelfilespath + @"\", archFile));
            //Application.Run(new Authentication(dataSource100A, "لبيب محمد أحمد",  archFile, "لبيب محمد أحمد", Modelfilespath, "16-06-1444", "01-08-2023"));

            //Application.Run(new Form4(0, -1, EmployeeName, dataSource57, Modelfilespath + @"\", archFile, JobPossition));

            //Application.Run(new NoteVerbal("29-06-2022","28-11-1443",JobPossition, dataSource56,  @"\\192.168.100.56\Users\Public\Documents\ModelFiles", archFile, EmployeeName, 1, false));
            //Application.Run(new MerriageDoc(dataSource100A,false, EmployeeName,2, "29-06-2022", "28-11-1443", @"D:\PrimariFiles\ModelFiles\", archFile));
            //Application.Run(new PassAway(2,dataSource100A,  @"D:\PrimariFiles\ModelFiles\", archFile, JobPossition,EmployeeName, "29-06-2022", "28-11-1443"));
            //Application.Run(new FormAuth(2, -1, "", dataSource100A, @"D:\PrimariFiles\ModelFiles\", archFile, EmployeeName, JobPossition, "05-14-2023", "10-23-1444", true));
            //Application.Run(new FormCollection(2, -1, 0, EmployeeName,dataSource100A, @"D:\PrimariFiles\ModelFiles\", archFile, JobPossition, "04-12-2023", "11-09-1444"));
            string Cdate = "08-02-2023";

            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            Cdate = DateTime.Now.ToString("MM-dd-yyyy");
            //MessageBox.Show(Cdate);
            //Application.Run(new MainForm("موظف ارشفة", 3, "57", "لبيب محمد أحمد", "نائب قنصل", dataSource100S, dataSource100A, @"D:\PrimariFiles\ModelFiles\", archFile, @"D:\PrimariFiles\FormData\", FormDataFile + @"\", true, Cdate, Modelfilespath, FormDataFile, false));
            //Application.Run(new Form8(dataSource100A, archFile));
            //string[] str = new string[1] { "" };
            //Application.Run(new FormPics("57", EmployeeName, "لبيب محمد أحمد", "نائب قنصل", dataSource100A, 0, FormDataFile, archFile, 10, str, str, false, str, str));
            //Application.Run(new Settings("57", false, dataSource100S, dataSource100A, false, Modelfilespath + @"\", archFile, archFile, FormDataFile + @"\", ""));
            //////Application.Run(new SignUp("جديد", "موظف محلي", dataSource100A, "احوال شخصية"));
            //Application.Run(new SignUp("جديد", "نائب قنصل", dataSource100A, "احوال شخصية","01-05-2023"));

            //if (source56)
            //{
            //    Console.WriteLine("server is 56");
            //    Application.Run(new FormDataBase("56", dataSource100S, dataSource100A, Modelfilespath + @"\", archFile, archFile, FormDataFile + @"\", newFiles));
            //}
            //else if (source57)
            //{
            //    Console.WriteLine(dataSource100A);
            //Application.Run(new FormDataBase("57", dataSource100S, dataSource100A, Modelfilespath + @"\", archFile, archFile, FormDataFile + @"\", newFiles));
            Application.Run(new Accountant(dataSource100A, Cdate, "محمد علي محمد", "مدير"));
            //}
        }
    }
}
