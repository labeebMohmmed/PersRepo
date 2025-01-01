using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Globalization;
using System.Configuration;
using System.Reflection;
using System.Threading;
using System.IO;
using System.Net.NetworkInformation;

namespace PersAhwal
{
    //https://www.youtube.com/watch?v=-2UcDV4uUu8
    public partial class FormDataBase : Form
    {

        static string DataSource56, DataSource57, Employee, LocalModelFiles,localModelForms, FilepathOut, ArchFile, ServerModelForms;

        // MainForm mainForm;
        string fileVersio;
        string GregorianDate = "";
        Thread th;
        bool inProcess = false;
        string ServerModelFiles ;
        string formsFile ;
        string primeryLink = "";
        string Server = "M";
        string DataSource;
        bool readyToClose = false;
        
        string file; 
        string IP;
        DataTable userTable;
        string currentVersion = "0.0.0.0.O";
        string currentStatus = "done";
        string cVersion56 = "";
            int currentVersion56 = 0;
        int currentVersion57 = 0;
        string NewFiles = "";
        string workingDisk = "";
        string MacAdd = "";

        string mainWorkingFile = @"E:\";
        

        public FormDataBase(string server,string dataSource56, string dataSource57, string modelFiles, string filepathOut, string archFile, string modelForms, string newFiles)
        {
            InitializeComponent();
            getDate();
            string hostname = Dns.GetHostName();
            Console.WriteLine(IP);
            IP = Dns.GetHostByName(hostname).AddressList[0].ToString();
            MacAdd =
                (
                from nic in NetworkInterface.GetAllNetworkInterfaces()
                where nic.OperationalStatus == OperationalStatus.Up
                select nic.GetPhysicalAddress().ToString()
                ).FirstOrDefault();
            
            NewFiles = newFiles;
            DataSource56 = dataSource56;
            DataSource57 = dataSource57;
            DataSource = DataSource57;
            //MessageBox.Show(IP);
            
            //MessageBox.Show("1");
            FilepathOut = filepathOut.Replace("D", BackupDisck("workingDisk"));
            ServerModelFiles = modelFiles.Replace("D", BackupDisck("workingDisk"));
            ServerModelForms = modelForms.Replace("D", BackupDisck("workingDisk"));
            mainWorkingFile = BackupDisck("archDisk") + @":\";
            //MessageBox.Show(mainWorkingFile);
            workingDisk = BackupDisck("workingDisk") + @":\";

            if (!Directory.Exists(archFile + @"\formUpdated"))
            {
                System.IO.Directory.CreateDirectory(archFile + @"\formUpdated");
            }
            if (IP == "192.168.100.100")
            {
                if (!Directory.Exists(mainWorkingFile + "sqlbackup"))
                {
                    Console.WriteLine(mainWorkingFile + "sqlbackup");
                    System.IO.Directory.CreateDirectory(mainWorkingFile + "sqlbackup");
                }
            }
           

            if (Directory.Exists(workingDisk))
            {
                primeryLink = workingDisk + @"PrimariFiles\";
               
            }
            else
            {
                string appFileName = Environment.GetCommandLineArgs()[0];
                string directory = Path.GetDirectoryName(appFileName);
                directory = directory + @"\";                
                primeryLink = directory + @"PrimariFiles\";                
            }
            
            if (Directory.Exists(primeryLink + "New folder"))
            { try{
                    Directory.Delete(primeryLink + "New folder");
                }catch (Exception ex) {
                    MessageBox.Show("يوجد ملفات");

                }
            }

            if (!Directory.Exists(primeryLink + "ModelFiles"))
            {
                System.IO.Directory.CreateDirectory(primeryLink + "ModelFiles");
            }
            
            if (!Directory.Exists(primeryLink + "FormData"))
            {
                System.IO.Directory.CreateDirectory(primeryLink + "FormData");
            }
            LocalModelFiles = primeryLink + @"ModelFiles\";
            localModelForms = primeryLink + @"FormData\";

            Server = server;
            
            cVersion56 = getVersio(DataSource56);
           
             if (!File.Exists(primeryLink + @"\Personnel\getVersio.txt"))
            {
                dataSourceWrite(primeryLink + @"\Personnel\getVersio.txt", getVersio(DataSource57));
            }
            if (!File.Exists(primeryLink + @"\SuddaneseAffairs\getVersio.txt"))
            {
                dataSourceWrite(primeryLink + @"\SuddaneseAffairs\getVersio.txt", getVersio(DataSource56));
            }

            if (!File.Exists(primeryLink + @"\updatingStatus.txt"))
                dataSourceWrite(primeryLink + @"\updatingStatus.txt", "done");

            else

            dataSourceWrite(primeryLink + @"\updatingStatus.txt", "Allowed");
            
            if (!File.Exists(primeryLink + @"\updatingSetup.txt"))
                dataSourceWrite(primeryLink + @"\updatingSetup.txt", "done");

            else

            dataSourceWrite(primeryLink + @"\updatingSetup.txt", "Allowed");

            
            currentVersion56 = Convert.ToInt32(cVersion56.Split('.')[3]);
            Console.WriteLine("currentVersion56 " + currentVersion56);
            

            versionUpdateInfo("SuddaneseAffairs");
            
                //MessageBox.Show(IP);
                if (Server == "57")
            {
                btnLog.BackColor = System.Drawing.SystemColors.ButtonShadow;
                button1.BackColor = System.Drawing.SystemColors.ButtonShadow;
            }
            else DataSource = DataSource56;
            userTable = new DataTable();

            //if (IP == "192.168.100.100")
            //    backup(mainWorkingFile);
            //else if (IP == "192.168.100.67")
            //    backup(BackupDisck("backupFile") + @":\");

            Console.WriteLine(Server);
            file = archFile + @"\dataSource.txt";
            
            if (!green57.Visible)
            {
                fillDatagrid(userTable, DataSource57, green57, red57, "الاتصال مع مخدم قسمي الاحوال الشخصية وشؤون الرعايا غير متاح يرجى التواصل مع مشغل النظام");
                fillDatagrid(userTable, DataSource56, green57, red57, "الاتصال مع مخدم قسمي الاحوال الشخصية وشؤون الرعايا غير متاح يرجى التواصل مع مشغل النظام");
            }

            username.Select();
            
            


        }

        
        
        private string getRenew()
        {
            string query = "";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            try
            {
                query = "select  renew from TableSettings";
            }
            catch (Exception ex) { return ""; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
                sqlCon.Close();

                foreach (DataRow row in dtbl.Rows)
                {
                    return row["renew"].ToString();                    
                }
            }
            catch (Exception ex) {
                return ""; }
            return "";

        }
        private string BackupDisck(string colName)
        {
            string infoDet = "";
            string query = "select "+ colName+" from TableSettings";
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return ""; }
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            try
            {
                sqlDa.Fill(dtbl);
            }
            catch (Exception ex)
            {
                return "";
            }



            sqlCon.Close();

            foreach (DataRow dataRow in dtbl.Rows)
            {
                try
                {
                    infoDet = dataRow[colName].ToString();
                }
                catch (Exception ex)
                {

                }
            }
            return infoDet;
        }

        public static int MaxDate(int[] dates)
        {
            int max = 0;
            for (int y = 0; y < dates.Length; y++)
            {
                if (max < dates[y])
                    max = dates[y];
            }
            return max;
        }

        private void backup() {
            string sqlInfoBackup = "AhwalDataBase_backup_";
            string sqlArchBackup = "ArchFilesDB_backup_";
            string sqlAffairBackup = "SudaneseAffairs_backup_";
            //Console.WriteLine(workFile + "sqlbackup");
            
            string year = "";
            string month = "";
            string date = "";
            string defaultArchFile = "E:\\SQL Server\\MSSQL16.MSSQLSERVER\\MSSQL\\Backup";
            
            //string[] backfiles = Directory.GetFiles(workFile + @"sqlbackup\");
            string[] backfiles = Directory.GetFiles(defaultArchFile); 
            int[] datesInfo = new int[backfiles.Length];
            int[] datesArch = new int[backfiles.Length];
            int[] datesAffa = new int[backfiles.Length];

            for (int i = 0; i < backfiles.Length; i++)
            {
                if (backfiles[i].Contains(sqlInfoBackup))
                {
                    var backfileinfo = new FileInfo(backfiles[i]);
                    year = backfiles[i].Split('_')[2];
                    month = backfiles[i].Split('_')[3];
                    if (month.Length == 1)
                        month = "0" + month;
                    date = backfiles[i].Split('_')[4];
                    if (date.Length == 1)
                        date = "0" + date;
                    datesInfo[i] = Convert.ToInt32(year + month + date);
                    //Console.WriteLine(backfiles[i]);
                    //Console.WriteLine("AhwalDataBase_backup_" + datesInfo[i]);
                }
                else if (backfiles[i].Contains(sqlArchBackup))
                {
                    var backfileinfo = new FileInfo(backfiles[i]);
                    year = backfiles[i].Split('_')[2];
                    month = backfiles[i].Split('_')[3];
                    if (month.Length == 1)
                        month = "0" + month;
                    date = backfiles[i].Split('_')[4];
                    if (date.Length == 1)
                        date = "0" + date;
                    datesArch[i] = Convert.ToInt32(year + month + date);
                    //Console.WriteLine(backfiles[i]); 
                    //Console.WriteLine("ArchFilesDB_backup_" + datesArch[i]);
                }
                else if (backfiles[i].Contains(sqlAffairBackup))
                {
                    var backfileinfo = new FileInfo(backfiles[i]);
                    year = backfiles[i].Split('_')[2];
                    month = backfiles[i].Split('_')[3];
                    if (month.Length == 1)
                        month = "0" + month;
                    date = backfiles[i].Split('_')[4];
                    if (date.Length == 1)
                        date = "0" + date;
                    datesAffa[i] = Convert.ToInt32(year + month + date);
                    //Console.WriteLine(backfiles[i]); 
                    //Console.WriteLine("SudaneseAffairs_" + datesAffa[i]);
                }
            }
            string maxdate = MaxDate(datesInfo).ToString().Split(' ')[0];
            year = SpecificDigit(maxdate,1,4);
            month = SpecificDigit(maxdate, 5, 6);
            date = SpecificDigit(maxdate, 7, 8);
            
            string newerFile = sqlInfoBackup + year + "_" + month + "_" + date + "_";
            //Console.WriteLine(newerFile);
            for (int i = 0; i < backfiles.Length; i++)
            {
                if (backfiles[i].Contains(sqlInfoBackup) && !backfiles[i].Contains(newerFile))
                {
                    Console.WriteLine("Files to Delete " + backfiles[i]);
                    try
                    {
                       File.Delete(backfiles[i]);
                    }
                    catch (Exception ex) { }
                }
            }
            
            maxdate = MaxDate(datesArch).ToString().Split(' ')[0];
            year = SpecificDigit(maxdate, 1, 4);
            month = SpecificDigit(maxdate, 5, 6);
            date = SpecificDigit(maxdate, 7, 8);

            newerFile = sqlArchBackup + year + "_" + month + "_" + date + "_";
            //Console.WriteLine(newerFile);
            for (int i = 0; i < backfiles.Length; i++)
            {
                if (backfiles[i].Contains(sqlArchBackup) && !backfiles[i].Contains(newerFile))
                {
                    Console.WriteLine("Files to Delete " + backfiles[i]);
                    try
                    {
                        File.Delete(backfiles[i]);
                    }
                    catch (Exception ex) { }
                }
            }
            
            maxdate = MaxDate(datesAffa).ToString().Split(' ')[0];
            year = SpecificDigit(maxdate, 1, 4);
            month = SpecificDigit(maxdate, 5, 6);
            date = SpecificDigit(maxdate, 7, 8);

            newerFile = sqlAffairBackup + year + "_" + month + "_" + date + "_";
            Console.WriteLine(newerFile);
            for (int i = 0; i < backfiles.Length; i++)
            {
                if (backfiles[i].Contains(sqlAffairBackup) && !backfiles[i].Contains(newerFile))
                {
                    Console.WriteLine("Files to Delete " + backfiles[i]);
                    //insertFileToDelete(backfiles[i], "AutoArchRecord");
                    try
                    {
                       File.Delete(backfiles[i]);
                    }
                    catch (Exception ex) { }
                }
            }
        }

        private void insertFileToDelete(string name, string table)
        {
            if(checkFiles(name, table)) return;
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }

            SqlCommand sqlCmd = new SqlCommand("INSERT INTO "+ table + " (Files) values (@Files)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@Files", name);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private bool checkFiles(string name, string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return false; }
            string settingData = "select * from " + table +" where Files = N'" + name + "'";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            if (dtbl.Rows.Count > 0)
                    return true;
            else 
                return false;
        }

        private string SpecificDigit(string text, int Firstdigits, int Lastdigits)
        {
            char[] characters = text.ToCharArray();
            string firstNchar = "";
            int z = 0;
            for (int x = Firstdigits - 1; x < Lastdigits && x < text.Length; x++)
            {
                firstNchar = firstNchar + characters[x];

            }
            return firstNchar;
        }

        public static void Copy(string sourceDirectory, string targetDirectory)
        {
            var diSource = new DirectoryInfo(sourceDirectory);
            var diTarget = new DirectoryInfo(targetDirectory);

            CopyAll(diSource, diTarget);
            //MessageBox.Show("تم نسخ الملفات الأولية");
        }
        public static void CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            Directory.CreateDirectory(target.FullName);

            // Copy each file into the new directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
                fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir =
                    target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir);
            }
        }
        private bool versionUpdateInfo(string text)
        {
            bool procede = true;
            
            fileVersio = primeryLink + text + @"\getVersio.txt";
            currentVersion = File.ReadAllText(fileVersio);
            Console.WriteLine("currentVersion " + currentVersion);
            if (!File.Exists(fileVersio))
            {
                dataSourceWrite(fileVersio, getVersio(DataSource56));
            }
            else
            {
                
                appversion.Text = getVersio(DataSource56);
                
                Console.WriteLine(currentVersion.Split('.')[3] + "--" + currentVersion56 + "--" + currentVersion.Split('.')[3] + "--" + currentVersion57  );
                if (currentVersion56 > Convert.ToInt32(currentVersion.Split('.')[3]))
                {
                    procede = false;
                    try
                    {
                        upDateClose();
                    }
                    catch (Exception exp)
                    {
                        
                    }
                }
            }
            return procede; 
        }

        private void closeToUpdate(string text)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("update TableSettings set closeToUpdate=@closeToUpdate where ID=@id", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@id", 1);
            sqlCmd.Parameters.AddWithValue("@closeToUpdate", text);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void upDateClose()
        {
            
            try
            {
                File.Delete(primeryLink + "fileUpdate.txt");
                System.Diagnostics.Process.Start(getAppFolder() + @"\setup.exe");
                File.Delete(fileVersio);
                File.Delete(primeryLink + @"\files.txt");
                VersionUpdate(cVersion56);
                dataSourceWrite(fileVersio = primeryLink + @"SuddaneseAffairs\getVersio.txt", cVersion56.ToString());
                dataSourceWrite(primeryLink + @"\updatingSetup.txt", "updating");
                this.Close();
            }
            catch (Exception e)
            {
                
            }
            
            
        }
        private void VersionUpdate(string version)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource56);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return ; }
            SqlCommand sqlCmd = new SqlCommand("update TableSettings set Version=@Version where ID='1'", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@Version", version);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }
        public bool fileIsOpen(string path)
        {
            System.IO.FileStream a = null;

            try
            {
                a = System.IO.File.Open(path,
                System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None);
                return false;
            }
            catch (System.IO.IOException ex)
            {
                return true;
            }

            finally
            {
                if (a != null)
                {
                    a.Close();
                    a.Dispose();
                }
            }
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

        private void fillDatagrid(DataTable dtbl,string DataSource, PictureBox green,PictureBox red, string text)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return; }
                    string settingData = "select * from TableUser";
                    SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
                    sqlDa.SelectCommand.CommandType = CommandType.Text;
                    sqlDa.Fill(dtbl);
                    dataGrid.DataSource = dtbl;
                    green.Visible = true;
                    red.Visible = false;
                }
                catch (Exception ex)
                {
                    green.Visible = false;
                    red.Visible = true;
                                       
                }
                finally
                {
                    sqlCon.Close();
                }

        }
        
        private bool checkRe(string mac)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)

                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return false; }
            string settingData = "select state from TableMac where macAdd = N'" + mac + "'";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            foreach (DataRow dataRow in dtbl.Rows)
            {
                if (dataRow["state"].ToString() == "authentic")
                    return true;
            }


            return false;
        }
        
        private string getAppFolder()
        {
            string DataSource = DataSource56;
            if (Server == "U")
                DataSource = DataSource57;

            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return ""; }
            string settingData = "select FolderApp from TableSettings where ID='1'";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            string ver = "";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                ver = dataRow["FolderApp"].ToString();

            }
            return ver;
        }

        

        
        private string getPassRest(string text, string dataSource)
        {            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return ""; }
            string settingData = "select RestPAss from TableUser where UserName=@UserName";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            sqlDa.SelectCommand.Parameters.AddWithValue("@UserName", text);            
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            string pass = "";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                pass = dataRow["RestPAss"].ToString();

            }
            return pass;
        }

        private string getVersio(string dataSource)
        {
           
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return ""; }
            string settingData = "select Version from TableSettings where ID='1'";
            SqlDataAdapter sqlDa = new SqlDataAdapter(settingData, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            string ver =  "1.0.0.0";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                ver = dataRow["Version"].ToString();

            }
            return ver;
        }
        private void getModelOutFiles(string dataSource)
        {

            SqlConnection Con = new SqlConnection(dataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select Modelfilespath,TempOutput,ServerName,Serverlogin,ServerPass,serverDatabase,FileArchive,FormDataFile  from TableSettings where ID=@id", Con);
            sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
            if (Con.State == ConnectionState.Closed)
                try
                {
                    try
                    {
                        Con.Open();
                    }
                    catch (Exception ex) { return; }

                    var reader = sqlCmd1.ExecuteReader();
                    if (reader.Read())
                    {
                        //FilespathIn = reader["Modelfilespath"].ToString();
                        //txtOutput.Text = reader["TempOutput"].ToString();
                        //txtServerIP.Text = reader["ServerName"].ToString();
                        //txtLogin.Text = reader["Serverlogin"].ToString();
                        //txtPass.Text = reader["ServerPass"].ToString();
                        //txtDatabase.Text = reader["serverDatabase"].ToString();
                        FilepathOut = ArchFile = reader["FileArchive"].ToString();
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


        private const string AssemblyName = "MyAssembly"; // Name of your assembly

       
        private void btnSave_Click(object sender, EventArgs e)
        {
            //string autoGenPass = 
            //if (checkRenew() && ) { 

            //}
            bool Procede = false;
            bool foundedUser = false;

            string userpass = "";
            string userApro = "";
            string name = "";
            int idNo = 0;
            string division = "";
            bool Pers_Peope = true;
            string joposition = "";
            string career = "";
            int x;
            if (username.Text == "")
                MessageBox.Show("أدخل اسم مستخدم صحيح أولا");
            else
            {
                for (int Xindex = 0; Xindex < dataGrid.Rows.Count - 1; Xindex++)
                {
                    try
                    {

                        if (dataGrid.Rows[Xindex].Cells["UserName"].Value.ToString().Trim() == username.Text.Trim())
                        {

                            idNo = Convert.ToInt32(dataGrid.Rows[Xindex].Cells["ID"].Value.ToString());
                            userpass = dataGrid.Rows[Xindex].Cells["Pass"].Value.ToString();
                            userApro = dataGrid.Rows[Xindex].Cells["Aproved"].Value.ToString();
                            //MessageBox.Show(userApro);
                            joposition = dataGrid.Rows[Xindex].Cells["JobPosition"].Value.ToString();
                            name = dataGrid.Rows[Xindex].Cells["EmployeeName"].Value.ToString();
                            division = dataGrid.Rows[Xindex].Cells["Purpose"].Value.ToString();
                            career = dataGrid.Rows[Xindex].Cells["Career"].Value.ToString();
                            foundedUser = true;
                            if (division == "شؤون رعايا")
                            {
                                Server = "56";
                                DataSource = DataSource56;
                                Pers_Peope = false;
                            }
                            else if (division == "احوال شخصية" || division == "الأرشفة العامة" || division == "محاسب")
                            {
                                Server = "57";
                                DataSource = DataSource57;
                                Pers_Peope = true;
                            }
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(dataGrid.RowCount.ToString());
                        Console.WriteLine(Xindex.ToString());
                        Console.WriteLine(dataGrid.Rows[Xindex].Cells[4].Value.ToString());

                    }

                }
                if (!foundedUser) 
                { 
                    MessageBox.Show("اسم المستخدم غير معرف في النظام");
                    return; 
                }
                else
                {
                    //MessageBox.Show("versionUpdateInfo");
                    Procede = versionUpdateInfo("SuddaneseAffairs");
                }
                if (userApro == "" || userApro.Contains("غير")) 
                {
                    //MessageBox.Show(userApro);
                    MessageBox.Show("حساب المستخدم غير مفعل أو غير معرف"); return; 
                }

                if (Password.Text == userpass)
                {
                    Employee = username.Text;
                    if (File.Exists(file) && !fileIsOpen(file))
                        File.Delete(file);
                    Console.WriteLine("pass login3");
                    Procede = true;
                    if (getPassRest(username.Text, DataSource) != "done" )
                    {
                        Console.WriteLine("pass login7");
                        if (!inProcess) 
                            MessageBox.Show(" يرجى إعادة تعيين كلمة المرور عبر الضغط على زر إعادة تعيين كلمة المرور");
                        inProcess = true;
                        if (!pass1.Visible)
                        {
                            Console.WriteLine("pass login6");
                            btnLog.Location = new System.Drawing.Point(486, 250);
                            button1.Location = new System.Drawing.Point(357, 250);
                            pass1.Visible = true;
                            pass2.Visible = true;
                            labelpass1.Visible = true;
                            labelpass2.Visible = true;
                            return;

                        }
                        else {
                            Console.WriteLine("pass login4");
                            if (pass1.Text == Password.Text)
                            {
                                MessageBox.Show("كلمة المرور الجديدة لا يمكن أن تطابق الكلمة السابقة");
                                return;
                            }
                            if (pass1.Text != pass2.Text)
                            {
                                MessageBox.Show("كلمة المرور غير متطابقة");
                                return;
                            }
                            if (pass1.Text.Length < 6)
                            {
                                MessageBox.Show("كلمة المرور يجب أن لا تقل عن ستة رموز");
                                return;
                            }
                            if (pass1.Text.All(char.IsDigit))
                            {
                                MessageBox.Show("كلمة المرور يجب أن تحتوي على أحرف");
                                return;
                            }
                            resetPass(idNo, pass1.Text);
                            btnLog.Location = new System.Drawing.Point(486, 168);
                            button1.Location = new System.Drawing.Point(357, 168);
                            pass1.Visible = false;
                            pass2.Visible = false;
                            labelpass1.Visible = false;
                            labelpass2.Visible = false;
                            inProcess = false;
                            Procede = true;
                            Console.WriteLine("pass login5");
                        }

                        //SignUp signUp = new SignUp(name, "موظف محلي", DataSource);
                        //signUp.Show();
                        //return;
                    }
                    if (Procede)
                    {
                        dataSourceWrite(file, DataSource);
                        getModelOutFiles(DataSource);
                        Password.Clear();
                        int userID = userLogInfo(name, IP, MacAdd);
                        if (Server == "57") 
                            btnLog.Enabled = checkRe(MacAdd);
                        //MessageBox.Show(MacAdd);
                        if (!btnLog.Enabled)
                        {
                            MessageBox.Show("الجهاز بالرقم( "+ MacAdd + " )غير معرف بالنظام، لا يمكن المتابعة");
                            insertUnauthenticDev(name, IP, MacAdd);
                            this.Close();
                        }
                        if (division == "الأرشفة العامة")
                        {
                            AllConsArchInfo allConsArchInfo = new AllConsArchInfo(DataSource57, name, GregorianDate, LocalModelFiles, ArchFile, joposition);
                            allConsArchInfo.ShowDialog();
                        }
                        if (division == "محاسب")
                        {
                            Accountant accountant = new Accountant(DataSource57, GregorianDate, name, career);
                            accountant.ShowDialog();
                        }
                        else
                        {
                            

                            MainForm mainForm = new MainForm(career, userID, Server, name, joposition, DataSource56, DataSource57, LocalModelFiles, FilepathOut, ArchFile, localModelForms, Pers_Peope, GregorianDate,  ServerModelFiles, ServerModelForms, true);
                            mainForm.ShowDialog();
                        }
                        timer1.Enabled = false;
                        Console.WriteLine("pass login1");
                    }
                }
                else MessageBox.Show("خطأ في كلمة مرور");
            }
            Console.WriteLine("pass login2");
        }
        private void resetPass(int id,string pass)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }
            SqlCommand sqlCmd = new SqlCommand("UPDATE TableUser SET Pass = @Pass,RestPAss=@RestPAss WHERE ID = @ID", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.Parameters.AddWithValue("@Pass", pass);
            sqlCmd.Parameters.AddWithValue("@RestPAss", "done");
            sqlCmd.ExecuteNonQuery();
        }

        private int loadIDNo(string table)
        {


            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return 0; }
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT top(1) ID from TableUserLog order by ID desc", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string rowCnt = "0";
            foreach (DataRow row in dtbl.Rows)
            {
                rowCnt = row["ID"].ToString();
            }

            return Convert.ToInt32(rowCnt);

        }
        private int userLogInfo(string name, string ip, string macAdd)
        {
            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return 0; }

            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableUserLog (userName,timeDateIn,timeDateOut,pcIP,macAdd) values (@userName,@timeDateIn,@timeDateOut,@pcIP,@macAdd)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@userName", name);
            sqlCmd.Parameters.AddWithValue("@timeDateIn", DateTime.Now.ToString("G"));
            sqlCmd.Parameters.AddWithValue("@timeDateOut", DateTime.Now.ToString("G"));
            sqlCmd.Parameters.AddWithValue("@pcIP", ip);
            sqlCmd.Parameters.AddWithValue("@macAdd", macAdd);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();

            return loadIDNo("TableUserLog");
        }
        
        private void insertUnauthenticDev(string name, string ip, string macAdd)
        {
            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return; }

            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableMac (macAdd,ipAdd,userName) values (@macAdd,@ipAdd,@userName)", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@userName", name);
            sqlCmd.Parameters.AddWithValue("@ipAdd", ip);
            sqlCmd.Parameters.AddWithValue("@macAdd", macAdd);
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        

        private void employeeName_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void employeeName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                if (Password.Text == "")
                    Password.Select();
                else
                    btnLog.PerformClick();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string DataSource = DataSource56;
            string serverType = "شؤون رعايا";
            if (Server == "57")
            {
                DataSource = DataSource57;
                serverType = "احوال شخصية";
            }

                SignUp signUp = new SignUp("جديد", "غير محدد", DataSource, serverType, GregorianDate,"no","", - 1);
            signUp.Show();
        }

        private void FormDataBase_MouseMove(object sender, MouseEventArgs e)
        {
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("en-US");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
        }

       
        private void FormDataBase_Load(object sender, EventArgs e)
        {

        }

        private void FormDataBase_MouseHover(object sender, EventArgs e)
        {
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("en-US");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                if (!File.Exists(primeryLink + @"\updatingSetup.txt"))
                {
                    dataSourceWrite(primeryLink + @"\updatingSetup.txt", "done");
                    currentStatus = "done";
                }
                else
                    currentStatus = File.ReadAllText(primeryLink + @"\updatingSetup.txt");
            }
            catch (Exception ex) { }
            try
            {
                string status = File.ReadAllText(primeryLink + @"\updatingStatus.txt");
                
                if (currentStatus == "updating" && status == "Allowed")
                {
                    dataSourceWrite(primeryLink + @"\updatingSetup.txt", "done");
                    this.Close();
                }
            }
            catch (Exception ex) { }
        }

        private void Password_TextChanged(object sender, EventArgs e)
        {
            if (actDate == username.Text)
            {
                SqlConnection sqlCon = new SqlConnection(DataSource);
                if (sqlCon.State == ConnectionState.Closed)
                    try
                    {
                        sqlCon.Open();
                    }
                    catch (Exception ex) { return; }
                SqlCommand sqlCmd = new SqlCommand("update TableSettings set renew=@renew where ID=@id", sqlCon);
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@id", 1);
                sqlCmd.Parameters.AddWithValue("@renew", GregorianDate);
                sqlCmd.ExecuteNonQuery();
                sqlCon.Close();
            }
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            
            if (!green57.Visible) {
                fillDatagrid(userTable, DataSource57, green57, red57, "الاتصال مع مخدم قسمي الاحوال الشخصية وشؤون الرعايا غير متاح يرجى التواصل مع مشغل النظام");
                fillDatagrid(userTable, DataSource56, green57, red57, "الاتصال مع مخدم قسمي الاحوال الشخصية وشؤون الرعايا غير متاح يرجى التواصل مع مشغل النظام");
            }
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
           
        }

        private int HijriDateDifferment(string source, bool daymonth)
        {
            int differment = 0;
            string query;
            using (SqlConnection saConn = new SqlConnection(source))
            {

                try
                {
                    saConn.Open();
                }
                catch (Exception ex) { return differment; }
                if (daymonth) query = "select hijriday from TableSettings";
                else query = "select hijrimonth from TableSettings";
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    if (daymonth) differment = Convert.ToInt32(reader["hijriday"].ToString());
                    else differment = Convert.ToInt32(reader["hijrimonth"].ToString());
                }

                saConn.Close();

            }
            return differment;
        }
        private void timer2_Tick(object sender, EventArgs e)
        {
            
        }

        private void appversion_MouseClick(object sender, MouseEventArgs e)
        {
            upDateClose();
        }

        private void timer5_Tick(object sender, EventArgs e)
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            if (GregorianDate != DateTime.Now.ToString("MM-dd-yyyy") && IP != "192.168.100.67")
                {
                MessageBox.Show("سيتم إغلاق النافذة نسبة لتجاوز القيد الزمني لليوم");
                this.Close();
            }
        }

        private void appversion_Click(object sender, EventArgs e)
        {

        }

        private void checkArchFiles_Tick(object sender, EventArgs e)
        {
            if ( IP == "192.168.100.67")
            backup();
            else checkArchFiles.Enabled = false;
            //else if (IP == "192.168.100.67")
            //    backup(BackupDisck("backupFile") + @":\");
        }

        string actDate = "";
        private void getDate()
        {
            CultureInfo arSA = new CultureInfo("ar-SA");
            arSA.DateTimeFormat.Calendar = new GregorianCalendar();


            Thread.CurrentThread.CurrentCulture = arSA;
            new System.Globalization.GregorianCalendar();
            GregorianDate = DateTime.Now.ToString("MM-dd-yyyy");
            int year = Convert.ToInt32(SpecificDigit(GregorianDate.Split('-')[2], 3, 4)); 
            int month = Convert.ToInt32(GregorianDate.Split('-')[0]);
            int day = Convert.ToInt32(GregorianDate.Split('-')[1]);
            //if (GregorianDate.Contains("-")) timer2.Enabled = false;

            actDate = "##314" + (7 * (year + day) + month).ToString() + "_@PErs" + (301 * month).ToString();
            //MessageBox.Show(actDate);
        }

       

        private void FormDataBase_MouseEnter_1(object sender, EventArgs e)
        {
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("en-US");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
           
            
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {

        }

        private void Password_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                //MessageBox.Show("enter");
                btnLog.PerformClick();
            }
        }

        private void SeePass_CheckedChanged(object sender, EventArgs e)
        {
            if (SeePass.CheckState == CheckState.Checked)
            {
                Password.UseSystemPasswordChar = false;
            }
            else Password.UseSystemPasswordChar = true;
        }



    }
}
