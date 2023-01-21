using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Configuration;
using System.Globalization;
using System.Threading;
using System.IO;
using System.Drawing.Printing;
using System.Drawing.Imaging;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Xceed.Words.NET;
using System.Diagnostics;
using Xceed.Document.NET;
using System.Globalization;
using System.Threading;
using Aspose.Words;
using System;
using System.Runtime.InteropServices;
using WIA;
using PersAhwal;
using System.Net;
using Image = System.Drawing.Image;
using Microsoft.Reporting.WinForms;
using Paragraph = Xceed.Document.NET.Paragraph;
using System.Xml;
using DocumentFormat.OpenXml.Office2010.Excel;
using ZXing;
using System.Data.SqlTypes;
using DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;
using System.Runtime.InteropServices.ComTypes;
using Aspose.Words.Settings;
using DocumentFormat.OpenXml.Office2016.Drawing.Charts;
using System.Text.RegularExpressions;
using static Azure.Core.HttpHeader;
using System.Security.Cryptography.X509Certificates;
using SautinSoft.Document;
using Color = System.Drawing.Color;
using DocumentFormat.OpenXml;

namespace PersAhwal
{
    public partial class AllConsArchInfo : Form
    {
        string DataSource = "";
        string EmpName = "";
        string AtVC = "";
        string GregorianDate = "";
        string HijriDate = "";
        string FilespathIn = "";
        string FilespathOut = "";
        string Jobposition = "";
        string updateAll = "";
        string insertAll = "";
        string itemAll = "";
        string[] allList;
        string[] PathImages = new string[100];
        int imagecount = 0;
        int unvalid = 0;
        PictureBox picUpdate;
        DeviceInfo AvailableScanner = null;
        bool autoCompleteMode = false;
        int PicID = 0;
        int genIDNo = 0;
        int gridDNo = 0;
        string ArchSerial = "";
        bool gridFill = false;  
        public AllConsArchInfo(string dataSource, string empName, string gregorianDate , string filespathIn, string filespathOut, string jobposition)
        {
            InitializeComponent();
            DataSource = dataSource;
            allList = getColList("AllConsArchInfo");
            FilespathIn = filespathIn;
            FilespathOut = filespathOut;            
            تاريخ_الأرشفة.Text = GregorianDate = gregorianDate;
            fillDataGrid();
            رقم_المكاتبة.Select();
            موظف_الأرشفة.Text = empName;
            EmpName = موظف_الأرشفة.Text = empName;
            Jobposition = jobposition;
            نوع_تاريخ_التوثيق.SelectedIndex = 4;
            if (jobposition.Contains("قنصل")) مسؤول_الأرشفة.Text = "معتمد " + empName;
            System.Globalization.CultureInfo TypeOfLanguage = new System.Globalization.CultureInfo("ar-SA");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(TypeOfLanguage);
        }

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
            string insertValues = "";
            string updateValues = "";
            foreach (DataRow row in dtbl.Rows)
            {

                if (row["name"].ToString() != "ID"&& row["name"].ToString() != "تعليق")
                {
                    allList[i] = row["name"].ToString();
                    //MessageBox.Show(row["name"].ToString());
                    if (i == 0)
                    {
                        itemAll = row["name"].ToString();
                        insertValues = "@" + row["name"].ToString();
                        updateValues = row["name"].ToString() + "=@" + row["name"].ToString();
                    }
                    else
                    {
                        itemAll = itemAll + "," + row["name"].ToString();
                        insertValues = insertValues + "," + "@" + row["name"].ToString();
                        updateValues = updateValues + "," + row["name"].ToString() + "=@" + row["name"].ToString();
                    }
                    i++;
                }
            }             
            updateAll = "UPDATE " + table + " SET " + updateValues + " where ID = @id";
            insertAll = "INSERT INTO " + table + "(" + itemAll + ") values (" + insertValues + "); SELECT @@IDENTITY as lastid";

            return allList;
        }
        private void fillDataGrid()
        {

            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query1 = "SELECT ID,"+ itemAll+ ",تعليق FROM AllConsArchInfo order by ID desc";
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query1, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            dataGridView1.DataSource = table;
            if (dataGridView1.Rows.Count > 1)
            {
                dataGridView1.BringToFront();
                //dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns["مقدم_الطلب"].Width = 200;
                dataGridView1.Columns["نوع_المكاتبة"].Width = 150;
                ColorFulGrid9();                
            }
            pictureBox1.Visible = false;    
            pictureBox1.Image = global::PersAhwal.Properties.Resources.noImage;
            dataGridView1.BringToFront();

        }

        private void AllConsArchInfo_Load(object sender, EventArgs e)
        {
            autoCompleteTextBox(نوع_المكاتبة, DataSource, "نوع_المكاتبة", "AllConsArchInfo");
            autoCompleteTextBox(اسم_المسؤول, DataSource, "اسم_المسؤول", "AllConsArchInfo");
            autoCompleteTextBox(مقدم_الطلب, DataSource, "الاسم", "TableGenNames");
            autoCompleteTextBox(رقم_المكاتبة, DataSource, "رقم_المكاتبة", "AllConsArchInfo");
            autoFillCompleteTextBox(القسم, DataSource, "القسم", "AllConsArchInfo");
            fillYears(combYear);
        }
        private void autoCompleteTextBox(TextBox textbox, string source, string comlumnName, string tableName)
        {

            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                DataTable Textboxtable = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(Textboxtable);
                AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();
                
                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    autoComplete.Add(dataRow[comlumnName].ToString());
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }
        private void autoCompleteTextBox(ComboBox  textbox, string source, string comlumnName, string tableName)
        {

            using (SqlConnection saConn = new SqlConnection(source))
            {
                saConn.Open();

                string query = "select " + comlumnName + " from " + tableName;
                SqlCommand cmd = new SqlCommand(query, saConn);
                cmd.ExecuteNonQuery();
                DataTable Textboxtable = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                dataAdapter.Fill(Textboxtable);
                AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();                
                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    autoComplete.Add(dataRow[comlumnName].ToString());
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
            }
        }
        private void autoFillCompleteTextBox(ComboBox textbox, string source, string comlumnName, string tableName)
        {
            string query = "select distinct " + comlumnName + " from " + tableName+" where "+ comlumnName+" <> ''";
            SqlConnection sqlCon = new SqlConnection(source);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable table = new DataTable();
            sqlDa.Fill(table);
            sqlCon.Close();
            foreach (DataRow dataRow in table.Rows)
            {
                textbox.Items.Add(dataRow[comlumnName].ToString());
            }
            try
            {
                textbox.SelectedIndex = 0;
            }
            catch (Exception ex) { }
        }
        private void fillYears(ComboBox combo)
        {
            combo.Items.Clear();
            string query = "select distinct DATENAME(YEAR, التاريخ_الميلادي)  as years from AllConsArchInfo order by DATENAME(YEAR, التاريخ_الميلادي) desc";
            SqlConnection Con = new SqlConnection(DataSource);
            if (Con.State == ConnectionState.Closed)
                try
                {
                    Con.Open();
                    SqlDataAdapter sqlDa = new SqlDataAdapter(query, Con);
                    sqlDa.SelectCommand.CommandType = CommandType.Text;
                    DataTable dtbl2 = new DataTable();
                    sqlDa.Fill(dtbl2);
                    Con.Close();
                    foreach (DataRow dataRow in dtbl2.Rows)
                    {
                        combo.Items.Add(dataRow["years"].ToString());
                    }
                }
                catch (Exception ex) { }
        }
        
        private void reSetPanel()
        {
            gridDNo = imagecount = 0;            
            PathImages = new string[100];
            for (int x = 0; x < 100; x++)
                PathImages[x] = "";
            
            foreach (Control control in panelpicTemp.Controls)
            {
                if (control is PictureBox)
                {
                    control.Name = "unvalid_" + unvalid.ToString();
                    control.Visible = false;
                    unvalid++;
                }
            }

        }
        private void drawTempPics(string location)
        {
            PictureBox picTemp = new PictureBox();
            picTemp.Dock = System.Windows.Forms.DockStyle.Top;
            picTemp.Location = new System.Drawing.Point(0, 0);
            picTemp.Name = "picTemp_" + imagecount.ToString();
            picTemp.Size = new System.Drawing.Size(123, 137);
            picTemp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            picTemp.TabIndex = 841;
            picTemp.TabStop = false;
            picTemp.Click += new System.EventHandler(this.viewDeletePic);
            picTemp.ImageLocation = location;
            panelpicTemp.Controls.Add(picTemp);
        }
        private void viewDeletePic(object sender, EventArgs e)
        {
            PictureBox pictureBox = (PictureBox)sender;
            picUpdate = pictureBox;
            PicID = Convert.ToInt32(pictureBox.Name.Split('_')[1]);
            pictureBox1.ImageLocation = PathImages[PicID];
            dataGridView1.SendToBack();
            pictureBox1.BringToFront();
            pictureBox1.Visible = true;
            if (PicID >= gridDNo)
            {
                fileUpdate.Visible = true;
            }else fileUpdate.Visible = false;
        }
        private void enableControl(bool enable)
        {
            foreach (Control control in this.Controls) {
                if(control.Name != "رقم_المكاتبة" &&control.Name != "txtSearch" && control.Name != "موظف_الأرشفة" && control.Name != "تاريخ_الأرشفة" && control.Name != "رقم_المكاتبة" && control is TextBox)
                    control.Enabled = enable;

            }
            //حفظ_وإنهاء_الارشفة.Enabled = enable;    
        }
        string lastInput1 = "";
        private void تاريخ_توقيع_المكاتبة_TextChanged(object sender, EventArgs e)
        {
            
            if (التاريخ_الميلادي.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(التاريخ_الميلادي.Text, 1, 2));
                if (month > 12)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //تاريخ_الميلاد.Text = "";
                    التاريخ_الميلادي.Text = SpecificDigit(التاريخ_الميلادي.Text, 3, 10);
                    return;
                }
            }

            if (التاريخ_الميلادي.Text.Length == 11)
            {
                التاريخ_الميلادي.Text = lastInput1; return;
            }
            if (التاريخ_الميلادي.Text.Length == 10) return;
            if (التاريخ_الميلادي.Text.Length == 4) التاريخ_الميلادي.Text = "-" + التاريخ_الميلادي.Text;
            else if (التاريخ_الميلادي.Text.Length == 7) التاريخ_الميلادي.Text = "-" + التاريخ_الميلادي.Text;
            lastInput1 = التاريخ_الميلادي.Text;
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

        private void رقم_المكاتبة_TextChanged(object sender, EventArgs e)
        {
            enableControl(false);
        }

        private void picVerify_Click(object sender, EventArgs e)
        {
            FillDatafromGenArch( رقم_المكاتبة.Text); 
            
        }
        private void FillAllConsArchDocs(string id)
        {
            reSetPanel();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from AllConsArchDocs where رقم_المرجع=N'" + id + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            foreach (DataRow reader in dtbl.Rows)
            {
                var name = reader["المستند"].ToString();
                var Data = (byte[])reader["Data1"];
                var ext = reader["Extension1"].ToString();

                string NewFileName = FilespathOut + name.Replace(ext, DateTime.Now.ToString("mmss"));
                // + ext;
                //MessageBox.Show(NewFileName);
                File.WriteAllBytes(NewFileName, Data);
                drawTempPics(NewFileName);
                PathImages[imagecount] = NewFileName;
                imagecount++;
                enableControl(false); picVerify.SendToBack();
                //return;
                //System.Diagnostics.Process.Start(NewFileName);
            }
            نوع_تاريخ_التوثيق.Select();
            enableControl(true); picVerify.SendToBack();
            sqlCon.Close();
        }
        private void FillDatafromGenArch(string id)
        {
            reSetPanel();
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableGeneralArch where  رقم_معاملة_القسم=N'" + id + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            
            foreach (DataRow reader in dtbl.Rows)
            {
                var name = reader["المستند"].ToString();
                var Data = (byte[])reader["Data1"];
                var ext = reader["Extension1"].ToString();

                string NewFileName = FilespathOut + name.Replace(ext, DateTime.Now.ToString("mmss"));
                // + ext;
                //MessageBox.Show(NewFileName);
                File.WriteAllBytes(NewFileName, Data);
                drawTempPics(NewFileName);
                PathImages[imagecount] = NewFileName;
                imagecount++;
                enableControl(false); picVerify.SendToBack();
                return;
                //System.Diagnostics.Process.Start(NewFileName);
            }

            enableControl(true); picVerify.BringToFront();
            sqlCon.Close();
        }

        private void حفظ_وإنهاء_الارشفة_Click(object sender, EventArgs e)
        {
            if (!ready()) return;
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(insertAll, sqlConnection);
            if(genIDNo != 0)
                sqlCommand = new SqlCommand(updateAll, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@id", genIDNo);
            for (int i = 0; i < allList.Length; i++)
            {
                
                    foreach (Control control in this.Controls)
                    {
                        if (control.Name == allList[i])
                        {
                            sqlCommand.Parameters.AddWithValue("@" + allList[i], control.Text);
                            break;
                        }
                    }
            }
            if (genIDNo == 0)
            {
                try
                {
                    var reader = sqlCommand.ExecuteReader();
                    if (reader.Read())
                    {
                        genIDNo = Convert.ToInt32(reader["lastid"].ToString());
                    }
                    sqlConnection.Close();
                }
                catch (Exception ex)
                {
                    return;
                }
            }
            else sqlCommand.ExecuteNonQuery();

            archDocs();
            updateComment(genIDNo);
            reSetPanel();
            fillDataGrid();
            رقم_المكاتبة.Text = مقدم_الطلب.Text = اسم_المسؤول.Text = نوع_المكاتبة.Text = التاريخ_الميلادي.Text = تعليق_جديد_Off.Text = التعليقات_السابقة_Off.Text = "";
            try
            {
                القسم.SelectedIndex = 0;
            }
            catch (Exception ex) { }

            MessageBox.Show("رقم الأرشفة " + genIDNo.ToString());
            genIDNo = 0;
            fillDataGrid();
            dataGridView1.Visible = txtSearch.Visible = btnSearch.Visible = txtSearch.Visible = btnSearch.Visible = labDescribed.Visible = true;
            pictureBox1.Visible = عرض_القائمة.Visible = false;
        }
        private void updateComment(int id) {
            string query = "UPDATE AllConsArchInfo SET تعليق=@تعليق where ID = @id";
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(insertAll, sqlConnection);
            if (genIDNo != 0)
                sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@id", id);
            sqlCommand.Parameters.AddWithValue("@تعليق", commentInfo());
            sqlCommand.ExecuteNonQuery();
        }

        private bool ready()
        {
            if (رقم_المكاتبة.Text == "") {
                var selectedOption = MessageBox.Show("رقم المكاتبة", " لا يمكن المتابعة بدون إضافة للمكاتبة، الاستمرار في الأرشفة بالرقم (بدون)؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                {
                    رقم_المكاتبة.Text = "بدون";
                    return true;
                }
                else if (selectedOption == DialogResult.No) {
                    return false;
                }
                if (القسم.Text == "") {
                    MessageBox.Show("يرجى تحديد القسم الذي تنتدرج تحته المكاتبة");
                    return false;
                }
                if (مقدم_الطلب.Text == "") {
                    MessageBox.Show("يرجى كتابة اسم مقدم الطلب");
                    return false;
                }
                if (اسم_المسؤول.Text == "") {
                    MessageBox.Show("يرجى كتابة اسم المسؤول الموقع على المكاتبة");
                    return false;
                }
                if (نوع_المكاتبة.Text == ""|| !نوع_المكاتبة.Text.Contains("-")) {
                    MessageBox.Show("يرجى تحديد نوع المكاتبة العام والنوع الخاص مفصولين بالعلامة (-)");
                    return false;
                }
                if (التاريخ_الميلادي.Text == "")
                {
                    MessageBox.Show("يرجى تحديد تاريخ إصدار المكاتبة");
                    return false;
                }
            }
            return true;
        }
        private void archDocs() {

            ArchSerial = "";
            for (int x = gridDNo; x < imagecount; x++)
            {
                //MessageBox.Show(location[x]);
                if (PathImages[x] != "")
                {
                    using (Stream stream = File.OpenRead(PathImages[x]))
                    {
                        byte[] buffer1 = new byte[stream.Length];
                        stream.Read(buffer1, 0, buffer1.Length);
                        var fileinfo1 = new FileInfo(PathImages[x]);
                        string extn1 = fileinfo1.Extension;
                        string DocName1 = fileinfo1.Name;
                        string retID = insertDoc(genIDNo.ToString(), GregorianDate, EmpName, DataSource, extn1, DocName1, رقم_المكاتبة.Text, buffer1);
                        if (x == 0) ArchSerial = "قام  " + موظف_الأرشفة.Text + " بأرشفة ملفات بالارقام " + retID;
                        else ArchSerial = ArchSerial + " و" + retID;                        
                    }
                }
            }
            if (ArchSerial != "") ArchSerial = ArchSerial + Environment.NewLine;
        }
        private string insertDoc(string id, string date, string employee, string dataSource, string extn1, string DocName1, string messNo, byte[] buffer1)
        {
            string query = "INSERT INTO AllConsArchDocs (Data1,Extension1,رقم_معاملة_القسم,المستند,الموظف,التاريخ,رقم_المرجع) values (@Data1,@Extension1,@رقم_معاملة_القسم,@المستند,@الموظف,@التاريخ,@رقم_المرجع);SELECT @@IDENTITY as lastid";
            SqlConnection sqlCon = new SqlConnection(dataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", messNo);            
            sqlCmd.Parameters.AddWithValue("@الموظف", employee);
            sqlCmd.Parameters.AddWithValue("@التاريخ", date);
            sqlCmd.Parameters.AddWithValue("@رقم_المرجع", id);
            sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
            sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
            sqlCmd.Parameters.Add("@المستند", SqlDbType.NVarChar).Value = DocName1;
            try
            {
                var reader = sqlCmd.ExecuteReader();
                if (reader.Read())
                {
                    return reader["lastid"].ToString();
                }
                sqlCon.Close();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(GenQuery);
            }
            return "";
        }
        private string commentInfo()
        {
            string comment = "";
            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text == "")
                comment = ArchSerial+  DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine;

            if (تعليق_جديد_Off.Text == "" && التعليقات_السابقة_Off.Text != "")
                comment = ArchSerial + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine + التعليقات_السابقة_Off.Text;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text == "")
                comment = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + ArchSerial + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine;

            if (تعليق_جديد_Off.Text != "" && التعليقات_السابقة_Off.Text != "")
                comment = تعليق_جديد_Off.Text.Trim() + Environment.NewLine + ArchSerial + DateTime.Now.ToString("G") + Environment.NewLine + "--------------" + Environment.NewLine + "*" + التعليقات_السابقة_Off.Text.Trim();

            return comment;
        }

        private void picVersio_Click(object sender, EventArgs e)
        {

            try

            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.

                    var ScanerItem = device.Items[1]; // select the scanner.


                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);

                    PathImages[imagecount] = FilespathOut + "ScanImg" + DateTime.Now.ToString("mmss") + (imagecount).ToString() + ".jpg";


                    if (File.Exists(PathImages[imagecount]))
                    {
                        File.Delete(PathImages[imagecount]);
                    }
                    imgFile.SaveFile(PathImages[imagecount]);

                    pictureBox1.ImageLocation = PathImages[imagecount];
                    drawTempPics(PathImages[imagecount]);
                    dataGridView1.Visible = txtSearch.Visible = btnSearch.Visible = false;
                    pictureBox1.Visible = true;
                    pictureBox1.BringToFront();
                    imagecount++;
                }
                else
                {
                    MessageBox.Show("لا يوجد جهاز ماسح متصل");
                    this.Close();
                }

            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void loadScanner()
        {
            try
            {
                var deviceManager = new DeviceManager();



                for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++) // Loop Through the get List Of Devices.
                {
                    if (deviceManager.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType) // Skip device If it is not a scanner
                    {
                        continue;
                    }
                    AvailableScanner = deviceManager.DeviceInfos[i];
                    break;
                }


            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Index != -1)
            {
                genIDNo = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                foreach (Control control in this.Controls)
                {
                    if ((control is TextBox || control is ComboBox || control is CheckBox) && !control.Name.Contains("Off"))
                    {
                        if (!control.Name.Contains("موظف"))
                            try
                            {
                                control.Text = dataGridView1.CurrentRow.Cells[control.Name].Value.ToString();
                            }
                            catch (Exception ex) { }
                    }

                }
                gridFill = false;
                التعليقات_السابقة_Off.Text = dataGridView1.CurrentRow.Cells["تعليق"].Value.ToString();
                
                if (dataGridView1.CurrentRow.Cells["مقدم_الطلب"].Value.ToString() != "")
                {
                    FillAllConsArchDocs(genIDNo.ToString());
                    gridDNo = imagecount;
                }
                labDescribed.Visible = dataGridView1.Visible = txtSearch.Visible = btnSearch.Visible = false;
                pictureBox1.Visible = true;
                عرض_القائمة.Visible = true;
                if(PathImages[0] != "")
                    pictureBox1.ImageLocation = PathImages[0];
            }
        }

        private void نوع_تاريخ_التوثيق_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13) مقدم_الطلب.Select(); 
        }

        private void مقدم_الطلب_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13) اسم_المسؤول.Select();
        }

        private void اسم_المسؤول_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13) التاريخ_الميلادي.Select();
        }

        private void تاريخ_توقيع_المكاتبة_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (e.KeyChar == (char)13) picVersio.Click = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (imagecount == 0)
            {
                حفظ_وإنهاء_الارشفة.Visible = false;
                panelpicTemp.Height = 638;
            }
            else
            {
                حفظ_وإنهاء_الارشفة.Visible = true;
                panelpicTemp.Height = 577;
            }
            if (Jobposition.Contains("قنصل")) مسؤول_الأرشفة.Text = "معتمد " + EmpName;
            ColorFulGrid9();
        }
        private void ColorFulGrid9()
        {
            int i = 0;
            int revised = 0;
            for (; i < dataGridView1.Rows.Count - 1; i++) 
            {
                //MessageBox.Show((dataGridView1.Rows.Count-1).ToString() +" - "+i.ToString());
                if (dataGridView1.Rows[i].Cells["مسؤول_الأرشفة"].Value.ToString() != "غير معتمد")
                {
                    // timerColor = false;
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    revised++;

                }
                else dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;

            }
            labDescribed.Text = "عدد (" + i.ToString() + ") مستند منها عدد (" + revised.ToString() + ") تم اعتماده من قبل مسؤول الأرشفة";

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(pictureBox1.ImageLocation);
        }

        private void fileUpdate_Click(object sender, EventArgs e)
        {
            try

            {
                if (AvailableScanner == null) loadScanner();
                if (AvailableScanner != null)
                {
                    var device = AvailableScanner.Connect(); //Connect to the available scanner.

                    var ScanerItem = device.Items[1]; // select the scanner.


                    var imgFile = (ImageFile)ScanerItem.Transfer(FormatID.wiaFormatJPEG);

                    PathImages[PicID] = FilespathOut + "ScanImg" + DateTime.Now.ToString("mmss") + PicID.ToString() + ".jpg";


                    if (File.Exists(PathImages[PicID]))
                    {
                        File.Delete(PathImages[PicID]);
                    }
                    imgFile.SaveFile(PathImages[PicID]);

                    pictureBox1.ImageLocation = PathImages[PicID];
                    picUpdate.ImageLocation = PathImages[PicID];
                    //drawTempPics(PathImages[PicID]);
                    dataGridView1.Visible = txtSearch.Visible = btnSearch.Visible = false;                    
                    fileUpdate.Enabled = false;
                }
                else
                {

                    MessageBox.Show("لا يوجد جهاز ماسح متصل");
                    this.Close();
                }

            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void عرض_القائمة_Click(object sender, EventArgs e)
        {
            
        }

        private void نوع_تاريخ_التوثيق_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (نوع_تاريخ_التوثيق.SelectedIndex == 0)
            {
                البحث_بتاريخ.Text = تاريخ_الأرشفة.Text;
                نوع_تاريخ_التوثيق.Width = 286;
                نوع_تاريخ_التوثيق.Location = new System.Drawing.Point(792, 109);
                btnYear.Visible = combYear.Visible = comboMonth.Visible = btnMonth.Visible = false;
            }

            else if (نوع_تاريخ_التوثيق.SelectedIndex == 1)
            {
                نوع_تاريخ_التوثيق.Width = 190;
                البحث_بتاريخ.Text = "";
                نوع_تاريخ_التوثيق.Location = new System.Drawing.Point(888, 109);
                btnYear.Visible = combYear.Visible = comboMonth.Visible = btnMonth.Visible = false;
            }
            else if (نوع_تاريخ_التوثيق.SelectedIndex == 2)
            {
                combYear.Visible = btnYear.Visible = true;
                comboMonth.Visible = btnMonth.Visible = false; 
            }
            else if (نوع_تاريخ_التوثيق.SelectedIndex == 3)
            {
                combYear.Visible = comboMonth.Visible = true;
                btnYear.Visible = btnMonth.Visible = true;
            }
            else if (نوع_تاريخ_التوثيق.SelectedIndex == 4)
            {
                fillDataGrid();
            }
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            if (txtSearch.Text.Length != 0)
            {                
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns["رقم_المكاتبة"].HeaderText.ToString() + " LIKE '" + txtSearch.Text + "'";
                dataGridView1.DataSource = bs;                
            }else fillDataGrid();
        }

        private void combYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            fillDataGrid();            
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns["التاريخ_الميلادي"].HeaderText.ToString() + " LIKE '%-" + combYear.Text + "'";
            dataGridView1.DataSource = bs;
        }

        private void comboMonth_SelectedIndexChanged(object sender, EventArgs e)
        {
            fillDataGrid();
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns["التاريخ_الميلادي"].HeaderText.ToString() + " LIKE '%-" + combYear.Text + "'";
            dataGridView1.DataSource = bs;

            if (نوع_تاريخ_التوثيق.SelectedIndex == 3)
            {
                bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns["التاريخ_الميلادي"].HeaderText.ToString() + " LIKE '" + comboMonth.Text + "%'";
                dataGridView1.DataSource = bs;
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            fillDataGrid();
            dataGridView1.Visible = txtSearch.Visible = btnSearch.Visible = txtSearch.Visible = btnSearch.Visible = labDescribed.Visible = true;
            pictureBox1.Visible = عرض_القائمة.Visible = false;
            
        }

        private void رقم_المكاتبة_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                //MessageBox.Show("enter");
                //btnLog.PerformClick();
            }
        }

        private void txtSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                //MessageBox.Show("enter");
                //btnLog.PerformClick();
            }
        }

        private void البحث_بتاريخ_TextChanged(object sender, EventArgs e)
        {
            if (البحث_بتاريخ.Text.Length == 10)
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns["التاريخ_الميلادي"].HeaderText.ToString() + " LIKE '" + البحث_بتاريخ.Text + "'";
                dataGridView1.DataSource = bs;
            }else fillDataGrid();
        }
    }
}
