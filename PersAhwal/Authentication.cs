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

namespace PersAhwal
{
    public partial class Authentication : Form
    {
        string DataSource = "";
        int Messid = 0;
        string appSex = "سوداني";
        int handIndex = 0;
        string رقم_معاملة_القسم = "";
        string[] PathImages = new string[100];
        int imagecount = 0;
        string FilespathIn = "";
        string FilespathOut = "";
        int MessageDocNo = 0;
        string HijriDate = "";
        string MessageNo = "ق س ج/80/01";
        string[] allList;
        string[] foundList;
        string updateAll = "";
        string insertAll = "";
        bool gridFill = false;
        PictureBox picUpdate;
        DeviceInfo AvailableScanner = null;
        bool autoCompleteMode = false;
        int PicID = 0;
        int unvalid = 0;
        bool نوع_المكاتبة_check = false;
        bool showGrid = false;
        string docStatus = "شهادة صحيحة";
        int picID = 0;
        public Authentication(string dataSource, string atvc, string filespathOut, string employee, string filespathIn, string hijriDate, string greDate)
        {
            InitializeComponent();
            DataSource = dataSource;
            جنسية_الدبلوماسي.SelectedIndex = 0;
            نوع_تاريخ_التوثيق.SelectedIndex = 0;
            عدد_المستندات_off.SelectedIndex = 0;
            نوع_تاريخ_التوثيق.SelectedIndex = 4;
            HijriDate = hijriDate;
            FilespathIn = filespathIn;
            FilespathOut = filespathOut+ @"\";    
            مدير_القسم.Text = atvc;
            تاريخ_الأرشفة.Text = greDate;
            موظف_الأرشقة.Text = employee;
            fillDataGrid("");
            allList = getColList("TableHandAuth");
        }

        private string[] getColList(string table)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('" + table + "') and  name <> 'ID' and name not like 'Data%'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();

            string[] allList = new string[dtbl.Rows.Count];
            for (int col = 0; col < dtbl.Rows.Count; col++)
                allList[col] = "";

            int i = 0;
            string insertItems = "";
            string insertValues = "";
            string updateValues = "";
            foreach (DataRow row in dtbl.Rows)
            {
                Console.WriteLine(row["name"].ToString());
                //MessageBox.Show(row["name"].ToString());
                allList[i] = row["name"].ToString();
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
            insertAll = "insert into " + table + " (" + insertItems + ") values (" + insertValues + ")";
            updateAll = "UPDATE " + table + " SET " + updateValues + " where ID = @id";
            return allList;

        }
        private void btnDeleteHand_Click(object sender, EventArgs e)
        {

        }
        private void fillDataGrid(string text)
        {
            
            SqlConnection sqlCon = new SqlConnection(DataSource);
            string query1 = "SELECT ID,اسم_موقع_المكاتبة ,نوع_المكاتبة,جنسية_الدبلوماسي,تاريخ_الأرشفة,Viewed,تاريخ_توقيع_المكاتبة,العدد,تعليق,مدير_القسم,موظف_الأرشقة,اسم_الجهة,اسم_صاحب_الشهادة,رقم_الشهادة,رقم_معاملة_القسم,الحالة FROM TableHandAuth order by ID desc";
            //try
            //{
                if (sqlCon.State == ConnectionState.Closed)
                    sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter(query1, sqlCon);
                //if (text != "")
                //    sqlDa = new SqlDataAdapter("SELECT ID,اسم_موقع_المكاتبة,نوع_المكاتبة,جنسية_الدبلوماسي,تاريخ_الأرشفة,Viewed,تاريخ_توقيع_المكاتبة,العدد,تعليق,مدير_القسم  FROM TableHandAuth where جنسية_الدبلوماسي=@جنسية_الدبلوماسي order by ID desc", sqlCon);
                sqlDa.SelectCommand.CommandType = CommandType.Text;
                //sqlDa.SelectCommand.Parameters.AddWithValue("@جنسية_الدبلوماسي", text);
                DataTable table = new DataTable();
                sqlDa.Fill(table);
                sqlCon.Close();
                dataGridView1.DataSource = table;
            //}
            //catch (Exception ex) { MessageBox.Show(DataSource + Environment.NewLine+ query1); return; }

            if (dataGridView1.Rows.Count > 1)
            {
                handIndex = 0;
                dataGridView1.BringToFront();
                dataGridView1.Columns[0].Visible = false;

                Messid = Convert.ToInt32(dataGridView1.Rows[handIndex].Cells[0].Value.ToString());
            }
        }

        private void button34_Click()
        {
            
        }

        private bool save2DataBase(bool insert)
        {
            
            string query = checkList(panel1, allList, "TableHandAuth", insert);
            SqlConnection sqlConnection = new SqlConnection(DataSource);
            if (sqlConnection.State == ConnectionState.Closed)
                sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand(query, sqlConnection);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Parameters.AddWithValue("@id", Messid);            
            for (int i = 0; i < foundList.Length; i++)
            {
                if (foundList[i] == "تعليق")
                {
                    if (التعليقات_السابقة_Off.Text != "")                        
                    sqlCommand.Parameters.AddWithValue("@تعليق", تعليق.Text);
                    else
                        sqlCommand.Parameters.AddWithValue("@تعليق", تعليق.Text + Environment.NewLine + " --------------" + تاريخ_الأرشفة + "--------------- " + Environment.NewLine + التعليقات_السابقة_Off.Text);
                }
                else
                    foreach (Control control in panel1.Controls)
                    {
                        string name = control.Name;
                        if (control is Label || control is Button || control is PictureBox) continue;

                        if (name == foundList[i])
                        {                            
                            if ((control is TextBox && control.Text == "") || (control is ComboBox && control.Text.Contains("إختر")))
                                foreach (Control Econtrol in panel1.Controls)
                                {
                                    if ((Econtrol is TextBox && control.Text == "") || (Econtrol is ComboBox && Econtrol.Text.Contains("ختر")))
                                        {                                            
                                            control.BackColor = System.Drawing.Color.MistyRose;
                                            MessageBox.Show("لا يمكن المتابعة يرجى إضافة بيانات الحقل اسم_المندوب ");
                                            return false;
                                        }
                                }
                            sqlCommand.Parameters.AddWithValue("@" + foundList[i], control.Text);
                            break;
                        }
                    }
            }
            sqlCommand.ExecuteNonQuery();

            return true;
        }
        private string checkList(Panel panel, string[] List, string table, bool insert)
        {
            string updateValues = "";
            string insertItems = "";
            string insertValues = "";

            foundList = new string[List.Length];
            for (int f = 0; f < List.Length; f++)
                foundList[f] = "";

            int found = 0;
            foreach (Control control in panel.Controls)
            {
                string name = control.Name;
                if (control is TextBox || control is ComboBox || control is CheckBox)
                    for (int col = 0; col < List.Length; col++)
                        if (name == List[col])
                        {
                            foundList[found] = name;
                            if (found == 0)
                            {
                                insertItems = name;
                                insertValues = "@" + name;
                                updateValues = name + "=@" + name;

                            }
                            else
                            {
                                insertItems = insertItems + "," + name;
                                insertValues = insertValues + "," + "@" + name;
                                updateValues = updateValues + "," + name + "=@" + name;
                            }
                            found++;
                        }
            }
            if(insert) 
                return "insert into " + table + " (" + insertItems + ") values (" + insertValues + ")";
            else 
                return "UPDATE " + table + " SET " + updateValues + " where ID = @id";

        }
        private int SubAuthData(int id, string نوع_المكاتبة, string اسم_موقع_المكاتبة, string جنسية_الدبلوماسي, string تاريخ_الأرشفة, string تاريخ_توقيع_المكاتبة,  string العدد, string مدير_القسم, string التعليق)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                }
                catch (Exception ex) { return 0; }
            SqlCommand sqlCmd = new SqlCommand("INSERT INTO TableHandAuth ( نوع_المكاتبة,اسم_موقع_المكاتبة,جنسية_الدبلوماسي,تاريخ_الأرشفة,تاريخ_توقيع_المكاتبة,العدد,تعليق,مدير_القسم,حالة_الارشفة,اسم_الجهة,رقم_الشهادة,اسم_صاحب_الشهادة,رقم_معاملة_القسم,الحالة,الجنسية) values (@نوع_المكاتبة,@اسم_موقع_المكاتبة,@جنسية_الدبلوماسي,@تاريخ_الأرشفة,@تاريخ_توقيع_المكاتبة,@العدد,@تعليق,@مدير_القسم,@حالة_الارشفة,@اسم_الجهة,@رقم_الشهادة,@اسم_صاحب_الشهادة,@رقم_معاملة_القسم,@الحالة,@الجنسية);SELECT @@IDENTITY as lastid", sqlCon);
            if (id != 1) sqlCmd = new SqlCommand("UPDATE TableHandAuth SET الجنسية=@الجنسية, نوع_المكاتبة=@نوع_المكاتبة,اسم_موقع_المكاتبة=@اسم_موقع_المكاتبة,جنسية_الدبلوماسي=@جنسية_الدبلوماسي,تاريخ_الأرشفة=@تاريخ_الأرشفة,تاريخ_توقيع_المكاتبة=@تاريخ_توقيع_المكاتبة,العدد=@العدد,تعليق=@تعليق,مدير_القسم=@مدير_القسم,حالة_الارشفة=@حالة_الارشفة,اسم_الجهة=@اسم_الجهة,اسم_صاحب_الشهادة=@اسم_صاحب_الشهادة,رقم_الشهادة=@رقم_الشهادة,الحالة=@الحالة where ID=@ID", sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@ID", id);
            sqlCmd.Parameters.AddWithValue("@نوع_المكاتبة", نوع_المكاتبة);
            sqlCmd.Parameters.AddWithValue("@اسم_موقع_المكاتبة", اسم_موقع_المكاتبة);
            sqlCmd.Parameters.AddWithValue("@جنسية_الدبلوماسي", جنسية_الدبلوماسي);
            sqlCmd.Parameters.AddWithValue("@تاريخ_توقيع_المكاتبة", تاريخ_توقيع_المكاتبة);
            sqlCmd.Parameters.AddWithValue("@العدد", العدد);
            sqlCmd.Parameters.AddWithValue("@الجنسية", appSex);
            sqlCmd.Parameters.AddWithValue("@الحالة", docStatus);
            sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", رقم_معاملة_القسم);
            sqlCmd.Parameters.AddWithValue("@اسم_الجهة", اسم_الجهة.Text);
            sqlCmd.Parameters.AddWithValue("@اسم_صاحب_الشهادة", اسم_صاحب_الشهادة.Text);
            sqlCmd.Parameters.AddWithValue("@رقم_الشهادة", رقم_الشهادة.Text);
            sqlCmd.Parameters.AddWithValue("@تاريخ_الأرشفة", تاريخ_الأرشفة);
            sqlCmd.Parameters.AddWithValue("@حالة_الارشفة", "مؤرشف نهائي");

            sqlCmd.Parameters.AddWithValue("@تعليق", تعليق.Text + Environment.NewLine + " --------------"+ تاريخ_الأرشفة + "--------------- "+ Environment.NewLine+ التعليقات_السابقة_Off.Text);

            sqlCmd.Parameters.AddWithValue("@مدير_القسم", مدير_القسم);
            

            if (id == 1)
            {
                var reader = sqlCmd.ExecuteReader();
                if (reader.Read())
                {
                    id = Convert.ToInt32(reader["lastid"].ToString());
                }
                sqlCon.Close();
            }
            else
                sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
            return id;
        }
        private string DocIDGenerator()
        {
            string formtype = "21";
            string year = DateTime.Now.Year.ToString().Replace("20", "");
            string query = "select max(cast (right(رقم_معاملة_القسم,LEN(رقم_معاملة_القسم) - 15) as int)) as newDocID from TableGeneralArch where رقم_معاملة_القسم like N'ق س ج/80/" + year + "/" + formtype + "%'";

            return "ق س ج/80/" + year + "/" + formtype + "/" + getUniqueID(query);
        }
        private string getUniqueID(string query)
        {
            SqlConnection sqlCon = new SqlConnection(DataSource.Replace("AhwalDataBase", "ArchFilesDB"));
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter(query, sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            string maxID = "1";
            foreach (DataRow dataRow in dtbl.Rows)
            {
                try
                {
                    maxID = (Convert.ToInt32(dataRow["newDocID"].ToString()) + 1).ToString();
                }
                catch (Exception ex)
                {
                    return maxID;
                }
            }
            return maxID;
        }
        private void CreatePic(string[] location, string id, string رقم_معاملة_القسم)
        {

            for (int x = picID; x < imagecount; x++)
            {
                //MessageBox.Show(location[x]);
                if (location[x] != "")
                {
                    using (Stream stream = File.OpenRead(location[x]))
                    {
                        byte[] buffer1 = new byte[stream.Length];
                        stream.Read(buffer1, 0, buffer1.Length);
                        var fileinfo1 = new FileInfo(location[x]);
                        string extn1 = fileinfo1.Extension;
                        string DocName1 = fileinfo1.Name;

                        insertDocx(id.ToString(), نوع_المكاتبة.Text, تاريخ_الأرشفة.Text, موظف_الأرشقة.Text, DataSource, extn1, DocName1, رقم_معاملة_القسم, "data2", buffer1);
                        //Console.WriteLine(docid);
                    }
                }
            }
        }

        private void insertDocx(string id, string name, string date, string employee, string dataSource, string extn1, string DocName1, string messNo, string docType, byte[] buffer1)
        {
            string query = "INSERT INTO TableGeneralArch (Data1,Extension1,نوع_المستند,رقم_معاملة_القسم,المستند,الموظف,التاريخ,رقم_المرجع,docTable,الاسم) values (@Data1,@Extension1,@نوع_المستند,@رقم_معاملة_القسم,@المستند,@الموظف,@التاريخ,@رقم_المرجع,@docTable,@الاسم)";
            SqlConnection sqlCon = new SqlConnection(dataSource.Replace("AhwalDataBase", "ArchFilesDB"));
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlCommand sqlCmd = new SqlCommand(query, sqlCon);
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@الاسم", name);
            sqlCmd.Parameters.AddWithValue("@رقم_معاملة_القسم", messNo);
            sqlCmd.Parameters.AddWithValue("@نوع_المستند", docType);
            sqlCmd.Parameters.AddWithValue("@الموظف", employee);
            sqlCmd.Parameters.AddWithValue("@التاريخ", date);
            sqlCmd.Parameters.AddWithValue("@رقم_المرجع", id);
            sqlCmd.Parameters.Add("@Data1", SqlDbType.VarBinary).Value = buffer1;
            sqlCmd.Parameters.Add("@Extension1", SqlDbType.Char).Value = extn1;
            sqlCmd.Parameters.Add("@المستند", SqlDbType.NVarChar).Value = DocName1;
            sqlCmd.Parameters.Add("@docTable", SqlDbType.NVarChar).Value = "TableHandAuth";
            sqlCmd.ExecuteNonQuery();
            sqlCon.Close();
        }

        private void عدد_المستندات_SelectedIndexChanged(object sender, EventArgs e)
        {
            العدد.Text = (عدد_المستندات_off.SelectedIndex + 1).ToString();
        }

        private void حفظ_وإنهاء_الارشفة_Click(object sender, EventArgs e)
        {
            string comment = "";
            if (!اسم_الجهة.Enabled)
            {
                var selectedOption = MessageBox.Show("", "هل الشهادة صحيحة؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption == DialogResult.Yes)
                {
                    docStatus = "شهادة صحيحة";
                }
                else if (selectedOption == DialogResult.No)
                    docStatus = "مستند غير صحيح"; 
            }
            
            var selectedOption1 = MessageBox.Show("", "صاحب الشهادة سوداني الجنسية؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectedOption1 == DialogResult.No)
                appSex = "أجنبي"; 
            else 
            if (selectedOption1 == DialogResult.Yes)
                appSex = "سوداني"; 
            
            if (اسم_الجهة.Text != "" && اسم_صاحب_الشهادة.Text != "" && رقم_الشهادة.Text != "")
            {
                CreateMessageAuthentication(تاريخ_الأرشفة.Text, HijriDate, مدير_القسم.Text);
            }
            if (imagecount == 0)
            {
                MessageBox.Show("يرجى أرشفة نموذج التوقيع أولا");
                return;
            }

            if (تاريخ_توقيع_المكاتبة.Text.Length == 0)
            {
                MessageBox.Show("يرجى توضيح تاريخ توثيق الشهادة من آخر جهة");
                return;
            }
            if (اسم_موقع_المكاتبة.Text.Length == 0)
            {
                MessageBox.Show("يرجى توضيح اسم موثق المكاتبة ");
                return;
            }
            int iD = 1;
            if (حفظ_وإنهاء_الارشفة.Text == "حفظ وإنهاء الارشفة")
            {
                // Console.WriteLine("حفظ وتأكيد");
                رقم_معاملة_القسم = DocIDGenerator();
                iD = SubAuthData(1, نوع_المكاتبة.Text, اسم_موقع_المكاتبة.Text, جنسية_الدبلوماسي.Text, تاريخ_الأرشفة.Text, تاريخ_توقيع_المكاتبة.Text, العدد.Text, مدير_القسم.Text, comment);
                if (iD != 0) CreatePic(PathImages, iD.ToString(), رقم_معاملة_القسم);
            }
            else if (حفظ_وإنهاء_الارشفة.Text == "تعديل وإنهاء الارشفة")
            {
                iD = Messid;
                SubAuthData(iD, نوع_المكاتبة.Text, اسم_موقع_المكاتبة.Text, جنسية_الدبلوماسي.Text, تاريخ_الأرشفة.Text, تاريخ_توقيع_المكاتبة.Text, العدد.Text, مدير_القسم.Text, comment);
                CreatePic(PathImages, Messid.ToString(), رقم_معاملة_القسم);
            }
            string PaperiD = رقم_معاملة_القسم.Split('/')[4]; ;
            MessageBox.Show("الرقم المرجعي " + PaperiD);
            this.Close();
        }
        private void CreateMessageAuthentication(string gregorianDate, string gijriDate, string ViseConsul) {


            string ActiveCopy;
            string ReportName = DateTime.Now.ToString("mmss");
            string routeDoc = FilespathIn + @"\MessageCapCheck.docx";
            loadMessageNo();
            ActiveCopy = FilespathOut + "Message" + اسم_صاحب_الشهادة.Text + ReportName + ".docx";
            if (!File.Exists(ActiveCopy))
            {
                System.IO.File.Copy(routeDoc, ActiveCopy);
                object oBMiss2 = System.Reflection.Missing.Value;
                Word.Application oBMicroWord2 = new Word.Application();



                Word.Document oBDoc2 = oBMicroWord2.Documents.Open(ActiveCopy, oBMiss2);


                Object ParaMApplicantName = "MarkApplicantName";
                Object ParaMassageIqrarNo = "MarkMassageIqrarNo";
                Object ParaMassageNo = "MarkMassageNo";
                Object ParaHijriDate = "MarkHijriDate";
                Object ParaDateGre = "MarkDateGre";
                Object ParaInstitute = "MarkInstitute";
                Object ParaViseConsul1 = "MarkViseConsul1";


                Word.Range BookMApplicantName = oBDoc2.Bookmarks.get_Item(ref ParaMApplicantName).Range;
                Word.Range BookMassageIqrarNo = oBDoc2.Bookmarks.get_Item(ref ParaMassageIqrarNo).Range;
                Word.Range BookMassageNo = oBDoc2.Bookmarks.get_Item(ref ParaMassageNo).Range;
                Word.Range BookDateGre = oBDoc2.Bookmarks.get_Item(ref ParaDateGre).Range;
                Word.Range BookHijriDate = oBDoc2.Bookmarks.get_Item(ref ParaHijriDate).Range;
                Word.Range BookInstitute = oBDoc2.Bookmarks.get_Item(ref ParaInstitute).Range;
                Word.Range BookViseConsul1 = oBDoc2.Bookmarks.get_Item(ref ParaViseConsul1).Range;

                string noID = MessageNo + (MessageDocNo + 1).ToString();

                BookMApplicantName.Text = اسم_صاحب_الشهادة.Text;
                BookMassageNo.Text = noID;
                BookMassageIqrarNo.Text = رقم_الشهادة.Text;
                BookDateGre.Text = gregorianDate;
                BookInstitute.Text = اسم_الجهة.Text;
                BookHijriDate.Text = HijriDate;
                BookViseConsul1.Text = ViseConsul;

                object rangeViseConsul1 = BookViseConsul1;
                object rangeMApplicantName = BookMApplicantName;
                object rangeMassageIqrarNo = BookMassageIqrarNo;
                object rangeMassageNo = BookMassageNo;
                object rangeDateGre = BookDateGre;
                object rangeHijriDate = BookHijriDate;
                object rangeInstitute = BookInstitute;


                oBDoc2.Bookmarks.Add("MarkViseConsul1", ref rangeViseConsul1);
                oBDoc2.Bookmarks.Add("MarkApplicantName", ref rangeMApplicantName);
                oBDoc2.Bookmarks.Add("MarkMassageIqrarNo", ref rangeMassageIqrarNo);
                oBDoc2.Bookmarks.Add("MarkMassageNo", ref rangeMassageNo);
                oBDoc2.Bookmarks.Add("MarkDateGre", ref rangeDateGre);
                oBDoc2.Bookmarks.Add("MarkInstitute", ref rangeInstitute);
                oBDoc2.Bookmarks.Add("MarkHijiData", ref rangeHijriDate);

                oBDoc2.Activate();
                oBDoc2.Save();
                //addMessageArch(ActiveCopy, noID);
                oBMicroWord2.Visible = true;
                NewMessageNo();
            }
            


        }
        private void loadMessageNo()
        {
            SqlConnection Con = new SqlConnection(DataSource);
            SqlCommand sqlCmd1 = new SqlCommand("select MessageNo  from TableSettings where ID=@id", Con);
            sqlCmd1.Parameters.Add("@id", SqlDbType.Int).Value = 1;
            if (Con.State == ConnectionState.Closed)
                Con.Open();

            var reader = sqlCmd1.ExecuteReader();

            if (reader.Read())
            {
                MessageDocNo = Convert.ToInt32(reader["MessageNo"].ToString());
            }

            Con.Close();


        }
        private void NewMessageNo()
        {
            SqlConnection sqlCon = new SqlConnection(DataSource);
            if (sqlCon.State == ConnectionState.Closed)
                try
                {
                    sqlCon.Open();
                    SqlCommand sqlCmd = new SqlCommand("UPDATE TableSettings SET MessageNo=@MessageNo WHERE ID=@ID", sqlCon);
                    sqlCmd.CommandType = CommandType.Text;
                    sqlCmd.Parameters.AddWithValue("@ID", 1);
                    sqlCmd.Parameters.AddWithValue("@MessageNo", MessageDocNo + 1);
                    sqlCmd.ExecuteNonQuery();
                    sqlCon.Close();

                }

                catch (Exception ex)
                {
                    MessageBox.Show("الوصول لقاعدة البيانات غير متاح");
                }
                finally
                {
                }
        }

        private void اسم_موقع_المكاتبة_TextChanged(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns["اسم_موقع_المكاتبة"].HeaderText.ToString() + " LIKE '%" + اسم_موقع_المكاتبة.Text + "%'";
            dataGridView1.DataSource = dataGridView1.DataSource = bs;
        }

        string lastInput2 = "";
        private void البحث_بتاريخ_TextChanged(object sender, EventArgs e)
        {
            if (البحث_بتاريخ.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(البحث_بتاريخ.Text, 1, 2));
                if (month > 12)
                {
                    //MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //تاريخ_الميلاد.Text = "";
                    البحث_بتاريخ.Text = SpecificDigit(البحث_بتاريخ.Text, 3, 10);
                    return;
                }
                //MessageBox.Show(dateAuth.Text);
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns[4].HeaderText.ToString() + " LIKE '" + البحث_بتاريخ.Text + "%'";
                dataGridView1.DataSource = dataGridView1.DataSource = bs;
            }

            if (البحث_بتاريخ.Text.Length == 11)
            {
                البحث_بتاريخ.Text = lastInput2; return;
            }
            if (البحث_بتاريخ.Text.Length == 10) return;
            if (البحث_بتاريخ.Text.Length == 4) البحث_بتاريخ.Text = "-" + البحث_بتاريخ.Text;
            else if (البحث_بتاريخ.Text.Length == 7) البحث_بتاريخ.Text = "-" + البحث_بتاريخ.Text;
            lastInput2 = البحث_بتاريخ.Text;
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
            
            if (dataGridView1.CurrentRow.Index != -1)
            {
                Messid = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                gridFill = true;
                fillInfo();
                رقم_معاملة_القسم = dataGridView1.CurrentRow.Cells["رقم_معاملة_القسم"].Value.ToString();
                FillDatafromGenArch("data2", Messid.ToString(), "TableHandAuth");
                التعليقات_السابقة_Off.Text = تعليق.Text;
                تعليق.Text = "";
                حفظ_وإنهاء_الارشفة.Text = "تعديل وإنهاء الارشفة";
            }
            gridFill = false;
            return;
        }
        
        private void singleInfo()
        {
            
            
            if (dataGridView1.Rows.Count>1)
            {
                Messid = Convert.ToInt32(dataGridView1.Rows[0].Cells[0].Value.ToString());
                gridFill = true;
                if (allList is null) return;
                foreach (Control control in panel1.Controls)
                {
                    for (int col = 0; col < allList.Length; col++)
                    {
                        if (control.Name == allList[col])
                        {
                            if (dataGridView1.Rows[0].Cells[allList[col]].Value.ToString() != "")
                            {
                                control.Text = dataGridView1.Rows[0].Cells[allList[col]].Value.ToString();
                            }

                        }
                    }
                }
                رقم_معاملة_القسم = dataGridView1.Rows[0].Cells["رقم_معاملة_القسم"].Value.ToString();
                FillDatafromGenArch("data2", Messid.ToString(), "TableHandAuth");
                التعليقات_السابقة_Off.Text = تعليق.Text;
                تعليق.Text = "";
                حفظ_وإنهاء_الارشفة.Text = "تعديل وإنهاء الارشفة";
            }
            gridFill = false;
            return;
        }
        void FillDatafromGenArch(string doc, string id, string table)
        {
            reSetPanel();
            SqlConnection sqlCon = new SqlConnection(DataSource.Replace("AhwalDataBase", "ArchFilesDB"));
            if (sqlCon.State == ConnectionState.Closed)
                sqlCon.Open();
            SqlDataAdapter sqlDa = new SqlDataAdapter("select * from TableGeneralArch where  رقم_المرجع='" + id + "' and نوع_المستند='" + doc + "' and docTable='" + table + "'", sqlCon);
            sqlDa.SelectCommand.CommandType = CommandType.Text;
            DataTable dtbl = new DataTable();
            sqlDa.Fill(dtbl);
            sqlCon.Close();
            picID = 0;
            foreach (DataRow reader in dtbl.Rows)
            {
                var name = reader["المستند"].ToString();
                var Data = (byte[])reader["Data1"];
                var ext = reader["Extension1"].ToString();
                
                string NewFileName = FilespathOut+ name.Replace(ext,"")+ imagecount.ToString()+ ext;
                // + ext;
                //MessageBox.Show(NewFileName);
                File.WriteAllBytes(NewFileName, Data);
                if (ext.Contains("docx")) 
                    drawTempDocx(NewFileName);
                else 
                    drawTempPics(NewFileName);
                PathImages[imagecount] = NewFileName;
                imagecount++;
                picID++;
                //System.Diagnostics.Process.Start(NewFileName);
            }


            sqlCon.Close();
        }
        


        private void fillInfo()
        {
            foreach (Control control in panel1.Controls)
            {
                panelFill(control);                
            }
        }

        public void panelFill(Control control)
        {
            if (allList is null) return;
            for (int col = 0; col < allList.Length; col++)
            {
                if (control.Name == allList[col])
                {
                    if (dataGridView1.CurrentRow.Cells[allList[col]].Value.ToString() != "")
                    {
                        control.Text = dataGridView1.CurrentRow.Cells[allList[col]].Value.ToString();
                    }

                }
            }
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
                    عرض_القائمة.Visible = pictureBox1.Visible = true;
                    
                    pictureBox1.BringToFront();
                    جنسية_الدبلوماسي.Location = new System.Drawing.Point(115, 0);
                    جنسية_الدبلوماسي.Width = 190;

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
        
        private void drawTempDocx(string location)
        {
            PictureBox picTemp = new PictureBox();
            picTemp.Dock = System.Windows.Forms.DockStyle.Top;
            picTemp.Location = new System.Drawing.Point(0, 0);
            picTemp.Name = "picTemp_" + imagecount.ToString();
            picTemp.Size = new System.Drawing.Size(123, 137);
            picTemp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            picTemp.TabIndex = 841;
            picTemp.TabStop = false;
            picTemp.Click += new System.EventHandler(this.viewDeleteDocx);
            picTemp.Image = global::PersAhwal.Properties.Resources.docx;
            panelpicTemp.Controls.Add(picTemp);
        }

        

        private void reSetPanel()
        {
            imagecount = 0;
            PathImages = new string[100];
            for(int x = 0;x < 100;x++)
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

        private void viewDeletePic(object sender, EventArgs e)
        {
            PictureBox pictureBox = (PictureBox)sender;
            picUpdate = pictureBox;
            PicID = Convert.ToInt32(pictureBox.Name.Split('_')[1]);
            //MessageBox.Show(PathImages[PicID]);
            pictureBox1.ImageLocation = PathImages[PicID];
            dataGridView1.SendToBack();
            pictureBox1.BringToFront();
            عرض_القائمة.Visible = pictureBox1.Visible = true;
            var selectedOption = MessageBox.Show("حذف المستند من قائمة الأرشفة؟", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                pictureBox.Visible = false;
                PathImages[Convert.ToInt32(pictureBox.Name.Split('_')[1])] = "";
            }
            fileUpdate.Enabled = true;
        }
        
        private void viewDeleteDocx(object sender, EventArgs e)
        {
            PictureBox pictureBox = (PictureBox)sender;
            picUpdate = pictureBox;
            PicID = Convert.ToInt32(pictureBox.Name.Split('_')[1]);
            //MessageBox.Show(PathImages[PicID]);
            System.Diagnostics.Process.Start(PathImages[PicID]);
            //pictureBox1.ImageLocation = PathImages[PicID];
            //dataGridView1.SendToBack();
            //pictureBox1.BringToFront();
            //عرض_القائمة.Visible = pictureBox1.Visible = true;
            var selectedOption = MessageBox.Show("حذف المستند من قائمة الأرشفة؟", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectedOption == DialogResult.Yes)
            {
                pictureBox.Visible = false;
                PathImages[Convert.ToInt32(pictureBox.Name.Split('_')[1])] = "";
            }
            fileUpdate.Enabled = true;
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (imagecount == 0)
            {
                حفظ_وإنهاء_الارشفة.Visible = false;
                panelpicTemp.Height = 638;
            }
            else {
                حفظ_وإنهاء_الارشفة.Visible = true;
                panelpicTemp.Height = 577;
            }
            ColorFulGrid9();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            fillDataGrid("");
            dataGridView1.Visible = txtSearch.Visible = btnSearch.Visible = true;
            pictureBox1.Visible = عرض_القائمة.Visible = false;
        }

        private void نوع_تاريخ_التوثيق_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (نوع_تاريخ_التوثيق.SelectedIndex == 0)
            {
                البحث_بتاريخ.Text = تاريخ_الأرشفة.Text;
                نوع_تاريخ_التوثيق.Width = 286;
                نوع_تاريخ_التوثيق.Location = new System.Drawing.Point(19, 44);
                btnYear.Visible = combYear.Visible = comboMonth.Visible = btnMonth.Visible = false;
            }

            else if (نوع_تاريخ_التوثيق.SelectedIndex == 1)
            {
                نوع_تاريخ_التوثيق.Width = 190;
                نوع_تاريخ_التوثيق.Location = new System.Drawing.Point(115, 44);
                btnYear.Visible = combYear.Visible = comboMonth.Visible = btnMonth.Visible = false;
            }
            else if (نوع_تاريخ_التوثيق.SelectedIndex == 2) {
                combYear.Visible = btnYear.Visible = true;
                comboMonth.Visible = btnMonth.Visible = false;
            }
            else if (نوع_تاريخ_التوثيق.SelectedIndex == 3) {
                combYear.Visible = comboMonth.Visible = true;
                btnYear.Visible = btnMonth.Visible = true;
            }
            else if (نوع_تاريخ_التوثيق.SelectedIndex == 4) {
                fillDataGrid("");
            }
        }

        private void fillYears(ComboBox combo)
        {
            combo.Items.Clear();
            string query = "select distinct DATENAME(YEAR, تاريخ_الأرشفة)  as years from TableHandAuth order by DATENAME(YEAR, تاريخ_الأرشفة) desc";
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

                    PathImages[imagecount] = FilespathOut + "ScanImg" + DateTime.Now.ToString("mmss") + imagecount.ToString() + ".jpg";


                    if (File.Exists(PathImages[imagecount]))
                    {
                        File.Delete(PathImages[imagecount]);
                    }
                    imgFile.SaveFile(PathImages[imagecount]);

                    pictureBox1.ImageLocation = PathImages[imagecount];
                    picUpdate.ImageLocation = PathImages[imagecount];
                    imagecount++;
                    //drawTempPics(PathImages[PicID]);
                    dataGridView1.Visible = txtSearch.Visible = btnSearch.Visible = false; 
                    جنسية_الدبلوماسي.Location = new System.Drawing.Point(115, 0);
                    جنسية_الدبلوماسي.Width = 190;
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

        string lastInput1 = "";
        private void تاريخ_توقيع_المكاتبة_TextChanged(object sender, EventArgs e)
        {
            if (تاريخ_توقيع_المكاتبة.Text.Length == 10)
            {
                int month = Convert.ToInt32(SpecificDigit(تاريخ_توقيع_المكاتبة.Text, 1, 2));
                if (month > 12 && !gridFill)
                {
                    MessageBox.Show("الشهر يحب أن يكون أقل من 12");
                    //تاريخ_الميلاد.Text = "";
                    تاريخ_توقيع_المكاتبة.Text = SpecificDigit(تاريخ_توقيع_المكاتبة.Text, 3, 10);
                    return;
                }
            }

            if (تاريخ_توقيع_المكاتبة.Text.Length == 11)
            {
                تاريخ_توقيع_المكاتبة.Text = lastInput1; return;
            }
            if (تاريخ_توقيع_المكاتبة.Text.Length == 10) return;
            if (تاريخ_توقيع_المكاتبة.Text.Length == 4) تاريخ_توقيع_المكاتبة.Text = "-" + تاريخ_توقيع_المكاتبة.Text;
            else if (تاريخ_توقيع_المكاتبة.Text.Length == 7) تاريخ_توقيع_المكاتبة.Text = "-" + تاريخ_توقيع_المكاتبة.Text;
            lastInput1 = تاريخ_توقيع_المكاتبة.Text;
        }

        private void Authentication_Load(object sender, EventArgs e)
        {
            autoCompleteTextBox(اسم_موقع_المكاتبة, DataSource, "اسم_موقع_المكاتبة", "TableHandAuth");
            autoCompleteTextBox(نوع_المكاتبة, DataSource, "نوع_المكاتبة", "TableHandAuth");
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
                bool newSrt = true;
                foreach (DataRow dataRow in Textboxtable.Rows)
                {
                    autoComplete.Add(dataRow[comlumnName].ToString());
                }
                textbox.AutoCompleteMode = AutoCompleteMode.Suggest;
                textbox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textbox.AutoCompleteCustomSource = autoComplete;
                saConn.Close();
                autoCompleteMode = true;
            }
        }

        private void Authentication_FormClosed(object sender, FormClosedEventArgs e)
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

        private void button35_Click(object sender, EventArgs e)
        {
            اسم_الجهة.Enabled = اسم_صاحب_الشهادة.Enabled = رقم_الشهادة.Enabled = true;
            docStatus = "في انتظار تأكيد صحتها";
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            string year  = SpecificDigit(txtSearch.Text, 1, 2);
            string docID = SpecificDigit(txtSearch.Text, 3, 10);
            if (docID.Length != 0)
            {
                string id = "ق س ج/80/" + year + "/21/" + docID;
                //MessageBox.Show(id);
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns["رقم_معاملة_القسم"].HeaderText.ToString() + " LIKE '" + id + "'";
                dataGridView1.DataSource = bs;
                Console.WriteLine(id);
                if (dataGridView1.Rows.Count == 2)
                    singleInfo();
                //ColorFulGrid9();

                    //MessageBox.Show(docID);
            }
        }

        private void combYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            fillDataGrid(""); 
            //if (نوع_تاريخ_التوثيق.SelectedIndex == 2) {
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns["تاريخ_الأرشفة"].HeaderText.ToString() + " LIKE '%-" + combYear.Text + "'";
                dataGridView1.DataSource = bs;
            //}
        }

        private void comboMonth_SelectedIndexChanged(object sender, EventArgs e)
        {
            fillDataGrid("");
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns["تاريخ_الأرشفة"].HeaderText.ToString() + " LIKE '%-" + combYear.Text + "'";
            dataGridView1.DataSource = bs;

            if (نوع_تاريخ_التوثيق.SelectedIndex == 3)
            {
                bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns["تاريخ_الأرشفة"].HeaderText.ToString() + " LIKE '" + comboMonth.Text+ "%'";
                dataGridView1.DataSource = bs;
            }
        }
        private void ColorFulGrid9()
        {
            int i = 0;
            int countSudan = 0;
            int countSaudi = 0;
            for (; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells["جنسية_الدبلوماسي"].Value.ToString() == "دبلوماسيون سودانيون")
                {
                    // timerColor = false;
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    countSudan++;

                }
                else if (dataGridView1.Rows[i].Cells["جنسية_الدبلوماسي"].Value.ToString() == "دبلوماسيون سعوديون")
                {
                    // timerColor = false;
                    //dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightPink;
                    countSaudi++;
                }
                if (dataGridView1.Rows[i].Cells["الحالة"].Value.ToString() == "في انتظار تأكيد صحتها")
                {
                    // timerColor = false;
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightPink;
                    
                }
                else if (dataGridView1.Rows[i].Cells["الحالة"].Value.ToString() == "مستند غير صحيح")
                {
                    // timerColor = false;
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    
                }

            }
            labDescribed.Text = "عدد (" + i.ToString() + ") مستند (" + countSudan.ToString() + "/"+ countSaudi.ToString()+")";

        }

        private void نوع_المكاتبة_TextChanged(object sender, EventArgs e)
        {
            if (نوع_المكاتبة_check)
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = dataGridView1.DataSource;
                bs.Filter = dataGridView1.Columns["نوع_المكاتبة"].HeaderText.ToString() + " LIKE '%" + نوع_المكاتبة.Text + "%'";
                dataGridView1.DataSource = dataGridView1.DataSource = bs;
                
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            نوع_المكاتبة_check = true;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(pictureBox1.ImageLocation);
            }
            catch (Exception ex) { }
        }
    }
}
