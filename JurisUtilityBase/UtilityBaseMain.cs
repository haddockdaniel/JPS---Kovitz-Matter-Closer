using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;
using Microsoft.VisualBasic.FileIO;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }
        public int activeTab = 0; // 0 = client, 1 = matter
        List<ErrorLog> errorList = new List<ErrorLog>();
        private string clientFilePath = "";
        private string matterFilePath = "";


        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods


        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }



        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {


            switch (activeTab) // 0 client, 1 matter
            {
                case 0:

                        if (checkClientBoxesForInfo())
                        {
                            CliMat cm = new CliMat();
                            cm.clicode = textBoxClientText.Text;
                            cm.name = textBoxCliName.Text.Replace("%", "").Replace("'", "");
                            cm.remarks = richTextBoxCliRemarks.Text.Replace("%", "").Replace("'", "");
                            cm.branch = textBoxCliBranch.Text.Replace("%", "").Replace("'", "");

                            processSingleClient(cm);
                            showFinish();
                    }
                    break;
                case 1:
                    if (checkMatterBoxesForInfo())
                    {
                        CliMat cm = new CliMat();
                        cm.clicode = textBoxClientMatter.Text;
                        cm.matcode = textBoxMatterText.Text;
                        cm.name = "";
                        if (!string.IsNullOrEmpty(richTextBoxRemarksMatter.Text.Replace("%", "").Replace("'", "")))
                            cm.remarks = richTextBoxRemarksMatter.Text.Replace("%", "").Replace("'", "");
                        else
                            cm.remarks = "";
                        cm.branch = "";
                        cm.matsys = getMatSysNbr(cm.clicode, cm.matcode);
                        cm.clisys = getCliSysNbr(cm.clicode);
                        if (cm.clisys != 0 && cm.matsys != 0)
                            processSingleMatter(cm.matsys, cm.remarks);
                        else
                        {
                            ErrorLog er = new ErrorLog();
                            er.client = cm.clicode;
                            er.matter = cm.matcode;
                            er.message = "Cannot close matter because client/matter " + cm.clicode + "/" + cm.matcode + " does not exist. \r\n" + "\r\n";
                            errorList.Add(er);

                        }
                        showFinish();
                    }// 0 is simply a place holder for method...means nothing
                    break;
            }


            errorList.Clear();
            clientFilePath = "";
            matterFilePath = "";
            textBoxClientText.Text = "";
            textBoxCliBranch.Text = "";
            textBoxCliName.Text = "";
            richTextBoxCliRemarks.Text = "";
            textBoxMatterText.Text = "";
            richTextBoxRemarksMatter.Text = "";
            textBoxClientMatter.Text = "";
        }

        private void showFinish()
        {
            UpdateStatus("Client(s)/Matter(s) updated.", 1, 1);

            if (errorList.Count == 0)
                MessageBox.Show("The process is complete and there were no errors.", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
            else
            {
                DialogResult ff = MessageBox.Show("The process is complete but there were errors." + "\r\n" + "Would you like to see the Error Log?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (ff == DialogResult.Yes)
                {
                    ReportDisplay rd = new ReportDisplay(errorList);
                    rd.showErrors();
                    rd.Show();
                }
            }
        }


        private bool checkMatterBoxesForInfo()
        {
            if (string.IsNullOrEmpty(textBoxClientMatter.Text)) 
            {
                MessageBox.Show("Please enter a Client Code", "Client Code Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            else
            {
                if (string.IsNullOrEmpty(textBoxMatterText.Text))
                {
                    MessageBox.Show("Please enter a new Branch", "Matter Code Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else
                    return true;

            }

        }

        private bool checkClientBoxesForInfo()
        {
            if (string.IsNullOrEmpty(textBoxClientText.Text))
            {
                MessageBox.Show("Please enter a Client Code", "Client Code Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            else
            {
                if (string.IsNullOrEmpty(textBoxCliBranch.Text))
                {
                    MessageBox.Show("Please enter a new Branch", "Branch Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else
                {
                    if (string.IsNullOrEmpty(textBoxCliName.Text))
                    {
                        MessageBox.Show("Please enter a new Name", "Name Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    else
                        return true;
                }
            }

        }

        private void processClientExcel()
        {
            //parse data from file
            OpenFileDialogOpen.InitialDirectory = @"C:\";
            OpenFileDialogOpen.Title = "Browse CSV Files";
            OpenFileDialogOpen.DefaultExt = "csv";
            OpenFileDialogOpen.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            OpenFileDialogOpen.CheckFileExists = true;
            OpenFileDialogOpen.CheckPathExists = true;
            OpenFileDialogOpen.Multiselect = false;

            List<CliMat> cliMatList = new List<CliMat>();

            if (OpenFileDialogOpen.ShowDialog() == DialogResult.OK)
            {
                int lineNum = 0;
                var fileStream = new FileStream(OpenFileDialogOpen.FileName, FileMode.Open, FileAccess.Read);
                using (var streamReader = new StreamReader(fileStream))
                {
                    string line;
                    while ((line = streamReader.ReadLine()) != null)
                    {
                        lineNum++;
                        if (lineNum > 1) // ignore header line
                        {
                            TextFieldParser parser = new TextFieldParser(new StringReader(line));

                            // You can also read from a file
                            // TextFieldParser parser = new TextFieldParser("mycsvfile.csv");

                            parser.HasFieldsEnclosedInQuotes = true;
                            parser.SetDelimiters(",");

                            string[] fields;

                            while (!parser.EndOfData)
                            {
                                fields = parser.ReadFields();

                                CliMat cm = new CliMat();
                                cm.clicode = fields[0];
                                cm.name = fields[1].Replace("%", "").Replace("'", "");
                                cm.clisys = getCliSysNbr(cm.clicode);
                                cm.remarks = fields[3].Replace("%", "").Replace("'", "");
                                cm.branch = fields[2].Replace("%", "").Replace("'", "");
                                cliMatList.Add(cm);
                            }

                            parser.Close();
                        }
                    }
                }

                //order all items by client and matter code
                var finalList = cliMatList.OrderBy(x => x.clicode).ToList();
                cliMatList.Clear();
                int total = finalList.Count;
                int runningTotal = 0;
                foreach (CliMat cc in finalList)
                {
                    if (cc.clisys != 0)
                        processSingleClient(cc);
                    else
                    {
                        ErrorLog er = new ErrorLog();
                        er.client = cc.clicode;
                        er.message = "Client " + cc.clicode + " does not appear to be a valid client. Check that the entered code matches what is displayed in Core exactly." + "\r\n" + "\r\n"; //still close client even with no matters
                        errorList.Add(er);

                    }
                    runningTotal++;
                    UpdateStatus("Updating....", runningTotal, total);
                }

            }
            showFinish();
            errorList.Clear();
            clientFilePath = "";
            matterFilePath = "";
        }

        private void processSingleClient(CliMat cc)
        {
            string sql = "";

            cc.clisys = getCliSysNbr(cc.clicode);
            if (cc.clisys == 0)
            {
                ErrorLog er = new ErrorLog();
                er.client = cc.clicode;
                er.message = "Client " + cc.clicode + " does not appear to be a valid client. Check that the entered code matches what is displayed in Core exactly." + "\r\n" + "\r\n"; //still close client even with no matters
                errorList.Add(er);
            }
            else
            { 
            //make changes to matters first
            sql = " select matsysnbr from matter where matclinbr = " + cc.clisys.ToString() + " and MatStatusFlag <> 'C'";
            DataSet ds = _jurisUtility.RecordsetFromSQL(sql);
                if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                {
                    ErrorLog er = new ErrorLog();
                    er.client = cc.clicode;
                    er.message = "Client " + cc.clicode + " does not have any open matters so no matters were changed. Client was still closed." + "\r\n" + "\r\n"; //still close client even with no matters
                    errorList.Add(er);
                }
                else
                {
                    foreach (DataRow dr in ds.Tables[0].Rows) // Lieke wants clients closed regardless of matter status or errors
                    {
                        processSingleMatter(Convert.ToInt32(dr[0].ToString()), "");
                    }

                }
                //close client regardless of matter status
                sql = "";
                sql = "update client set CliReportingName = left('" + cc.name + "', 30), CliNickName = left('" + cc.name + "', 30), branch = '" + cc.branch + "' where clisysnbr = " + cc.clisys.ToString();
                _jurisUtility.ExecuteNonQuery(0, sql);

                sql = "update client set AccountRep = '' where clisysnbr = " + cc.clisys.ToString();
                _jurisUtility.ExecuteNonQuery(0, sql);

                sql = " select CNNoteText from ClientNote where cnclient = " + cc.clisys + " and CNNoteIndex = 'Remarks'"; 
                DataSet ds1 = _jurisUtility.RecordsetFromSQL(sql);
                if (ds1 == null || ds1.Tables.Count == 0 || ds1.Tables[0].Rows.Count == 0)
                { //if no remarks notecard exists, create a new one
                    sql = "insert into ClientNote ([CNClient] ,[CNNoteIndex],[CNObject],[CNNoteText] ,[CNNoteObject]) " +
                        " values (" + cc.clisys.ToString() + ", 'Remarks', ' ', cast('" + cc.remarks + "' as nvarchar(max)), null)";
                    _jurisUtility.ExecuteNonQuery(0, sql);
                }
                else // if it does exist, put the new text at the top
                {
                    sql = "update ClientNote set CNNoteText = cast('" + cc.remarks + "'  + char(10) + char(13) + cast(CNNoteText as varchar(1000)) as nvarchar(max)) where CNClient = " + cc.clisys.ToString() + " and CNNoteIndex = 'Remarks'";
                    _jurisUtility.ExecuteNonQuery(0, sql);
                }
            }
        }

        private void processMatterExcel()
        {
            OpenFileDialogOpen.InitialDirectory = @"C:\";
            OpenFileDialogOpen.Title = "Browse CSV Files";
            OpenFileDialogOpen.DefaultExt = "csv";
            OpenFileDialogOpen.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            OpenFileDialogOpen.CheckFileExists = true;
            OpenFileDialogOpen.CheckPathExists = true;
            OpenFileDialogOpen.Multiselect = false;

            List<CliMat> cliMatList = new List<CliMat>();

            if (OpenFileDialogOpen.ShowDialog() == DialogResult.OK)
            {


                    int lineNum = 0;
                    var fileStream = new FileStream(OpenFileDialogOpen.FileName, FileMode.Open, FileAccess.Read);
                    using (var streamReader = new StreamReader(fileStream))
                    {
                        string line;
                        while ((line = streamReader.ReadLine()) != null)
                        {
                            lineNum++;
                            if (lineNum > 1) // ignore header line
                            {
                                TextFieldParser parser = new TextFieldParser(new StringReader(line));

                                // You can also read from a file
                                // TextFieldParser parser = new TextFieldParser("mycsvfile.csv");

                                parser.HasFieldsEnclosedInQuotes = true;
                                parser.SetDelimiters(",");

                                string[] fields;

                                while (!parser.EndOfData)
                                {
                                    fields = parser.ReadFields();

                                        CliMat cm = new CliMat();
                                        cm.clicode = fields[0];
                                        cm.matcode = fields[1];
                                        cm.clisys = getCliSysNbr(cm.clicode);
                                        cm.matsys = getMatSysNbr(cm.clicode, cm.matcode);
                                        cm.remarks = fields[2].Replace("%", "").Replace("'", "");
                                        cm.branch = "";
                                       cliMatList.Add(cm);                               
                                }

                                parser.Close();
                            }
                        }
                    }



                //order all items by client and matter code
                var finalList  = cliMatList.OrderBy(x => x.clicode).ThenBy(y => y.matcode).ToList();
                cliMatList.Clear();
                int total = finalList.Count;
                int runningTotal = 0;
                foreach (CliMat cc in finalList)
                {
                    bool notUsed = true;
                    if (cc.clisys !=0 && cc.matsys != 0)
                        notUsed = processSingleMatter(cc.matsys, cc.remarks);
                    else
                    {
                        ErrorLog er = new ErrorLog();
                        er.client = cc.clicode;
                        er.matter = cc.matcode;
                        er.message = "Cannot close matter because client/matter " + cc.clicode + "/" + cc.matcode + " does not exist. \r\n" + "\r\n";
                        errorList.Add(er);

                    }

                }
                runningTotal++;
                UpdateStatus("Updating....", runningTotal, total);
            }
            showFinish();
            errorList.Clear();
            clientFilePath = "";
            matterFilePath = "";
        }

        private bool processSingleMatter(int matsys, string comment)
        {
            if (checkForBalances(matsys)) // if comment is blank, then its a client closing call, else, matter closing direct call
            {
                string sql = "";
                if (string.IsNullOrEmpty(comment)) // this was a call from Client which means we do not upfdate remarks
                {
                    sql = "update matter set MatStatusFlag = 'C', MatLockFlag = 3, MatDateClosed = getdate() where matsysnbr = " + matsys.ToString();
                    _jurisUtility.ExecuteNonQuery(0, sql);
                }
                else // this was a call from the matter which means we do update the remark
                {
                    sql = "update matter set MatStatusFlag = 'C', MatLockFlag = 3, MatDateClosed = getdate(), matremarks = left('" + comment + "' + char(10) + char(13) + matremarks, 250) where matsysnbr = " + matsys.ToString();
                    _jurisUtility.ExecuteNonQuery(0, sql);
                }
                return true;
            }
            else
                return false;

        }

        private int getMatSysNbr(string clicode, string matcode)
        {
            int matsys = 0;
            string sql = "";
            sql ="select matsysnbr from matter " +
                "inner join client on clisysnbr = matclinbr " +
                " where dbo.jfn_FormatClientCode(clicode) = '" + clicode + "' and dbo.jfn_FormatMatterCode(MatCode) = " + matcode;

            DataSet ds = new DataSet();
            ds = _jurisUtility.RecordsetFromSQL(sql);
            if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0) // no matsys found
                return 0;
            else
            {
                matsys = Convert.ToInt32(ds.Tables[0].Rows[0][0].ToString());
                return matsys;
            }
        }

        private int getCliSysNbr(string clicode)
        {
            int id = 0;
            string sql = "";
            sql = "select clisysnbr from client " +
                " where dbo.jfn_FormatClientCode(clicode) = '" + clicode + "'";

            DataSet ds = new DataSet();
            ds = _jurisUtility.RecordsetFromSQL(sql);
            if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0) // no clisys found
                return 0;
            else
            {
                id = Convert.ToInt32(ds.Tables[0].Rows[0][0].ToString());
                return id;
            }


        }

        private bool checkForBalances(int matsysnbr) //true means some balance exists
        {
            string sql = "";
            sql =
                "  select dbo.jfn_FormatClientCode(clicode) as clicode, dbo.jfn_FormatMatterCode(MatCode) as matcode, cast(sum(ppd) as money) as ppd, cast(sum(UT) as money) as UT, cast(sum(UE) as money) as UE, cast(sum(AR) as money) as AR, cast(sum(Trust) as money) as Trust " +
                "from( " +
                "select matsysnbr as matsys, MatPPDBalance as ppd, 0 as UT, 0 as UE, 0 as AR, 0 as Trust " +
                  " from matter " +
                  " where MatPPDBalance <> 0 and matsysnbr = " + matsysnbr +
                  " union all " +
                   " select utmatter as matsys, 0 as ppd, sum(utamount) as UT, 0 as UE, 0 as AR, 0 as Trust " +
                  " from unbilledtime where utmatter = " + matsysnbr + 
                  " group by utmatter " +
                  " having sum(utamount) <> 0 " +
                  " union all " +
                  " select uematter as matsys, 0 as ppd, 0 as UT, sum(ueamount) as UE, 0 as AR, 0 as Trust " +
                 " from unbilledexpense where uematter = " + matsysnbr + 
                  "  group by uematter " +
                " having sum(ueamount) <> 0 " +
                 " union all " +
                "  select armmatter as matsys, 0 as ppd, 0 as UT, 0 as UE, sum(ARMBalDue) as AR, 0 as Trust " +
                "  from armatalloc where armmatter = " + matsysnbr + 
                "  group by armmatter " +
                "  having sum(ARMBalDue) <> 0 " +
                "  union all " +
                "  select tamatter as matsys, 0 as ppd, 0 as UT, 0 as UE, 0 as AR, sum(TABalance) as Trust " +
                 " from trustaccount where tamatter = " + matsysnbr + 
                "  group by tamatter " +
                "  having sum(TABalance) <> 0) hhg " +
                " inner join matter on hhg.matsys = matsysnbr " +
                " inner join client on clisysnbr = matclinbr " +
                "  group by dbo.jfn_FormatClientCode(clicode), dbo.jfn_FormatMatterCode(MatCode) " +
                "  having sum(ppd) <> 0 or sum(UT)  <> 0 or sum(UE)  <> 0 or sum(AR)  <> 0 or sum(Trust) <> 0";

            DataSet ds = new DataSet();
            ds = _jurisUtility.RecordsetFromSQL(sql);
            if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                return true;
            else
            {
                ErrorLog er = new ErrorLog();
                er.client = ds.Tables[0].Rows[0][0].ToString();
                er.matter = ds.Tables[0].Rows[0][1].ToString();
                er.message = "Cannot close matter " +er.client + "/" + er.matter + " because balance(s) exist. See below for more detail: \r\n" +
                    "Prepaid Balance: " + ds.Tables[0].Rows[0][2].ToString() + "\r\n" +
                    "Unbilled Time Balance: " + ds.Tables[0].Rows[0][3].ToString() + "\r\n" +
                    "Unbilled Expense Balance: " + ds.Tables[0].Rows[0][4].ToString() + "\r\n" +
                    "A/R Balance: " + ds.Tables[0].Rows[0][5].ToString() + "\r\n" +
                    "Trust Balance: " + ds.Tables[0].Rows[0][6].ToString() + "\r\n" + "\r\n";
                errorList.Add(er);
                return false;
            }
        }


        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }

        private void buttonMatterExcel_Click(object sender, EventArgs e)
        {
            processMatterExcel();
        }

        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (TabControl1.SelectedTab == TabControl1.TabPages["Client"])//your specific tabname
                activeTab = 0; //client is active tab
            else
                activeTab = 1; // means matter is active
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void buttonClientExcel_Click(object sender, EventArgs e)
        {
            processClientExcel();
        }
    }
}
