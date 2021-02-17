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
                    if (string.IsNullOrEmpty(textBoxClientText.Text))
                        processClientExcel();
                    else
                        processSingleClient(textBoxClientText.Text);
                    break;
                case 1:
                    if (string.IsNullOrEmpty(textBoxMatterText.Text))
                        processMatterExcel();
                    else
                        processSingleMatter(); // 0 is simply a place holder for method...means nothing
                    break;
            }

            UpdateStatus("Client(s)/Matter(s) updated.", 1, 1);

            MessageBox.Show("The process is complete.", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);

            errorList.Clear();
            clientFilePath = "";
            matterFilePath = "";
            textBoxClientText.Text = "";
            textBoxMatterText.Text = "";
        }

        private void processClientExcel()
        {
            //parse data from file



            //go one client at a time
            processSingleClient(clicode);

        }

        private void processSingleClient(string clicode)
        {
            string sql = "";
            //make changes to matters first
            sql = " select matsysnbr from matter inner join client on clisysnbr = matclinbr where dbo.jfn_FormatClientCode(clicode) = '" + clicode + "'";
            DataSet ds = _jurisUtility.RecordsetFromSQL(sql);
            if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
            {
                ErrorLog er = new ErrorLog();
                er.client = clicode;
                er.message = "Client does not have any matters so it cannot be processed" + "\r\n" + "\r\n";
                errorList.Add(er);
            }
            else
            {
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    processSingleMatter(Convert.ToInt32(dr[0].ToString()));
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

            if (OpenFileDialogOpen.ShowDialog() == DialogResult.OK)
            {
                using (TextFieldParser parser = new TextFieldParser(OpenFileDialogOpen.FileName))
                {
                    parser.CommentTokens = new string[] { "|" };
                    parser.SetDelimiters(new string[] { "," });
                    parser.HasFieldsEnclosedInQuotes = true;

                    // Skip over header line.
                    parser.ReadLine();

                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        CliMat cm = new CliMat();
                        {
                            Name = fields[0],
                            FactoryLocation = fields[1],
                            EstablishedYear = int.Parse(fields[2]),
                            Profit = double.Parse(fields[3], swedishCulture)
                        };
                    }
                }

            }





            string sql = "";

            sql = "  insert into trustledger ([TLMatter],[TLBank] ,[TLType] ,[TLDate] ,[TLCheckNbr] ,[TLAmount] ,[TLMemo], tlsysnbr) " +
        " values(65258, 'T10', 1, convert(DateTime, '02/13/2020', 102), 0, 442.24, 'Added To Ledger By Tool on: ' + convert(varchar, getdate(), 101), (select max(tlsysnbr) + 1 from trustledger)) ";
            _jurisUtility.ExecuteNonQuery(0, sql);

        }

        private void processSingleMatter(int matsys)
        {
            string sql = "";

            sql = "  ";
            _jurisUtility.ExecuteNonQuery(0, sql);


        }

        private bool checkForBalances(int matsysnbr) //true means some balance exists
        {
            string sql = "";
            sql =
                "  select dbo.jfn_FormatClientCode(clicode) as clicode, dbo.jfn_FormatMatterCode(MatCode) as matcode, sum(ppd) as ppd, sum(UT) as UT, sum(UE) as UE, sum(AR) as AR, sum(Trust) as Trust " +
                "from( " +
                "select matsysnbr as matsys, MatPPDBalance as ppd, 0 as UT, 0 as UE, 0 as AR, 0 as Trust " +
                  "from matter " +
                  "where MatPPDBalance <> 0 and matsysnbr = " + matsysnbr +
                  " union all " +
                  "select utmatter as matsys, 0 as ppd, sum(utamount) as UT, 0 as UE, 0 as AR, 0 as Trust " +
                  "from unbilledtime where utmatter = " + matsysnbr + 
                  " group by utmatter " +
                  "having sum(utamount) <> 0 " +
                  "union all " +
                  "select uematter as matsys, 0 as ppd, 0 as UT, sum(ueamount) as UE, 0 as AR, 0 as Trust " +
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
                "inner join matter on hhg.matsys = matsysnbr " +
                "inner join client on clisysnbr = matsysnbr " +
                "  group by matsys ";

            DataSet ds = new DataSet();
            ds = _jurisUtility.RecordsetFromSQL(sql);
            if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                return true;
            else
            {
                ErrorLog er = new ErrorLog();
                er.client = ds.Tables[0].Rows[0][0].ToString();
                er.matter = ds.Tables[0].Rows[0][1].ToString();
                er.message = "Cannot close matter because balance(s) exist: See Below for more detail: \r\n" +
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

        }

        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (TabControl1.SelectedTab == TabControl1.TabPages["Client"])//your specific tabname
                activeTab = 0; //client is active tab
            else
                activeTab = 1; // means matter is active
        }
    }
}
