using OfficeOpenXml;
using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Bookings
{
    enum ExitCode { UnhandledException = 1, SftpError };

    class SFTP
    {
        public static SftpClient Connect(NameValueCollection appSettings, log4net.ILog log)
        {
            ConnectionInfo connectionInfo = new ConnectionInfo(appSettings["SFTP_Host"].ToString(), appSettings["SFTP_User"].ToString(), new PasswordAuthenticationMethod(appSettings["SFTP_User"], appSettings["SFTP_Pass"]));
            SftpClient client = new SftpClient(connectionInfo);

            try
            {
                client.Connect();
            }

            catch (Exception error)
            {
                log.Fatal("Unable to connect to SFTP site.", error);
                Environment.Exit((int)ExitCode.SftpError);
            }

            return client;
        }
    }

    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static readonly NameValueCollection appSettings = ConfigurationManager.AppSettings;

        static DataTable ConvertToDataTable(ExcelWorksheet workSheet, bool hasHeader)
        {
            DataTable dataTable = new DataTable();

            foreach (ExcelRangeBase firstRowCell in workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column])
            {
                dataTable.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }

            int startRow = hasHeader ? 2 : 1;

            for (int rowNum = startRow; rowNum <= workSheet.Dimension.End.Row; rowNum++)
            {
                ExcelRange wsRow = workSheet.Cells[rowNum, 1, rowNum, workSheet.Dimension.End.Column];
                DataRow row = dataTable.Rows.Add();

                foreach (ExcelRangeBase cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
            }

            return dataTable;
        }

        static void Main(string[] args)
        {
            SftpClient client = SFTP.Connect(appSettings, log);
            IEnumerable<Renci.SshNet.Sftp.SftpFile> files = client.ListDirectory(appSettings["InputFolder"]);
            List<Renci.SshNet.Sftp.SftpFile> filesToProcess = new List<Renci.SshNet.Sftp.SftpFile>();
            Regex inputPattern = new Regex(appSettings["InputPattern"], RegexOptions.Compiled | RegexOptions.IgnoreCase);

            foreach (Renci.SshNet.Sftp.SftpFile file in files)
            {
                if (inputPattern.IsMatch(file.Name))
                {
                    filesToProcess.Add(file);
                }
            }

            ExcelPackage targetFile = new ExcelPackage();
            ExcelWorksheet targetWorkSheet = targetFile.Workbook.Worksheets.Add("Sheet1");
            DataTable targetDataTable = new DataTable();
            string[] targetColumnNames = { "DealerName", "DealerCode", "JobCardNumber", "BookingCreatedDate", "JobStartDateTime", "JobEndDateTime", "Regno", "Vin", "MakeModel", "BookingAdvisor", "JobSequence", "OperationCode", "OpCodeDesc", "TotalJobTime", "isMOT", "isWarranty",  "Mobile", "Email", "NextServiceDate", "NextMotDate" };

            foreach (string columnName in targetColumnNames)
            {
                targetDataTable.Columns.Add(new DataColumn(columnName));
            }

            foreach (Renci.SshNet.Sftp.SftpFile file in filesToProcess)
            {
                MemoryStream ms = new MemoryStream();
                client.DownloadFile(String.Join("/", appSettings["InputFolder"], file.Name), ms);
                ExcelPackage sourceFile = new ExcelPackage(ms);
                ExcelWorksheet sourceWorkSheet = sourceFile.Workbook.Worksheets.First();
                DataTable sourceDataTable = ConvertToDataTable(sourceWorkSheet, true);

                foreach (DataRow sourceRow in sourceDataTable.Rows)
                {
                    DataRow targetRow = targetDataTable.NewRow();
                    targetRow["DealerName"] = sourceRow["SalesLocationName"].ToString();
                    targetRow["DealerCode"] = sourceRow["SalesLocationName"].ToString();
                    targetRow["JobCardNumber"] = "";
                    targetRow["BookingCreatedDate"] = sourceRow["BookingCreateDate"].ToString();
                    targetRow["JobStartDateTime"] = sourceRow["BookingDate"].ToString();
                    targetRow["JobEndDateTime"] = sourceRow["BookingDate"].ToString();
                    targetRow["Regno"] = sourceRow["Regno"].ToString();
                    targetRow["Vin"] = "";
                    targetRow["MakeModel"] = sourceRow["VehicleMake"].ToString();
                    targetRow["BookingAdvisor"] = "";
                    targetRow["JobSequence"] = "";
                    targetRow["OperationCode"] = "";
                    targetRow["OpCodeDesc"] = sourceRow["BookingType"].ToString().ToUpper().Contains("SERVICE") ? "SERVICE" : "";
                    targetRow["TotalJobTime"] = "0";
                    targetRow["isMOT"] = sourceRow["BookingType"].ToString().ToUpper().Contains("MOT") ? "1" : "0";
                    targetRow["isWarranty"] = "";
                    targetRow["Mobile"] = "";
                    targetRow["Email"] = sourceRow["Email"].ToString();
                    targetRow["NextServiceDate"] = "";
                    targetRow["NextMotDate"] = "";
                    targetDataTable.Rows.Add(targetRow);
                }

                file.MoveTo(String.Join("/", appSettings["InputFolder"], "archived", file.Name));
            }

            if (filesToProcess.Count > 0)
            {
                targetWorkSheet.Cells["A1"].LoadFromDataTable(targetDataTable, true);
                byte[] output = EpplusCsvConverter.ConvertToCsv(targetFile);

                try
                {
                    client.RenameFile(appSettings["OutputFolder"] + appSettings["OutputFile"], appSettings["OutputFolder"] + "/archived.bookings.csv", true);
                }

                catch(Exception error)
                {
                    log.Error("Unable to rename existing file.", error);
                }

            client.WriteAllBytes(appSettings["OutputFolder"] + appSettings["OutputFile"], output);
        }

        client.Disconnect();
        }
    }
}
