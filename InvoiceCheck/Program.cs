using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace InvoiceCheck
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path = "D:\\TIF_eInvoice\\tmp\\";

            Stopwatch sw = Stopwatch.StartNew();
            int totalMissing = 0;
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine($"Start: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

            CommonTool.TruncateTempInv();

            for (int i = 1; i <= 4; i++)
            {
                DateTime dateTime = DateTime.Now.AddDays(-i);

                DirectoryInfo dirInfo = new DirectoryInfo(path + $"deq\\{dateTime:yyyyMMdd}");
                if (!dirInfo.Exists) continue;

                DirectoryInfo[] hourfileInfos = dirInfo.GetDirectories();

                DataTable ackInvoices = CommonTool.CreateInvoiceTable();

                foreach (var dir in hourfileInfos)
                {
                    FileInfo[] ackFiles = dir.GetFiles("*_ACK.txt");
                    if (ackFiles == null || ackFiles.Length == 0) continue;

                    foreach (var fileInfo in ackFiles)
                    {
                        if (fileInfo == null || !fileInfo.Exists) continue;

                        DataTable invoiceData = CommonTool.GetTradevanInvoiceDataTable(fileInfo.FullName);
                        ackInvoices.Merge(invoiceData);
                    }
                }

                // 將 ACK 資料寫入 TempInv，供後續 DB 比對
                CommonTool.SaveInvoicesToTempInv(ackInvoices);

                // TempInv 尋找的缺漏比對
                DataTable missingInvoices = CommonTool.GetMissingInvoices(dateTime);
                logBuilder.AppendLine($"Date {dateTime:yyyyMMdd}: missing {missingInvoices.Rows.Count}");
                totalMissing += missingInvoices.Rows.Count;

                if (missingInvoices.Rows.Count > 0)
                {
                    string missDir = Path.Combine(path, "Miss");
                    Directory.CreateDirectory(missDir);
                    string missingReportPath = Path.Combine(missDir, $"MissingInvoices_{dateTime:yyyyMMdd}.csv");
                    CommonTool.SaveMissingInvoices(missingReportPath, missingInvoices);
                }
            }

            sw.Stop();
            logBuilder.AppendLine($"Total missing: {totalMissing}");
            logBuilder.AppendLine($"Elapsed: {sw.Elapsed}");
            logBuilder.AppendLine($"End: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

            string logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            Directory.CreateDirectory(logDir);
            string logPath = Path.Combine(logDir, $"{DateTime.Now:yyyyMMdd}.log");
            string content = logBuilder.ToString();
            if (File.Exists(logPath))
            {
                content = Environment.NewLine + content;
            }
            File.AppendAllText(logPath, content);
        }
    }
}
