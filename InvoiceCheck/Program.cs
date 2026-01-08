using System;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;

namespace InvoiceCheck
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path = ConfigurationManager.AppSettings["TIF_Path"].ToString();
            string daysSetting = ConfigurationManager.AppSettings["Days"];
            int dayCount = 1;
            int.TryParse(daysSetting, out dayCount);
            if (dayCount < 1) dayCount = 1;

            Stopwatch sw = Stopwatch.StartNew();
            int totalMissing = 0;
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine($"Start: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

            for (int i = 1; i <= dayCount; i++)
            {
                DateTime dateTime = DateTime.Now.AddDays(-i);
                Console.WriteLine($"正在執行 {dateTime:yyyy/MM/dd} 檢測中...");

                string targetDir = path + $"deq\\{dateTime:yyyyMMdd}";
                DirectoryInfo dirInfo = new DirectoryInfo(targetDir);
                Console.WriteLine($"檢測資料夾: {dirInfo.FullName}");
                if (!dirInfo.Exists)
                {
                    Console.WriteLine("  資料夾不存在，跳過。");
                    continue;
                }

                DirectoryInfo[] hourfileInfos = dirInfo.GetDirectories();
                Console.WriteLine($"  發現子資料夾數: {hourfileInfos.Length}");

                DataTable ackInvoices = CommonTool.CreateInvoiceTable();

                foreach (var dir in hourfileInfos)
                {
                    Console.WriteLine($"  讀取子資料夾: {dir.FullName}");
                    FileInfo[] ackFiles = dir.GetFiles("*_ACK.txt");
                    if (ackFiles == null || ackFiles.Length == 0)
                    {
                        Console.WriteLine("    沒有找到 *_ACK.txt 檔案");
                        continue;
                    }

                    Console.WriteLine($"    找到 {ackFiles.Length} 個 ACK 檔案");

                    foreach (var fileInfo in ackFiles)
                    {
                        if (fileInfo == null || !fileInfo.Exists) continue;

                        Console.WriteLine($"    處理檔案: {fileInfo.FullName}");
                        DataTable invoiceData = CommonTool.GetTradevanInvoiceDataTable(fileInfo.FullName);
                        ackInvoices.Merge(invoiceData);
                    }
                }

                // 將 ACK 資料寫入 [TradeVanInvoice]，供後續 DB 比對（先刪除該日舊資料避免累積）
                if (ackInvoices.Rows.Count > 0)
                {
                    Console.WriteLine($"  寫入 {ackInvoices.Rows.Count} 筆 ACK 資料到資料庫");
                    //CommonTool.DeleteTradeVanInvoiceByDate(dateTime);
                    CommonTool.SaveInvoicesToTradeVanInvoice(ackInvoices);
                }
                else
                {
                    Console.WriteLine("  無 ACK 資料可寫入");
                }

                // [TradeVanInvoice] 尋找的缺漏比對
                DataTable missingInvoices = CommonTool.GetMissingInvoices(dateTime);
                logBuilder.AppendLine($"Date {dateTime:yyyyMMdd}: missing {missingInvoices.Rows.Count}");
                Console.WriteLine($"  缺漏比對完成，筆數: {missingInvoices.Rows.Count}");
                totalMissing += missingInvoices.Rows.Count;

                if (missingInvoices.Rows.Count > 0)
                {
                    string missDir = Path.Combine(path, "Miss");
                    Directory.CreateDirectory(missDir);
                    string missingReportPath = Path.Combine(missDir, $"MissingInvoices_{dateTime:yyyyMMdd}.csv");
                    CommonTool.SaveMissingInvoices(missingReportPath, missingInvoices);
                    Console.WriteLine($"  已輸出缺漏清單: {missingReportPath}");
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

            Console.WriteLine("執行完成...");
            Thread.Sleep(TimeSpan.FromMinutes(1));
        }
    }
}
