using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

namespace InvoiceCheck
{
    internal class CommonTool
    {
        /// <summary>
        /// 建立發票資料表的共用欄位結構
        /// </summary>
        public static DataTable CreateInvoiceTable()
        {
            DataTable dt = new DataTable("InvoiceData");
            dt.Columns.Add("InvoiceType", typeof(string)); // 例如: IG, VOID
            dt.Columns.Add("ShopNo", typeof(string));     // 例如: PUB
            dt.Columns.Add("EcrHdKey", typeof(string));   // 例如: PUB0220260100000256
            dt.Columns.Add("InvoiceDate", typeof(string)); // 例如: 2026/01/03
            dt.Columns.Add("InvoiceNumber", typeof(string)); // 例如: XL84791838
            dt.Columns.Add("ProcessDateTime", typeof(string)); // 例如: 2026/01/03
            return dt;
        }

        /// <summary>
        /// 清空 [TradeVanInvoice]，避免重複累計
        /// </summary>
        public static void TruncateTradeVanInvoice()
        {
            string connString = ConfigurationManager.ConnectionStrings["ConnString_TR_DATA"]?.ConnectionString;
            if (string.IsNullOrWhiteSpace(connString))
            {
                throw new InvalidOperationException("Connection string 'ConnString_TR_DATA' not found.");
            }

            using (SqlConnection conn = new SqlConnection(connString))
            using (SqlCommand cmd = new SqlCommand("TRUNCATE TABLE dbo.[TradeVanInvoice]", conn))
            {
                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 依指定發票日期刪除 [TradeVanInvoice] 資料
        /// </summary>
        public static void DeleteTradeVanInvoiceByDate(DateTime invoiceDate)
        {
            string connString = ConfigurationManager.ConnectionStrings["ConnString_TR_DATA"]?.ConnectionString;
            if (string.IsNullOrWhiteSpace(connString))
            {
                throw new InvalidOperationException("Connection string 'ConnString_TR_DATA' not found.");
            }

            using (SqlConnection conn = new SqlConnection(connString))
            using (SqlCommand cmd = new SqlCommand("DELETE FROM dbo.[TradeVanInvoice] WHERE InvoiceDate = @InvoiceDate", conn))
            {
                cmd.Parameters.Add("@InvoiceDate", SqlDbType.NVarChar, 50).Value = invoiceDate.ToString("yyyy/MM/dd");

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 從資料庫取得交易端有開立、但 [TradeVanInvoice]（ACK 匯入）尚未覆蓋的發票
        /// </summary>
        public static DataTable GetMissingInvoices(DateTime targetDate)
        {
            string connString = ConfigurationManager.ConnectionStrings["ConnString_TR_DATA"]?.ConnectionString;
            if (string.IsNullOrWhiteSpace(connString))
            {
                throw new InvalidOperationException("Connection string 'ConnString_TR_DATA' not found.");
            }

            DataTable dt = new DataTable();

            using (SqlConnection conn = new SqlConnection(connString))
            using (SqlCommand cmd = new SqlCommand(@"SELECT TK.ShopNo, TK.EcrHdKey, TK.HTK_NO AS InvoiceNumber
                                                     FROM [RX_V4_TRN].[dbo].EcrTKHs TK
                                                     INNER JOIN [RX_V4_TRN].[dbo].Shop S WITH (NOLOCK)
                                                        ON TK.ShopNo = S.ShopNo AND S.SpGpNo = '6000'
                                                     WHERE TK.HTK_Date = @HtkDate
                                                       and HTK_NO not in (select InvoiceNumber from [TR_DATA].[dbo].[TradeVanInvoice]) and HTK_NO <> '' ", conn))
            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
            {
                cmd.Parameters.Add("@HtkDate", SqlDbType.Date).Value = targetDate.Date;
                adapter.Fill(dt);
            }

            return dt;
        }

        /// <summary>
        /// 解析新版格式檔案，回傳包含門市、完整編號與發票號碼的 DataTable
        /// </summary>
        public static DataTable GetTradevanInvoiceDataTable(string filePath)
        {
            DataTable dt = CreateInvoiceTable();

            try
            {
                if (!File.Exists(filePath)) return dt;

                FileInfo file = new FileInfo(filePath);
                string processType = file.Name.Split('_')[0];


                // 建議指定 Encoding。如果是台灣常見舊系統請用 Encoding.GetEncoding("Big5")
                using (StreamReader sr = new StreamReader(filePath, Encoding.UTF8))
                {
                    string line;
                    if (processType == "IG")
                    {
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (string.IsNullOrWhiteSpace(line)) continue;

                            string[] columns = line.Split('^');

                            // 索引說明：
                            // [3]: 60499576_PUB_20260103... (用來切出 PUB 與日期)
                            // [4]: PUB0220260100000256      (完整編號)
                            // [6]: XL84791838               (發票號碼)

                            if (columns.Length >= 7)
                            {
                                // 1. 提取門市簡稱
                                string[] storeParts = columns[3].Split('_');
                                string storeName = storeParts.Length > 1 ? storeParts[1] : "N/A";

                                // 2. 提取完整編號
                                string storeFullID = columns[4];

                                // 3. 提取發票號碼
                                string invoiceNo = columns[6];

                                // 4. 發票日期 (格式 yyyy/MM/dd)
                                string invoiceDate = string.Empty;
                                if (storeParts.Length > 2)
                                {
                                    string datePart = storeParts[2].Length >= 8 ? storeParts[2].Substring(0, 8) : storeParts[2];
                                    DateTime parsedDate;
                                    if (DateTime.TryParseExact(datePart, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out parsedDate))
                                    {
                                        invoiceDate = parsedDate.ToString("yyyy/MM/dd");
                                    }
                                }

                                string processDateTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                                dt.Rows.Add(processType, storeName, storeFullID, invoiceDate, invoiceNo, processDateTime);
                            }
                        }

                    }
                    else if (processType == "VOID")
                    {
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (string.IsNullOrWhiteSpace(line)) continue;

                            string[] columns = line.Split('^');
                            // 格式範例：
                            // 00001^Y^^R^60499576_PUB_20260107_PUB0220260100001185^PUB0220260100001185^20260107^XL84792751^^^^^^
                            // 索引對應：
                            // [5]: 60499576_PUB_20260107_PUB0220260100001185 (含門市與日期)
                            // [6]: PUB0220260100001185                         (完整編號)
                            // [7]: 20260107                                    (發票日期)
                            // [8]: XL84792751                                   (發票號碼)

                            if (columns.Length >= 9)
                            {
                                // 門市
                                string[] storeParts = columns[4].Split('_');
                                string storeName = storeParts.Length > 1 ? storeParts[1] : "N/A";

                                // 完整編號
                                string storeFullID = columns[5];

                                // 發票號碼
                                string invoiceNo = columns[7];

                                // 發票日期 (格式 yyyy/MM/dd)
                                string invoiceDate = columns[6];

                                string processDateTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                                dt.Rows.Add(processType, storeName, storeFullID, invoiceDate, invoiceNo, processDateTime);
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw new Exception("解析失敗：" + ex.Message);
            }

            return dt;
        }

        /// <summary>
        /// 將發票資料寫入 [TradeVanInvoice] 資料表（保留原有功能）
        /// </summary>
        public static void SaveInvoicesToTradeVanInvoice(DataTable invoiceData)
        {
            if (invoiceData == null || invoiceData.Rows.Count == 0) return;

            string connString = ConfigurationManager.ConnectionStrings["ConnString_TR_DATA"]?.ConnectionString;
            if (string.IsNullOrWhiteSpace(connString))
            {
                throw new InvalidOperationException("Connection string 'ConnString_TR_DATA' not found.");
            }

            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                using (SqlTransaction tran = conn.BeginTransaction())
                using (SqlCommand cmd = new SqlCommand(@"INSERT INTO [TradeVanInvoice] (InvoiceType, ShopNo, EcrHdKey, InvoiceDate, InvoiceNumber, ProcessDateTime)
                                                         VALUES (@InvoiceType, @ShopNo, @EcrHdKey, @InvoiceDate, @InvoiceNumber, @ProcessDateTime)", conn, tran))
                {
                    cmd.Parameters.Add("@InvoiceType", SqlDbType.NVarChar, 10);
                    cmd.Parameters.Add("@ShopNo", SqlDbType.NVarChar, 50);
                    cmd.Parameters.Add("@EcrHdKey", SqlDbType.NVarChar, 100);
                    cmd.Parameters.Add("@InvoiceDate", SqlDbType.NVarChar, 50);
                    cmd.Parameters.Add("@InvoiceNumber", SqlDbType.NVarChar, 50);
                    cmd.Parameters.Add("@ProcessDateTime", SqlDbType.NVarChar, 50);

                    try
                    {
                        foreach (DataRow row in invoiceData.Rows)
                        {
                            cmd.Parameters["@InvoiceType"].Value = ToDbValue(row, "InvoiceType");
                            cmd.Parameters["@ShopNo"].Value = ToDbValue(row, "ShopNo");
                            cmd.Parameters["@EcrHdKey"].Value = ToDbValue(row, "EcrHdKey");
                            cmd.Parameters["@InvoiceDate"].Value = ToDbValue(row, "InvoiceDate");
                            cmd.Parameters["@InvoiceNumber"].Value = ToDbValue(row, "InvoiceNumber");
                            cmd.Parameters["@ProcessDateTime"].Value = ToDbValue(row, "ProcessDateTime");
                            cmd.ExecuteNonQuery();
                        }

                        tran.Commit();
                    }
                    catch
                    {
                        tran.Rollback();
                        throw;
                    }
                }
            }
        }

        private static object ToDbValue(DataRow row, string columnName)
        {
            if (row == null || !row.Table.Columns.Contains(columnName)) return DBNull.Value;

            object value = row[columnName];
            if (value == null || value == DBNull.Value) return DBNull.Value;

            string strValue = value.ToString();
            return string.IsNullOrWhiteSpace(strValue) ? (object)DBNull.Value : value;
        }
        /// <summary>
        /// 取得指定日期、SpGpNo=6000 的交易端發票資料
        /// </summary>
        public static DataTable GetEcrInvoices(DateTime targetDate)
        {
            string connString = ConfigurationManager.ConnectionStrings["ConnString_TR_DATA"]?.ConnectionString;
            if (string.IsNullOrWhiteSpace(connString))
            {
                throw new InvalidOperationException("Connection string 'ConnString_TR_DATA' not found.");
            }

            DataTable dt = CreateInvoiceTable();

            using (SqlConnection conn = new SqlConnection(connString))
            using (SqlCommand cmd = new SqlCommand(@"SELECT TK.ShopNo, TK.EcrHdKey, TK.HTK_NO AS InvoiceNumber
                                                     FROM [RX_V4_TRN].[dbo].EcrTKHs TK
                                                     INNER JOIN [RX_V4_TRN].[dbo].Shop S WITH (NOLOCK)
                                                         ON TK.ShopNo = S.ShopNo AND S.SpGpNo = '6000'
                                                     WHERE TK.HTK_Date = @HtkDate", conn))
            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
            {
                cmd.Parameters.Add("@HtkDate", SqlDbType.Date).Value = targetDate.Date;
                adapter.Fill(dt);
            }

            return dt;
        }

        /// <summary>
        /// 取得交易端有開立但 ACK 檔未出現的發票（純記憶體比對，不寫 DB）
        /// </summary>
        public static DataTable FindMissingInvoices(DataTable ecrInvoices, DataTable ackInvoices)
        {
            if (ecrInvoices == null) return CreateInvoiceTable();
            if (ackInvoices == null) ackInvoices = CreateInvoiceTable();

            HashSet<string> ackKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow row in ackInvoices.Rows)
            {
                var key = row["EcrHdKey"]?.ToString();
                if (!string.IsNullOrWhiteSpace(key)) ackKeys.Add(key);
            }

            DataTable missing = ecrInvoices.Clone();

            foreach (DataRow row in ecrInvoices.Rows)
            {
                var key = row["EcrHdKey"]?.ToString();
                if (string.IsNullOrWhiteSpace(key)) continue;

                if (!ackKeys.Contains(key))
                {
                    missing.ImportRow(row);
                }
            }

            return missing;
        }

        /// <summary>
        /// 匯出缺漏清單至檔案，便於人工檢查
        /// </summary>
        public static void SaveMissingInvoices(string filePath, DataTable missingInvoices)
        {
            if (missingInvoices == null || missingInvoices.Rows.Count == 0) return;

            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                string[] columns = new[] { "ShopNo", "EcrHdKey", "InvoiceNumber" };

                // 標題列
                sw.WriteLine(string.Join(",", columns));

                // 資料列
                foreach (DataRow row in missingInvoices.Rows)
                {
                    var values = columns.Select(col =>
                    {
                        if (missingInvoices.Columns.Contains(col))
                            return row[col]?.ToString()?.Replace(",", " ") ?? string.Empty;
                        return string.Empty;
                    });

                    sw.WriteLine(string.Join(",", values));
                }
            }
        }
    }
}
