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
            dt.Columns.Add("ShopNo", typeof(string));     // 例如: PUB
            dt.Columns.Add("EcrHdKey", typeof(string));   // 例如: PUB0220260100000256
            dt.Columns.Add("InvoiceNumber", typeof(string)); // 例如: XL84791838
            return dt;
        }

        /// <summary>
        /// 清空 TempInv，避免重複累計
        /// </summary>
        public static void TruncateTempInv()
        {
            string connString = ConfigurationManager.ConnectionStrings["ConnString_TR_DATA"]?.ConnectionString;
            if (string.IsNullOrWhiteSpace(connString))
            {
                throw new InvalidOperationException("Connection string 'ConnString_TR_DATA' not found.");
            }

            using (SqlConnection conn = new SqlConnection(connString))
            using (SqlCommand cmd = new SqlCommand("TRUNCATE TABLE dbo.TempInv", conn))
            {
                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 從資料庫取得交易端有開立、但 TempInv（ACK 匯入）尚未覆蓋的發票
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
                                                       and HTK_NO not in (select InvoiceNumber from [TR_DATA].[dbo].TempInv) and HTK_NO <> '' ", conn))
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

                // 建議指定 Encoding。如果是台灣常見舊系統請用 Encoding.GetEncoding("Big5")
                using (StreamReader sr = new StreamReader(filePath, Encoding.UTF8))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (string.IsNullOrWhiteSpace(line)) continue;

                        string[] columns = line.Split('^');

                        // 索引說明：
                        // [3]: 60499576_PUB_20260103... (用來切出 PUB)
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

                            dt.Rows.Add(storeName, storeFullID, invoiceNo);
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
        /// 將發票資料寫入 TempInv 資料表（保留原有功能）
        /// </summary>
        public static void SaveInvoicesToTempInv(DataTable invoiceData)
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
                using (SqlCommand cmd = new SqlCommand(@"INSERT INTO TempInv (ShopNo, EcrHdKey, InvoiceNumber, CreateDate)
                                                         VALUES (@ShopNo, @EcrHdKey, @InvoiceNumber, @CreateDate)", conn, tran))
                {
                    cmd.Parameters.Add("@ShopNo", SqlDbType.NVarChar, 50);
                    cmd.Parameters.Add("@EcrHdKey", SqlDbType.NVarChar, 100);
                    cmd.Parameters.Add("@InvoiceNumber", SqlDbType.NVarChar, 50);
                    cmd.Parameters.Add("@CreateDate", SqlDbType.NVarChar, 50);

                    try
                    {
                        foreach (DataRow row in invoiceData.Rows)
                        {
                            cmd.Parameters["@ShopNo"].Value = row["ShopNo"] ?? DBNull.Value;
                            cmd.Parameters["@EcrHdKey"].Value = row["EcrHdKey"] ?? DBNull.Value;
                            cmd.Parameters["@InvoiceNumber"].Value = row["InvoiceNumber"] ?? DBNull.Value;
                            cmd.Parameters["@CreateDate"].Value = DateTime.Now.ToString("yyyy/MM/dd");
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
