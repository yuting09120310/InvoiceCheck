using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;

namespace InvoiceCheck
{
    internal class CommonTool
    {

        /// <summary>
        /// 解析新版格式檔案，回傳包含門市、完整編號與發票號碼的 DataTable
        /// </summary>
        public static DataTable GetTradevanInvoiceDataTable(string filePath)
        {
            DataTable dt = new DataTable("InvoiceData");
            dt.Columns.Add("ShopNo", typeof(string));     // 例如: PUB
            dt.Columns.Add("EcrHdKey", typeof(string));   // 例如: PUB0220260100000256
            dt.Columns.Add("InvoiceNumber", typeof(string)); // 例如: XL84791838

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
        /// 將發票資料寫入 TempInv 資料表
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

    }
}
