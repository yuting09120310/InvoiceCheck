using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceCheck
{
    internal class Program
    {
        static void Main(string[] args)
        {
            CommonTool commonTool = new CommonTool();
            string filePath = "D:\\TEMP\\IG_3.0_60499576_PUB_20260103_005_ACK.txt"; // 替換為實際檔案路徑
            DataTable invoiceData = CommonTool.GetTradevanInvoiceDataTable(filePath);
        }
    }
}
