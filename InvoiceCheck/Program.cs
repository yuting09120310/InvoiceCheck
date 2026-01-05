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
            DateTime dateTime = DateTime.Now.AddDays(-2);

            DirectoryInfo dirInfo = new DirectoryInfo($"D:\\TIF_eInvoice\\tmp\\deq\\{dateTime.ToString("yyyyMMdd")}");
            if (!dirInfo.Exists) return;

            DirectoryInfo[] hourfileInfos = dirInfo.GetDirectories();

            foreach (var dir in hourfileInfos)
            {
                FileInfo fileInfo = dir.GetFiles("*_ACK.txt").FirstOrDefault();
                if (fileInfo == null || !fileInfo.Exists) continue;

                DataTable invoiceData = CommonTool.GetTradevanInvoiceDataTable(fileInfo.FullName);
                CommonTool.SaveInvoicesToTempInv(invoiceData);
            }
        }
    }
}
