using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWriting
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = excelApp.ActiveSheet;
            worksheet.Cells[1, "A"] = "This is written from C#";
            worksheet.Cells[2, "B"] = "by SalmanMKC";
            worksheet.Cells[3, "C"] = "Hello Twitter 🐱‍👤";
        }
    }
}
