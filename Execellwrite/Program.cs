using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Execellwrite
{
    public class Program
    {
        static void Main(string[] args)
        {
            Employee employee = new Employee();
            employee.ExcelFile();
            
        }
    }
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public DateTime DOB { get; set; }

        public void ExcelFile()
        {

            // Creating an instance
            // of ExcelPackage
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage excel = new ExcelPackage();

            // name of the sheet
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

            // setting the properties
            // of the work sheet 
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            // Setting the properties
            // of the first row
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            // Header of the Excel sheet
            workSheet.Cells[1, 1].Value = "Name";
            workSheet.Cells[1, 2].Value = "Age";
            workSheet.Cells[1, 3].Value = "DOB";


            // file name with .xlsx extension 
            string p_strPath = @"C:\Users\user\source\repos\Execellwrite\Execellwrite\StaticFile\Testing.xlsx";

            if (File.Exists(p_strPath))
                File.Delete(p_strPath);

            // Create excel file on physical disk 
            FileStream objFileStrm = File.Create(p_strPath);
            objFileStrm.Close();

            // Write content to excel file 
            File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
            //Close Excel package
            excel.Dispose();
            Console.WriteLine("Submit");
            Console.ReadKey();
        }
   }
}