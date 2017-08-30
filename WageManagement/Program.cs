using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace WageManagement
{
    class Program
    {
        static void Main(string[] args)
        {


            ObservableCollection<Employee> employeeList = new ObservableCollection<Employee>();

           //Application xlApp = new Application();

           //if (xlApp == null)
           //{
           //    Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
           //    return;
           //}
           //xlApp.Visible = true;
           // xlApp.Workbooks.Open("..\\EmployeeAttendance.xls");


            Application exlApp= new Application() ;
            Workbook xlWorkBook = exlApp.Workbooks.Open(@"d:\EmployeeAttendance.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0); ;
            Worksheet xlWorkSheet = xlWorkBook.Worksheets.get_Item(1); 


        }
    }
}
