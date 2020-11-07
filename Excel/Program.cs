using System;
using System.Linq;
using IronXL;

namespace Excel
{
    class Program
    {

        static void Main(string[] args)
        {
            WorkBook wb = WorkBook.Load("sample.xlsx"); //load Excel file 
            WorkSheet ws = wb.GetWorkSheet("Sheet1"); //Get sheet1 of sample.xlsx
            ws["A1"].Value = "FirstName"; //access A1 cell and edit the value
            ws["B1"].Value = "LastName";
            for (int i = 2; i < 4; i++)
            {
                ws["A" + i].Value = Console.ReadLine();
                ws["B" + i].Value = Console.ReadLine();
            }
            wb.SaveAs("sample.xlsx");   //save changes

            wb = WorkBook.Load("sample.xlsx");
            ws = wb.GetWorkSheet("Sheet1");

            for (int i = 2; i < 4; i++)
            {
                foreach (var firstname in ws["A"+i])
                {
                    foreach (var lastname in ws["B"+i])
                    {
                        Console.WriteLine("FirstName: {0},LastName:{1}", firstname.Text,lastname.Text);
                    }
                    
                }
            }
            
        }

    }
}
