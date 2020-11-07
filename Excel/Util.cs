using IronXL;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excel
{
    class Util
    {
        public static void WriteExcel(out WorkBook wb, out WorkSheet ws)
        {
            wb = WorkBook.Load("sample.xlsx");
            ws = wb.GetWorkSheet("Sheet1");
            ws["A1"].Value = "FirstName"; //access A1 cell and edit the value
            ws["B1"].Value = "LastName";
            for (int i = 2; i < 4; i++)
            {
                ws["A" + i].Value = Console.ReadLine();
                ws["B" + i].Value = Console.ReadLine();
            }
            wb.SaveAs("sample.xlsx");   //save changes
        }
        public static void ReadExcel(out WorkBook wb, out WorkSheet ws)
        {
            wb = WorkBook.Load("sample.xlsx");
            ws = wb.GetWorkSheet("Sheet1");

            for (int i = 2; i < 4; i++)
            {
                foreach (var firstname in ws["A" + i])
                {
                    foreach (var lastname in ws["B" + i])
                    {
                        Console.WriteLine("FirstName: {0},LastName:{1}", firstname.Text, lastname.Text);
                    }

                }
            }
        }


    }

}
