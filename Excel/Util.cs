using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;
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
            for (int i = ws.Rows.Count()+1; i <= ws.Rows.Count()+2; i++)
            {                
               
                Console.WriteLine("Do you want to add a record? y/n");
                var answer = Console.ReadLine();
                if (answer == "n")
                {
                    wb.SaveAs("sample.xlsx");   //save changes
                    Console.WriteLine("Saved");
                    for (i = 2; i <= ws.Rows.Count(); i++)
                    {
                        foreach (var firstname in ws["A" + i])
                        {
                            foreach (var lastname in ws["B" + i])
                            {
                                Console.WriteLine("FirstName: {0},LastName:{1}", firstname.Text, lastname.Text);
                                
                            }

                        }                       
                    }
                    Environment.Exit(0);
                }
                Console.WriteLine("What is your firstname?");
                ws["A" + i].Value = Console.ReadLine();
                Console.WriteLine("What is your lastname?");
                ws["B" + i].Value = Console.ReadLine();
            }           
        }
    }

}
