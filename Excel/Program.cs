//Author: Phuong Nguyen
//Date: 9:36PM 10/06/2020
//Project: Writing and Reading in C# into Excel File with IronXL
using System;
using System.Linq;
using IronXL;

namespace Excel
{
    class Program
    {

        static void Main(string[] args)
        {
            WorkBook wb;
            WorkSheet ws;
            Util.WriteExcel(out wb, out ws);
            Util.ReadExcel(out wb, out ws);

        }




    }
}
