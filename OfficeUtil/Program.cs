using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace Office
{
    class Program
    {

        static void Main(string[] args)
        {
            string root = @"C:\Users\Jingtang Zhang\Desktop\excel";

            Utils.ExcelPasswordRemover.RemovePassword(root, "0123456789", 3);
            
        }

    }
    
}
