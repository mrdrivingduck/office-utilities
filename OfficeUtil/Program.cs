using Office.Utils.PasswdStructure;
using System.Collections.Generic;
using System;

namespace Office
{
    class Program
    {

        static void Main(string[] args)
        {
            string root = @"C:\Users\Jingtang Zhang\Desktop\应收预收明细表.xls";
            /*List<string> passwds = new List<string>();
            passwds.Add("168168");
            passwds.Add("88");
            passwds.Add("168");
            passwds.Add("071010");
            passwds.Add("1230");
            passwds.Add("1988");
            passwds.Add("7756");
            passwds.Add("668");
            passwds.Add("888");
            passwds.Add("833");
            passwds.Add("050901");
            passwds.Add("202");
            passwds.Add("202202");
            passwds.Add("fin88");
            passwds.Add("1002");
            passwds.Add("abc");
            passwds.Add("0717");
            passwds.Add("1108");
            passwds.Add("1103");
            passwds.Add("2008");
            Utils.ExcelPasswordRemover.RemovePassword(root, passwds);*/

            Utils.ExcelPasswordRemover.RemovePassword(root, "88");

        }

    }
    
}
