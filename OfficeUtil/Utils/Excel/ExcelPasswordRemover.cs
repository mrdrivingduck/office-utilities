///
/// Author - Mr Dk.
/// Version - 2019/08/29
/// Description -
///     Remove the document with an existing password
///     or with an approximate range of characters
///     and return true/false for successful or not
///     

using Office.Utils.PasswdStructure;
using System;
using System.IO;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace Office.Utils
{

    class ExcelPasswordRemover
    {
        const int SUCCESS = 0;
        const int FAILED = 1;
        const int IGNORED = 2;

        public static void RemovePassword(string path, string password)
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="legalCharacters"></param>
        /// <param name="length"></param>
        public static void RemovePassword(string path, string legalCharacters, int length)
        {
            Excel.Application app = new Excel.Application();
            PasswdGenerator store = new PasswdGenerator(legalCharacters, length);

            if (Directory.Exists(path))
            {
                DoDirPassword(app, path, store);
            }
            else if (IsExcelFile(path))
            {
                DoFilePassword(app, path, store);
            }
            else
            {
                Console.WriteLine("Invalid file format");
            }
        }

        /// <summary>
        ///     Remove the password of specific Excel book
        /// </summary>
        /// <param name="app"></param>
        /// <param name="path"></param>
        /// <param name="passwd"></param>
        /// <returns></returns>
        private static int CrackFile(Excel.Application app, string path, string passwd)
        {
            int status = FAILED;
            
            Thread unlockThread = new Thread(new ThreadStart(() => {

                Excel.Workbook book = null;
                try
                {
                    book = app.Workbooks.Open(path, Password: passwd);
                    if (book.Password == "********" && passwd != "********")
                    {
                        status = IGNORED;
                    }
                    else if (book.Password != "")
                    {
                        book.Password = "";
                        book.Save();
                        status = SUCCESS;
                    }
                    book.Close();
                }
                catch (Exception)
                {
                    // Console.WriteLine(e);
                    // Console.Write("\r    password [" + passwd + "] failed, retrying...");
                }

            }));
            unlockThread.IsBackground = true;
            unlockThread.Start();
            unlockThread.Join();
            return status;
        }

        /// <summary>
        ///     Remove the password of file with a list of password
        /// </summary>
        /// <param name="app"></param>
        /// <param name="path"></param>
        /// <param name="store"></param>
        /// <returns></returns>
        private static int DoFilePassword(Excel.Application app, string path, PasswordStore store)
        {
            Console.WriteLine(path + " : start cracking ... ");
            store.Reset();
            while (store.HasNext())
            {
                string passwd = store.Next();
                int status = CrackFile(app, path, passwd);
                if (status == IGNORED)
                {
                    Console.WriteLine("\r" + path + " : no password");
                    return IGNORED;
                }
                else if (status == SUCCESS)
                {
                    Console.WriteLine("\r" + path + " : Crack success with password - [" + passwd + "]");
                    return SUCCESS;
                }
                else {
                    Console.Write("\r    password [" + passwd + "] failed, retrying...");
                }
            }
            Console.WriteLine("\r" + path + " : Failed to crack ... ");
            return FAILED;
        }

        /// <summary>
        ///     Remove password of all the files with a list of password
        /// </summary>
        /// <param name="app"></param>
        /// <param name="path"></param>
        /// <param name="store"></param>
        private static void DoDirPassword(Excel.Application app, string path, PasswordStore store)
        {
            Console.WriteLine("In path - " + path);

            string[] subPaths = Directory.GetFileSystemEntries(path);
            foreach (string subPath in subPaths)
            {
                if (Directory.Exists(subPath))
                {
                    DoDirPassword(app, subPath, store);
                }
                else if (IsExcelFile(subPath))
                {
                    DoFilePassword(app, subPath, store);
                }
            }

            Console.WriteLine("Out path - " + path);
        }

        private static bool IsExcelFile(string path)
        {
            return File.Exists(path) && (path.EndsWith(".xls") || path.EndsWith(".xlsx"));
        }
    }
}
