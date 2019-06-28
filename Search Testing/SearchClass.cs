using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Search_Testing
{
    class SearchClass
    {
        static List<string> filesThatConstainSSN = new List<string>();

        public static void Main(string[] args)
        {
            String directorySearch;
            List<string> docFiles = new List<string>();
            List<string> excelFiles = new List<string>();
            List<string> pdfFiles = new List<string>();

            //random comment

            //Console Enter
            Console.WriteLine("Enter File Path");
            directorySearch = Console.ReadLine();
            Console.WriteLine("***************************************");



            //Word Documents
            docFiles = Directory.GetFiles(directorySearch, "*.doc", SearchOption.AllDirectories).ToList();

            foreach (string file in docFiles)
            {
                Word.Application app = new Word.Application();

                //^?^?^?  -   ^?^?  -  ^?^?^?^?
                //the ^? finds any digit and the dash makes sure you get the correct form
                FindWord(app, "^#^#^#-^#^#-^#^#^#^#", file);
                Console.WriteLine(file);

            }


            //Excel Documents
            excelFiles = Directory.GetFiles(directorySearch, "*.xlsx", SearchOption.AllDirectories).ToList();

            foreach (string file in excelFiles)
            {
                //? is used for any charcter in excel
                FindExcel("???-??-????", file);

                Console.WriteLine(file);
            }




            pdfFiles = Directory.GetFiles(directorySearch, "*.pdf", SearchOption.AllDirectories).ToList();

            foreach (string file in pdfFiles)
            {
                Console.WriteLine(file);
            }

            Console.WriteLine("***************************************");
            Console.WriteLine("Files That Contain SSN Formating: ");
            Console.WriteLine("");
            filesThatConstainSSN.ForEach(Console.WriteLine);

            Console.ReadLine();
            Console.ReadLine();
        }






        public static void FindWord(Word.Application WordApp, object findText, string Wfile)
        {

            Word.Application objWordApp = new Word.Application();

            objWordApp.Visible = false;
            object missing = System.Reflection.Missing.Value;

            object filename = Wfile;

            Microsoft.Office.Interop.Word.Document objDoc;
            objDoc = objWordApp.Documents.Open(ref filename, ref missing, ref missing, ref missing,
         ref missing, ref missing, ref missing, ref missing,
         ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing);

            objDoc.Content.Find.ClearFormatting();
            try
            {
                if (objDoc.Content.Find.Execute(ref findText,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing))
                {
                    //Finds Text
                    filesThatConstainSSN.Add(Wfile);
                }
                else
                {
                    //Moves on
                }
                objDoc.Close(ref missing, ref missing, ref missing);
                objWordApp.Application.Quit(ref missing, ref missing, ref missing);
            }
            catch (Exception ex)
            {
                objDoc.Close(ref missing, ref missing, ref missing);
                objWordApp.Application.Quit(ref missing, ref missing, ref missing);
            }

        }




        private static void FindExcel(string findText, string Wfile)
        {
            
            string File_name = Wfile;
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;

            Application _excelApp = new Application();
            Workbook workBook = _excelApp.Workbooks.Open(Wfile);
            
                        int numSheets = workBook.Sheets.Count;

            while (numSheets > 0)
            {
                try
                {
                    object missing = System.Reflection.Missing.Value;
                    oWB = oXL.Workbooks.Open(File_name, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing);


                    oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.Worksheets[numSheets];
                    Microsoft.Office.Interop.Excel.Range oRng = GetSpecifiedRange(findText, oSheet);
                    if (oRng != null)
                    {
                        filesThatConstainSSN.Add(Wfile);

                        oWB.Close(false, missing, missing);

                        oSheet = null;
                        oWB = null;
                        oXL.Quit();
                    }
                    else
                    {

                    }
                    oWB.Close(false, missing, missing);

                    oSheet = null;
                    oWB = null;
                    oXL.Quit();
                }
                catch (Exception ex)
                {

                }
                numSheets--;
            }
        }

        private static Microsoft.Office.Interop.Excel.Range GetSpecifiedRange(string matchStr, Microsoft.Office.Interop.Excel.Worksheet objWs)
        {
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Range currentFind = null;
            Microsoft.Office.Interop.Excel.Range firstFind = null;
            currentFind = objWs.get_Range("A1", "AM100").Find(matchStr, missing,
                           Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                           Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
                           Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                           Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, false, missing, missing);
            return currentFind;
        }
    }
}

