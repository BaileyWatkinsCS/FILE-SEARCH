using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using Extractor2;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Search_Testing
{
    class SearchClass
    {
        static List<string> filesThatConstainSSN = new List<string>();
        static List<string> accessDenied = new List<string>();
        static List<string> corrupted = new List<string>();
        static List<string> generalErrors = new List<string>();

        //Need this for dialog box to save dont know why?
        [STAThread]

        public static void Main(string[] args)
        {
            String directorySearch;
            List<string> docFiles = new List<string>();
            List<string> excelFiles = new List<string>();
            List<string> pdfFiles = new List<string>();
            //Console Enter
            Console.WriteLine("Enter File Path");
            directorySearch = Console.ReadLine();
            Console.WriteLine("***************************************");
            // doc files
            docFiles = DirSearchWord(directorySearch);
            List<string> DirSearchWord(string ds)
            {
                try
                {
                    foreach (string file in Directory.GetFiles(ds, "*.doc"))
                    {
                        Word.Application app = new Word.Application();
                        try
                        {
                            //^?^?^?  -   ^?^?  -  ^?^?^?^?
                            //the ^? finds any digit and the dash makes sure you get the correct form
                            FindWord(app, "^#^#^#-^#^#-^#^#^#^#", file);
                            Console.WriteLine(file);
                        }
                        catch (System.Exception excpt)
                        {
                            Console.WriteLine(excpt.Message + " : " + file);
                        }
                        app.Quit();
                    }
                    foreach (string directory in Directory.GetDirectories(ds))
                    {
                        docFiles.AddRange(DirSearchWord(directory));
                    }
                }
                catch (System.Exception excpt)
                {
                    Console.WriteLine(excpt.Message);
                }
                return docFiles;
            }
            //Excel Documents
            excelFiles = DirSearchExcel(directorySearch);
            List<string> DirSearchExcel(string ds)
            {
                try
                {
                    foreach (string file in Directory.GetFiles(ds, "*.xlsx"))
                    {
                        Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                        //? is used for any charcter in excel
                        FindExcel(oXL,"???-??-????", file);
                        Console.WriteLine(file);
                    }
                    foreach (string directory in Directory.GetDirectories(ds))
                    {
                        excelFiles.AddRange(DirSearchExcel(directory));
                    }
                }
                catch (System.Exception excpt)
                {
                    Console.WriteLine(excpt.Message);
                }
                return excelFiles;
            }
            //pdf documents
            pdfFiles = DirSearchPDF(directorySearch);
            string currentPdfText;
            ExtractPDF currPdf;
            List<string> DirSearchPDF(string ds)
            {
                try
                {
                    foreach (string file in Directory.GetFiles(ds, "*.pdf"))
                    {
                        currPdf = new ExtractPDF(file);
                        currentPdfText = currPdf.extract();
                        if(currentPdfText == "Filename is incorrect or cannot be found.")
                        {
                            Console.WriteLine("Couldn't read {0}.",file);
                        }
                        else
                        {
                            Console.WriteLine(file);
                            FindPdf(currentPdfText, file);
                        }
                    }
                    foreach (string directory in Directory.GetDirectories(ds))
                    {
                        pdfFiles.AddRange(DirSearchPDF(directory));
                    }
                }
                catch (System.Exception excpt)
                {
                    Console.WriteLine(excpt.Message);
                }
                return pdfFiles;
            }

            Console.WriteLine("***************************************");
            Console.WriteLine("Files That Contain SSN Formating: ");
            Console.WriteLine("");
            filesThatConstainSSN.ForEach(Console.WriteLine);
            Console.ReadLine();


            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = @"*\Documents";
            saveFileDialog1.Filter = "Microsoft Word Documents|*.DOC | txt files (*.txt)|*.txt|All files (*.*)|*.*"; // or just "txt files (*.txt)|*.txt" if you only want to save text files
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

                using (StreamWriter writer = new StreamWriter(saveFileDialog1.FileName))
                {
                    foreach (String s in filesThatConstainSSN)
                        writer.WriteLine(s);
                    writer.Close();
                }
            }
            Console.ReadLine();
        }

        public static void FindWord(Word.Application WordApp, object findText, string Wfile)
        {
            WordApp.Visible = false;
            object missing = System.Reflection.Missing.Value;
            object filename = Wfile;
            Microsoft.Office.Interop.Word.Document objDoc;
            objDoc = WordApp.Documents.Open(ref filename, ref missing, ref missing, ref missing,
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
                objDoc.Close(ref missing, ref missing, ref missing);
                WordApp.Application.Quit(ref missing, ref missing, ref missing);
            }
            catch (Exception)
            {
                objDoc.Close(ref missing, ref missing, ref missing);
                WordApp.Application.Quit(ref missing, ref missing, ref missing);
            }
        }

        private static void FindExcel(Excel.Application oXL,string findText, string Wfile)
        {
            string File_name = Wfile;
            Microsoft.Office.Interop.Excel.Workbook oWB = null;
            Microsoft.Office.Interop.Excel.Worksheet oSheet = null;
            Application _excelApp = new Application();
            Workbook workBook = _excelApp.Workbooks.Open(Wfile);            
            int numSheets = workBook.Sheets.Count;
            bool ExitNow = true;
            while (numSheets > 0 && ExitNow != false)
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
                        ExitNow = false;
                    }

                }
                catch (Exception)
                {
                    oXL.DisplayAlerts = false;
                    workBook.Close(null, null, null);
                    oXL.Workbooks.Close();
                    oXL.Quit();
                    if (oSheet != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet);
                    if (workBook != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook);
                    if (oXL != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                    oSheet = null;
                    workBook = null;
                    oXL = null;
                    GC.Collect();
                }
                numSheets--;
            }

            //final cleanup
            oXL.DisplayAlerts = false;
            workBook.Close(null, null, null);
            oXL.Workbooks.Close();
            oXL.Quit();
            if (oSheet != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet);
            if (workBook != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook);
            if (oXL != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
            oSheet = null;
            workBook = null;
            oXL = null;
            GC.Collect();
        }

        private static void FindPdf(string text, string fileName)
        {
            bool containsSSN = Regex.IsMatch(text, @"\d\d\d-\d\d-\d\d\d\d");
            if (containsSSN)
            {
               filesThatConstainSSN.Add(fileName);
            }
        }

        private static Microsoft.Office.Interop.Excel.Range GetSpecifiedRange(string matchStr, Microsoft.Office.Interop.Excel.Worksheet objWs)
        {
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Range currentFind = null;
            currentFind = objWs.get_Range("A1", "AM100").Find(matchStr, missing,
                           Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                           Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
                           Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                           Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, false, missing, missing);
            return currentFind;
        } 
    }
}

