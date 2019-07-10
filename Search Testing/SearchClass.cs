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
using Microsoft.VisualBasic.FileIO;
using System.Diagnostics;
using System.Text;

namespace Search_Testing
{
    class SearchClass
    {
        static List<string> filesThatConstainSSN = new List<string>();
        static List<string> accessDenied = new List<string>();
        static List<string> corrupted = new List<string>();
        static List<string> generalErrors = new List<string>();

        //Need this for dialog box to save
        [STAThread]

        public static void Main(string[] args)
        {
            String directorySearch;
            List<string> docFiles = new List<string>();
            List<string> excelFiles = new List<string>();
            List<string> pdfFiles = new List<string>();
            List<string> textFiles = new List<string>();
            string Print = "";

            //Console Enter
            Console.WriteLine("Enter File Path");
            directorySearch = Console.ReadLine();
            while (!Directory.Exists(directorySearch))
            {
                Console.WriteLine("Enter a VALID File Path");
                directorySearch = Console.ReadLine();
            }
            Console.WriteLine("***************************************");

            // doc files
            docFiles = DirSearchWord(directorySearch);
            List<string> DirSearchWord(string ds)
            {
                try
                {
                    //Searches all directories and checks their file extentions
                    var files = Directory.EnumerateFiles(ds, "*.*")
                     .Where(s => s.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) || s.EndsWith(".docx", StringComparison.OrdinalIgnoreCase));

                    foreach (string file in files)
                    {
                        //opens word app here so it can be easily disposed
                        Word.Application app = new Word.Application();
                        try
                        {
                            //calls get text word and find text doc which do all the checks
                            FindTextDoc(GetTextFromWord(app, file), file);

                            //Prints the file that is opened (doesnt for any files that are: corrupted, acces denied or any other erros
                            Console.WriteLine(file);
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            corrupted.Add(file);
                        }
                        app.Quit();
                    }
                    foreach (string directory in Directory.GetDirectories(ds))
                    {
                        //goes into subdirectores/folders
                        docFiles.AddRange(DirSearchWord(directory));
                    }
                }
                catch (System.UnauthorizedAccessException)
                {
                    accessDenied.Add(ds);
                }
                catch (System.Exception)
                {
                    generalErrors.Add(ds);
                }
                return docFiles;
            }

            //Excel Documents
            excelFiles = DirSearchExcel(directorySearch);

            List<string> DirSearchExcel(string ds)
            {
                try
                {
                    //Searches all directories and checks their file extentions
                    var files = Directory.EnumerateFiles(ds, "*.*")
                     .Where(s => s.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || s.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) || s.EndsWith(".xltx", StringComparison.OrdinalIgnoreCase) ||
                      s.EndsWith(".xltm", StringComparison.OrdinalIgnoreCase) || s.EndsWith(".csv", StringComparison.OrdinalIgnoreCase));
                    foreach (string file in files)
                    {
                        try
                        { //opens word app here so it can be easily disposed
                            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();

                            //? is used for any charcter in excel
                            FindExcel(oXL, file);

                            Console.WriteLine(file);
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            corrupted.Add(file);
                        }
                    }
                    foreach (string directory in Directory.GetDirectories(ds))
                    {
                        //goes into subdirectores/folders
                        excelFiles.AddRange(DirSearchExcel(directory));
                    }
                }
                catch (System.UnauthorizedAccessException)
                {
                    accessDenied.Add(ds);
                }
                catch (System.Exception)
                {
                    generalErrors.Add(ds);
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
                { //Searches all directories and checks their file extentions
                    var files = Directory.EnumerateFiles(ds, "*.*")
                     .Where(s => s.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase));
                    foreach (string file in files)
                    {
                        try
                        {
                            currPdf = new ExtractPDF(file);
                            currentPdfText = currPdf.extract();
                            if (currentPdfText == "Filename is incorrect or cannot be found.")
                            {
                                Console.WriteLine("Couldn't read {0}.", file);
                            }
                            else
                            {
                                //finds text in pdf after extracted from pdf
                                FindTextDoc(currentPdfText, file);
                                Console.WriteLine(file);
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            corrupted.Add(file);
                        }

                    }
                    foreach (string directory in Directory.GetDirectories(ds))
                    {
                        //goes into subdirectores/folders
                        pdfFiles.AddRange(DirSearchPDF(directory));
                    }
                }
                catch (System.UnauthorizedAccessException)
                {
                    accessDenied.Add(ds);
                }
                catch (System.Exception)
                {
                    generalErrors.Add(ds);
                }
                return pdfFiles;
            }

            textFiles = DirSearchTextFiles(directorySearch);
            List<string> DirSearchTextFiles(string ds)
            {
                try
                {
                    //Searches all directories and checks their file extentions
                    var files = Directory.EnumerateFiles(ds, "*.*")
                     .Where(s => s.EndsWith(".txt", StringComparison.OrdinalIgnoreCase));

                    foreach (string file in files)
                    {
                        try
                        {
                            //extracts all the text from a textfile into a string which we then parse through
                            string readText = File.ReadAllText(file);
                            FindTextDoc(readText, file);
                            Console.WriteLine(file);
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            corrupted.Add(file);
                        }
                    }
                    foreach (string directory in Directory.GetDirectories(ds))
                    {
                        //goes into subdirectores/folders
                        textFiles.AddRange(DirSearchTextFiles(directory));
                    }
                }
                catch (System.UnauthorizedAccessException)
                {
                    accessDenied.Add(ds);
                }
                catch (System.Exception)
                {
                    generalErrors.Add(ds);
                }
                return excelFiles;
            }
            Console.WriteLine("***************************************");
            accessDenied = accessDenied.Distinct().ToList();
            Console.WriteLine("Files That Contain SSN Formating: ");
            Console.WriteLine("");
            filesThatConstainSSN.ForEach(Console.WriteLine);

            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                InitialDirectory = @"*\Documents", //intial location for save
                Filter = "Microsoft Word Documents|*.DOC | txt files (*.txt)|*.txt|All files (*.*)|*.*", // or just "txt files (*.txt)|*.txt" if you only want to save text files
                FilterIndex = 2,
                RestoreDirectory = true
            };

            //Writes to a textfile with headers
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Print = saveFileDialog1.FileName;
                StreamWriter writer = new StreamWriter(saveFileDialog1.FileName);
                using (writer)
                {
                    writer.WriteLine("****Files that contain SSN Numbers****");
                    foreach (String s in filesThatConstainSSN)
                    {
                        writer.WriteLine(s);
                    }
                    writer.WriteLine();
                    writer.WriteLine();
                    writer.WriteLine("********Unkown/General Errors*********");
                    foreach (String ge in generalErrors)
                    {
                        writer.WriteLine(ge);
                    }
                    writer.WriteLine();
                    writer.WriteLine("***********Corrupted File*************");
                    foreach (String c in corrupted)
                    {
                        writer.WriteLine(c);
                    }
                    writer.WriteLine();
                    writer.WriteLine("***********Acces Was Denied***********");
                    foreach (String ad in accessDenied)
                    {
                        writer.WriteLine(ad);
                    }
                    writer.Close();
                }
            }
            try
            {
                //tries to opne save document if it wasnt saved we catch the excepion 
                Process.Start(Print);
            }
            catch
            {
                Console.WriteLine("File Was Not Saved!");
                Console.ReadLine();
            }
        }

        private static void FindExcel(Excel.Application oXL, string Wfile)
        {
            Microsoft.Office.Interop.Excel.Workbook oWB = null;
            Microsoft.Office.Interop.Excel.Worksheet oSheet = null;
            Application _excelApp = new Application();
            Workbook workBook = _excelApp.Workbooks.Open(Wfile);
            int numSheets = workBook.Sheets.Count;

            //until a the file is added or until it runs out of sheets to check
            while (numSheets > 0 && filesThatConstainSSN.Contains(Wfile) != true)
            {
                try
                {
                    object missing = System.Reflection.Missing.Value;
                    oWB = oXL.Workbooks.Open(Wfile, missing, missing, missing, missing,
                     missing, missing, missing, missing, missing, missing,
                     missing, missing, missing, missing);
                    //used for number of worksheets
                    oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.Worksheets[numSheets];
                    //gets range of cells where format is similar
                    Microsoft.Office.Interop.Excel.Range oRng = GetSpecifiedRange(oSheet);
                    //checks ti see if anything exists, makes a string then checks exact format
                    if (oRng != null)
                    {
                        string str = oRng.Text.ToString();
                        FindTextDoc(str, Wfile);
                    }
                }
                //closes up if there is an exception
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
        private static Microsoft.Office.Interop.Excel.Range GetSpecifiedRange(Microsoft.Office.Interop.Excel.Worksheet objWs)
        {
            //gets range of cells with certain format then return the object
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Range pattern = null;
            Microsoft.Office.Interop.Excel.Range ssn = null;
            Microsoft.Office.Interop.Excel.Range ssnum = null;
            Microsoft.Office.Interop.Excel.Range socsecnum = null;
            Microsoft.Office.Interop.Excel.Range merger = null;
            Excel.Range last = objWs.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            //finds the format of numbers that are ssn
            pattern = objWs.get_Range("A1", last).Find("???-??-????", missing,
             Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
             Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
             Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
             Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, false, missing, missing);
            ssn = objWs.get_Range("A1", last).Find("SSN", missing,
             Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
             Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
             Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
             Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, false, missing, missing);
            ssnum = objWs.get_Range("A1", last).Find("ss#", missing,
             Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
             Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
             Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
             Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, false, missing, missing);
            socsecnum = objWs.get_Range("A1", last).Find("social security number", missing,
             Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
             Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
             Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
             Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, false, missing, missing);
            //if it finds a matching pattern it will either add or merge if there is an object that matches the above find.
            if (pattern != null)
            {
                if (merger == null)
                {
                    merger = pattern;
                }
                else
                {
                    merger.Application.Union(merger, pattern);
                }
            }
            if (ssn != null)
            {
                if (merger == null)
                {
                    merger = ssn;
                }
                else
                {
                    merger.Application.Union(merger, ssn);
                }
            }
            if (ssnum != null)
            {
                if (merger == null)
                {
                    merger = ssnum;
                }
                else
                {
                    merger.Application.Union(merger, ssnum);
                }
            }
            if (socsecnum != null)
            {
                if (merger == null)
                {
                    merger = socsecnum;
                }
                else
                {
                    merger.Application.Union(merger, socsecnum);
                }
            }
            return merger;
        }
        private static void FindTextDoc(string text, string fileName)
        {
            //checks the past in text to see if it fits format
            text.Replace('\n', ' ');
            text.Replace('\r', ' ');
            // \D is not digit, \d is digit
            bool containsSSN = Regex.IsMatch(text, @"\D\d\d\d-\d\d-\d\d\d\d\D");
            if (!containsSSN)
            {
                containsSSN = Regex.IsMatch(text, @"\d\d\d-\d\d-\d\d\d\d\D");
            }
            if (!containsSSN)
            {
                containsSSN = Regex.IsMatch(text, @"\D\d\d\d-\d\d-\d\d\d\d");
            }
            if (!containsSSN && Regex.IsMatch(text, @"\d\d\d-\d\d-\d\d\d\d") && text.Length == 11)
            {
                containsSSN = true;
            }
            if (text.ToString().ToLower().Contains("social security number") || text.ToString().ToLower().Contains(" ssn ") || text.ToString().ToLower().Contains(" ss# ") ||
             (text.ToString().ToLower().Contains("ssn") && text.Length == 3) || (text.ToString().ToLower().Contains("ss#") && text.Length == 3))
            {
                containsSSN = true;
            }
            if (containsSSN)
            {
                filesThatConstainSSN.Add(fileName);
            }
        }

        private static string GetTextFromWord(Word.Application WordApp, string file)
        {
            //converts document to string
            StringBuilder text = new StringBuilder();
            object miss = System.Reflection.Missing.Value;
            object path = file;
            object readOnly = true;
            Microsoft.Office.Interop.Word.Document docs = WordApp.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            string WordText = docs.Range().Text;
            docs.Application.Quit();
            return WordText;
        }
    }
}