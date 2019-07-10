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

        //Need this for dialog box to save dont know why?
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
                    //Fixes depth with search option. all directories and also fixes if it is capitalized with StringComparison.OrdinalIgnoreCase
                    var files = Directory.EnumerateFiles(ds, "*.*")
                        .Where(s => s.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) || s.EndsWith(".docx", StringComparison.OrdinalIgnoreCase));

                    foreach (string file in files)
                    {
                        Word.Application app = new Word.Application();
                        try
                        {
                            FindTextDoc(GetTextFromWord(app, file), file);

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
                    var files = Directory.EnumerateFiles(ds, "*.*")
                        .Where(s => s.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || s.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) || s.EndsWith(".xltx", StringComparison.OrdinalIgnoreCase)
                        || s.EndsWith(".xltm", StringComparison.OrdinalIgnoreCase) || s.EndsWith(".csv", StringComparison.OrdinalIgnoreCase));
                    foreach (string file in files)
                    {
                        try
                        {
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
                {
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
                                Console.WriteLine(file);
                                FindTextDoc(currentPdfText, file);
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            corrupted.Add(file);
                        }

                    }
                    foreach (string directory in Directory.GetDirectories(ds))
                    {
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
                    var files = Directory.EnumerateFiles(ds, "*.*")
                        .Where(s => s.EndsWith(".txt", StringComparison.OrdinalIgnoreCase));

                    foreach (string file in files)
                    {
                        try
                        {
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

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.InitialDirectory = @"*\Documents";
            saveFileDialog1.Filter = "Microsoft Word Documents|*.DOC | txt files (*.txt)|*.txt|All files (*.*)|*.*"; // or just "txt files (*.txt)|*.txt" if you only want to save text files
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

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
                    writer.WriteLine("***********Acces Was Denied***********");
                    foreach (String ad in accessDenied)
                    {
                        writer.WriteLine(ad);
                    }
                    writer.WriteLine();
                    writer.WriteLine("***********Corrupted File*************");
                    foreach (String c in corrupted)
                    {
                        writer.WriteLine(c);
                    }
                    writer.WriteLine();
                    writer.WriteLine("********Unkown/General Errors*********");
                    foreach (String ge in generalErrors)
                    {
                        writer.WriteLine(ge);
                    }
                    writer.Close();
                }
            }
            try
            {
                Process.Start(Print);
            }
            catch
            {
                Console.WriteLine("File Was Not Saved!");
            }
        }

        private static void FindExcel(Excel.Application oXL, string Wfile)
        {
            string File_name = Wfile;
            Microsoft.Office.Interop.Excel.Workbook oWB = null;
            Microsoft.Office.Interop.Excel.Worksheet oSheet = null;
            Application _excelApp = new Application();
            Workbook workBook = _excelApp.Workbooks.Open(Wfile);
            int numSheets = workBook.Sheets.Count;

            while (numSheets > 0 && filesThatConstainSSN.Contains(Wfile) != true)
            {
                try
                {
                    object missing = System.Reflection.Missing.Value;
                    oWB = oXL.Workbooks.Open(File_name, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing);
                    oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.Worksheets[numSheets];
                    Microsoft.Office.Interop.Excel.Range oRng = GetSpecifiedRange(oSheet);
                    if (oRng != null)
                    {
                        string str = oRng.Text.ToString();
                        FindTextDoc(str, Wfile);
                    }
                }
                catch (Exception ex)
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
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Range pattern = null;
            Microsoft.Office.Interop.Excel.Range ssn = null;
            Microsoft.Office.Interop.Excel.Range ssnum = null;
            Microsoft.Office.Interop.Excel.Range socsecnum = null;
            Microsoft.Office.Interop.Excel.Range merger = null;
            Excel.Range last = objWs.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
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
            if (pattern != null)
            {
                if(merger == null)
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
            text.Replace('\n', ' ');
            text.Replace('\r', ' ');
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
            if (text.ToString().ToLower().Contains("social security number") || text.ToString().ToLower().Contains(" ssn ") || text.ToString().ToLower().Contains(" ss# ")
                || (text.ToString().ToLower().Contains("ssn") && text.Length == 3) || (text.ToString().ToLower().Contains("ss#") && text.Length == 3))
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

