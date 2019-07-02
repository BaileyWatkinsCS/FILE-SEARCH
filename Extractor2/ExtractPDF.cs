using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;


namespace Extractor2
{
    public class ExtractPDF
    {
        string filename;

        public ExtractPDF(string filename)
        {
            this.filename = filename;
        }

        public string extract()
        {
            StringBuilder text = new StringBuilder();
            //try catch installed to handle unknown files
            try
            {
                PdfReader reader = new PdfReader(filename);
                for (int page = 1; page <= reader.NumberOfPages; page++) //gets the text from each page and adds it to the text variable
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(reader, page, strategy);
                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    text.Append(currentText);
                }
                reader.Close();
            }
            catch
            {
                text.Append("Filename is incorrect or cannot be found.");
            }
            return text.ToString();
        }
    }
}
