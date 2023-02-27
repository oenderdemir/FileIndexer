// See https://aka.ms/new-console-template for more information
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.ExtendedProperties;


Console.WriteLine("Hello, World!");
IndeksDOC(@"C:\Autopsy\2023_02\Export\İŞE İZİNSİZ GELMEME.docx");

void IndeksPDF(string fileName)
{
    using (PdfDocument document = PdfDocument.Open(fileName))
    {
        foreach (Page page in document.GetPages())
        {
            string pageText = page.Text;

            foreach (Word word in page.GetWords())
            {
                Console.WriteLine(word.Text);
            }
        }
    }
}

void IndeksDOC(string fileName)
{


    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(fileName, true))
    {
        Body body = wordDoc.MainDocumentPart.Document.Body;
        string contents = "";

        foreach (Paragraph co in
                    wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>())
        {
            contents += co.InnerText;
        }
      
        Console.WriteLine(contents);
    }
}