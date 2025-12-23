using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Docx2Md.Tests;

/// <summary>
/// Helper class to create test DOCX documents
/// </summary>
public static class TestDocxHelper
{
    public static void CreateSampleDocument(string filePath)
    {
        using (var wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            // Add title
            AddHeading(body, "Sample Document", "Heading1");
            
            // Add paragraph
            AddParagraph(body, "This is a sample paragraph with some content.");
            
            // Add subheading
            AddHeading(body, "Section 1", "Heading2");
            
            // Add more content
            AddParagraph(body, "This section contains important information.");
            AddParagraph(body, "Here is another paragraph with more details.");
            
            // Add another section
            AddHeading(body, "Section 2", "Heading2");
            AddParagraph(body, "This is the second section of the document.");

            mainPart.Document.Save();
        }
    }

    private static void AddHeading(Body body, string text, string styleId)
    {
        var paragraph = body.AppendChild(new Paragraph());
        
        var props = paragraph.AppendChild(new ParagraphProperties());
        props.ParagraphStyleId = new ParagraphStyleId { Val = styleId };
        
        var run = paragraph.AppendChild(new Run());
        run.AppendChild(new Text(text));
    }

    private static void AddParagraph(Body body, string text)
    {
        var paragraph = body.AppendChild(new Paragraph());
        var run = paragraph.AppendChild(new Run());
        run.AppendChild(new Text(text));
    }
}
