using Docx2Md.Core;
using Docx2Md.Core.Models;
using Docx2Md.Core.Conversion;
using Xunit;

namespace Docx2Md.Tests;

public class MarkdownConverterTests
{
    [Fact]
    public void ConvertHeading_GeneratesCorrectMarkdown()
    {
        // Arrange
        var converter = new MarkdownConverter();
        var segment = new Segment
        {
            Type = SegmentType.Heading,
            Content = "Test Heading",
            Metadata = new SourceMetadata
            {
                StyleName = "Heading1"
            }
        };
        var document = new DocumentModel();

        // Act
        var result = converter.ConvertSegment(segment, document);

        // Assert
        Assert.StartsWith("# ", result);
        Assert.Contains("Test Heading", result);
    }

    [Fact]
    public void ConvertParagraph_ReturnsContent()
    {
        // Arrange
        var converter = new MarkdownConverter();
        var segment = new Segment
        {
            Type = SegmentType.Paragraph,
            Content = "This is a test paragraph."
        };
        var document = new DocumentModel();

        // Act
        var result = converter.ConvertSegment(segment, document);

        // Assert
        Assert.Equal("This is a test paragraph.", result);
    }

    [Fact]
    public void ConvertDocument_ProcessesAllSegments()
    {
        // Arrange
        var converter = new MarkdownConverter();
        var document = new DocumentModel
        {
            Segments = new List<Segment>
            {
                new Segment
                {
                    Type = SegmentType.Heading,
                    Content = "Heading",
                    Metadata = new SourceMetadata { StyleName = "Heading1" }
                },
                new Segment
                {
                    Type = SegmentType.Paragraph,
                    Content = "Paragraph content"
                }
            }
        };

        // Act
        converter.ConvertDocument(document);

        // Assert
        Assert.All(document.Segments, s => Assert.NotEmpty(s.MarkdownOutput));
    }

    [Fact]
    public void ConvertListItem_GeneratesBulletList()
    {
        // Arrange
        var converter = new MarkdownConverter();
        var segment = new Segment
        {
            Type = SegmentType.ListItem,
            Content = "List item text",
            Metadata = new SourceMetadata
            {
                NumberingLevel = 0
            }
        };
        var document = new DocumentModel();

        // Act
        var result = converter.ConvertSegment(segment, document);

        // Assert
        Assert.Contains("-", result);
        Assert.Contains("List item text", result);
    }
}
