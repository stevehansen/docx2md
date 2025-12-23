using Docx2Md.Core;
using Docx2Md.Core.Models;
using Xunit;

namespace Docx2Md.Tests;

public class IntegrationTests
{
    [Fact]
    public void ConvertFile_EndToEnd_Success()
    {
        // Arrange
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        try
        {
            var inputPath = Path.Combine(tempDir, "test.docx");
            var outputPath = Path.Combine(tempDir, "test.md");

            // Create a test DOCX file
            TestDocxHelper.CreateSampleDocument(inputPath);

            var converter = new Docx2MdConverter();

            // Act
            var document = converter.ConvertFile(inputPath, outputPath);

            // Assert
            Assert.True(File.Exists(outputPath), "Output markdown file should exist");
            Assert.NotEmpty(document.Segments);
            
            var markdown = File.ReadAllText(outputPath);
            Assert.Contains("Sample Document", markdown);
            Assert.Contains("Section 1", markdown);
            Assert.Contains("Section 2", markdown);
            
            // Verify headings are properly formatted
            Assert.Contains("# Sample Document", markdown);
            Assert.Contains("## Section 1", markdown);
        }
        finally
        {
            // Cleanup
            if (Directory.Exists(tempDir))
            {
                Directory.Delete(tempDir, recursive: true);
            }
        }
    }

    [Fact]
    public void ParseFile_ExtractsSegments()
    {
        // Arrange
        var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        try
        {
            var inputPath = Path.Combine(tempDir, "test.docx");
            TestDocxHelper.CreateSampleDocument(inputPath);

            var converter = new Docx2MdConverter();

            // Act
            var document = converter.ParseFile(inputPath);

            // Assert
            Assert.NotNull(document);
            Assert.NotEmpty(document.Segments);
            
            // Should have headings and paragraphs
            Assert.Contains(document.Segments, s => s.Type == SegmentType.Heading);
            Assert.Contains(document.Segments, s => s.Type == SegmentType.Paragraph);
        }
        finally
        {
            if (Directory.Exists(tempDir))
            {
                Directory.Delete(tempDir, recursive: true);
            }
        }
    }
}
