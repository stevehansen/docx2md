namespace Docx2Md.Core.Models;

/// <summary>
/// Types of document segments as defined in PRD Section 5.1
/// </summary>
public enum SegmentType
{
    Heading,
    Paragraph,
    ListItem,
    Table,
    Image,
    PageBreak,
    SectionBreak,
    Unknown
}
