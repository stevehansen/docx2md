namespace Docx2Md.Core.Models;

/// <summary>
/// Represents an embedded image in the document
/// </summary>
public class ImageInfo
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string RelationshipId { get; set; } = string.Empty;
    public byte[] Data { get; set; } = Array.Empty<byte>();
    public string ContentType { get; set; } = string.Empty;
    public string? SuggestedFileName { get; set; }
    public string? AltText { get; set; }
    public int? Width { get; set; }
    public int? Height { get; set; }

    /// <summary>
    /// Get file extension based on content type
    /// </summary>
    public string GetFileExtension()
    {
        return ContentType.ToLowerInvariant() switch
        {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/jpg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/svg+xml" => ".svg",
            _ => ".bin"
        };
    }

    /// <summary>
    /// Get suggested file name or generate one
    /// </summary>
    public string GetFileName()
    {
        if (!string.IsNullOrEmpty(SuggestedFileName))
            return SuggestedFileName;

        return $"image_{Id}{GetFileExtension()}";
    }
}
