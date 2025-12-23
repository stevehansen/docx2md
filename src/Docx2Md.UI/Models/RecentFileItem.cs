using System.IO;

namespace Docx2Md.UI.Models;

/// <summary>
/// Represents a recently opened file for the Recent Files menu.
/// </summary>
public class RecentFileItem
{
    public string FilePath { get; set; } = string.Empty;

    /// <summary>
    /// Gets just the filename for display in the menu.
    /// </summary>
    public string FileName => Path.GetFileName(FilePath);

    /// <summary>
    /// Gets a shortened display name showing the last two path segments.
    /// </summary>
    public string DisplayName
    {
        get
        {
            if (string.IsNullOrEmpty(FilePath))
                return string.Empty;

            var parts = FilePath.Split(Path.DirectorySeparatorChar);
            if (parts.Length > 2)
            {
                return $"...{Path.DirectorySeparatorChar}{string.Join(Path.DirectorySeparatorChar, parts[^2..])}";
            }
            return FilePath;
        }
    }

    /// <summary>
    /// Gets whether the file still exists on disk.
    /// </summary>
    public bool Exists => File.Exists(FilePath);
}
