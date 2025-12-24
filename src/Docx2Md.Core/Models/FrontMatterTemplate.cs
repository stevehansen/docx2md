namespace Docx2Md.Core.Models;

/// <summary>
/// Defines a front matter template for static site generators
/// </summary>
public class FrontMatterTemplate
{
    /// <summary>
    /// Name of the template (e.g., "Hugo", "Jekyll")
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Format of the front matter (YAML or TOML)
    /// </summary>
    public FrontMatterFormat Format { get; set; } = FrontMatterFormat.Yaml;

    /// <summary>
    /// Fields to include in the front matter
    /// </summary>
    public List<FrontMatterField> Fields { get; set; } = new();
}

/// <summary>
/// A single field in the front matter
/// </summary>
public class FrontMatterField
{
    /// <summary>
    /// The key name in the front matter (e.g., "title", "date")
    /// </summary>
    public string Key { get; set; } = string.Empty;

    /// <summary>
    /// Where to get the value from
    /// </summary>
    public FrontMatterSource Source { get; set; } = FrontMatterSource.Custom;

    /// <summary>
    /// Default value if source is not available
    /// </summary>
    public string? DefaultValue { get; set; }

    /// <summary>
    /// Custom value (used when Source is Custom)
    /// </summary>
    public string? CustomValue { get; set; }
}

/// <summary>
/// Format of the front matter block
/// </summary>
public enum FrontMatterFormat
{
    /// <summary>
    /// YAML format with --- delimiters
    /// </summary>
    Yaml,

    /// <summary>
    /// TOML format with +++ delimiters
    /// </summary>
    Toml
}

/// <summary>
/// Source for front matter field values
/// </summary>
public enum FrontMatterSource
{
    /// <summary>
    /// Document title from DOCX properties
    /// </summary>
    DocumentTitle,

    /// <summary>
    /// Document author from DOCX properties
    /// </summary>
    DocumentAuthor,

    /// <summary>
    /// Document creation date
    /// </summary>
    DateCreated,

    /// <summary>
    /// Document last modified date
    /// </summary>
    DateModified,

    /// <summary>
    /// Current date/time
    /// </summary>
    CurrentDate,

    /// <summary>
    /// Source file name without extension
    /// </summary>
    FileName,

    /// <summary>
    /// Custom user-specified value
    /// </summary>
    Custom
}

/// <summary>
/// Built-in front matter templates for common static site generators
/// </summary>
public static class BuiltInTemplates
{
    /// <summary>
    /// Hugo-compatible front matter template
    /// </summary>
    public static FrontMatterTemplate Hugo => new()
    {
        Name = "Hugo",
        Format = FrontMatterFormat.Yaml,
        Fields = new List<FrontMatterField>
        {
            new() { Key = "title", Source = FrontMatterSource.DocumentTitle, DefaultValue = "Untitled" },
            new() { Key = "date", Source = FrontMatterSource.DateCreated },
            new() { Key = "draft", Source = FrontMatterSource.Custom, CustomValue = "false" }
        }
    };

    /// <summary>
    /// Jekyll-compatible front matter template
    /// </summary>
    public static FrontMatterTemplate Jekyll => new()
    {
        Name = "Jekyll",
        Format = FrontMatterFormat.Yaml,
        Fields = new List<FrontMatterField>
        {
            new() { Key = "layout", Source = FrontMatterSource.Custom, CustomValue = "post" },
            new() { Key = "title", Source = FrontMatterSource.DocumentTitle, DefaultValue = "Untitled" },
            new() { Key = "date", Source = FrontMatterSource.DateCreated }
        }
    };

    /// <summary>
    /// Get all built-in templates
    /// </summary>
    public static List<FrontMatterTemplate> GetAll() => new() { Hugo, Jekyll };
}
