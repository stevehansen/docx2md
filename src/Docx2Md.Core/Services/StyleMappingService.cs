using Docx2Md.Core.Models;
using System.Text.Json;

namespace Docx2Md.Core.Services;

/// <summary>
/// Service for loading and saving custom style mappings
/// Style mappings are stored in %APPDATA%/docx2md/style-mappings.json
/// </summary>
public class StyleMappingService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true
    };

    private readonly string _filePath;

    public StyleMappingService()
    {
        var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var configDir = Path.Combine(appDataPath, "docx2md");
        _filePath = Path.Combine(configDir, "style-mappings.json");
    }

    /// <summary>
    /// Load style mappings from the configuration file
    /// </summary>
    public List<StyleMapping> Load()
    {
        try
        {
            if (!File.Exists(_filePath))
            {
                // Create default file with example mappings
                var defaults = GetDefaultMappings();
                Save(defaults);
                return defaults;
            }

            var json = File.ReadAllText(_filePath);
            var mappings = JsonSerializer.Deserialize<List<StyleMapping>>(json, JsonOptions);
            return mappings ?? new List<StyleMapping>();
        }
        catch
        {
            // If file is corrupted, return empty list
            return new List<StyleMapping>();
        }
    }

    /// <summary>
    /// Save style mappings to the configuration file
    /// </summary>
    public void Save(List<StyleMapping> mappings)
    {
        try
        {
            var directory = Path.GetDirectoryName(_filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonSerializer.Serialize(mappings, JsonOptions);
            File.WriteAllText(_filePath, json);
        }
        catch
        {
            // Best effort - ignore save failures
        }
    }

    /// <summary>
    /// Get the path to the style mappings file
    /// </summary>
    public string GetFilePath() => _filePath;

    /// <summary>
    /// Get default example mappings
    /// </summary>
    private static List<StyleMapping> GetDefaultMappings()
    {
        return new List<StyleMapping>
        {
            new StyleMapping
            {
                WordStyleName = "Code",
                Action = StyleMappingAction.CodeBlock,
                CodeLanguage = ""
            },
            new StyleMapping
            {
                WordStyleName = "CodeBlock",
                Action = StyleMappingAction.CodeBlock,
                CodeLanguage = ""
            },
            new StyleMapping
            {
                WordStyleName = "Quote",
                Action = StyleMappingAction.Blockquote
            },
            new StyleMapping
            {
                WordStyleName = "BlockQuote",
                Action = StyleMappingAction.Blockquote
            }
        };
    }
}
