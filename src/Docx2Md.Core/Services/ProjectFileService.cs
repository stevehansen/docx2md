using Docx2Md.Core.Models;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Docx2Md.Core.Services;

/// <summary>
/// Service for saving and loading project files (.docx2md)
/// </summary>
public class ProjectFileService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    /// <summary>
    /// File extension for project files
    /// </summary>
    public const string FileExtension = ".docx2md";

    /// <summary>
    /// Create a project file from a document model
    /// </summary>
    public ProjectFile CreateFromDocument(DocumentModel document, ConversionSettings settings)
    {
        var project = new ProjectFile
        {
            SourceDocxPath = document.SourcePath,
            LastModified = DateTime.Now,
            Settings = settings
        };

        // Only save segments that have overrides
        foreach (var segment in document.Segments)
        {
            if (HasOverrides(segment))
            {
                project.SegmentOverrides.Add(new SegmentOverride
                {
                    SegmentId = segment.Id,
                    OrderIndex = segment.OrderIndex,
                    ContentHash = ComputeContentHash(segment.Content),
                    ExcludeFromOutput = segment.ExcludeFromOutput ? true : null,
                    OverrideHeadingLevel = segment.OverrideHeadingLevel,
                    OverrideType = segment.OverrideType,
                    ManualMarkdownOverride = segment.ManualMarkdownOverride
                });
            }
        }

        return project;
    }

    /// <summary>
    /// Apply project overrides to a document model
    /// </summary>
    public void ApplyToDocument(ProjectFile project, DocumentModel document)
    {
        foreach (var segmentOverride in project.SegmentOverrides)
        {
            var segment = FindMatchingSegment(document, segmentOverride);
            if (segment != null)
            {
                ApplyOverride(segment, segmentOverride);
            }
        }
    }

    /// <summary>
    /// Save a project file to disk
    /// </summary>
    public void Save(ProjectFile project, string filePath)
    {
        project.LastModified = DateTime.Now;
        var json = JsonSerializer.Serialize(project, JsonOptions);
        File.WriteAllText(filePath, json);
    }

    /// <summary>
    /// Load a project file from disk
    /// </summary>
    public ProjectFile? Load(string filePath)
    {
        if (!File.Exists(filePath))
            return null;

        try
        {
            var json = File.ReadAllText(filePath);
            return JsonSerializer.Deserialize<ProjectFile>(json, JsonOptions);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Check if a segment has any overrides
    /// </summary>
    private static bool HasOverrides(Segment segment)
    {
        return segment.ExcludeFromOutput ||
               segment.OverrideHeadingLevel.HasValue ||
               segment.OverrideType.HasValue ||
               !string.IsNullOrEmpty(segment.ManualMarkdownOverride);
    }

    /// <summary>
    /// Find a segment matching the override
    /// </summary>
    private static Segment? FindMatchingSegment(DocumentModel document, SegmentOverride ovr)
    {
        // Try by ID first
        var segment = document.Segments.FirstOrDefault(s => s.Id == ovr.SegmentId);
        if (segment != null)
            return segment;

        // Fallback to order index + content hash
        if (ovr.OrderIndex >= 0 && ovr.OrderIndex < document.Segments.Count)
        {
            var candidate = document.Segments[ovr.OrderIndex];
            if (string.IsNullOrEmpty(ovr.ContentHash) ||
                ComputeContentHash(candidate.Content) == ovr.ContentHash)
            {
                return candidate;
            }
        }

        // Try to find by content hash alone
        if (!string.IsNullOrEmpty(ovr.ContentHash))
        {
            return document.Segments.FirstOrDefault(s =>
                ComputeContentHash(s.Content) == ovr.ContentHash);
        }

        return null;
    }

    /// <summary>
    /// Apply override to a segment
    /// </summary>
    private static void ApplyOverride(Segment segment, SegmentOverride ovr)
    {
        if (ovr.ExcludeFromOutput == true)
            segment.ExcludeFromOutput = true;

        if (ovr.OverrideHeadingLevel.HasValue)
            segment.OverrideHeadingLevel = ovr.OverrideHeadingLevel;

        if (ovr.OverrideType.HasValue)
            segment.OverrideType = ovr.OverrideType;

        if (!string.IsNullOrEmpty(ovr.ManualMarkdownOverride))
            segment.ManualMarkdownOverride = ovr.ManualMarkdownOverride;
    }

    /// <summary>
    /// Compute a hash of content for change detection
    /// </summary>
    private static string ComputeContentHash(string content)
    {
        if (string.IsNullOrEmpty(content))
            return "";

        var bytes = Encoding.UTF8.GetBytes(content);
        var hash = SHA256.HashData(bytes);
        return Convert.ToBase64String(hash).Substring(0, 16);
    }
}
