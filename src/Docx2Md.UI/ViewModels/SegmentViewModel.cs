using System;
using System.Collections.Generic;
using CommunityToolkit.Mvvm.ComponentModel;
using Docx2Md.Core.Models;

namespace Docx2Md.UI.ViewModels;

/// <summary>
/// ViewModel wrapper for Segment that provides change notification for UI binding.
/// Enables live updates when override properties change.
/// </summary>
public partial class SegmentViewModel : ViewModelBase
{
    private readonly Segment _segment;

    /// <summary>
    /// Event raised when any override property changes, signaling markdown regeneration is needed.
    /// </summary>
    public event EventHandler? OverrideChanged;

    /// <summary>
    /// Whether this segment is currently selected in the UI.
    /// Used for visual highlighting in preview panes.
    /// </summary>
    [ObservableProperty]
    private bool _isSelected;

    public SegmentViewModel(Segment segment)
    {
        _segment = segment ?? throw new ArgumentNullException(nameof(segment));
    }

    /// <summary>
    /// The underlying Segment model
    /// </summary>
    public Segment Segment => _segment;

    // Read-only properties delegated to the underlying Segment
    public string Id => _segment.Id;
    public int OrderIndex => _segment.OrderIndex;
    public SegmentType Type => _segment.Type;
    public SourceMetadata Metadata => _segment.Metadata;
    public string Content => _segment.Content;
    public string MarkdownOutput => _segment.MarkdownOutput;
    public List<Diagnostic> Diagnostics => _segment.Diagnostics;
    public int DiagnosticCount => _segment.Diagnostics.Count;

    /// <summary>
    /// Effective type considering override
    /// </summary>
    public SegmentType EffectiveType => _segment.EffectiveType;

    /// <summary>
    /// Effective markdown considering override
    /// </summary>
    public string EffectiveMarkdown => _segment.EffectiveMarkdown;

    /// <summary>
    /// Whether to exclude this segment from output
    /// </summary>
    public bool ExcludeFromOutput
    {
        get => _segment.ExcludeFromOutput;
        set
        {
            if (_segment.ExcludeFromOutput != value)
            {
                _segment.ExcludeFromOutput = value;
                OnPropertyChanged();
                RaiseOverrideChanged();
            }
        }
    }

    /// <summary>
    /// Override heading level (1-6, or null for no override)
    /// </summary>
    public int? OverrideHeadingLevel
    {
        get => _segment.OverrideHeadingLevel;
        set
        {
            if (_segment.OverrideHeadingLevel != value)
            {
                _segment.OverrideHeadingLevel = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(EffectiveMarkdown));
                RaiseOverrideChanged();
            }
        }
    }

    /// <summary>
    /// Override segment type (or null for no override)
    /// </summary>
    public SegmentType? OverrideType
    {
        get => _segment.OverrideType;
        set
        {
            if (_segment.OverrideType != value)
            {
                _segment.OverrideType = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(EffectiveType));
                OnPropertyChanged(nameof(EffectiveMarkdown));
                RaiseOverrideChanged();
            }
        }
    }

    /// <summary>
    /// Manual markdown override (or null for no override)
    /// </summary>
    public string? ManualMarkdownOverride
    {
        get => _segment.ManualMarkdownOverride;
        set
        {
            if (_segment.ManualMarkdownOverride != value)
            {
                _segment.ManualMarkdownOverride = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(EffectiveMarkdown));
                OnPropertyChanged(nameof(HasManualOverride));
                RaiseOverrideChanged();
            }
        }
    }

    /// <summary>
    /// Whether this segment has a manual markdown override
    /// </summary>
    public bool HasManualOverride => !string.IsNullOrEmpty(_segment.ManualMarkdownOverride);

    private void RaiseOverrideChanged()
    {
        OverrideChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Refreshes computed properties after markdown regeneration
    /// </summary>
    public void RefreshMarkdown()
    {
        OnPropertyChanged(nameof(MarkdownOutput));
        OnPropertyChanged(nameof(EffectiveMarkdown));
    }
}
