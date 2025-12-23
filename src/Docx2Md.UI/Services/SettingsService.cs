using System;
using System.IO;
using System.Text.Json;
using System.Timers;

namespace Docx2Md.UI.Services;

/// <summary>
/// Handles loading and saving application settings to disk.
/// Settings are stored as JSON in %APPDATA%/docx2md/settings.json.
/// Uses debounced saving to avoid excessive I/O on rapid changes.
/// </summary>
public class SettingsService
{
    private static readonly string SettingsFolder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "docx2md");

    private static readonly string SettingsFile = Path.Combine(SettingsFolder, "settings.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private Timer? _saveDebounceTimer;
    private AppSettings? _pendingSettings;
    private readonly object _saveLock = new();

    /// <summary>
    /// Load settings from disk. Returns default settings if file doesn't exist or is corrupted.
    /// </summary>
    public AppSettings Load()
    {
        try
        {
            if (File.Exists(SettingsFile))
            {
                var json = File.ReadAllText(SettingsFile);
                var settings = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions);
                return settings ?? new AppSettings();
            }
        }
        catch (Exception)
        {
            // Settings file corrupted or unreadable - return defaults
        }

        return new AppSettings();
    }

    /// <summary>
    /// Save settings to disk with debouncing (500ms delay).
    /// Multiple rapid calls will only result in one disk write.
    /// </summary>
    public void Save(AppSettings settings)
    {
        lock (_saveLock)
        {
            _pendingSettings = settings;

            // Cancel existing timer if any
            _saveDebounceTimer?.Stop();
            _saveDebounceTimer?.Dispose();

            // Start new debounce timer
            _saveDebounceTimer = new Timer(500);
            _saveDebounceTimer.Elapsed += OnSaveTimerElapsed;
            _saveDebounceTimer.AutoReset = false;
            _saveDebounceTimer.Start();
        }
    }

    /// <summary>
    /// Save settings immediately without debouncing.
    /// Use this on application exit.
    /// </summary>
    public void SaveImmediate(AppSettings settings)
    {
        lock (_saveLock)
        {
            _saveDebounceTimer?.Stop();
            _saveDebounceTimer?.Dispose();
            _saveDebounceTimer = null;

            WriteSettingsToFile(settings);
        }
    }

    private void OnSaveTimerElapsed(object? sender, ElapsedEventArgs e)
    {
        lock (_saveLock)
        {
            if (_pendingSettings != null)
            {
                WriteSettingsToFile(_pendingSettings);
                _pendingSettings = null;
            }
        }
    }

    private void WriteSettingsToFile(AppSettings settings)
    {
        try
        {
            Directory.CreateDirectory(SettingsFolder);
            var json = JsonSerializer.Serialize(settings, JsonOptions);
            File.WriteAllText(SettingsFile, json);
        }
        catch (Exception)
        {
            // Failed to save settings - ignore (best effort)
        }
    }
}
