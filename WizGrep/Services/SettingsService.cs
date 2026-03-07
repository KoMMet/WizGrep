using System;
using System.IO;
using System.Text.Json;
using WizGrep.Helpers;
using WizGrep.Models;

namespace WizGrep.Services;

/// <summary>
/// Persists and restores <see cref="GrepSettings"/> and <see cref="WizGrepSettings"/>
/// to/from a JSON file (<c>wizgrepsetting_hozon.txt</c>) located next to the application executable.
/// Silently returns default instances when loading fails.
/// </summary>
public class SettingsService
{
    /// <summary>Settings file name placed in the same directory as the EXE.</summary>
    private const string SettingsFileName = "wizgrepsetting_hozon.txt";

    /// <summary>Absolute path to the settings file.</summary>
    private static readonly string SettingsFilePath =
        Path.Combine(AppContext.BaseDirectory, SettingsFileName);

    /// <summary>Serializer options shared across all read/write operations.</summary>
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    /// <summary>In-memory cache so the file is read at most once per app session.</summary>
    private SettingsData? _cache;

    /// <summary>
    /// Wrapper that holds both settings objects in a single JSON file.
    /// </summary>
    private sealed class SettingsData
    {
        public GrepSettings? GrepSettings { get; set; }
        public WizGrepSettings? WizGrepSettings { get; set; }
    }

    /// <summary>
    /// Loads the combined settings from the file (or returns the cached copy).
    /// </summary>
    private SettingsData LoadAll()
    {
        if (_cache is not null) return _cache;

        try
        {
            if (File.Exists(SettingsFilePath))
            {
                var json = File.ReadAllText(SettingsFilePath);
                _cache = JsonSerializer.Deserialize<SettingsData>(json);
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error loading settings file: {e.StackTrace}");
        }

        _cache ??= new SettingsData();
        return _cache;
    }

    /// <summary>
    /// Writes the combined settings to the file and updates the cache.
    /// </summary>
    private void SaveAll(SettingsData data)
    {
        try
        {
            _cache = data;
            var json = JsonSerializer.Serialize(data, JsonOptions);
            File.WriteAllText(SettingsFilePath, json);
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error saving settings file: {e.StackTrace}");
        }
    }

    /// <summary>
    /// Serializes and saves the given <see cref="GrepSettings"/> to the settings file.
    /// Any serialization or I/O errors are silently swallowed to keep the application running.
    /// </summary>
    /// <param name="settings">The grep configuration to persist.</param>
    public void SaveGrepSettings(GrepSettings settings)
    {
        var data = LoadAll();
        data.GrepSettings = settings;
        SaveAll(data);
    }

    /// <summary>
    /// Deserializes <see cref="GrepSettings"/> from the settings file.
    /// Returns a default instance if the file is absent, the JSON is corrupt, or deserialization fails.
    /// </summary>
    /// <returns>The persisted settings, or a default <see cref="GrepSettings"/> when loading fails.</returns>
    public GrepSettings LoadGrepSettings()
    {
        return LoadAll().GrepSettings ?? new GrepSettings();
    }

    /// <summary>
    /// Serializes and saves the given <see cref="WizGrepSettings"/> to the settings file.
    /// Any serialization or I/O errors are silently swallowed to keep the application running.
    /// </summary>
    /// <param name="settings">The WizGrep-specific configuration to persist.</param>
    public void SaveWizGrepSettings(WizGrepSettings settings)
    {
        var data = LoadAll();
        data.WizGrepSettings = settings;
        SaveAll(data);
    }

    /// <summary>
    /// Deserializes <see cref="WizGrepSettings"/> from the settings file.
    /// Returns a default instance if the file is absent, the JSON is corrupt, or deserialization fails.
    /// </summary>
    /// <returns>The persisted settings, or a default <see cref="WizGrepSettings"/> when loading fails.</returns>
    public WizGrepSettings LoadWizGrepSettings()
    {
        return LoadAll().WizGrepSettings ?? new WizGrepSettings();
    }
}