using System;
using System.Text.Json;
using Windows.Storage;
using WizGrep.Helpers;
using WizGrep.Models;

namespace WizGrep.Services;

/// <summary>
/// Persists and restores <see cref="GrepSettings"/> and <see cref="WizGrepSettings"/>
/// to/from the WinUI <see cref="ApplicationDataContainer"/> (local app data) as JSON.
/// Silently returns default instances when loading fails.
/// </summary>
public class SettingsService
{
    /// <summary>Storage key for the serialized <see cref="GrepSettings"/>.</summary>
    private const string GrepSettingsKey = "GrepSettings";

    /// <summary>Storage key for the serialized <see cref="WizGrepSettings"/>.</summary>
    private const string WizGrepSettingsKey = "WizGrepSettings";

    /// <summary>Lazily-initialized handle to the local application data container.</summary>
    private ApplicationDataContainer LocalSettings
    {
        get
        {
            field ??= ApplicationData.Current.LocalSettings;
            return field;
        }
    }

    /// <summary>
    /// Serializes and saves the given <see cref="GrepSettings"/> to local app storage as JSON.
    /// Any serialization or storage errors are silently swallowed to keep the application running.
    /// </summary>
    /// <param name="settings">The grep configuration to persist.</param>
    public void SaveGrepSettings(GrepSettings settings)
    {
        try
        {
            LocalSettings.Values[GrepSettingsKey] = JsonSerializer.Serialize(settings);
        }
        catch(Exception e)
        {
            LoggerHelper.Instance.LogError($"Error saving GrepSettings: {e.Message}");
        }
    }

    /// <summary>
    /// Deserializes <see cref="GrepSettings"/> from local app storage.
    /// Returns a default instance if the key is absent, the JSON is corrupt, or deserialization fails.
    /// </summary>
    /// <returns>The persisted settings, or a default <see cref="GrepSettings"/> when loading fails.</returns>
    public GrepSettings LoadGrepSettings()
    {
        try
        {
            if (LocalSettings.Values.TryGetValue(GrepSettingsKey, out var value) && value is string json)
            {
                return JsonSerializer.Deserialize<GrepSettings>(json) ?? new GrepSettings();
            }
        }
        catch(Exception e)
        {
            LoggerHelper.Instance.LogError($"Error loading GrepSettings: {e.Message}");
        }

        return new GrepSettings();
    }

    /// <summary>
    /// Serializes and saves the given <see cref="WizGrepSettings"/> to local app storage as JSON.
    /// Any serialization or storage errors are silently swallowed to keep the application running.
    /// </summary>
    /// <param name="settings">The WizGrep-specific configuration to persist.</param>
    public void SaveWizGrepSettings(WizGrepSettings settings)
    {
        try
        {
            LocalSettings.Values[WizGrepSettingsKey] = JsonSerializer.Serialize(settings);
        }
        catch(Exception e)
        {
            LoggerHelper.Instance.LogError($"Error saving WizGrepSettings: {e.Message}");
        }
    }

    /// <summary>
    /// Deserializes <see cref="WizGrepSettings"/> from local app storage.
    /// Returns a default instance if the key is absent, the JSON is corrupt, or deserialization fails.
    /// </summary>
    /// <returns>The persisted settings, or a default <see cref="WizGrepSettings"/> when loading fails.</returns>
    public WizGrepSettings LoadWizGrepSettings()
    {
        try
        {
            if (LocalSettings.Values.TryGetValue(WizGrepSettingsKey, out var value) && value is string json)
            {
                return JsonSerializer.Deserialize<WizGrepSettings>(json) ?? new WizGrepSettings();
            }
        }
        catch(Exception e)
        {
            LoggerHelper.Instance.LogError($"Error loading WizGrepSettings: {e.Message}");
        }

        return new WizGrepSettings();
    }
}