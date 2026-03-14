namespace WizGrep.Models;

/// <summary>
/// Persistent settings for the WizGrep indexing feature.
/// Serialized to / deserialized from local application storage by <see cref="Services.SettingsService"/>.
/// </summary>
public class WizGrepSettings
{
    /// <summary>
    /// Root directory where index and timestamp files are stored.
    /// A subfolder tree mirroring the target folder path is created beneath this base path
    /// by <see cref="Services.IndexService"/>.
    /// </summary>
    public string IndexBasePath { get; set; } = string.Empty;

    /// <summary>
    /// When <c>true</c>, the existing index for the target folder is deleted and rebuilt
    /// from scratch at the start of the next grep operation.
    /// </summary>
    public bool RebuildIndex { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, search conditions are prepended to the exported file
    /// (both file list and grep results exports).
    /// </summary>
    public bool ShowSearchConditionsInExport { get; set; } = false;
}