namespace WizGrep.Models;

/// <summary>
/// Represents a single search keyword along with an enabled/disabled flag.
/// Up to five instances are stored in <see cref="GrepSettings.Keywords"/>,
/// allowing users to define multiple independent search terms for a grep operation.
/// </summary>
public class SearchKeyword
{
    /// <summary>
    /// The search term text. An empty string means no keyword is specified for this slot.
    /// </summary>
    public string Keyword { get; set; } = string.Empty;

    /// <summary>
    /// Whether this keyword participates in the search. Disabled keywords are ignored
    /// by <see cref="GrepSettings.GetActiveKeywords"/>.
    /// </summary>
    public bool IsEnabled { get; set; } = true;
}