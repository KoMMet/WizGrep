using CommunityToolkit.Mvvm.ComponentModel;

namespace WizGrep.Helpers;

/// <summary>
/// Observable model that tracks the current pixel widths of the resizable columns
/// in the search-results list header.
/// </summary>
/// <remarks>
/// Declared as a XAML resource (<c>ResultColumnWidthsHelper</c>) in <c>MainWindow.xaml</c>.
/// When the user drags a <see cref="ColumnSplitterHelper"/>, the code-behind updates
/// <see cref="FileNameWidth"/> or <see cref="LocationWidth"/>, and the change
/// propagates to each <c>DataTemplate</c> row via one-way binding so that the
/// data rows stay aligned with the header columns.
/// </remarks>
public partial class ResultColumnWidthsHelper : ObservableObject
{
    /// <summary>Current width (in pixels) of the "File Name" column. Default is 200.</summary>
    [ObservableProperty]
    private double _fileNameWidth = 200;y

    /// <summary>Current width (in pixels) of the "Location" column. Default is 120.</summary>
    [ObservableProperty]
    private double _locationWidth = 120;
}
