using Microsoft.UI;
using Microsoft.UI.Input;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

namespace WizGrep.Helpers;

/// <summary>
/// A thin vertical grip control placed between columns in the results header grid.
/// Users drag this splitter to resize the adjacent columns (file name, location, content).
/// </summary>
/// <remarks>
/// Inherits from <see cref="Grid"/> so it can participate in the WinUI visual tree.
/// The <see cref="Microsoft.UI.Input.InputSystemCursor"/> is set to a horizontal resize
/// cursor to give the user a visual affordance that dragging is supported.
/// Pointer events are handled in <c>MainWindow.xaml.cs</c>.
/// </remarks>
public sealed class ColumnSplitterHelper : Grid
{
    public ColumnSplitterHelper()
    {
        Width = 8;
        Background = new SolidColorBrush(Colors.WhiteSmoke);
        ProtectedCursor = InputSystemCursor.Create(InputSystemCursorShape.SizeWestEast);
    }
}
