using System.Text;
using Windows.Globalization;
using Microsoft.UI.Xaml;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace WizGrep;

/// <summary>
/// Application entry point that supplements the default WinUI <see cref="Application"/> class.
/// Registers the code-pages encoding provider required by Shift_JIS and Windows-1252
/// detection/reading throughout the application.
/// </summary>
public partial class App : Application
{
    /// <summary>Reference to the single application window instance.</summary>
    private Window? _window;

    /// <summary>
    /// Initializes the singleton application object. This is the first line of authored code
    /// executed, equivalent to <c>Main()</c> or <c>WinMain()</c>.
    /// Registers <see cref="System.Text.CodePagesEncodingProvider"/> so that legacy
    /// encodings (Shift_JIS, Windows-1252, etc.) are available application-wide.
    /// </summary>
    public App()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        InitializeComponent();

    }

    /// <summary>
    /// Invoked when the application is launched. Creates and activates the main window.
    /// </summary>
    /// <param name="args">Details about the launch request and process.</param>
    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        _window = new MainWindow();
        _window.Activate();
    }
}