using System.Windows;
using PolyDoc.App.Services;

namespace PolyDoc.App.Views;

public partial class SettingsWindow : Window
{
    public SettingsWindow()
    {
        InitializeComponent();
        // 현재 테마 반영
        switch (ThemeService.Current)
        {
            case ThemeService.Theme.Dark: ThemeDark.IsChecked  = true; break;
            case ThemeService.Theme.Soft: ThemeSoft.IsChecked  = true; break;
            default:                      ThemeLight.IsChecked = true; break;
        }
    }

    private void OnThemeChecked(object sender, RoutedEventArgs e)
    {
        if (ThemeDark.IsChecked == true)       ThemeService.Apply(ThemeService.Theme.Dark);
        else if (ThemeSoft.IsChecked == true)  ThemeService.Apply(ThemeService.Theme.Soft);
        else                                   ThemeService.Apply(ThemeService.Theme.Light);
    }

    private void OnClose(object sender, RoutedEventArgs e) => Close();
}
