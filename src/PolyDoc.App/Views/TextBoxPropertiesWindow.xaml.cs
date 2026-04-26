using System;
using System.Windows;
using System.Windows.Media;

namespace PolyDoc.App.Views;

/// <summary>글상자 테두리 색·두께·배경색 편집 대화상자.</summary>
public partial class TextBoxPropertiesWindow : Window
{
    /// <summary>대화상자 확인 후의 테두리 색 hex (null/빈 = 검정).</summary>
    public string? ResultBorderColor { get; private set; }

    /// <summary>대화상자 확인 후의 테두리 두께 (pt).</summary>
    public double ResultBorderThicknessPt { get; private set; }

    /// <summary>대화상자 확인 후의 배경색 hex (null/빈 = 투명).</summary>
    public string? ResultBackgroundColor { get; private set; }

    public TextBoxPropertiesWindow(string? borderColor, double borderThicknessPt, string? backgroundColor)
    {
        InitializeComponent();
        TxtBorderColor.Text     = borderColor ?? "#000000";
        TxtBorderThickness.Text = borderThicknessPt.ToString("0.##");
        TxtBackgroundColor.Text = backgroundColor ?? "";
        UpdatePreview(BorderColorPreview, TxtBorderColor.Text);
        UpdatePreview(BackgroundColorPreview, TxtBackgroundColor.Text);
    }

    private void OnBorderColorChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        => UpdatePreview(BorderColorPreview, TxtBorderColor.Text);

    private void OnBackgroundColorChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        => UpdatePreview(BackgroundColorPreview, TxtBackgroundColor.Text);

    private static void UpdatePreview(System.Windows.Controls.Border preview, string? hex)
    {
        if (string.IsNullOrWhiteSpace(hex))
        {
            preview.Background = Brushes.Transparent;
            return;
        }
        try
        {
            var s = hex.Trim();
            if (!s.StartsWith('#')) s = '#' + s;
            var color = (Color)ColorConverter.ConvertFromString(s)!;
            preview.Background = new SolidColorBrush(color);
        }
        catch
        {
            preview.Background = Brushes.Transparent;
        }
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        ResultBorderColor = string.IsNullOrWhiteSpace(TxtBorderColor.Text)
            ? null : TxtBorderColor.Text.Trim();

        ResultBorderThicknessPt = double.TryParse(TxtBorderThickness.Text, out double pt) && pt >= 0
            ? pt : 1.0;

        ResultBackgroundColor = string.IsNullOrWhiteSpace(TxtBackgroundColor.Text)
            ? null : TxtBackgroundColor.Text.Trim();

        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
