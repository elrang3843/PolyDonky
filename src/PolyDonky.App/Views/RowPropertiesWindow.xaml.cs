using System.Windows;
using PolyDonky.App.Controls;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>표의 특정 행 속성을 편집하는 다이얼로그 (높이, 배경색, 세로 정렬 기본값, 머리글 여부).</summary>
public partial class RowPropertiesWindow : Window
{
    private readonly Table _table;
    private readonly int _rowIndex;

    public RowPropertiesWindow(Table table, int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= table.Rows.Count)
            throw new ArgumentOutOfRangeException(nameof(rowIndex));

        InitializeComponent();
        _table    = table;
        _rowIndex = rowIndex;
        LoadValues();
        WireUpEvents();
    }

    private void LoadValues()
    {
        var row = _table.Rows[_rowIndex];
        HeightBox.Text          = row.HeightMm > 0 ? row.HeightMm.ToString("F1") : "0";
        RowBgColorPicker.ColorText = row.BackgroundColor ?? string.Empty;
        IsHeaderCheck.IsChecked = row.IsHeader;

        switch (row.VerticalAlign)
        {
            case CellVerticalAlign.Top:    RVAlignTopRadio.IsChecked    = true; break;
            case CellVerticalAlign.Middle: RVAlignMiddleRadio.IsChecked = true; break;
            case CellVerticalAlign.Bottom: RVAlignBottomRadio.IsChecked = true; break;
            default:                       RVAlignAutoRadio.IsChecked   = true; break;
        }
    }

    private void WireUpEvents()
    {
        OkButton.Click     += OkButton_Click;
        CancelButton.Click += CancelButton_Click;
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        if (!double.TryParse(HeightBox.Text.Trim(), out double height))
        {
            MessageBox.Show(this, "높이는 숫자여야 합니다.", "행 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            HeightBox.Focus();
            return;
        }

        string bgColor = RowBgColorPicker.ColorText.Trim();
        if (bgColor.Length > 0 && !ColorPickerBox.TryParseColor(bgColor, out _))
        {
            MessageBox.Show(this, "배경색을 올바른 색상 값으로 입력하세요 (예: #FFCC00).", "행 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            RowBgColorPicker.Focus();
            return;
        }

        var row = _table.Rows[_rowIndex];
        row.HeightMm       = height;
        row.BackgroundColor = bgColor.Length > 0 ? bgColor : null;
        row.IsHeader        = IsHeaderCheck.IsChecked ?? false;

        row.VerticalAlign = RVAlignTopRadio.IsChecked    == true ? CellVerticalAlign.Top
                          : RVAlignMiddleRadio.IsChecked == true ? CellVerticalAlign.Middle
                          : RVAlignBottomRadio.IsChecked == true ? CellVerticalAlign.Bottom
                          : (CellVerticalAlign?)null;

        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
