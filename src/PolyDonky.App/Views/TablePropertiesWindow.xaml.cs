using System.Windows;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>표 속성 다이얼로그. 표 수평 정렬을 편집한다.</summary>
public partial class TablePropertiesWindow : Window
{
    private readonly Table _table;

    public TablePropertiesWindow(Table table)
    {
        InitializeComponent();
        _table = table;

        // 현재 값으로 초기화
        switch (table.HAlign)
        {
            case TableHAlign.Center: AlignCenterRadio.IsChecked = true; break;
            case TableHAlign.Right:  AlignRightRadio.IsChecked  = true; break;
            default:                 AlignLeftRadio.IsChecked   = true; break;
        }
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        _table.HAlign = AlignCenterRadio.IsChecked == true ? TableHAlign.Center
                      : AlignRightRadio.IsChecked  == true ? TableHAlign.Right
                      : TableHAlign.Left;
        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e) => Close();
}
