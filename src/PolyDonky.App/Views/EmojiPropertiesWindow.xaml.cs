using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using PolyDonky.App.Services;
using PolyDonky.Core;
using CoreRun = PolyDonky.Core.Run;

namespace PolyDonky.App.Views;

/// <summary>
/// 이모지 속성 편집 다이얼로그. 크기(pt) 와 기준선 정렬을 변경해 즉시 Image 에 반영.
/// </summary>
public partial class EmojiPropertiesWindow : Window
{
    private readonly Image              _img;
    private readonly InlineUIContainer  _iuc;
    private readonly CoreRun            _run;

    private static readonly (EmojiAlignment Value, string Label)[] AlignOptions =
    {
        (EmojiAlignment.TextTop,    "위 (TextTop)"),
        (EmojiAlignment.Center,     "가운데 (Center)"),
        (EmojiAlignment.TextBottom, "아래 (TextBottom)"),
        (EmojiAlignment.Baseline,   "기준선 (Baseline)"),
    };

    public EmojiPropertiesWindow(Image img, InlineUIContainer iuc, CoreRun run)
    {
        InitializeComponent();
        _img = img;
        _iuc = iuc;
        _run = run;
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        foreach (var (_, label) in EmojiWindow.SizeOptions)
            SizeCombo.Items.Add(label);

        double curPt = _run.Style.FontSizePt > 0 ? _run.Style.FontSizePt : 16.0;
        int sizeIdx  = Array.FindIndex(EmojiWindow.SizeOptions, s => s.Pt == curPt);
        SizeCombo.SelectedIndex = sizeIdx >= 0 ? sizeIdx : 1;

        foreach (var (_, label) in AlignOptions)
            AlignCombo.Items.Add(label);

        var curAlign = _run.EmojiAlignment ?? EmojiAlignment.Center;
        AlignCombo.SelectedIndex = Array.FindIndex(AlignOptions, a => a.Value == curAlign) is >= 0 and var i ? i : 1;
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        // 크기 업데이트
        if (SizeCombo.SelectedIndex >= 0)
        {
            double newPt  = EmojiWindow.SizeOptions[SizeCombo.SelectedIndex].Pt;
            double newDip = FlowDocumentBuilder.PtToDip(newPt);
            _img.Width  = newDip;
            _img.Height = newDip;
            _run.Style.FontSizePt = newPt;
        }

        // 기준선 정렬 업데이트
        if (AlignCombo.SelectedIndex >= 0)
        {
            var newAlign = AlignOptions[AlignCombo.SelectedIndex].Value;
            _run.EmojiAlignment   = newAlign;
            _iuc.BaselineAlignment = newAlign switch
            {
                EmojiAlignment.TextTop    => BaselineAlignment.TextTop,
                EmojiAlignment.TextBottom => BaselineAlignment.TextBottom,
                EmojiAlignment.Baseline   => BaselineAlignment.Baseline,
                _                         => BaselineAlignment.Center,
            };
        }

        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e) => Close();
}
