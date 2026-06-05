using System;
using System.Windows;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>
/// 주석(CommentEntry) 추가·편집 다이얼로그.
/// DialogResult = true 이면 <see cref="ResultAuthor"/> / <see cref="ResultContent"/> 를 사용한다.
/// </summary>
public partial class CommentEditDialog : Window
{
    public string ResultAuthor  { get; private set; } = string.Empty;
    public string ResultContent { get; private set; } = string.Empty;

    /// <summary>새 주석 추가 모드.</summary>
    public CommentEditDialog(string defaultAuthor)
    {
        InitializeComponent();
        AuthorBox.Text = defaultAuthor;
        DateLabel.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm",
            System.Globalization.CultureInfo.InvariantCulture);
        Loaded += (_, _) => ContentBox.Focus();
    }

    /// <summary>기존 주석 편집 모드.</summary>
    public CommentEditDialog(CommentEntry entry)
    {
        InitializeComponent();
        AuthorBox.Text = entry.Author ?? string.Empty;
        DateLabel.Text = (entry.Date ?? DateTimeOffset.Now).ToString("yyyy-MM-dd HH:mm",
            System.Globalization.CultureInfo.InvariantCulture);

        // 본문 블록들을 평문으로 합쳐 TextBox 에 표시
        var sb = new System.Text.StringBuilder();
        bool first = true;
        foreach (var b in entry.Blocks)
        {
            if (b is Paragraph para)
            {
                if (!first) sb.Append('\n');
                sb.Append(para.GetPlainText());
                first = false;
            }
        }
        ContentBox.Text = sb.ToString();
        Loaded += (_, _) => ContentBox.Focus();
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        ResultAuthor  = AuthorBox.Text.Trim();
        ResultContent = ContentBox.Text ?? string.Empty;
        DialogResult  = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e) => Close();
}
