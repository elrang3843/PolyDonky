using System.Windows;
using SR = PolyDoc.App.Properties.Resources;

namespace PolyDoc.App.Views;

/// <summary>
/// 비밀번호 설정/변경/제거 다이얼로그. 두 입력란이 일치하면 확인 가능.
/// 새 비밀번호를 비워두면 제거 의도로 해석 (ResultPassword == null, ResultRemove == true).
/// </summary>
public partial class PasswordChangeWindow : Window
{
    public PasswordChangeWindow()
    {
        InitializeComponent();
        Loaded += (_, _) => NewPasswordInput.Focus();
    }

    /// <summary>설정할 새 비밀번호. 제거인 경우 null.</summary>
    public string? ResultPassword { get; private set; }

    /// <summary>true 면 기존 비밀번호를 제거하라는 의도.</summary>
    public bool ResultRemove { get; private set; }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        var pwd     = NewPasswordInput.Password;
        var confirm = ConfirmPasswordInput.Password;

        if (pwd != confirm)
        {
            ErrorText.Text = SR.PwdMismatch;
            ErrorText.Visibility = Visibility.Visible;
            return;
        }

        if (string.IsNullOrEmpty(pwd))
        {
            ResultPassword = null;
            ResultRemove   = true;
        }
        else
        {
            ResultPassword = pwd;
            ResultRemove   = false;
        }
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
