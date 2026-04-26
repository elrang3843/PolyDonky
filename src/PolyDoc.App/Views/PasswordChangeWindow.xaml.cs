using System.Windows;
using PolyDoc.Iwpf;
using SR = PolyDoc.App.Properties.Resources;

namespace PolyDoc.App.Views;

/// <summary>
/// 비밀번호 보호 모드 설정/변경 다이얼로그.
/// 모드 라디오버튼 + 비밀번호 입력란 (None 선택 시 숨김)으로 구성.
/// </summary>
public partial class PasswordChangeWindow : Window
{
    public PasswordChangeWindow(PasswordMode currentMode = PasswordMode.None)
    {
        InitializeComponent();
        SelectMode(currentMode);
        Loaded += (_, _) => NewPasswordInput.Focus();
    }

    /// <summary>사용자가 선택한 보호 모드.</summary>
    public PasswordMode ResultMode { get; private set; }

    /// <summary>설정할 비밀번호. None 모드이면 null.</summary>
    public string? ResultPassword { get; private set; }

    private void SelectMode(PasswordMode mode)
    {
        switch (mode)
        {
            case PasswordMode.Read:  RadioRead.IsChecked  = true; break;
            case PasswordMode.Write: RadioWrite.IsChecked = true; break;
            case PasswordMode.Both:  RadioBoth.IsChecked  = true; break;
            default:                 RadioNone.IsChecked  = true; break;
        }
        UpdatePasswordPanelVisibility();
    }

    private void OnModeChanged(object sender, RoutedEventArgs e)
        => UpdatePasswordPanelVisibility();

    private void UpdatePasswordPanelVisibility()
    {
        PasswordPanel.Visibility = RadioNone.IsChecked == true
            ? Visibility.Collapsed
            : Visibility.Visible;
    }

    private PasswordMode SelectedMode()
    {
        if (RadioRead.IsChecked  == true) return PasswordMode.Read;
        if (RadioWrite.IsChecked == true) return PasswordMode.Write;
        if (RadioBoth.IsChecked  == true) return PasswordMode.Both;
        return PasswordMode.None;
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        var mode = SelectedMode();

        if (mode == PasswordMode.None)
        {
            ResultMode     = PasswordMode.None;
            ResultPassword = null;
            DialogResult   = true;
            return;
        }

        var pwd     = NewPasswordInput.Password;
        var confirm = ConfirmPasswordInput.Password;

        if (string.IsNullOrEmpty(pwd))
        {
            ErrorText.Text       = SR.PwdEmpty;
            ErrorText.Visibility = Visibility.Visible;
            return;
        }

        if (pwd != confirm)
        {
            ErrorText.Text       = SR.PwdMismatch;
            ErrorText.Visibility = Visibility.Visible;
            return;
        }

        ResultMode     = mode;
        ResultPassword = pwd;
        DialogResult   = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
