using System.Windows;
using PolyDoc.App.Models;
using SR = PolyDoc.App.Properties.Resources;

namespace PolyDoc.App.Views;

public partial class DocumentInfoWindow : Window
{
    private readonly DocumentInfoModel _model;

    public DocumentInfoWindow(DocumentInfoModel model)
    {
        _model = model;
        InitializeComponent();
        DataContext = model;
        UpdatePasswordStatus();
    }

    private void UpdatePasswordStatus()
        => PasswordStatusText.Text = _model.HasPassword ? SR.DocInfoPwdSet : SR.DocInfoPwdNone;

    private void OnPasswordChange(object sender, RoutedEventArgs e)
    {
        var dlg = new PasswordChangeWindow { Owner = this };
        if (dlg.ShowDialog() != true) return;

        _model.PasswordChanged = true;
        if (dlg.ResultRemove)
        {
            _model.NewPassword = null;
            _model.HasPassword = false;
        }
        else
        {
            _model.NewPassword = dlg.ResultPassword;
            _model.HasPassword = !string.IsNullOrEmpty(dlg.ResultPassword);
        }
        UpdatePasswordStatus();
    }

    private void OnOk(object sender, RoutedEventArgs e)     => DialogResult = true;
    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
