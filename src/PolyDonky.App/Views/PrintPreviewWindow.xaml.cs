using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Xps.Packaging;
using PolyDonky.Core;
using Fdb = PolyDonky.App.Services.FlowDocumentBuilder;

namespace PolyDonky.App.Views;

public partial class PrintPreviewWindow : Window
{
    private XpsDocument? _xpsDoc;
    private string? _tmpXpsPath;
    private double _pageWidthDip;
    private double _pageHeightDip;
    private bool _initialZoomApplied;

    public PrintPreviewWindow(PolyDonkyument doc)
    {
        InitializeComponent();
        Loaded += (_, _) => BuildPreview(doc);
        SizeChanged += OnViewerSizeChanged;
    }

    private void BuildPreview(PolyDonkyument doc)
    {
        try
        {
            var page = doc.Sections.FirstOrDefault()?.Page ?? new PageSettings();

            _pageWidthDip  = Fdb.MmToDip(page.EffectiveWidthMm);
            _pageHeightDip = Fdb.MmToDip(page.EffectiveHeightMm);
            double padL = Fdb.MmToDip(page.MarginLeftMm);
            double padT = Fdb.MmToDip(page.MarginTopMm);
            double padR = Fdb.MmToDip(page.MarginRightMm);
            double padB = Fdb.MmToDip(page.MarginBottomMm);

            var fd = Fdb.Build(doc);
            fd.PageWidth   = _pageWidthDip;
            fd.PageHeight  = _pageHeightDip;
            fd.PagePadding = new Thickness(padL, padT, padR, padB);
            fd.ColumnWidth = double.MaxValue;

            _tmpXpsPath = Path.Combine(Path.GetTempPath(), $"pdpreview_{Guid.NewGuid():N}.xps");

            var paginator = ((System.Windows.Documents.IDocumentPaginatorSource)fd).DocumentPaginator;
            paginator.PageSize = new Size(_pageWidthDip, _pageHeightDip);

            using (var xpsWrite = new XpsDocument(_tmpXpsPath, FileAccess.ReadWrite))
                XpsDocument.CreateXpsDocumentWriter(xpsWrite).Write(paginator);

            _xpsDoc = new XpsDocument(_tmpXpsPath, FileAccess.Read);

            var fixedSeq = _xpsDoc.GetFixedDocumentSequence();
            PreviewViewer.Document = fixedSeq;

            int pageCount = fixedSeq.References.Count > 0
                ? fixedSeq.References[0].GetDocument(false)?.Pages.Count ?? 0
                : 0;
            PageInfoText.Text = pageCount > 0 ? $"총 {pageCount}페이지" : string.Empty;

            // 레이아웃 완료 후 페이지 전체 맞춤 적용 (Loaded 우선순위로 디스패치)
            Dispatcher.BeginInvoke(new Action(ApplyFitPage), DispatcherPriority.Loaded);
        }
        catch (Exception ex)
        {
            PageInfoText.Text = $"미리보기 생성 실패: {ex.Message}";
        }
        finally
        {
            LoadingOverlay.Visibility = Visibility.Collapsed;
        }
    }

    /// <summary>DocumentViewer 가용 영역에 페이지 한 장이 통째로 들어가도록 줌 계산.</summary>
    private void ApplyFitPage()
    {
        if (_pageWidthDip <= 0 || _pageHeightDip <= 0) return;
        if (PreviewViewer.ActualWidth < 50 || PreviewViewer.ActualHeight < 50)
        {
            // 레이아웃이 아직 준비되지 않았으면 한 번 더 디스패치 (단발).
            if (!_initialZoomApplied)
            {
                _initialZoomApplied = true;
                Dispatcher.BeginInvoke(new Action(() => { _initialZoomApplied = false; ApplyFitPage(); }),
                                       DispatcherPriority.Loaded);
            }
            return;
        }

        const double padding = 32;  // 스크롤바·여백 고려
        double availW = PreviewViewer.ActualWidth - padding;
        double availH = PreviewViewer.ActualHeight - padding;

        double zoom = Math.Min(availW / _pageWidthDip, availH / _pageHeightDip) * 100;
        zoom = Math.Max(10, Math.Min(500, zoom));
        SetZoom(zoom);
        _initialZoomApplied = true;
    }

    private void ApplyFitWidth()
    {
        if (_pageWidthDip <= 0) return;
        if (PreviewViewer.ActualWidth < 50) return;
        const double padding = 32;
        double availW = PreviewViewer.ActualWidth - padding;
        double zoom = (availW / _pageWidthDip) * 100;
        zoom = Math.Max(10, Math.Min(500, zoom));
        SetZoom(zoom);
    }

    private void SetZoom(double zoomPercent)
    {
        PreviewViewer.Zoom = zoomPercent;
        ZoomText.Text = $"{Math.Round(zoomPercent)}%";
    }

    private void OnViewerSizeChanged(object sender, SizeChangedEventArgs e)
    {
        // 사용자가 창 크기를 변경했을 때, 초기 fit-page 가 한 번 적용된 이후에는 자동 재계산하지 않는다
        // (사용자가 직접 줌을 조절했을 수 있으므로). 단 초기 적용이 아직 안 됐으면 다시 시도.
        if (!_initialZoomApplied) ApplyFitPage();
    }

    // ── 버튼 핸들러 ──────────────────────────────────────────────

    private void OnZoomInClick(object sender, RoutedEventArgs e)
        => SetZoom(Math.Min(500, PreviewViewer.Zoom * 1.25));

    private void OnZoomOutClick(object sender, RoutedEventArgs e)
        => SetZoom(Math.Max(10, PreviewViewer.Zoom / 1.25));

    private void OnFitPageClick(object sender, RoutedEventArgs e)
        => ApplyFitPage();

    private void OnFitWidthClick(object sender, RoutedEventArgs e)
        => ApplyFitWidth();

    private void OnPrintClick(object sender, RoutedEventArgs e)
        => PreviewViewer.Print();

    private void OnCloseClick(object sender, RoutedEventArgs e)
        => Close();

    private void OnWindowClosed(object? sender, EventArgs e)
    {
        PreviewViewer.Document = null;
        _xpsDoc?.Close();
        _xpsDoc = null;
        if (_tmpXpsPath is not null)
        {
            try { File.Delete(_tmpXpsPath); } catch { /* best-effort */ }
            _tmpXpsPath = null;
        }
    }
}
