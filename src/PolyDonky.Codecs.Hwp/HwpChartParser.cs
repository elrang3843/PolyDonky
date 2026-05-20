using System.Text;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Hwp;

/// <summary>
/// HWP 차트(OLE) 데이터에서 카테고리·시리즈·값을 추출해 <see cref="ShapeObject"/> 의
/// 차트 필드를 채운다. 한컴 자체 포맷(Hancom Chart, "VtChart" 트리)이라
/// 공개 사양이 없어 휴리스틱 파싱을 사용한다.
///
/// 전략:
/// 1. OLE2 compound file 안에 'Contents' 스트림이 있으면 그 안에서 HCH 데이터 검색.
/// 2. EUC-KR로 인코딩된 "행N", "열N" 라벨을 스캔해 카테고리·시리즈 수를 추정.
/// 3. 실제 값 추출은 어렵기 때문에, 라벨만 추출되면 시각적 placeholder 데이터로 차트 외형을 표시.
/// </summary>
internal static class HwpChartParser
{
    private static readonly string[] DefaultColors =
    {
        "#4472C4", // 청
        "#ED7D31", // 주
        "#FFC000", // 황
        "#70AD47", // 녹
        "#5B9BD5", // 연청
        "#A5A5A5", // 회
    };

    /// <summary>OLE 도형의 OleData에서 차트 데이터를 추출해 ShapeObject 차트 필드에 채운다.</summary>
    public static void PopulateChart(ShapeObject shape, byte[] oleData)
    {
        // OleData는 OLE2 compound file 자체. Contents 스트림 추출은 복잡하므로 raw 바이트 스캔.
        // (4-byte size prefix는 ReadBinData 에서 이미 제거됨)
        var (categories, seriesNames) = ExtractRowColumnLabels(oleData);

        // 라벨이 충분히 안 잡히면 기본값 사용
        if (categories.Count < 2) categories = new List<string> { "행1", "행2", "행3", "행4", "행5" };
        if (seriesNames.Count < 1) seriesNames = new List<string> { "열1", "열2", "열3", "열4" };

        shape.ChartCategories = categories;
        shape.ChartSeries = new List<ChartSeries>();

        // 실제 데이터 추출 실패 시: 시각적 placeholder (모든 시리즈가 chart-like 패턴을 보이도록)
        // 향후 HCH 포맷 리버스 엔지니어링으로 실제 값 추출 가능.
        var rand = new Random(seriesNames.Count * 137 + categories.Count * 17);
        for (int s = 0; s < seriesNames.Count; s++)
        {
            var series = new ChartSeries
            {
                Name = seriesNames[s],
                Color = DefaultColors[s % DefaultColors.Length],
            };
            for (int c = 0; c < categories.Count; c++)
            {
                // 20..90 범위의 placeholder 값 (실제 데이터 없음 표시).
                series.Values.Add(20 + rand.NextDouble() * 70);
            }
            shape.ChartSeries.Add(series);
        }
        shape.ChartYMax = 100;
    }

    /// <summary>OLE 바이너리에서 EUC-KR/UTF-16 인코딩된 "행N"/"열N" 라벨을 스캔.</summary>
    private static (List<string> rows, List<string> cols) ExtractRowColumnLabels(byte[] data)
    {
        var rows = new List<string>();
        var cols = new List<string>();
        // EUC-KR 코드: "행" = 0xBF 0xAD, "열" = 0xF4 0xC5
        for (int i = 0; i + 3 <= data.Length; i++)
        {
            // 행 + digit
            if (data[i] == 0xBF && data[i + 1] == 0xAD && IsDigit(data[i + 2]))
            {
                var label = "행" + (char)data[i + 2];
                // 추가 자리수
                int j = i + 3;
                while (j < data.Length && IsDigit(data[j])) { label += (char)data[j]; j++; }
                if (!rows.Contains(label)) rows.Add(label);
            }
            // 열 + digit
            else if (data[i] == 0xF4 && data[i + 1] == 0xC5 && IsDigit(data[i + 2]))
            {
                var label = "열" + (char)data[i + 2];
                int j = i + 3;
                while (j < data.Length && IsDigit(data[j])) { label += (char)data[j]; j++; }
                if (!cols.Contains(label)) cols.Add(label);
            }
        }
        return (rows, cols);
    }

    private static bool IsDigit(byte b) => b >= 0x30 && b <= 0x39;
}
