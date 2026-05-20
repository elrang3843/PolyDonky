namespace PolyDonky.Convert.Common;

/// <summary>
/// CLI 변환기 공통 인자 파서.
/// 위치 인자(positional)와 옵션 플래그를 분리해 반환한다.
///
/// 지원 옵션:
///   --debug | -d | DEBUG  — 상세 진단 로그 활성화
/// </summary>
public sealed class ConverterArgs
{
    /// <summary>상세 진단 로그 활성화 여부.</summary>
    public bool DebugLog { get; }

    /// <summary>옵션 플래그를 제거한 순수 위치 인자 배열.</summary>
    public string[] Positional { get; }

    private ConverterArgs(bool debugLog, string[] positional)
    {
        DebugLog   = debugLog;
        Positional = positional;
    }

    /// <summary>
    /// <paramref name="args"/> 에서 알려진 옵션 플래그를 추출하고
    /// 나머지를 <see cref="Positional"/> 에 담아 반환한다.
    /// </summary>
    public static ConverterArgs Parse(string[] args)
    {
        bool debug = false;
        var  pos   = new List<string>(args.Length);

        foreach (var a in args)
        {
            if (a is "--debug" or "-d" or "DEBUG")
                debug = true;
            else
                pos.Add(a);
        }

        return new ConverterArgs(debug, pos.ToArray());
    }
}
