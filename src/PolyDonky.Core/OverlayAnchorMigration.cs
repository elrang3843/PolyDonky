namespace PolyDonky.Core;

/// <summary>
/// 옛 빌드의 "연속 Y" 좌표(PaperBorder 좌상단 기준 누적) 를 새 페이지-로컬 좌표
/// (<see cref="IOverlayAnchored.AnchorPageIndex"/> + 페이지 모서리 기준 X/Y) 로
/// 변환하는 일회성 마이그레이션.
///
/// 트리거 조건: 어떤 오버레이의 <c>OverlayYMm</c> 가 그 섹션의
/// <c>Page.EffectiveHeightMm</c> 보다 크면 옛 좌표계로 간주하고 분해한다.
/// 새로 저장된 파일은 항상 <c>OverlayYMm &lt; pageHeight</c> 를 만족하므로
/// 이 함수가 다시 호출돼도 멱등(idempotent).
///
/// 주의: 사용자가 옛 파일 저장 이후 페이지 크기/여백을 바꿔 다시 저장한 경우,
/// 페이지 분할이 이전과 달라 마이그레이션이 부정확할 수 있다. 이 경우엔
/// 사용자가 수동으로 위치를 맞춰야 한다 (드물게 발생).
/// </summary>
public static class OverlayAnchorMigration
{
    /// <summary>
    /// 문서 전체를 훑어 연속 Y 좌표를 가진 오버레이 객체를 찾고, 페이지 인덱스와
    /// 페이지 로컬 Y 로 변환한다. 마이그레이션이 적용된 경우 true 반환.
    /// </summary>
    public static bool MigrateContinuousAnchors(PolyDonkyument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        bool anyMigrated = false;
        foreach (var section in document.Sections)
        {
            anyMigrated |= MigrateSection(section);
        }
        return anyMigrated;
    }

    private static bool MigrateSection(Section section)
    {
        double pageH = section.Page.EffectiveHeightMm;
        if (pageH <= 0) return false;

        bool migrated = false;
        Walk(section.Blocks);
        return migrated;

        void Walk(IList<Block> blocks)
        {
            foreach (var block in blocks)
            {
                switch (block)
                {
                    case IOverlayAnchored anch when anch.AnchorPageIndex == 0 && anch.OverlayYMm > pageH:
                    {
                        int pageIndex = (int)Math.Floor(anch.OverlayYMm / pageH);
                        anch.AnchorPageIndex = pageIndex;
                        anch.OverlayYMm     -= pageIndex * pageH;
                        migrated = true;
                        break;
                    }
                }
                if (block is Table table)
                    foreach (var row in table.Rows)
                        foreach (var cell in row.Cells)
                            Walk(cell.Blocks);
                if (block is TextBoxObject tb)
                    Walk(tb.Content);
            }
        }
    }
}
