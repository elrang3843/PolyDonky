using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using PolyDonky.Core;
using Wpf = System.Windows.Documents;

namespace PolyDonky.App.Services;

/// <summary>
/// 선택 영역에 RunStyle 변경을 적용하는 공유 헬퍼.
///
/// <para><b>왜 필요한가?</b><br/>
/// 단순히 <see cref="TextSelection.ApplyPropertyValue"/> 만 호출하면 두 가지 문제가 있다.</para>
/// <list type="number">
///   <item>
///     <b>per-char IUC 텍스트의 시각 갱신 누락.</b>
///     글자폭 ≠ 100% 이거나 자간 ≠ 0 인 텍스트, 글상자 내부 텍스트 등은
///     <see cref="FlowDocumentBuilder.BuildScaledContainer"/> 가 부모 <see cref="Span"/> + per-char
///     <see cref="InlineUIContainer"/> 로 렌더링한다. 이때 자식 <see cref="System.Windows.Controls.TextBlock"/>
///     에 직접 <c>FontWeight</c>·<c>Foreground</c> 등이 박혀 있어 ApplyPropertyValue 가 Span 에 속성을
///     설정해도 자식 TextBlock 의 explicit value 가 우선되어 시각이 갱신되지 않는다.
///   </item>
///   <item>
///     <b>Tag.Style stale.</b>
///     <see cref="FlowDocumentParser"/> 의 IUC 경로(<c>iuc.Tag is Run origRun → Add(origRun.Clone())</c>)
///     는 Tag 의 <see cref="RunStyle"/> 을 verbatim 복원한다. ApplyPropertyValue 가 WPF 속성만
///     갱신하고 Tag 를 그대로 두면 다음 파스 사이클에 모델이 stale 상태로 되돌아간다.
///   </item>
/// </list>
///
/// <para><b>해법.</b> <see cref="ApplyRunStyleChange"/> 가</para>
/// <list type="number">
///   <item>ApplyPropertyValue 로 WPF 속성 적용 (Wpf.Run 즉시 시각 갱신)</item>
///   <item>영향받은 inlines 의 Tag.Style 을 <c>syncTag</c> 델리게이트로 동기화</item>
///   <item>Span/IUC 인라인은 갱신된 Tag.Style 로 통째 rebuild</item>
/// </list>
/// <para>특수 인라인 (LatexSource·EmojiKey·FootnoteId·EndnoteId·Field) 은 rebuild 대상에서 제외 — 텍스트가
/// 아닌 이미지/필드 렌더링이라 일반 글자 서식 변경 의미가 없음.</para>
/// </summary>
public static class SelectionFormatApplier
{
    /// <summary>
    /// 선택 영역에 글자 서식 변경을 일관 적용. ApplyPropertyValue + Tag 동기화 + Span/IUC rebuild 를 한 번에 처리.
    /// </summary>
    /// <param name="sel">대상 selection (RichTextBox.Selection).</param>
    /// <param name="prop">적용할 WPF DependencyProperty (예: <see cref="TextElement.FontWeightProperty"/>).</param>
    /// <param name="value">적용할 값.</param>
    /// <param name="syncTag">영향받은 각 inline 의 <see cref="RunStyle"/> 을 갱신할 델리게이트. WPF 속성과 같은 의미로
    /// Tag.Style 의 해당 필드를 설정해야 한다.</param>
    public static void ApplyRunStyleChange(TextSelection sel, DependencyProperty prop, object value, Action<RunStyle> syncTag)
    {
        if (sel.IsEmpty) return;

        // 1. WPF property — Wpf.Run 의 즉시 시각 갱신
        sel.ApplyPropertyValue(prop, value);

        // 2. 영향받은 inlines 수집 + Tag.Style 동기화
        var inlines = CollectLeafInlines(sel);
        foreach (var inline in inlines)
        {
            if (TryGetTagRun(inline) is { } tagRun) syncTag(tagRun.Style);
        }

        // 3. Span/IUC rebuild (per-char IUC 의 자식 TextBlock 은 ApplyPropertyValue 효과 못 받음)
        foreach (var inline in inlines)
        {
            if (inline is not (Span or InlineUIContainer)) continue;
            if (TryGetTagRun(inline) is not { } pr) continue;
            // 수식·이모지·각주·미주·필드 참조 런 등 특수 인라인은 rebuild 대상 제외 — 텍스트 외 렌더링.
            if (pr.LatexSource is { Length: > 0 })  continue;
            if (pr.EmojiKey    is { Length: > 0 })  continue;
            if (pr.FootnoteId  is { Length: > 0 })  continue;
            if (pr.EndnoteId   is { Length: > 0 })  continue;
            if (pr.Field is not null)                continue;

            var newInline = FlowDocumentBuilder.BuildInline(pr);
            ReplaceInline(inline, newInline);
        }
    }

    private static PolyDonky.Core.Run? TryGetTagRun(Wpf.Inline inline) => inline switch
    {
        Wpf.Run r             => r.Tag as PolyDonky.Core.Run,
        Span sp               => sp.Tag as PolyDonky.Core.Run,
        InlineUIContainer iuc => iuc.Tag as PolyDonky.Core.Run,
        _                     => null,
    };

    /// <summary>
    /// 선택 영역의 모든 Wpf.Run / Span (Tag=Run) / InlineUIContainer 를 중복 없이 수집.
    /// per-char IUC 가 Tag=Run-tagged Span 안에 있으면 Span 단위로 묶어 한 번만 추가.
    /// </summary>
    private static List<Wpf.Inline> CollectLeafInlines(TextSelection sel)
    {
        var result = new List<Wpf.Inline>();

        void TryAdd(object? obj)
        {
            if (obj is InlineUIContainer iuc
                && iuc.Parent is Span parentSpan
                && parentSpan.Tag is PolyDonky.Core.Run
                && ReferenceEquals(parentSpan.Tag, iuc.Tag))
            {
                if (!result.Contains(parentSpan)) result.Add(parentSpan);
                return;
            }
            if (obj is Wpf.Run r && !result.Contains(r))                                                  result.Add(r);
            else if (obj is InlineUIContainer i && !result.Contains(i))                                   result.Add(i);
            else if (obj is Span s && s.Tag is PolyDonky.Core.Run && !result.Contains(s))                 result.Add(s);
        }

        TryAdd(sel.Start.Parent);

        var ptr = sel.Start;
        while (ptr != null && ptr.CompareTo(sel.End) < 0)
        {
            if (ptr.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.ElementStart)
                TryAdd(ptr.GetAdjacentElement(LogicalDirection.Forward));
            ptr = ptr.GetNextContextPosition(LogicalDirection.Forward);
        }

        return result;
    }

    private static void ReplaceInline(Wpf.Inline old, Wpf.Inline replacement)
    {
        if (old.Parent is Wpf.Paragraph para)
        {
            para.Inlines.InsertBefore(old, replacement);
            para.Inlines.Remove(old);
        }
        else if (old.Parent is Span span)
        {
            span.Inlines.InsertBefore(old, replacement);
            span.Inlines.Remove(old);
        }
    }
}
