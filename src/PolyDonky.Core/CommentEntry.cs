namespace PolyDonky.Core;

/// <summary>리뷰어 코멘트(주석) 항목. <see cref="Run.CommentId"/> 가 이 Id 를 참조한다.</summary>
public sealed class CommentEntry
{
    /// <summary>참조 ID — Run.CommentId 와 1:1 매핑.</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>작성자 이름. null 이면 미상.</summary>
    public string? Author { get; set; }

    /// <summary>코멘트 작성/수정 시각. null 이면 미상.</summary>
    public System.DateTimeOffset? Date { get; set; }

    /// <summary>코멘트 본문 블록 목록.</summary>
    public IList<Block> Blocks { get; set; } = new List<Block>();
}
