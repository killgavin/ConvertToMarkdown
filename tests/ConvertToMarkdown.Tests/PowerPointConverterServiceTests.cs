using FluentAssertions;
using Xunit;

namespace ConvertToMarkdown.Tests;

/// <summary>
/// PowerPointConverterService 的單元測試 - 涵蓋 EscapeMarkdownCell 內部方法，
/// 並驗證其行為與 ExcelConverterService.EscapeMarkdownCell 一致。
/// </summary>
public class PowerPointConverterServiceTests
{
    // ═══════════════════════════════════════════════════════════
    // EscapeMarkdownCell 測試
    // ═══════════════════════════════════════════════════════════

    /// <summary>
    /// 測試 EscapeMarkdownCell 當輸入為空字串時應回傳空白字串。
    /// </summary>
    [Fact]
    public void EscapeMarkdownCell_空字串_回傳空白字串()
    {
        // Arrange
        string value = string.Empty;

        // Act
        string result = PowerPointConverterService.EscapeMarkdownCell(value);

        // Assert
        result.Should().Be(" ");
    }

    /// <summary>
    /// 測試 EscapeMarkdownCell 當輸入為 null 時應回傳空白字串。
    /// </summary>
    [Fact]
    public void EscapeMarkdownCell_null_回傳空白字串()
    {
        // Arrange：強制傳入 null（方法宣告為 string，需要型別轉換）
        string value = null!;

        // Act
        string result = PowerPointConverterService.EscapeMarkdownCell(value);

        // Assert
        result.Should().Be(" ");
    }

    /// <summary>
    /// 測試 EscapeMarkdownCell 對管線符號 | 進行跳脫。
    /// </summary>
    [Fact]
    public void EscapeMarkdownCell_含管線符號_跳脫為反斜線管線()
    {
        // Arrange
        string value = "A|B";

        // Act
        string result = PowerPointConverterService.EscapeMarkdownCell(value);

        // Assert
        result.Should().Be(@"A\|B");
    }

    /// <summary>
    /// 測試 EscapeMarkdownCell 對換行字元進行替換。
    /// </summary>
    [Theory]
    [InlineData("Line1\nLine2", "Line1 Line2")]
    [InlineData("Line1\r\nLine2", "Line1 Line2")]
    [InlineData("Line1\rLine2", "Line1 Line2")]
    public void EscapeMarkdownCell_含換行字元_替換為空格(string input, string expected)
    {
        // Arrange（已在 InlineData 中定義）

        // Act
        string result = PowerPointConverterService.EscapeMarkdownCell(input);

        // Assert
        result.Should().Be(expected);
    }

    /// <summary>
    /// 測試 EscapeMarkdownCell 對正常文字不做任何修改。
    /// </summary>
    [Fact]
    public void EscapeMarkdownCell_正常文字_不變()
    {
        // Arrange
        string value = "投影片內容";

        // Act
        string result = PowerPointConverterService.EscapeMarkdownCell(value);

        // Assert
        result.Should().Be("投影片內容");
    }

    /// <summary>
    /// 驗證 PowerPointConverterService.EscapeMarkdownCell 與 ExcelConverterService.EscapeMarkdownCell 行為一致。
    /// </summary>
    [Theory]
    [InlineData("")]
    [InlineData("Hello")]
    [InlineData("A|B")]
    [InlineData("Line1\nLine2")]
    [InlineData("A|B\r\nC")]
    public void EscapeMarkdownCell_與ExcelConverterService行為一致(string input)
    {
        // Arrange（已在 InlineData 中定義）

        // Act：分別呼叫兩個類別的相同方法
        string pptResult = PowerPointConverterService.EscapeMarkdownCell(input);
        string excelResult = ExcelConverterService.EscapeMarkdownCell(input);

        // Assert：兩者結果應完全相同
        pptResult.Should().Be(excelResult);
    }
}
