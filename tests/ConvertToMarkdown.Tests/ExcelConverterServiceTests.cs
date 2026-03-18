using FluentAssertions;
using Xunit;

namespace ConvertToMarkdown.Tests;

/// <summary>
/// ExcelConverterService 的單元測試 - 涵蓋 FormatCellValue、EscapeMarkdownCell、SanitizeFileName 三個內部方法。
/// </summary>
public class ExcelConverterServiceTests
{
    // ═══════════════════════════════════════════════════════════
    // FormatCellValue 測試
    // ═══════════════════════════════════════════════════════════

    /// <summary>
    /// 測試 FormatCellValue 當輸入為 null 時應回傳空白字串。
    /// </summary>
    [Fact]
    public void FormatCellValue_null_回傳空白字串()
    {
        // Arrange
        object? value = null;

        // Act
        string result = ExcelConverterService.FormatCellValue(value);

        // Assert
        result.Should().Be(" ");
    }

    /// <summary>
    /// 測試 FormatCellValue 當輸入為字串時應正確回傳。
    /// </summary>
    [Theory]
    [InlineData("Hello", "Hello")]
    [InlineData("測試文字", "測試文字")]
    [InlineData("  spaces  ", "  spaces  ")]
    public void FormatCellValue_字串值_回傳原始字串(string input, string expected)
    {
        // Arrange
        object value = input;

        // Act
        string result = ExcelConverterService.FormatCellValue(value);

        // Assert
        result.Should().Be(expected);
    }

    /// <summary>
    /// 測試 FormatCellValue 當輸入為數值時應回傳其字串表示。
    /// </summary>
    [Theory]
    [InlineData(42.0, "42")]
    [InlineData(3.14, "3.14")]
    [InlineData(0.0, "0")]
    public void FormatCellValue_數值_回傳字串表示(double input, string expected)
    {
        // Arrange
        object value = input;

        // Act
        string result = ExcelConverterService.FormatCellValue(value);

        // Assert
        result.Should().Be(expected);
    }

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
        string result = ExcelConverterService.EscapeMarkdownCell(value);

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
        string result = ExcelConverterService.EscapeMarkdownCell(value);

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
        string value = "A|B|C";

        // Act
        string result = ExcelConverterService.EscapeMarkdownCell(value);

        // Assert
        result.Should().Be(@"A\|B\|C");
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
        string result = ExcelConverterService.EscapeMarkdownCell(input);

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
        string value = "Hello World 123";

        // Act
        string result = ExcelConverterService.EscapeMarkdownCell(value);

        // Assert
        result.Should().Be("Hello World 123");
    }

    // ═══════════════════════════════════════════════════════════
    // SanitizeFileName 測試
    // ═══════════════════════════════════════════════════════════

    /// <summary>
    /// 測試 SanitizeFileName 對正常名稱不做任何修改。
    /// </summary>
    [Fact]
    public void SanitizeFileName_正常名稱_不變()
    {
        // Arrange
        string name = "Sheet1";

        // Act
        string result = ExcelConverterService.SanitizeFileName(name);

        // Assert
        result.Should().Be("Sheet1");
    }

    /// <summary>
    /// 測試 SanitizeFileName 對含不合法字元的名稱進行替換。
    /// </summary>
    [Theory]
    [InlineData("Sheet/1", "Sheet_1")]
    [InlineData("Sheet:1", "Sheet_1")]
    [InlineData("Sheet*1", "Sheet_1")]
    [InlineData("Sheet?1", "Sheet_1")]
    public void SanitizeFileName_含不合法字元_替換為底線(string input, string expected)
    {
        // Arrange（已在 InlineData 中定義）

        // Act
        string result = ExcelConverterService.SanitizeFileName(input);

        // Assert
        result.Should().Be(expected);
    }

    /// <summary>
    /// 測試 SanitizeFileName 當輸入為空白字串時應回傳預設值 "Sheet"。
    /// </summary>
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SanitizeFileName_空白字串_回傳預設值Sheet(string input)
    {
        // Arrange（已在 InlineData 中定義）

        // Act
        string result = ExcelConverterService.SanitizeFileName(input);

        // Assert
        result.Should().Be("Sheet");
    }

    /// <summary>
    /// 測試 SanitizeFileName 對含中文的名稱正確保留。
    /// </summary>
    [Fact]
    public void SanitizeFileName_含中文名稱_正確保留()
    {
        // Arrange
        string name = "工作表一";

        // Act
        string result = ExcelConverterService.SanitizeFileName(name);

        // Assert
        result.Should().Be("工作表一");
    }
}
