using FluentAssertions;
using HtmlAgilityPack;
using Xunit;

namespace ConvertToMarkdown.Tests;

/// <summary>
/// ConverterService 的單元測試 - 涵蓋 ParseSpanAttr、NormalizeTables、FlattenTable 三個內部方法。
/// </summary>
public class ConverterServiceTests
{
    // ═══════════════════════════════════════════════════════════
    // ParseSpanAttr 測試
    // ═══════════════════════════════════════════════════════════

    /// <summary>
    /// 測試 ParseSpanAttr 解析正常 colspan/rowspan 整數值。
    /// </summary>
    [Theory]
    [InlineData("colspan", "2", 2)]
    [InlineData("colspan", "3", 3)]
    [InlineData("rowspan", "2", 2)]
    [InlineData("rowspan", "5", 5)]
    public void ParseSpanAttr_正常值_回傳對應整數(string attrName, string attrValue, int expected)
    {
        // Arrange：建立含有指定屬性的 HtmlNode
        var doc = new HtmlDocument();
        doc.LoadHtml($"<td {attrName}=\"{attrValue}\">cell</td>");
        var node = doc.DocumentNode.SelectSingleNode("//td")!;

        // Act
        int result = ConverterService.ParseSpanAttr(node, attrName);

        // Assert
        result.Should().Be(expected);
    }

    /// <summary>
    /// 測試 ParseSpanAttr 當屬性不存在時應回傳預設值 1。
    /// </summary>
    [Fact]
    public void ParseSpanAttr_屬性不存在_回傳1()
    {
        // Arrange：建立沒有 colspan/rowspan 屬性的節點
        var doc = new HtmlDocument();
        doc.LoadHtml("<td>cell</td>");
        var node = doc.DocumentNode.SelectSingleNode("//td")!;

        // Act
        int result = ConverterService.ParseSpanAttr(node, "colspan");

        // Assert
        result.Should().Be(1);
    }

    /// <summary>
    /// 測試 ParseSpanAttr 當屬性值無效時應回傳預設值 1。
    /// </summary>
    [Theory]
    [InlineData("colspan", "0")]
    [InlineData("colspan", "-1")]
    [InlineData("colspan", "abc")]
    [InlineData("colspan", "")]
    public void ParseSpanAttr_無效值_回傳1(string attrName, string attrValue)
    {
        // Arrange
        var doc = new HtmlDocument();
        doc.LoadHtml($"<td {attrName}=\"{attrValue}\">cell</td>");
        var node = doc.DocumentNode.SelectSingleNode("//td")!;

        // Act
        int result = ConverterService.ParseSpanAttr(node, attrName);

        // Assert
        result.Should().Be(1);
    }

    // ═══════════════════════════════════════════════════════════
    // NormalizeTables 測試
    // ═══════════════════════════════════════════════════════════

    /// <summary>
    /// 測試 NormalizeTables 當 HTML 不含表格時應原樣回傳。
    /// </summary>
    [Fact]
    public void NormalizeTables_無表格HTML_原樣回傳()
    {
        // Arrange
        string html = "<p>Hello World</p>";
        var progress = new Progress<string>();

        // Act
        string result = ConverterService.NormalizeTables(html, progress);

        // Assert：結果應包含原始文字內容（HtmlAgilityPack 可能稍微調整 HTML 結構）
        result.Should().Contain("Hello World");
    }

    /// <summary>
    /// 測試 NormalizeTables 對含 colspan 的表格正確展開。
    /// </summary>
    [Fact]
    public void NormalizeTables_含colspan表格_正確展開()
    {
        // Arrange：建立一個含 colspan="2" 的表格（表頭兩格合併）
        string html = "<table><tr><th colspan=\"2\">Header</th></tr><tr><td>A</td><td>B</td></tr></table>";
        var progress = new Progress<string>();

        // Act
        string result = ConverterService.NormalizeTables(html, progress);

        // Assert：展開後第一行應有 2 個 <th> 欄位
        var doc = new HtmlDocument();
        doc.LoadHtml(result);
        var headerRow = doc.DocumentNode.SelectNodes("//tr[1]/th");
        headerRow.Should().NotBeNull();
        headerRow!.Count.Should().Be(2);
    }

    /// <summary>
    /// 測試 NormalizeTables 對含 rowspan 的表格正確展開。
    /// </summary>
    [Fact]
    public void NormalizeTables_含rowspan表格_正確展開()
    {
        // Arrange：建立一個含 rowspan="2" 的表格（第一欄跨兩列）
        string html = "<table><tr><th rowspan=\"2\">Merged</th><th>Col2</th></tr><tr><td>Data</td></tr></table>";
        var progress = new Progress<string>();

        // Act
        string result = ConverterService.NormalizeTables(html, progress);

        // Assert：展開後每列應有相同的欄位數
        var doc = new HtmlDocument();
        doc.LoadHtml(result);
        var rows = doc.DocumentNode.SelectNodes("//tr");
        rows.Should().NotBeNull();

        // 第一列（th 節點）和第二列（td 節點）各應有 2 個儲存格
        var firstRowCells = doc.DocumentNode.SelectNodes("//tr[1]/th | //tr[1]/td");
        var secondRowCells = doc.DocumentNode.SelectNodes("//tr[2]/th | //tr[2]/td");
        firstRowCells.Should().NotBeNull();
        secondRowCells.Should().NotBeNull();
        firstRowCells!.Count.Should().Be(secondRowCells!.Count);
    }

    /// <summary>
    /// 測試 NormalizeTables 對含 colspan + rowspan 混合的表格正確展開。
    /// </summary>
    [Fact]
    public void NormalizeTables_含colspan與rowspan混合表格_正確展開()
    {
        // Arrange：建立含混合合併的表格
        string html = "<table>" +
                      "<tr><th colspan=\"2\" rowspan=\"2\">TopLeft</th><th>Col3</th></tr>" +
                      "<tr><td>R2C3</td></tr>" +
                      "<tr><td>R3C1</td><td>R3C2</td><td>R3C3</td></tr>" +
                      "</table>";
        var progress = new Progress<string>();

        // Act
        string result = ConverterService.NormalizeTables(html, progress);

        // Assert：結果應為有效 HTML，包含表格內容
        result.Should().Contain("TopLeft");
        result.Should().Contain("R3C1");

        // 確認每列都有相同欄位數（展開後為 3 欄）
        var doc = new HtmlDocument();
        doc.LoadHtml(result);
        var rows = doc.DocumentNode.SelectNodes("//tr");
        rows.Should().NotBeNull();
        foreach (var row in rows)
        {
            var cells = row.SelectNodes("td|th");
            cells.Should().NotBeNull();
            cells!.Count.Should().Be(3);
        }
    }

    // ═══════════════════════════════════════════════════════════
    // FlattenTable 測試
    // ═══════════════════════════════════════════════════════════

    /// <summary>
    /// 測試 FlattenTable 對標準無合併表格不改變結構。
    /// </summary>
    [Fact]
    public void FlattenTable_標準無合併表格_結構不變()
    {
        // Arrange：建立標準 2x2 表格
        var doc = new HtmlDocument();
        doc.LoadHtml("<table><tr><th>H1</th><th>H2</th></tr><tr><td>A</td><td>B</td></tr></table>");
        var tableNode = doc.DocumentNode.SelectSingleNode("//table")!;

        // Act
        ConverterService.FlattenTable(tableNode);

        // Assert：展開後仍應有 2 列，每列 2 個儲存格
        var rows = tableNode.SelectNodes(".//tr");
        rows.Should().NotBeNull();
        rows!.Count.Should().Be(2);

        var firstRowCells = rows[0].SelectNodes("td|th");
        var secondRowCells = rows[1].SelectNodes("td|th");
        firstRowCells.Should().NotBeNull();
        secondRowCells.Should().NotBeNull();
        firstRowCells!.Count.Should().Be(2);
        secondRowCells!.Count.Should().Be(2);
    }

    /// <summary>
    /// 測試 FlattenTable 對空表格不崩潰。
    /// </summary>
    [Fact]
    public void FlattenTable_空表格_不崩潰()
    {
        // Arrange：建立空表格（無 <tr>）
        var doc = new HtmlDocument();
        doc.LoadHtml("<table></table>");
        var tableNode = doc.DocumentNode.SelectSingleNode("//table")!;

        // Act：不應擲出任何例外
        var act = () => ConverterService.FlattenTable(tableNode);

        // Assert
        act.Should().NotThrow();
    }
}
