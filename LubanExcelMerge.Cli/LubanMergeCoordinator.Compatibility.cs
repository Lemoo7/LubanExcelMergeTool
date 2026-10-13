using LubanExcelMerge.Core;
using LubanExcelMerge.Luban;
using LubanExcelMerge.OpenXml;
using System.IO.Compression;
using System.Xml.Linq;

namespace LubanExcelMerge.Cli;

public sealed partial class LubanMergeCoordinator
{
    private static string HeaderMarker(OpenXmlRowSnapshot row) =>
        (row.Cells.FirstOrDefault(cell => !string.IsNullOrWhiteSpace(cell.Payload.RawValue))?.Payload.RawValue ?? "")
        .Trim().TrimStart('\uFEFF');

    private static bool HasSchemaMarker(SheetSnapshot sheet) => sheet.Rows.Any(row =>
        HeaderMarker(row).StartsWith("##", StringComparison.Ordinal));

    private static bool IsColumnLayout(SheetSnapshot sheet) => sheet.Rows.Any(row => HeaderMarker(row) == "##column##var");

    private static bool SheetContentEquals(SheetSnapshot left, SheetSnapshot right)
    {
        var cells = left.Rows.SelectMany(row => row.Cells).Concat(right.Rows.SelectMany(row => row.Cells));
        return cells.All(cell => (left.GetCell(cell.Address)?.Payload ?? CellPayload.Blank)
            .ContentEquals(right.GetCell(cell.Address)?.Payload ?? CellPayload.Blank));
    }

    private static bool AuxiliaryPartsEqual(WorkbookSnapshot leftBook, SheetSnapshot left, WorkbookSnapshot rightBook, SheetSnapshot right)
    {
        using var a = ZipFile.OpenRead(leftBook.SourcePath);
        using var b = ZipFile.OpenRead(rightBook.SourcePath);
        var parts = new[] { left.PartPath, "xl/styles.xml" }
            .Concat(a.Entries.Concat(b.Entries).Select(entry => entry.FullName).Where(name =>
                name.StartsWith("xl/comments", StringComparison.Ordinal) || name.StartsWith("xl/drawings/", StringComparison.Ordinal) ||
                name.StartsWith("xl/worksheets/_rels/", StringComparison.Ordinal))).Distinct();
        foreach (var part in parts)
        {
            var ae = a.GetEntry(part);
            var be = b.GetEntry(part == left.PartPath ? right.PartPath : part);
            if (ae is null || be is null) { if (ae != be) return false; continue; }
            using var sa = ae.Open(); using var sb = be.Open();
            if (part.EndsWith(".xml", StringComparison.Ordinal) || part.EndsWith(".rels", StringComparison.Ordinal))
            {
                if (!XNode.DeepEquals(XDocument.Load(sa), XDocument.Load(sb))) return false;
            }
            else
            {
                if (!System.Security.Cryptography.SHA256.HashData(sa).SequenceEqual(System.Security.Cryptography.SHA256.HashData(sb))) return false;
            }
        }
        return true;
    }

    private static (LubanSchema Base, LubanSchema Local, LubanSchema Remote) ParseColumnSchemas(
        SheetSnapshot @base, SheetSnapshot local, SheetSnapshot remote)
    {
        var sheets = new[] { @base, local, remote };
        var headers = sheets.Select(sheet => sheet.Rows.FirstOrDefault(row => HeaderMarker(row) == "##column##var")).ToArray();
        if (headers.Any(row => row is null) || headers.Any(row => row!.RowNumber != headers[0]!.RowNumber) ||
            headers.Any(row => row!.GetCell(0)?.Payload.RawValue?.Trim().TrimStart('\uFEFF') != "##column##var"))
            throw new UnsafeWorkbookException("BASE、LOCAL、REMOTE 的纵向表布局不同；##column##var 必须位于同一行的 A 列。");
        var headerRow = headers[0]!.RowNumber;
        var width = sheets.SelectMany(sheet => sheet.Rows).SelectMany(row => row.Cells)
            .Where(HasCellContent).Select(cell => cell.ColumnIndex).DefaultIfEmpty(0).Max() + 1;
        for (var column = 0; column < width; column++)
        {
            var original = headers[0]!.GetCell(column)?.Payload ?? CellPayload.Blank;
            if (headers.Skip(1).Any(row => !original.ContentEquals(row!.GetCell(column)?.Payload ?? CellPayload.Blank)))
                throw new UnsafeWorkbookException($"纵向表第 {headerRow} 行的列布局发生变化，请先统一三方的列布局再合并。");
        }
        // Physical rows are field definitions; their names are stable identities.
        // Do not transpose the workbook: formulas, styles and cell addresses stay native.
        var fields = Enumerable.Range(0, width).Select(column =>
            new LubanField(column, column == 0 ? "字段名" : $"列{CellReference.Create(1, column)[..^1]}", "string")).ToArray();
        LubanSchema Parse(SheetSnapshot sheet)
        {
            var firstField = sheet.Rows.Where(row => row.RowNumber > headerRow)
                .FirstOrDefault(row => !string.IsNullOrWhiteSpace(row.GetCell(0)?.Payload.RawValue) &&
                    !row.GetCell(0)!.Payload.RawValue!.StartsWith("##", StringComparison.Ordinal));
            var start = firstField?.RowNumber ?? headerRow + 1;
            var prefix = Enumerable.Range(1, start - 1).Select(number =>
            {
                var values = new string?[width];
                foreach (var cell in sheet.GetRow(number)?.Cells ?? Array.Empty<OpenXmlCellSnapshot>())
                    if (cell.ColumnIndex < width) values[cell.ColumnIndex] = cell.Payload.RawValue;
                return new LubanRawRow(number, values);
            }).ToArray();
            return new LubanSchema(fields, prefix, headerRow, null, start, false, Array.Empty<string>());
        }
        var schemas = sheets.Select(Parse).ToArray();
        if (schemas.Any(schema => schema.DataStartRowNumber != schemas[0].DataStartRowNumber))
            throw new UnsafeWorkbookException("纵向表数据起始行发生变化，请先统一顶部说明区的行数。");
        var typeColumn = headers[0]!.Cells.FirstOrDefault(cell => cell.Payload.RawValue?.Trim() == "##type")?.ColumnIndex;
        if (typeColumn is int typeIndex)
        {
            var originals = @base.Rows.Where(row => row.RowNumber >= schemas[0].DataStartRowNumber && !string.IsNullOrEmpty(row.GetCell(0)?.Payload.RawValue))
                .GroupBy(row => row.GetCell(0)!.Payload.RawValue!).ToDictionary(group => group.Key, group => group.First());
            foreach (var sheet in sheets.Skip(1))
            foreach (var row in sheet.Rows.Where(row => row.RowNumber >= schemas[0].DataStartRowNumber))
            {
                if (originals.TryGetValue(row.GetCell(0)?.Payload.RawValue ?? "", out var original) &&
                    !(original.GetCell(typeIndex)?.Payload ?? CellPayload.Blank).ContentEquals(row.GetCell(typeIndex)?.Payload ?? CellPayload.Blank))
                    throw new UnsafeWorkbookException($"纵向表字段 {row.GetCell(0)!.Payload.RawValue} 的类型发生变化，请先确认类型与数据兼容后统一类型定义。");
            }
        }
        return (schemas[0], schemas[1], schemas[2]);
    }
}
