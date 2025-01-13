namespace RxBim.Tools.TableBuilder;

using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel;
using JetBrains.Annotations;

/// <summary>
/// Defines a <see cref="Table"/> converter to an Excle workbook.
/// </summary>
[UsedImplicitly]
internal class ExcelTableConverter : IExcelTableConverter
{
    /// <inheritdoc />
    public IXLWorkbook Convert(Table table, ExcelTableConverterParameters parameters)
    {
        var workbook = parameters.Workbook ?? new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(parameters.WorksheetName ?? "Sheet1");

        var mergedCells = new List<CellRange>();

        var sheetColumnIndex = 1;
        for (var currentColumnIndex = 0;
             currentColumnIndex < table.Columns.Count;
             currentColumnIndex++, sheetColumnIndex++)
        {
            var column = worksheet.Column(sheetColumnIndex);
            if (table.Columns[currentColumnIndex].IsAdjustedToContent)
            {
                column.AdjustToContents();
            }
            else
            {
                column.Width = table.Columns[currentColumnIndex].Width ?? table.GetAverageColumnWidth();
            }

            FillRow(table, worksheet, currentColumnIndex, sheetColumnIndex, mergedCells);
        }

        if (parameters.FreezeRows > 0)
            worksheet.SheetView.Freeze(parameters.FreezeRows, table.Columns.Count);

        var (fromRow, fromColumn, toRow, toColumn) = parameters.AutoFilterRange;

        if (fromRow > 0 && fromColumn > 0)
            worksheet.Range(fromRow, fromColumn, toRow, toColumn).SetAutoFilter(true);

        return workbook;
    }

    private void FillRow(
        Table table,
        IXLWorksheet worksheet,
        int currentColumnIndex,
        int sheetColumnIndex,
        List<CellRange> mergedCells)
    {
        var sheetRowIndex = 1;
        for (var currentRowIndex = 0; currentRowIndex < table.Rows.Count; currentRowIndex++, sheetRowIndex++)
        {
            var row = worksheet.Row(sheetRowIndex);
            var cell = table[currentRowIndex, currentColumnIndex];
            var sheetCell = worksheet.Cell(sheetRowIndex, sheetColumnIndex);

            SetFormat(sheetCell, cell.GetComposedFormat());
            SetData(sheetCell, cell.Content);

            if (table.Rows[currentRowIndex].IsAdjustedToContent)
                row.AdjustToContents();
            else
                row.Height = table.Rows[currentRowIndex].Height ?? table.GetAverageRowHeight();

            // Merge logic
            if (!cell.Merged || cell.MergeArea == null || mergedCells.Exists(x => Equals(x, cell.MergeArea)))
                continue;

            mergedCells.Add(cell.MergeArea.Value);

            worksheet.Range(
                cell.MergeArea.Value.TopRow + 1,
                cell.MergeArea.Value.LeftColumn + 1,
                cell.MergeArea.Value.BottomRow + 1,
                cell.MergeArea.Value.RightColumn + 1).Merge();
        }
    }

    private void SetData(IXLCell cell, ICellContent content)
    {
        switch (content)
        {
            case FormulaCellContent formula:
                cell.FormulaA1 = GetFormula(cell.Worksheet, formula);
                break;

            case NumericCellContent numeric:
                cell.Value = numeric.ValueObject;
                cell.Style.NumberFormat.Format = numeric.Format;
                break;

            case ImageCellContent image:
                SetImage(cell, image);
                break;

            default:
                cell.SetValue(content.ValueObject);
                break;
        }
    }

    private void SetImage(IXLCell cell, ImageCellContent image)
    {
        var pictureCell = cell.Worksheet
            .AddPicture(image.ImageStream)
            .MoveTo(cell)
            .Scale(image.Scale, true);

        // Константы определены опытным путем.
        const double scaleWidthToPixel = 11.8;
        const double scaleHeightToPixel = 2;

        var left = cell.Style.Alignment.Horizontal switch
        {
            XLAlignmentHorizontalValues.Center =>
                (cell.WorksheetColumn().Width * scaleWidthToPixel - pictureCell.Width) / 2,
            XLAlignmentHorizontalValues.Right =>
                cell.WorksheetColumn().Width * scaleWidthToPixel - pictureCell.Width,
            _ => 0
        };

        var top = cell.Style.Alignment.Vertical switch
        {
            XLAlignmentVerticalValues.Center =>
                (cell.WorksheetRow().Height * scaleHeightToPixel - pictureCell.Height) / 2,
            XLAlignmentVerticalValues.Bottom =>
                cell.WorksheetRow().Height * scaleHeightToPixel - pictureCell.Height,
            _ => 0
        };

        pictureCell.MoveTo(cell, (int)left, (int)top);
    }

    private string GetFormula(IXLWorksheet ws, FormulaCellContent formula)
    {
        var sFormula = new StringBuilder();

        var (fromRow, fromColumn, toRow, toColumn) = formula.CellRange;

        sFormula.Append(formula.Formula switch
            {
                Formulas.Sum => "SUM",
                _ => throw new NotImplementedException(formula.Formula.ToString())
            })
            .Append("(")
            .Append(ws.Range(fromRow, fromColumn, toRow, toColumn).RangeAddress)
            .Append(")");

        return sFormula.ToString();
    }

    private void SetFormat(IXLCell sheetCell, CellFormatStyle cellFormat)
    {
        var cellStyle = sheetCell.Style;

        SetBackgroundFormat(cellFormat, cellStyle);
        SetTextFormat(cellFormat, cellStyle);
        SetAlignmentFormat(cellFormat, cellStyle);
        SetBordersFormat(cellFormat, cellStyle);
    }

    private void SetBackgroundFormat(CellFormatStyle cellFormat, IXLStyle cellStyle)
    {
        if (cellFormat.BackgroundColor != null)
        {
            cellStyle.Fill.PatternType = XLFillPatternValues.Solid;
            cellStyle.Fill.SetBackgroundColor(XLColor.FromColor(cellFormat.BackgroundColor.Value));
        }
    }

    private void SetAlignmentFormat(CellFormatStyle cellFormat, IXLStyle cellStyle)
    {
        if (cellFormat.ContentHorizontalAlignment != null)
            cellStyle.Alignment.SetHorizontal(GetExcelHorizontalAlignment(cellFormat.ContentHorizontalAlignment.Value));

        if (cellFormat.ContentVerticalAlignment != null)
            cellStyle.Alignment.SetVertical(GetExcelVerticalAlignment(cellFormat.ContentVerticalAlignment.Value));
    }

    private void SetBordersFormat(CellFormatStyle cellFormat, IXLStyle cellStyle)
    {
        if (cellFormat.Borders.Top != null)
            cellStyle.Border.TopBorder = GetExcelBorderStyle(cellFormat.Borders.Top.Value);

        if (cellFormat.Borders.Right != null)
            cellStyle.Border.RightBorder = GetExcelBorderStyle(cellFormat.Borders.Right.Value);

        if (cellFormat.Borders.Bottom != null)
            cellStyle.Border.BottomBorder = GetExcelBorderStyle(cellFormat.Borders.Bottom.Value);

        if (cellFormat.Borders.Left != null)
            cellStyle.Border.LeftBorder = GetExcelBorderStyle(cellFormat.Borders.Left.Value);
    }

    private void SetTextFormat(CellFormatStyle cellFormat, IXLStyle cellStyle)
    {
        if (!string.IsNullOrEmpty(cellFormat.TextFormat.FontFamily))
            cellStyle.Font.FontName = cellFormat.TextFormat.FontFamily;

        if (cellFormat.TextFormat.Bold != null)
            cellStyle.Font.Bold = cellFormat.TextFormat.Bold.Value;

        if (cellFormat.TextFormat.Italic != null)
            cellStyle.Font.Italic = cellFormat.TextFormat.Italic.Value;

        if (cellFormat.TextFormat.TextColor != null)
            cellStyle.Font.SetFontColor(XLColor.FromColor(cellFormat.TextFormat.TextColor.Value));

        if (cellFormat.TextFormat.TextSize != null)
            cellStyle.Font.FontSize = cellFormat.TextFormat.TextSize.Value;

        if (cellFormat.TextFormat.WrapText != null)
            cellStyle.Alignment.WrapText = cellFormat.TextFormat.WrapText.Value;

        if (cellFormat.TextFormat.ShrinkToFit != null)
            cellStyle.Alignment.ShrinkToFit = cellFormat.TextFormat.ShrinkToFit.Value;
    }

    private XLAlignmentVerticalValues GetExcelVerticalAlignment(CellContentVerticalAlignment verticalAlignment)
    {
        return verticalAlignment switch
        {
            CellContentVerticalAlignment.Top => XLAlignmentVerticalValues.Top,
            CellContentVerticalAlignment.Middle => XLAlignmentVerticalValues.Center,
            CellContentVerticalAlignment.Bottom => XLAlignmentVerticalValues.Bottom,
            _ => throw new NotImplementedException(verticalAlignment.ToString())
        };
    }

    private XLAlignmentHorizontalValues GetExcelHorizontalAlignment(CellContentHorizontalAlignment horizontalAlignment) =>
        horizontalAlignment switch
        {
            CellContentHorizontalAlignment.Center => XLAlignmentHorizontalValues.Center,
            CellContentHorizontalAlignment.Left => XLAlignmentHorizontalValues.Left,
            CellContentHorizontalAlignment.Right => XLAlignmentHorizontalValues.Right,
            _ => throw new NotImplementedException(horizontalAlignment.ToString())
        };

    private XLBorderStyleValues GetExcelBorderStyle(CellBorderType borderType) =>
        borderType switch
        {
            CellBorderType.Hidden => XLBorderStyleValues.None,
            CellBorderType.Bold => XLBorderStyleValues.Medium,
            CellBorderType.Thin => XLBorderStyleValues.Thin,
            _ => throw new NotImplementedException(borderType.ToString())
        };
}