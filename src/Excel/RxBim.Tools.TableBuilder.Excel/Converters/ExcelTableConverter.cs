namespace RxBim.Tools.TableBuilder;

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using ClosedXML.Excel;
using JetBrains.Annotations;

/// <inheritdoc />
[UsedImplicitly]
internal class ExcelTableConverter : IExcelTableConverter
{
    /// <summary>
    /// Стандартное разрешение экрана.
    /// </summary>
    private const double StandardDpi = 96;

    /// <summary>
    /// Разрешение экрана определено.
    /// </summary>
    private static bool _isDpiCalculated;

    /// <summary>
    /// Разрешение экрана по X.
    /// </summary>
    private static uint _dpiX;

    /// <summary>
    /// Разрешение экрана по Y.
    /// </summary>
    private static uint _dpiY;

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

    /// <inheritdoc/>
    public double ConvertWidthToPixels(double width)
    {
        CalculateDpi();

        // Магические константы определения ширины столбца и высоты строки в пикселях (из гугла).
        return (width * 7 + 5) / (StandardDpi / _dpiX);
    }

    /// <inheritdoc/>
    public double ConvertHeightToPixels(double height)
    {
        CalculateDpi();
        return height / 0.75 / (StandardDpi / _dpiY);
    }

    [DllImport("shcore.dll")]
    private static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

    [DllImport("user32.dll")]
    private static extern IntPtr GetDesktopWindow();

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

        // DPI (разрешение экрана) нужно для определения ширины и высоты Excel в пикселях.
        CalculateDpi();

        var left = cell.Style.Alignment.Horizontal switch
        {
            XLAlignmentHorizontalValues.Center =>
                (ConvertWidthToPixels(cell.WorksheetColumn().Width) - pictureCell.Width) / 2,
            XLAlignmentHorizontalValues.Right =>
                ConvertWidthToPixels(cell.WorksheetColumn().Width) - pictureCell.Width,
            _ => 0
        };

        var top = cell.Style.Alignment.Vertical switch
        {
            XLAlignmentVerticalValues.Center =>
                (ConvertHeightToPixels(cell.WorksheetRow().Height) - pictureCell.Height) / 2,
            XLAlignmentVerticalValues.Bottom =>
                ConvertHeightToPixels(cell.WorksheetRow().Height) - pictureCell.Height,
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

    private XLAlignmentVerticalValues GetExcelVerticalAlignment(CellContentVerticalAlignment verticalAlignment) =>
        verticalAlignment switch
        {
            CellContentVerticalAlignment.Top => XLAlignmentVerticalValues.Top,
            CellContentVerticalAlignment.Middle => XLAlignmentVerticalValues.Center,
            CellContentVerticalAlignment.Bottom => XLAlignmentVerticalValues.Bottom,
            _ => throw new NotImplementedException(verticalAlignment.ToString())
        };

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

    private void CalculateDpi()
    {
        if (_isDpiCalculated)
            return;

        // Определение dpi через Graphics не рабатает (не находит сборку System.Drawing.Common).
        var hwnd = GetDesktopWindow();
        var hMonitor = MonitorFromWindow(hwnd, 0);
        GetDpiForMonitor(hMonitor, 0, out _dpiX, out _dpiY);
        _isDpiCalculated = true;
    }
}