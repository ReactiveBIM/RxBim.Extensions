namespace RxBim.Tools.TableBuilder;

using ClosedXML.Excel;
using JetBrains.Annotations;

/// <summary>
/// Defines an interface of a <see cref="Table"/> converter to an Excel workbook.
/// </summary>
[PublicAPI]
public interface IExcelTableConverter
    : ITableConverter<Table, ExcelTableConverterParameters, IXLWorkbook>
{
    /// <summary>
    /// Возвращает ширину в пикселях.
    /// </summary>
    /// <param name="width">Ширина в поинтах Excel.</param>
    double ConvertWidthToPixels(double width);

    /// <summary>
    /// Возвращает высоту в пикселях.
    /// </summary>
    /// <param name="height">Высота в поинтах Excel.</param>
    double ConvertHeightToPixels(double height);
}