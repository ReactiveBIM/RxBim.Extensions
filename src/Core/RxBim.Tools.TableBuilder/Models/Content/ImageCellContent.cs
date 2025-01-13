namespace RxBim.Tools.TableBuilder;

using System.IO;
using JetBrains.Annotations;

/// <summary>
/// The image content of cell.
/// </summary>
[UsedImplicitly]
public class ImageCellContent : ICellContent
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ImageCellContent"/> class.
    /// </summary>
    /// <param name="imageStream">Stream of image.</param>
    /// <param name="scale">Scale.</param>
    public ImageCellContent(
        Stream imageStream,
        double scale = 1)
    {
        ImageStream = imageStream;
        Scale = scale;
    }

    /// <inheritdoc />
    public object? ValueObject => null;

    /// <summary>
    /// Stream of image.
    /// </summary>
    public Stream ImageStream { get; }

    /// <summary>
    /// Scale.
    /// </summary>
    public double Scale { get; }
}