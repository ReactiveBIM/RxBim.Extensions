﻿namespace RxBim.Tools.Autocad.Extensions.TableBuilder
{
    using System;
    using Autodesk.AutoCAD.DatabaseServices;
    using Serializers;
    using Tools.TableBuilder.Services;

    /// <summary>
    /// Extensions for <see cref="CellBuilder"/>
    /// </summary>
    public static class CellBuilderExtensions
    {
        /// <summary>
        /// Sets the rotated text content.
        /// </summary>
        /// <param name="builder"><see cref="CellBuilder"/> object.</param>
        /// <param name="text">Content text value.</param>
        /// <param name="angle">Text rotation angle.</param>
        /// <param name="adjustCellSize"><see cref="AutocadTextCellContent.AdjustCellSize"/> property value.</param>
        public static CellBuilder SetAcadTableText(
            this CellBuilder builder,
            string text,
            RotationAngle angle = RotationAngle.Degrees000,
            bool adjustCellSize = false)
        {
            var content = new AutocadTextCellContent(text)
            {
                Rotation = angle switch
                {
                    RotationAngle.DegreesUnknown => 0,
                    RotationAngle.Degrees000 => 0,
                    RotationAngle.Degrees090 => Math.PI / 2,
                    RotationAngle.Degrees180 => Math.PI,
                    RotationAngle.Degrees270 => Math.PI * 3 / 2,
                    _ => throw new ArgumentOutOfRangeException(nameof(angle), angle, null)
                },
                AdjustCellSize = adjustCellSize
            };
            return builder.SetContent(content);
        }
    }
}