﻿namespace RxBim.Tools.TableBuilder.Services
{
    using System;
    using Models;

    /// <summary>
    /// The builder of a single <see cref="Column"/> of a <see cref="Table"/>.
    /// </summary>
    public class ColumnBuilder : CellsSetBuilder<Column, ColumnBuilder>
    {
        /// <inheritdoc />
        public ColumnBuilder(Column column)
            : base(column)
        {
        }

        /// <summary>
        /// Sets the width of the column.
        /// </summary>
        /// <param name="width">Column width.</param>
        public ColumnBuilder SetWidth(double width)
        {
            if (width <= 0)
                throw new ArgumentException("Must be a positive number.", nameof(width));

            ObjectForBuild.OwnWidth = width;
            return this;
        }
    }
}