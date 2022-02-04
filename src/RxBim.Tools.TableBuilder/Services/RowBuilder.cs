﻿namespace RxBim.Tools.TableBuilder.Services
{
    using System;
    using System.Linq;
    using Models;

    /// <summary>
    /// The builder of a single <see cref="Row"/> of a <see cref="Table"/>.
    /// </summary>
    public class RowBuilder : CellsSetBuilder<Row, RowBuilder>
    {
        /// <inheritdoc />
        public RowBuilder(Row row)
            : base(row)
        {
        }

        /// <summary>
        /// Sets the height of the row.
        /// </summary>
        /// <param name="height">Row height value.</param>
        public RowBuilder SetHeight(double height)
        {
            if (height <= 0)
                throw new ArgumentException("Must be a positive number.", nameof(height));

            ObjectForBuild.OwnHeight = height;
            return this;
        }

        /// <summary>
        /// Merges all cells in the row.
        /// </summary>
        /// <param name="action">Delegate, applied to the cells to be merged.</param>
        public RowBuilder MergeRow(Action<CellBuilder, CellBuilder>? action = null)
        {
            ((CellBuilder)ObjectForBuild.Cells.First()).MergeNext(ObjectForBuild.Cells.Count() - 1, action);
            return this;
        }
    }
}