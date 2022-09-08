﻿namespace RxBim.Tools.Revit
{
    using Autodesk.Revit.DB;

    /// <summary>
    /// The transaction context for <see cref="Document"/>.
    /// </summary>
    public class DocumentContext : TransactionContext<Document>
    {
        /// <inheritdoc />
        public DocumentContext(Document contextObject)
            : base(contextObject)
        {
        }
    }
}