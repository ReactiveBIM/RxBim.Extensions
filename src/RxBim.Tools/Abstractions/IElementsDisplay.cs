﻿namespace RxBim.Tools
{
    using System.Collections.Generic;
    using JetBrains.Annotations;

    /// <summary>
    /// Object display service.
    /// </summary>
    /// <typeparam name="T">Object identifier type.</typeparam>
    [PublicAPI]
    public interface IElementsDisplay<in T>
        where T : struct
    {
        /// <summary>
        /// Makes the objects selected.
        /// </summary>
        /// <param name="ids">List of object identifiers.</param>
        void SetSelectedElements(IEnumerable<T> ids);

        /// <summary>
        /// Makes an object selected.
        /// </summary>
        /// <param name="id">Object identifier.</param>
        void SetSelectedElement(T id);

        /// <summary>
        /// Resets the current selection of objects.
        /// </summary>
        void ResetSelection();

        /// <summary>
        /// Sets the view on the object.
        /// </summary>
        /// <param name="id">Object identifier.</param>
        /// <param name="zoomFactor">
        /// Factor by which to zoom in or out. Values greater than 1 zooms in, less than 1 zooms out.
        /// </param>
        void Zoom(T id, double zoomFactor = 0.25);
    }
}