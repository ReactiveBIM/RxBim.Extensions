﻿namespace RxBim.Tools.Revit.Services
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Abstractions;
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using JetBrains.Annotations;

    /// <inheritdoc />
    [UsedImplicitly]
    internal class ElementsDisplayService : IElementsDisplay
    {
        private readonly UIApplication _uiApplication;

        /// <summary>
        /// Initializes a new instance of the <see cref="ElementsDisplayService"/> class.
        /// </summary>
        /// <param name="uiApplication">Current <see cref="UIApplication"/></param>
        public ElementsDisplayService(UIApplication uiApplication)
        {
            _uiApplication = uiApplication;
        }

        /// <inheritdoc />
        public void SetSelectedElements(IEnumerable<int> elementIds)
        {
            _uiApplication.ActiveUIDocument.Selection.SetElementIds(elementIds.Select(e => new ElementId(e)).ToList());
        }

        /// <inheritdoc />
        public void SetSelectedElement(int elementId)
        {
            _uiApplication.ActiveUIDocument.Selection.SetElementIds(new List<ElementId> { new ElementId(elementId) });
        }

        /// <inheritdoc />
        public void ResetSelection()
        {
            _uiApplication.ActiveUIDocument.Selection.SetElementIds(new List<ElementId>());
        }

        /// <inheritdoc />
        public void Zoom(int elementId, double zoomFactor = 0.25)
        {
            var activeView = _uiApplication.ActiveUIDocument.ActiveView;
            if (activeView == null)
                return;

            var openUiViews = _uiApplication.ActiveUIDocument.GetOpenUIViews();

            var currentUiView = openUiViews
                .FirstOrDefault(x => x.ViewId == activeView.Id);
            if (currentUiView == null)
                return;

            var document = activeView.Document;

            var element = document.GetElement(new ElementId(elementId));
            var boundingBox = element?.get_BoundingBox(null);
            if (boundingBox == null)
                return;

            var (bottomLeft, upperRight) = GetTransformedRectangleCorners(boundingBox);

            currentUiView.ZoomAndCenterRectangle(bottomLeft, upperRight);
            currentUiView.Zoom(zoomFactor);
        }

        private (XYZ BottomLeft, XYZ UpperRight) GetTransformedRectangleCorners(BoundingBoxXYZ boundingBox)
        {
            var transform = Transform.Identity;

            var minTransformed = transform.OfPoint(boundingBox.Min);
            var maxTransformed = transform.OfPoint(boundingBox.Max);

            var bottomLeft = CombineCoords(minTransformed, maxTransformed, Math.Min);
            var upperRight = CombineCoords(minTransformed, maxTransformed, Math.Max);

            return (bottomLeft, upperRight);
        }

        private XYZ CombineCoords(XYZ point1, XYZ point2, Func<double, double, double> combineFunc)
        {
            return new XYZ(
                combineFunc(point1.X, point2.X),
                combineFunc(point1.Y, point2.Y),
                combineFunc(point1.Z, point2.Z));
        }
    }
}