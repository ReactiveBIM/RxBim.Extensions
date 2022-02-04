namespace RxBim.Tools.Autocad.Serializers
{
    using Autodesk.AutoCAD.DatabaseServices;

    /// <summary>
    /// Serialization parameters to AutoCAD table
    /// </summary>
    public class AutocadTableSerializerParameters
    {
        /// <summary>
        /// The default height of a table rows.
        /// </summary>
        public double DefaultRowHeight { get; set; } = 8;

        /// <summary>
        /// The style identifier for a table.
        /// </summary>
        public ObjectId TableStyleId { get; set; }

        /// <summary>
        /// The text style identifier for a table text.
        /// </summary>
        public ObjectId TextStyleId { get; set; }

        /// <summary>
        /// The drawing database into which a table is inserted.
        /// </summary>
        public Database? TargetDatabase { get; set; }

        /// <summary>
        /// Bold line weight.
        /// </summary>
        public LineWeight BoldLine { get; set; } = LineWeight.LineWeight050;

        /// <summary>
        /// Thin line weight.
        /// </summary>
        public LineWeight ThinLine { get; set; } = LineWeight.LineWeight018;
    }
}