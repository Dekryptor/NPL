// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class LineProperties : ControlSpecificPropertiesBase
    {
        public Vb6Color BorderColor { get; set; } = Vb6Color.Black; // Default is usually black
        public LineBorderStyleConstants BorderStyle { get; set; } = LineBorderStyleConstants.Solid; // Or Transparent, Dashed etc.
        public int BorderWidth { get; set; } = 1; // In pixels
        public DrawModeConstants DrawMode { get; set; } = DrawModeConstants.CopyPen;
        public int Index { get; set; } = -1; // Contextual
        public string Tag { get; set; } = string.Empty;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int X1 { get; set; } = 0; // Starting X-coordinate
        public int Y1 { get; set; } = 0; // Starting Y-coordinate
        public int X2 { get; set; } = 0; // Ending X-coordinate (e.g., 1200)
        public int Y2 { get; set; } = 0; // Ending Y-coordinate (e.g., 0 for a horizontal line)

        // Line controls do not have Font, BackColor, ForeColor, Caption, TabStop, Enabled (meaningfully) etc.
        // Left, Top, Height, Width are implicit from X1, Y1, X2, Y2.
        // The base Font property will be nullified.

        public LineProperties()
        {
            this.Font = null; // Line controls don't have a Font property.
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            BorderColor = PropertyParsingHelpers.GetColor(textualProps, "BorderColor", this.BorderColor);
            BorderStyle = PropertyParsingHelpers.GetEnum(textualProps, "BorderStyle", this.BorderStyle);
            BorderWidth = PropertyParsingHelpers.GetInt32(textualProps, "BorderWidth", this.BorderWidth);
            DrawMode = PropertyParsingHelpers.GetEnum(textualProps, "DrawMode", this.DrawMode);
            // Index is typically handled by Vb6Control, but if explicitly set:
            // if (textualProps.TryGetValue("Index", out string indexValStr) && int.TryParse(indexValStr, out int parsedIndex)) { /* this.Index = parsedIndex; */ }
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", this.Visible);
            X1 = PropertyParsingHelpers.GetInt32(textualProps, "X1", this.X1);
            Y1 = PropertyParsingHelpers.GetInt32(textualProps, "Y1", this.Y1);
            X2 = PropertyParsingHelpers.GetInt32(textualProps, "X2", this.X2);
            Y2 = PropertyParsingHelpers.GetInt32(textualProps, "Y2", this.Y2);

            // No binary properties expected for a standard Line control.
            // No standard Font property group.
        }

        protected override void PopulateFontProperty(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            // Line controls do not have a Font property.
            this.Font = null;
        }
    }

    // This enum might be specific to Line/Shape or shared if values overlap with DrawStyleConstants
    public enum LineBorderStyleConstants
    {
        Transparent = 0,
        Solid = 1,
        Dash = 2,
        Dot = 3,
        DashDot = 4,
        DashDotDot = 5,
        InsideSolid = 6 // For Shapes, might not apply to Line directly but VB uses a shared property editor sometimes
    }
}