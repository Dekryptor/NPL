// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For DrawMode, DrawStyle, Visibility enums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 Line control.
    /// </summary>
    public class LineProperties
    {
        public Vb6Color BorderColor { get; set; }
        public DrawStyle BorderStyle { get; set; } // Refers to the line's pattern
        public int BorderWidth { get; set; }   // Thickness of the line in pixels
        public DrawMode DrawMode { get; set; }
        public Visibility Visible { get; set; }
        public int X1 { get; set; } // Starting X-coordinate (in Twips, relative to container)
        public int Y1 { get; set; } // Starting Y-coordinate (in Twips, relative to container)
        public int X2 { get; set; } // Ending X-coordinate (in Twips, relative to container)
        public int Y2 { get; set; } // Ending Y-coordinate (in Twips, relative to container)

        /// <summary>
        /// Initializes a new instance of the <see cref="LineProperties"/> class with default VB6 values.
        /// </summary>
        public LineProperties()
        {
            BorderColor = Vb6Color.FromOleColor(0x80000008); // System Window Text (Black by default on white bg)
            BorderStyle = DrawStyle.Solid; // Default
            BorderWidth = 1; // Default thickness
            DrawMode = DrawMode.CopyPen; // Default
            Visible = Visibility.Visible; // Default
            X1 = 0;  // Example Default (Rust: 0)
            Y1 = 0;  // Example Default (Rust: 0)
            X2 = 1200; // Example Default (Rust: 100, which is very short for Twips)
            Y2 = 0;  // Example Default (Rust: 100, makes it diagonal; common default is horizontal or vertical)
                     // For a horizontal line, Y1 and Y2 would be same.
                     // For a vertical line, X1 and X2 would be same.
                     // VB IDE often places a default horizontal line.
        }
    }
}