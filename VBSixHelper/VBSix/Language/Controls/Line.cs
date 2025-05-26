using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 Line control.
    /// The Line control is used to draw horizontal, vertical, or diagonal lines on a Form or other container.
    /// </summary>
    public class LineProperties
    {
        public VB6Color BorderColor { get; set; } = VB6Color.VbWindowText; // Default: &H80000008& (System Color: Window Text)
        
        /// <summary>
        /// Gets or sets the style of the line's border. Uses the <see cref="Controls.DrawStyle"/> enum.
        /// VB6 Property: BorderStyle. Note: This is different from the common BorderStyle enum used by other controls.
        /// </summary>
        public DrawStyle BorderStyle { get; set; } = Controls.DrawStyle.Solid; // Default for Line is Solid (0)
        
        public int BorderWidth { get; set; } = 1; // In pixels
        
        /// <summary>
        /// Gets or sets how the line is drawn relative to existing graphics.
        /// VB6 Property: DrawMode.
        /// </summary>
        public DrawMode DrawMode { get; set; } = Controls.DrawMode.CopyPen;
        
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        
        /// <summary>
        /// The X-coordinate of the starting point of the line, in container's scale units (e.g., Twips).
        /// </summary>
        public int X1 { get; set; } = 0; 
        
        /// <summary>
        /// The Y-coordinate of the starting point of the line, in container's scale units.
        /// </summary>
        public int Y1 { get; set; } = 0; 
        
        /// <summary>
        /// The X-coordinate of the ending point of the line, in container's scale units.
        /// </summary>
        public int X2 { get; set; } = 1440; // Default 1 inch (1440 Twips)
        
        /// <summary>
        /// The Y-coordinate of the ending point of the line, in container's scale units.
        /// </summary>
        public int Y2 { get; set; } = 1440; // Default 1 inch (1440 Twips)

        // Index, Name, Tag are part of the parent VB6Control.
        // Line controls do not have properties like Left, Top, Width, Height directly;
        // their position and size are defined by X1, Y1, X2, Y2.
        // They also don't typically have Font, BackColor, ForeColor (BorderColor is used for line color).

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for a Line control.
        /// </summary>
        public LineProperties() { }

        /// <summary>
        /// Initializes LineProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public LineProperties(IDictionary<string, PropertyValue> rawProps)
        {
            BorderColor = rawProps.GetVB6Color("BorderColor", BorderColor);
            BorderStyle = rawProps.GetEnum("BorderStyle", BorderStyle); // Uses DrawStyle enum
            BorderWidth = rawProps.GetInt32("BorderWidth", BorderWidth);
            DrawMode = rawProps.GetEnum("DrawMode", DrawMode);
            Visible = rawProps.GetEnum("Visible", Visible);
            X1 = rawProps.GetInt32("X1", X1);
            Y1 = rawProps.GetInt32("Y1", Y1);
            X2 = rawProps.GetInt32("X2", X2);
            Y2 = rawProps.GetInt32("Y2", Y2);
        }
    }
}