using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary

namespace VBSix.Language.Controls
{
    // ShapeType enum is already defined in VB6ControlEnums.cs
    // public enum ShapeType { Rectangle = 0, Square = 1, ... }

    /// <summary>
    /// Represents the properties specific to a VB6 Shape control.
    /// The Shape control is used to draw various graphical shapes like rectangles, ovals, etc.
    /// </summary>
    public class ShapeProperties
    {
        public VB6Color BackColor { get; set; } = VB6Color.VbWindowBackground;
        public BackStyle BackStyle { get; set; } = Controls.BackStyle.Opaque; // VB6 default is Opaque (1).
        public VB6Color BorderColor { get; set; } = VB6Color.VbWindowText;   // Default: &H80000008& (System Color: Window Text)
        
        /// <summary>
        /// Gets or sets the style of the shape's border. Uses the <see cref="Controls.DrawStyle"/> enum.
        /// VB6 Property: BorderStyle.
        /// </summary>
        public DrawStyle BorderStyle { get; set; } = Controls.DrawStyle.Solid; // Default for Shape is Solid (0)
        
        public int BorderWidth { get; set; } = 1; // In pixels
        
        /// <summary>
        /// Gets or sets how the shape is drawn relative to existing graphics.
        /// VB6 Property: DrawMode.
        /// </summary>
        public DrawMode DrawMode { get; set; } = Controls.DrawMode.CopyPen;
        
        public VB6Color FillColor { get; set; } = VB6Color.VbBlack; // Default: &H00000000& (Black) if FillStyle is not Transparent
        
        /// <summary>
        /// Gets or sets the pattern used to fill the shape. Uses the <see cref="Controls.DrawStyle"/> enum
        /// (despite the name, FillStyle for Shape uses DrawStyle values).
        /// VB6 Property: FillStyle.
        /// </summary>
        public DrawStyle FillStyle { get; set; } = Controls.DrawStyle.Transparent; // Default for Shape is Transparent (5)
        
        public int Height { get; set; } = 495;  // Typical default height in Twips
        public int Left { get; set; } = 120;   // Typical default Left in Twips
        
        /// <summary>
        /// Gets or sets the type of shape to draw (Rectangle, Oval, etc.).
        /// VB6 Property: Shape.
        /// </summary>
        public ShapeType Shape { get; set; } = ShapeType.Rectangle; // Default is Rectangle (0)
        
        public int Top { get; set; } = 120;    // Typical default Top in Twips
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int Width { get; set; } = 1215;   // Typical default width in Twips

        // Shape controls do not have properties like:
        // Caption, Font, TabIndex, TabStop, CausesValidation, DataField, etc.
        // They are purely graphical elements.
        // Index, Name, Tag are part of the parent VB6Control.

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for a Shape control.
        /// </summary>
        public ShapeProperties() { }

        /// <summary>
        /// Initializes ShapeProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public ShapeProperties(IDictionary<string, PropertyValue> rawProps)
        {
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            BackStyle = rawProps.GetEnum("BackStyle", BackStyle);
            BorderColor = rawProps.GetVB6Color("BorderColor", BorderColor);
            BorderStyle = rawProps.GetEnum("BorderStyle", BorderStyle); // Uses DrawStyle enum
            BorderWidth = rawProps.GetInt32("BorderWidth", BorderWidth);
            DrawMode = rawProps.GetEnum("DrawMode", DrawMode);
            FillColor = rawProps.GetVB6Color("FillColor", FillColor);
            FillStyle = rawProps.GetEnum("FillStyle", FillStyle); // Uses DrawStyle enum
            Height = rawProps.GetInt32("Height", Height);
            Left = rawProps.GetInt32("Left", Left);
            Shape = rawProps.GetEnum("Shape", Shape); // Uses ShapeType enum
            Top = rawProps.GetInt32("Top", Top);
            Visible = rawProps.GetEnum("Visible", Visible);
            Width = rawProps.GetInt32("Width", Width);
        }
    }
}