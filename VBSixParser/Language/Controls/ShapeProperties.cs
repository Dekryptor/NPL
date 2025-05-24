// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For BackStyle, DrawMode, DrawStyle, FillStyleType, Visibility enums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Defines the type of shape for a Shape control.
    /// </summary>
    public enum ShapeType
    {
        /// <summary>Rectangle. (Default)</summary>
        Rectangle = 0,
        /// <summary>Square.</summary>
        Square = 1,
        /// <summary>Oval.</summary>
        Oval = 2,
        /// <summary>Circle.</summary>
        Circle = 3,
        /// <summary>Rounded Rectangle.</summary>
        RoundedRectangle = 4,
        /// <summary>Rounded Square.</summary>
        RoundedSquare = 5 // Note: VB6 name is "Rounded Square"
    }

    /// <summary>
    /// Properties specific to a VB6 Shape control.
    /// </summary>
    public class ShapeProperties
    {
        public Vb6Color BackColor { get; set; } // Color used for patterned fills or if BackStyle is Opaque
        public BackStyle BackStyle { get; set; } // Opaque or Transparent
        public Vb6Color BorderColor { get; set; }
        public DrawStyle BorderStyle { get; set; } // Pattern of the border line
        public int BorderWidth { get; set; }   // Thickness of the border in pixels
        public DrawMode DrawMode { get; set; }
        public Vb6Color FillColor { get; set; } // Color used for solid or patterned fills
        public FillStyleType FillStyle { get; set; } // Pattern used to fill the shape
        public int Height { get; set; } // Typically in Twips
        public int Left { get; set; }   // Typically in Twips
        public ShapeType Shape { get; set; } // The type of shape to draw
        public int Top { get; set; }     // Typically in Twips
        public Visibility Visible { get; set; }
        public int Width { get; set; }   // Typically in Twips

        /// <summary>
        /// Initializes a new instance of the <see cref="ShapeProperties"/> class with default VB6 values.
        /// </summary>
        public ShapeProperties()
        {
            // VB6 defaults for Shape control:
            BackColor = Vb6Color.FromOleColor(0x80000005); // System Window Background (used if BackStyle=Opaque)
                                                            // Rust uses System { index: 5 } which is Window Background
            BackStyle = BackStyle.Opaque;                   // VB6 Shape default is Opaque. Rust has Transparent.
                                                            // Let's stick to VB6 behavior.
            BorderColor = Vb6Color.FromOleColor(0x80000008); // System Window Text (Black by default on white bg)
                                                            // Rust uses System { index: 8 } which is Window Text
            BorderStyle = DrawStyle.Solid; // Default
            BorderWidth = 1; // Default
            DrawMode = DrawMode.CopyPen; // Default
            FillColor = Vb6Color.FromRGB(0, 0, 0); // Black. Used if FillStyle is not Transparent.
                                                   // Rust has this as black. VB6 default FillColor is 0 (Black).
            FillStyle = FillStyleType.Transparent; // VB6 Shape default FillStyle is Transparent. Rust has Transparent.
            Height = 495;  // Example Default (Rust: 355). VB6 IDE default is often related to grid size.
            Left = 0;      // Example Default (Rust: 30)
            Shape = ShapeType.Rectangle; // Default
            Top = 0;       // Example Default (Rust: 200)
            Visible = Visibility.Visible; // Default
            Width = 1215;  // Example Default (Rust: 355)
        }
    }
}