// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class ShapeProperties : ControlSpecificPropertiesBase
    {
        public Vb6Color BackColor { get; set; } // Color for pattern background or if FillStyle is Transparent
        public BackStyleConstants BackStyle { get; set; } = BackStyleConstants.Opaque; // Opaque or Transparent background for patterned fills
        public Vb6Color BorderColor { get; set; } = Vb6Color.Black;
        public LineBorderStyleConstants BorderStyle { get; set; } = LineBorderStyleConstants.Solid; // Reusing LineBorderStyleConstants
        public int BorderWidth { get; set; } = 1; // In pixels
        public DrawModeConstants DrawMode { get; set; } = DrawModeConstants.CopyPen;
        public Vb6Color FillColor { get; set; } = Vb6Color.Black; // Color for the shape's fill
        public FillStyleConstants FillStyle { get; set; } = FillStyleConstants.Transparent; // Solid, Transparent, patterns
        public int Height { get; set; } = 735; // Default in Twips
        public int Index { get; set; } = -1; // Contextual
        public int Left { get; set; } = 0;
        public ShapeConstants Shape { get; set; } = ShapeConstants.Rectangle;
        public string Tag { get; set; } = string.Empty;
        public int Top { get; set; } = 0;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int Width { get; set; } = 1215; // Default in Twips

        // Shape controls do not have Font, Caption, TabStop, Enabled (meaningfully for interaction) etc.
        // The base Font property will be nullified.

        public ShapeProperties()
        {
            this.Font = null; // Shape controls don't have a Font property.
            // Default BackColor for Shape is often the container's color, or a system color.
            // Using a common default like ButtonFace or WindowBackground.
            BackColor = Vb6Color.DefaultSystemColors.ButtonFace;
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            BackColor = PropertyParsingHelpers.GetColor(textualProps, "BackColor", this.BackColor);
            BackStyle = PropertyParsingHelpers.GetEnum(textualProps, "BackStyle", this.BackStyle);
            BorderColor = PropertyParsingHelpers.GetColor(textualProps, "BorderColor", this.BorderColor);
            BorderStyle = PropertyParsingHelpers.GetEnum(textualProps, "BorderStyle", this.BorderStyle);
            BorderWidth = PropertyParsingHelpers.GetInt32(textualProps, "BorderWidth", this.BorderWidth);
            DrawMode = PropertyParsingHelpers.GetEnum(textualProps, "DrawMode", this.DrawMode);
            FillColor = PropertyParsingHelpers.GetColor(textualProps, "FillColor", this.FillColor);
            FillStyle = PropertyParsingHelpers.GetEnum(textualProps, "FillStyle", this.FillStyle);
            Height = PropertyParsingHelpers.GetInt32(textualProps, "Height", this.Height);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", this.Left);
            Shape = PropertyParsingHelpers.GetEnum(textualProps, "Shape", this.Shape);
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", this.Top);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", this.Visible);
            Width = PropertyParsingHelpers.GetInt32(textualProps, "Width", this.Width);

            // No binary properties expected for a standard Shape control.
            // No standard Font property group.
        }
        
        protected override void PopulateFontProperty(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            // Shape controls do not have a Font property.
            this.Font = null;
        }
    }

    public enum ShapeConstants
    {
        Rectangle = 0,
        Square = 1,
        Oval = 2,
        Circle = 3,
        RoundedRectangle = 4,
        RoundedSquare = 5
    }

    // BackStyleConstants might already exist or be similar to Appearance.
    // For Shape, BackStyle (Opaque/Transparent) affects how patterned fills interact with BackColor.
    // public enum BackStyleConstants // Often 0 for Transparent, 1 for Opaque
    // {
    //     Transparent = 0,
    //     Opaque = 1
    // }
    // FillStyleConstants and DrawModeConstants would be defined elsewhere (e.g., with PictureBox)
    // LineBorderStyleConstants is reused from LineProperties.
}