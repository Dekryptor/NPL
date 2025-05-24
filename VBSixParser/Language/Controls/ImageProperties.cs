// Namespace: VB6Parse.Language.Controls
// For common enums like Appearance, BorderStyle, etc. they are in this namespace or CommonControlEnums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 Image control.
    /// </summary>
    public class ImageProperties
    {
        public Appearance Appearance { get; set; } // Typically Flat (0) for Image control
        public BorderStyle BorderStyle { get; set; } // None or FixedSingle
        public string DataField { get; set; }
        // public string DataFormat { get; set; } // StdDataFormat object in VB6. String for simplicity.
        public string DataMember { get; set; }
        public string DataSource { get; set; }
        // public byte[] DragIcon { get; set; } // Placeholder for image data
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        public int Height { get; set; } // Typically in Twips
        public int Left { get; set; }   // Typically in Twips
        // public byte[] MouseIcon { get; set; } // Placeholder for image data
        public MousePointer MousePointer { get; set; }
        public OLEDragMode OLEDragMode { get; set; }
        public OLEDropMode OLEDropMode { get; set; }
        // public byte[] Picture { get; set; } // Placeholder for image data
        public bool Stretch { get; set; } // True to resize picture to fit control, False to clip.
        public string ToolTipText { get; set; }
        public int Top { get; set; }     // Typically in Twips
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }   // Typically in Twips

        /// <summary>
        /// Initializes a new instance of the <see cref="ImageProperties"/> class with default VB6 values.
        /// </summary>
        public ImageProperties()
        {
            Appearance = Appearance.Flat; // VB6 Image default Appearance is Flat (0)
            BorderStyle = BorderStyle.None; // Default
            DataField = string.Empty;
            // DataFormat = string.Empty;
            DataMember = string.Empty;
            DataSource = string.Empty;
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            Height = 975; // Example default from Rust. VB actual default depends on image loaded or is small.
            Left = 1080;  // Example default from Rust.
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            OLEDragMode = OLEDragMode.Manual;
            OLEDropMode = OLEDropMode.None; // Image control default
            // Picture = null;
            Stretch = false; // Default is False (clip image)
            ToolTipText = string.Empty;
            Top = 960;    // Example default from Rust.
            Visible = Visibility.Visible;
            WhatsThisHelpId = 0;
            Width = 615;  // Example default from Rust.
        }
    }
}