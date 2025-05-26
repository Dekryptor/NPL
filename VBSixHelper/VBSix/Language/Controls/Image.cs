using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 Image control.
    /// The Image control is used to display pictures. It's a lightweight alternative to PictureBox
    /// for simple image display without container capabilities or advanced graphics methods.
    /// </summary>
    public class ImageProperties
    {
        public Appearance Appearance { get; set; } = Controls.Appearance.Flat; // VB6 Image control default is Flat (0)
        public BorderStyle BorderStyle { get; set; } = Controls.BorderStyle.None; // Default is None (0)
        public string DataField { get; set; } = ""; // For data binding
        public string DataFormat { get; set; } = ""; // For data binding
        public string DataMember { get; set; } = ""; // For data binding
        public string DataSource { get; set; } = ""; // For data binding
        // DragIcon, MouseIcon, Picture are FRX-based.
        // Their resolved byte[] data will be stored in the VB6ImageKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        public int Height { get; set; } = 300;  // Typical default height in Twips, can vary
        public int Left { get; set; } = 120;   // Typical default Left in Twips
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public OLEDragMode OleDragMode { get; set; } = Controls.OLEDragMode.Manual; // Image can be an OLE drag source
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None;   // Image can be an OLE drop target
        public bool Stretch { get; set; } = false; // If True, picture resizes to fit control; if False, control resizes or clips.
        public string ToolTipText { get; set; } = "";
        public int Top { get; set; } = 120;    // Typical default Top in Twips
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 480;   // Typical default width in Twips

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for an Image control.
        /// </summary>
        public ImageProperties() { }

        /// <summary>
        /// Initializes ImageProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public ImageProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            BorderStyle = rawProps.GetEnum("BorderStyle", BorderStyle);
            DataField = rawProps.GetString("DataField", DataField);
            DataFormat = rawProps.GetString("DataFormat", DataFormat); 
            DataMember = rawProps.GetString("DataMember", DataMember);
            DataSource = rawProps.GetString("DataSource", DataSource);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            Height = rawProps.GetInt32("Height", Height);
            Left = rawProps.GetInt32("Left", Left);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            OleDragMode = rawProps.GetEnum("OLEDragMode", OleDragMode);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            Stretch = rawProps.GetBoolean("Stretch", Stretch);
            ToolTipText = rawProps.GetString("ToolTipText", ToolTipText);
            Top = rawProps.GetInt32("Top", Top);
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);
            Width = rawProps.GetInt32("Width", Width);
        }
    }
}