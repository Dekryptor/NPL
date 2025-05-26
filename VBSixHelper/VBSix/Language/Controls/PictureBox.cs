using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 PictureBox control.
    /// The PictureBox control can display graphics, act as a container for other controls,
    /// and serve as a target for graphics methods.
    /// </summary>
    public class PictureBoxProperties
    {
        public Align Align { get; set; } = Controls.Align.None; // How the PictureBox aligns if its container supports it
        public Appearance Appearance { get; set; } = Controls.Appearance.ThreeD; // Default is 3D
        public AutoRedraw AutoRedraw { get; set; } = Controls.AutoRedraw.Manual;
        public AutoSize AutoSize { get; set; } = Controls.AutoSize.Fixed; // If True, resizes to fit its Picture
        public VB6Color BackColor { get; set; } = VB6Color.VbButtonFace; // Default: &H8000000F&
        public BorderStyle BorderStyle { get; set; } = Controls.BorderStyle.FixedSingle; // Default is FixedSingle (1)
        public CausesValidation CausesValidation { get; set; } = Controls.CausesValidation.Yes;
        public ClipControls ClipControls { get; set; } = Controls.ClipControls.Clipped;
        public string DataField { get; set; } = ""; // For data binding
        public string DataFormat { get; set; } = ""; // For data binding
        public string DataMember { get; set; } = ""; // For data binding
        public string DataSource { get; set; } = ""; // For data binding
        // DragIcon, MouseIcon, Picture are FRX-based.
        // Their resolved byte[] data will be stored in the VB6PictureBoxKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public DrawMode DrawMode { get; set; } = Controls.DrawMode.CopyPen;
        public DrawStyle DrawStyle { get; set; } = Controls.DrawStyle.Solid;
        public int DrawWidth { get; set; } = 1; // In pixels
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        public VB6Color FillColor { get; set; } = VB6Color.VbBlack; // Default: &H00000000&
        public FillStyle FillStyle { get; set; } = Controls.FillStyle.Transparent;
        // Font is a group property.
        public FontTransparency FontTransparent { get; set; } = Controls.FontTransparency.Transparent;
        public VB6Color ForeColor { get; set; } = VB6Color.VbWindowText; // Default: &H80000008&
        public HasDeviceContext HasDC { get; set; } = Controls.HasDeviceContext.Yes; // PictureBoxes usually have a DC
        public int Height { get; set; } = 1200; // Typical default height in Twips
        public int HelpContextID { get; set; } = 0;
        // Image property (runtime, for GDI drawing) is distinct from Picture property (design-time image)
        public int Left { get; set; } = 120;   // Typical default Left in Twips
        public string LinkItem { get; set; } = ""; // For DDE
        public LinkMode LinkMode { get; set; } = Controls.LinkMode.None; // For DDE
        public int LinkTimeout { get; set; } = 50; // For DDE
        public string LinkTopic { get; set; } = ""; // For DDE
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public bool Negotiate { get; set; } = false; // For OLE container features when PictureBox hosts an OLE object
        public OLEDragMode OleDragMode { get; set; } = Controls.OLEDragMode.Manual;
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None;
        public TextDirection RightToLeft { get; set; } = Controls.TextDirection.LeftToRight; // For text drawn on PictureBox
        public int ScaleHeight { get; set; } = 1140; // Inner height for scaling, often Height - border
        public int ScaleLeft { get; set; } = 0;
        public ScaleMode ScaleMode { get; set; } = Controls.ScaleMode.Twip;
        public int ScaleTop { get; set; } = 0;
        public int ScaleWidth { get; set; } = 2220; // Inner width for scaling
        public int TabIndex { get; set; } = 0; // Placeholder
        public TabStop TabStop { get; set; } = Controls.TabStop.Included; // Can receive focus
        public string ToolTipText { get; set; } = "";
        public int Top { get; set; } = 120;    // Typical default Top in Twips
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 2280;   // Typical default width in Twips
        // CurrentX, CurrentY are runtime graphics properties.

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for a PictureBox control.
        /// </summary>
        public PictureBoxProperties() { }

        /// <summary>
        /// Initializes PictureBoxProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public PictureBoxProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Align = rawProps.GetEnum("Align", Align);
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            AutoRedraw = rawProps.GetEnum("AutoRedraw", AutoRedraw);
            AutoSize = rawProps.GetEnum("AutoSize", AutoSize);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            BorderStyle = rawProps.GetEnum("BorderStyle", BorderStyle);
            CausesValidation = rawProps.GetEnum("CausesValidation", CausesValidation);
            ClipControls = rawProps.GetEnum("ClipControls", ClipControls);
            DataField = rawProps.GetString("DataField", DataField);
            DataFormat = rawProps.GetString("DataFormat", DataFormat);
            DataMember = rawProps.GetString("DataMember", DataMember);
            DataSource = rawProps.GetString("DataSource", DataSource);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            DrawMode = rawProps.GetEnum("DrawMode", DrawMode);
            DrawStyle = rawProps.GetEnum("DrawStyle", DrawStyle);
            DrawWidth = rawProps.GetInt32("DrawWidth", DrawWidth);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            FillColor = rawProps.GetVB6Color("FillColor", FillColor);
            FillStyle = rawProps.GetEnum("FillStyle", FillStyle);
            FontTransparent = rawProps.GetEnum("FontTransparent", FontTransparent);
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor);
            HasDC = rawProps.GetEnum("HasDC", HasDC);
            Height = rawProps.GetInt32("Height", Height);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            Left = rawProps.GetInt32("Left", Left);
            LinkItem = rawProps.GetString("LinkItem", LinkItem);
            LinkMode = rawProps.GetEnum("LinkMode", LinkMode);
            LinkTimeout = rawProps.GetInt32("LinkTimeout", LinkTimeout);
            LinkTopic = rawProps.GetString("LinkTopic", LinkTopic);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            Negotiate = rawProps.GetBoolean("Negotiate", Negotiate);
            OleDragMode = rawProps.GetEnum("OLEDragMode", OleDragMode);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            RightToLeft = rawProps.GetEnum("RightToLeft", RightToLeft);
            ScaleHeight = rawProps.GetInt32("ScaleHeight", ScaleHeight);
            ScaleLeft = rawProps.GetInt32("ScaleLeft", ScaleLeft);
            ScaleMode = rawProps.GetEnum("ScaleMode", ScaleMode);
            ScaleTop = rawProps.GetInt32("ScaleTop", ScaleTop);
            ScaleWidth = rawProps.GetInt32("ScaleWidth", ScaleWidth);
            TabIndex = rawProps.GetInt32("TabIndex", TabIndex);
            TabStop = rawProps.GetEnum("TabStop", TabStop);
            ToolTipText = rawProps.GetString("ToolTipText", ToolTipText);
            Top = rawProps.GetInt32("Top", Top);
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);
            Width = rawProps.GetInt32("Width", Width);
        }
    }
}