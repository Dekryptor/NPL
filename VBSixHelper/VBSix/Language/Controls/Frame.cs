using System.Text;
using VBSix.Utilities;
using VBSix.Parsers;

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 Frame control.
    /// The Frame control is used as a container to group other controls visually and functionally.
    /// </summary>
    public class FrameProperties
    {
        public Appearance Appearance { get; set; } = Controls.Appearance.ThreeD; // VB6 Frame default is 3D
        public VB6Color BackColor { get; set; } = VB6Color.VbButtonFace; // Default: &H8000000F&
        public BorderStyle BorderStyle { get; set; } = Controls.BorderStyle.FixedSingle; // Default for Frame is FixedSingle (1)
        public string Caption { get; set; } = "Frame1"; // Default caption can vary
        public ClipControls ClipControls { get; set; } = Controls.ClipControls.Clipped; // Default is True (1)
        // DragIcon, MouseIcon are FRX-based.
        // Their resolved byte[] data will be stored in the VB6FrameKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        // Font is a group property.
        public VB6Color ForeColor { get; set; } = VB6Color.VbWindowText; // Default: &H80000008& (Text color for Caption)
        public int Height { get; set; } = 1500; // Typical default height in Twips
        public int HelpContextID { get; set; } = 0;
        public int Left { get; set; } = 120;   // Typical default Left in Twips
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None; // Frames can be OLE drop targets
        public TextDirection RightToLeft { get; set; } = Controls.TextDirection.LeftToRight; // Affects Caption display
        public int TabIndex { get; set; } = 0; // Frames themselves don't receive focus via Tab, but affect child tab order
        public string ToolTipText { get; set; } = "";
        public int Top { get; set; } = 120;    // Typical default Top in Twips
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 2055;   // Typical default width in Twips

        // Properties not typically directly on Frame but affect it or children:
        // Container, Parent (runtime)
        // Controls (collection of child controls, managed by VB6FrameKind)

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for a Frame control.
        /// </summary>
        public FrameProperties() { }

        /// <summary>
        /// Initializes FrameProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public FrameProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            BorderStyle = rawProps.GetEnum("BorderStyle", BorderStyle);
            Caption = rawProps.GetString("Caption", Caption);
            ClipControls = rawProps.GetEnum("ClipControls", ClipControls);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            // Font handled by PropertyGroup
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor);
            Height = rawProps.GetInt32("Height", Height);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            Left = rawProps.GetInt32("Left", Left);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            RightToLeft = rawProps.GetEnum("RightToLeft", RightToLeft);
            TabIndex = rawProps.GetInt32("TabIndex", TabIndex); // VB6 Frames have a TabIndex
            ToolTipText = rawProps.GetString("ToolTipText", ToolTipText);
            Top = rawProps.GetInt32("Top", Top);
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);
            Width = rawProps.GetInt32("Width", Width);
        }
    }
}
