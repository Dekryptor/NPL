using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 DirListBox control.
    /// The DirListBox control displays a list of directories and allows navigation.
    /// </summary>
    public class DirListBoxProperties
    {
        public Appearance Appearance { get; set; } = Controls.Appearance.ThreeD;
        public VB6Color BackColor { get; set; } = VB6Color.VbWindowBackground; // Default: &H80000005&
        public CausesValidation CausesValidation { get; set; } = Controls.CausesValidation.Yes;
        // DragIcon, MouseIcon are FRX-based.
        // Their resolved byte[] data will be stored in the VB6DirListBoxKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        // Font is a group property.
        public VB6Color ForeColor { get; set; } = VB6Color.VbWindowText; // Default: &H80000008&
        public int Height { get; set; } = 3195; // Typical default height in Twips
        public int HelpContextID { get; set; } = 0;
        public int Left { get; set; } = 120;  // Typical default Left in Twips
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public OLEDragMode OleDragMode { get; set; } = Controls.OLEDragMode.Manual;
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None; // VB6 default
        public int TabIndex { get; set; } = 0; // Placeholder
        public TabStop TabStop { get; set; } = Controls.TabStop.Included;
        public string ToolTipText { get; set; } = "";
        public int Top { get; set; } = 120; // Typical default Top in Twips
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 2055; // Typical default width in Twips

        // Runtime properties not stored in .frm typically:
        // Path, List, ListCount, ListIndex, TopIndex

        /// <summary>
        /// Default constructor initializing with common VB6 defaults.
        /// </summary>
        public DirListBoxProperties() { }

        /// <summary>
        /// Initializes DirListBoxProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public DirListBoxProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            CausesValidation = rawProps.GetEnum("CausesValidation", CausesValidation);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            // Font handled by PropertyGroup
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor);
            Height = rawProps.GetInt32("Height", Height);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            Left = rawProps.GetInt32("Left", Left);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            OleDragMode = rawProps.GetEnum("OLEDragMode", OleDragMode);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
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