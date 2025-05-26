using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 FileListBox control.
    /// The FileListBox control displays a list of files in a specified directory.
    /// </summary>
    public class FileListBoxProperties
    {
        public Appearance Appearance { get; set; } = Controls.Appearance.ThreeD;
        public ArchiveAttribute Archive { get; set; } = ArchiveAttribute.Include;
        public VB6Color BackColor { get; set; } = VB6Color.VbWindowBackground; // Default: &H80000005&
        public CausesValidation CausesValidation { get; set; } = Controls.CausesValidation.Yes;
        // DragIcon, MouseIcon are FRX-based.
        // Their resolved byte[] data will be stored in the VB6FileListBoxKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        // Font is a group property.
        public VB6Color ForeColor { get; set; } = VB6Color.VbWindowText; // Default: &H80000008&
        public int Height { get; set; } = 2175; // Typical default height in Twips
        public int HelpContextID { get; set; } = 0;
        public HiddenAttribute Hidden { get; set; } = HiddenAttribute.Exclude;
        public int Left { get; set; } = 120;  // Typical default Left in Twips
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public MultiSelect MultiSelect { get; set; } = Controls.MultiSelect.None;
        public NormalAttribute Normal { get; set; } = NormalAttribute.Include;
        public OLEDragMode OleDragMode { get; set; } = Controls.OLEDragMode.Manual;
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None; // VB6 default
        public string Pattern { get; set; } = "*.*"; // Default pattern
        public ReadOnlyAttribute ReadOnly { get; set; } = Controls.ReadOnlyAttribute.Include;
        public SystemAttribute System { get; set; } = Controls.SystemAttribute.Exclude; // VB6 default for System files is often Exclude
        public int TabIndex { get; set; } = 0; // Placeholder
        public TabStop TabStop { get; set; } = Controls.TabStop.Included;
        public string ToolTipText { get; set; } = "";
        public int Top { get; set; } = 120; // Typical default Top in Twips
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 2055; // Typical default width in Twips

        // Runtime properties not stored in .frm typically:
        // Path, FileName, List, ListCount, ListIndex, Selected, TopIndex

        /// <summary>
        /// Default constructor initializing with common VB6 defaults.
        /// </summary>
        public FileListBoxProperties() { }

        /// <summary>
        /// Initializes FileListBoxProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public FileListBoxProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            Archive = rawProps.GetEnum("Archive", Archive);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            CausesValidation = rawProps.GetEnum("CausesValidation", CausesValidation);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            // Font handled by PropertyGroup
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor);
            Height = rawProps.GetInt32("Height", Height);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            Hidden = rawProps.GetEnum("Hidden", Hidden);
            Left = rawProps.GetInt32("Left", Left);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            MultiSelect = rawProps.GetEnum("MultiSelect", MultiSelect);
            Normal = rawProps.GetEnum("Normal", Normal);
            OleDragMode = rawProps.GetEnum("OLEDragMode", OleDragMode);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            Pattern = rawProps.GetString("Pattern", Pattern);
            ReadOnly = rawProps.GetEnum("ReadOnly", ReadOnly); // Property is "ReadOnly", maps to our ReadOnlyAttribute
            System = rawProps.GetEnum("System", System);       // Property is "System", maps to our SystemAttribute
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