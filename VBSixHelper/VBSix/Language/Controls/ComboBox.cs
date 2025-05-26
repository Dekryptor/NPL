using VBSix.Language;
using VBSix.Utilities;
using VBSix.Parsers;

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 ComboBox control.
    /// These properties are typically found in a .frm file.
    /// </summary>
    public class ComboBoxProperties
    {
        public Appearance Appearance { get; set; } = Controls.Appearance.ThreeD;
        public VB6Color BackColor { get; set; } = VB6Color.VbWindowBackground; // Default: &H80000005&
        public CausesValidation CausesValidation { get; set; } = Controls.CausesValidation.Yes;
        public string DataField { get; set; } = "";
        public string DataFormat { get; set; } = "";
        public string DataMember { get; set; } = "";
        public string DataSource { get; set; } = "";
        // DragIcon, MouseIcon are FRX-based.
        // ItemData and List (for initial items) are also FRX-based.
        // Their resolved data will be stored in the VB6ComboBoxKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        public VB6Color ForeColor { get; set; } = VB6Color.VbWindowText; // Default: &H80000008&
        public int Height { get; set; } = 315; // Typical default height in Twips for a drop-down combo
        public int HelpContextID { get; set; } = 0;
        public bool IntegralHeight { get; set; } = true; // Default is True
        public int Left { get; set; } = 120; // Typical default Left in Twips
        public bool Locked { get; set; } = false; // Default is False
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public OLEDragMode OleDragMode { get; set; } = Controls.OLEDragMode.Manual;
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None; // VB6 default
        public TextDirection RightToLeft { get; set; } = Controls.TextDirection.LeftToRight;
        public bool Sorted { get; set; } = false; // Default is False
        public ComboBoxStyle Style { get; set; } = ComboBoxStyle.DropDownCombo; // Default style
        public int TabIndex { get; set; } = 0; // Placeholder
        public TabStop TabStop { get; set; } = Controls.TabStop.Included;
        public string Text { get; set; } = ""; // The initial text in the edit portion (if applicable)
        public string ToolTipText { get; set; } = "";
        public int Top { get; set; } = 120; // Typical default Top in Twips
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1875; // Typical default width in Twips

        /// <summary>
        /// Default constructor initializing with common VB6 defaults.
        /// </summary>
        public ComboBoxProperties() { }

        /// <summary>
        /// Initializes ComboBoxProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public ComboBoxProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            CausesValidation = rawProps.GetEnum("CausesValidation", CausesValidation);
            DataField = rawProps.GetString("DataField", DataField);
            DataFormat = rawProps.GetString("DataFormat", DataFormat);
            DataMember = rawProps.GetString("DataMember", DataMember);
            DataSource = rawProps.GetString("DataSource", DataSource);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor);
            Height = rawProps.GetInt32("Height", Height);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            IntegralHeight = rawProps.GetBoolean("IntegralHeight", IntegralHeight);
            Left = rawProps.GetInt32("Left", Left);
            Locked = rawProps.GetBoolean("Locked", Locked);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            OleDragMode = rawProps.GetEnum("OLEDragMode", OleDragMode);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            RightToLeft = rawProps.GetEnum("RightToLeft", RightToLeft);
            Sorted = rawProps.GetBoolean("Sorted", Sorted);
            Style = rawProps.GetEnum("Style", Style);
            TabIndex = rawProps.GetInt32("TabIndex", TabIndex);
            TabStop = rawProps.GetEnum("TabStop", TabStop);
            Text = rawProps.GetString("Text", Text); // .frm stores initial text if Style allows typing
            ToolTipText = rawProps.GetString("ToolTipText", ToolTipText);
            Top = rawProps.GetInt32("Top", Top);
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);
            Width = rawProps.GetInt32("Width", Width);
        }
    }
}