using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 ListBox control.
    /// The ListBox control displays a list of items from which the user can select one or more.
    /// </summary>
    public class ListBoxProperties
    {
        public Appearance Appearance { get; set; } = Controls.Appearance.ThreeD;
        public VB6Color BackColor { get; set; } = VB6Color.VbWindowBackground; // Default: &H80000005&
        public CausesValidation CausesValidation { get; set; } = Controls.CausesValidation.Yes;
        public int Columns { get; set; } = 0; // 0 for single column, >0 for multi-column (snaking)
        public string DataField { get; set; } = ""; // For data binding
        public string DataFormat { get; set; } = ""; // For data binding
        public string DataMember { get; set; } = ""; // For data binding
        public string DataSource { get; set; } = ""; // For data binding
        // DragIcon, MouseIcon are FRX-based.
        // List (initial items) and ItemData are also FRX-based for ListBox.
        // Their resolved data will be stored in the VB6ListBoxKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        // Font is a group property.
        public VB6Color ForeColor { get; set; } = VB6Color.VbWindowText; // Default: &H80000008&
        public int Height { get; set; } = 990;  // Typical default height in Twips
        public int HelpContextID { get; set; } = 0;
        public bool IntegralHeight { get; set; } = true; // If True, resizes to avoid partial items
        public int Left { get; set; } = 120;   // Typical default Left in Twips
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public MultiSelect MultiSelect { get; set; } = Controls.MultiSelect.None;
        public OLEDragMode OleDragMode { get; set; } = Controls.OLEDragMode.Manual;
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None; // VB6 default
        public TextDirection RightToLeft { get; set; } = Controls.TextDirection.LeftToRight;
        public bool Sorted { get; set; } = false; // If True, items are sorted alphabetically
        public ListBoxStyle Style { get; set; } = ListBoxStyle.Standard; // Standard or Checkbox style
        public int TabIndex { get; set; } = 0; // Placeholder
        public TabStop TabStop { get; set; } = Controls.TabStop.Included;
        public string ToolTipText { get; set; } = "";
        public int Top { get; set; } = 120;    // Typical default Top in Twips
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1875;   // Typical default width in Twips

        // Runtime properties not stored in .frm typically:
        // ListCount, ListIndex, NewIndex, SelCount, Selected, TopIndex, Text

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for a ListBox control.
        /// </summary>
        public ListBoxProperties() { }

        /// <summary>
        /// Initializes ListBoxProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public ListBoxProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            CausesValidation = rawProps.GetEnum("CausesValidation", CausesValidation);
            Columns = rawProps.GetInt32("Columns", Columns);
            DataField = rawProps.GetString("DataField", DataField);
            DataFormat = rawProps.GetString("DataFormat", DataFormat);
            DataMember = rawProps.GetString("DataMember", DataMember);
            DataSource = rawProps.GetString("DataSource", DataSource);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            // Font handled by PropertyGroup
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor);
            Height = rawProps.GetInt32("Height", Height);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            IntegralHeight = rawProps.GetBoolean("IntegralHeight", IntegralHeight);
            Left = rawProps.GetInt32("Left", Left);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            MultiSelect = rawProps.GetEnum("MultiSelect", MultiSelect);
            OleDragMode = rawProps.GetEnum("OLEDragMode", OleDragMode);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            RightToLeft = rawProps.GetEnum("RightToLeft", RightToLeft);
            Sorted = rawProps.GetBoolean("Sorted", Sorted);
            Style = rawProps.GetEnum("Style", Style);
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