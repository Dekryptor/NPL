using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary

namespace VBSix.Language.Controls
{
    // ScrollBars and MultiLine enums are already defined in VB6ControlEnums.cs
    // public enum ScrollBars { None = 0, Horizontal = 1, Vertical = 2, Both = 3 }
    // public enum MultiLine { SingleLine = 0, MultiLine = -1 }

    /// <summary>
    /// Represents the properties specific to a VB6 TextBox control.
    /// The TextBox control is used for displaying or editing text.
    /// </summary>
    public class TextBoxProperties
    {
        public Alignment Alignment { get; set; } = Controls.Alignment.LeftJustify;
        public Appearance Appearance { get; set; } = Controls.Appearance.ThreeD; // VB6 TextBox default is 3D
        public VB6Color BackColor { get; set; } = VB6Color.VbWindowBackground; // Default: &H80000005&
        public BorderStyle BorderStyle { get; set; } = Controls.BorderStyle.FixedSingle; // Default is FixedSingle (1)
        public CausesValidation CausesValidation { get; set; } = Controls.CausesValidation.Yes;
        public string DataField { get; set; } = ""; // For data binding
        public string DataFormat { get; set; } = ""; // For data binding
        public string DataMember { get; set; } = ""; // For data binding
        public string DataSource { get; set; } = ""; // For data binding
        // DragIcon, MouseIcon are FRX-based.
        // Their resolved byte[] data will be stored in the VB6TextBoxKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        // Font is a group property.
        public VB6Color ForeColor { get; set; } = VB6Color.VbWindowText; // Default: &H80000008&
        public int Height { get; set; } = 285;  // Typical default height in Twips for single-line
        public int HelpContextID { get; set; } = 0;
        public bool HideSelection { get; set; } = true; // If True, selected text doesn't show highlighting when focus is lost
        public int Left { get; set; } = 120;   // Typical default Left in Twips
        public string LinkItem { get; set; } = ""; // For DDE
        public LinkMode LinkMode { get; set; } = Controls.LinkMode.None; // For DDE
        public int LinkTimeout { get; set; } = 50; // For DDE
        public string LinkTopic { get; set; } = ""; // For DDE
        public bool Locked { get; set; } = false; // If True, text cannot be edited by user
        public int MaxLength { get; set; } = 0; // 0 means no limit (up to ~32K for single-line, ~2GB for multi-line)
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public MultiLine MultiLine { get; set; } = Controls.MultiLine.SingleLine; // Default is False (0) for single-line
        public OLEDragMode OleDragMode { get; set; } = Controls.OLEDragMode.Manual; // TextBox can be OLE drag source
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None;   // TextBox can be OLE drop target
        public char? PasswordChar { get; set; } = null; // Character to display instead of actual text
        public TextDirection RightToLeft { get; set; } = Controls.TextDirection.LeftToRight;
        public ScrollBars ScrollBars { get; set; } = Controls.ScrollBars.None; // None, Horizontal, Vertical, Both
        public int TabIndex { get; set; } = 0; // Placeholder
        public TabStop TabStop { get; set; } = Controls.TabStop.Included;
        public string Text { get; set; } = "Text1"; // Default text can vary (e.g., "Text1")
        public string ToolTipText { get; set; } = "";
        public int Top { get; set; } = 120;    // Typical default Top in Twips
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1215;   // Typical default width in Twips

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for a TextBox control.
        /// </summary>
        public TextBoxProperties() { }

        /// <summary>
        /// Initializes TextBoxProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public TextBoxProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Alignment = rawProps.GetEnum("Alignment", Alignment);
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            BorderStyle = rawProps.GetEnum("BorderStyle", BorderStyle);
            CausesValidation = rawProps.GetEnum("CausesValidation", CausesValidation);
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
            HideSelection = rawProps.GetBoolean("HideSelection", HideSelection);
            Left = rawProps.GetInt32("Left", Left);
            LinkItem = rawProps.GetString("LinkItem", LinkItem);
            LinkMode = rawProps.GetEnum("LinkMode", LinkMode);
            LinkTimeout = rawProps.GetInt32("LinkTimeout", LinkTimeout);
            LinkTopic = rawProps.GetString("LinkTopic", LinkTopic);
            Locked = rawProps.GetBoolean("Locked", Locked);
            MaxLength = rawProps.GetInt32("MaxLength", MaxLength);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            MultiLine = rawProps.GetEnum("MultiLine", MultiLine);
            OleDragMode = rawProps.GetEnum("OLEDragMode", OleDragMode);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            PasswordChar = rawProps.GetOptionalChar("PasswordChar"); // Returns null if property not found or empty
            RightToLeft = rawProps.GetEnum("RightToLeft", RightToLeft);
            ScrollBars = rawProps.GetEnum("ScrollBars", ScrollBars);
            TabIndex = rawProps.GetInt32("TabIndex", TabIndex);
            TabStop = rawProps.GetEnum("TabStop", TabStop);
            Text = rawProps.GetString("Text", Text); // Initial text content
            ToolTipText = rawProps.GetString("ToolTipText", ToolTipText);
            Top = rawProps.GetInt32("Top", Top);
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);
            Width = rawProps.GetInt32("Width", Width);
        }
    }
}