using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 Label control.
    /// The Label control is used to display text that the user cannot directly edit.
    /// </summary>
    public class LabelProperties
    {
        public Alignment Alignment { get; set; } = Controls.Alignment.LeftJustify;
        public Appearance Appearance { get; set; } = Controls.Appearance.Flat; // VB6 Label default is Flat (0)
        public AutoSize AutoSize { get; set; } = Controls.AutoSize.Fixed; // If True, label resizes to fit Caption
        public VB6Color BackColor { get; set; } = VB6Color.VbButtonFace; // Default: &H8000000F&
        public BackStyle BackStyle { get; set; } = Controls.BackStyle.Opaque; // Can be Transparent
        public BorderStyle BorderStyle { get; set; } = Controls.BorderStyle.None; // Default is None (0)
        public string Caption { get; set; } = "Label1"; // Default caption can vary (e.g., "Label1")
        public string DataField { get; set; } = ""; // For data binding
        public string DataFormat { get; set; } = ""; // For data binding
        public string DataMember { get; set; } = ""; // For data binding
        public string DataSource { get; set; } = ""; // For data binding
        // DragIcon, MouseIcon are FRX-based.
        // Their resolved byte[] data will be stored in the VB6LabelKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        // Font is a group property.
        public VB6Color ForeColor { get; set; } = VB6Color.VbWindowText; // Default: &H80000008&
        public int Height { get; set; } = 255;  // Typical default height in Twips
        public int Left { get; set; } = 120;   // Typical default Left in Twips
        public string LinkItem { get; set; } = ""; // For DDE
        public LinkMode LinkMode { get; set; } = Controls.LinkMode.None; // For DDE
        public int LinkTimeout { get; set; } = 50; // For DDE
        public string LinkTopic { get; set; } = ""; // For DDE
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None; // Labels can be OLE drop targets
        public TextDirection RightToLeft { get; set; } = Controls.TextDirection.LeftToRight;
        public int TabIndex { get; set; } = 0; // Labels are not focusable, so TabIndex is for order but not TabStop
        public string ToolTipText { get; set; } = "";
        public int Top { get; set; } = 120;    // Typical default Top in Twips
        public bool UseMnemonic { get; set; } = true; // If True, '&' in Caption creates an access key
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1275;   // Typical default width in Twips
        public WordWrap WordWrap { get; set; } = Controls.WordWrap.NonWrapping; // If True, text wraps to multiple lines

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for a Label control.
        /// </summary>
        public LabelProperties() { }

        /// <summary>
        /// Initializes LabelProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public LabelProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Alignment = rawProps.GetEnum("Alignment", Alignment);
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            AutoSize = rawProps.GetEnum("AutoSize", AutoSize);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            BackStyle = rawProps.GetEnum("BackStyle", BackStyle);
            BorderStyle = rawProps.GetEnum("BorderStyle", BorderStyle);
            Caption = rawProps.GetString("Caption", Caption);
            DataField = rawProps.GetString("DataField", DataField);
            DataFormat = rawProps.GetString("DataFormat", DataFormat);
            DataMember = rawProps.GetString("DataMember", DataMember);
            DataSource = rawProps.GetString("DataSource", DataSource);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            // Font handled by PropertyGroup
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor);
            Height = rawProps.GetInt32("Height", Height);
            Left = rawProps.GetInt32("Left", Left);
            LinkItem = rawProps.GetString("LinkItem", LinkItem);
            LinkMode = rawProps.GetEnum("LinkMode", LinkMode);
            LinkTimeout = rawProps.GetInt32("LinkTimeout", LinkTimeout);
            LinkTopic = rawProps.GetString("LinkTopic", LinkTopic);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            RightToLeft = rawProps.GetEnum("RightToLeft", RightToLeft);
            TabIndex = rawProps.GetInt32("TabIndex", TabIndex); // Labels have TabIndex but not TabStop
            ToolTipText = rawProps.GetString("ToolTipText", ToolTipText);
            Top = rawProps.GetInt32("Top", Top);
            UseMnemonic = rawProps.GetBoolean("UseMnemonic", UseMnemonic);
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);
            Width = rawProps.GetInt32("Width", Width);
            WordWrap = rawProps.GetEnum("WordWrap", WordWrap);
        }
    }
}