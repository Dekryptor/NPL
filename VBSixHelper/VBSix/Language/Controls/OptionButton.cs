using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 OptionButton control.
    /// OptionButtons (radio buttons) are typically used in groups to select one option from several.
    /// </summary>
    public class OptionButtonProperties
    {
        public JustifyAlignment Alignment { get; set; } = Controls.JustifyAlignment.LeftJustify;
        public Appearance Appearance { get; set; } = Controls.Appearance.ThreeD;
        public VB6Color BackColor { get; set; } = VB6Color.VbButtonFace; // Default: &H8000000F&
        public string Caption { get; set; } = "Option1"; // Default caption can vary
        public CausesValidation CausesValidation { get; set; } = Controls.CausesValidation.Yes;
        // DisabledPicture, DownPicture, DragIcon, MouseIcon, Picture are FRX-based.
        // Their resolved byte[] data will be stored in the VB6OptionButtonKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        // Font is a group property.
        public VB6Color ForeColor { get; set; } = VB6Color.VbButtonText; // Default: &H80000012&
        public int Height { get; set; } = 255;  // Typical default height in Twips
        public int HelpContextID { get; set; } = 0;
        public int Left { get; set; } = 120;   // Typical default Left in Twips
        public VB6Color MaskColor { get; set; } = VB6Color.FromRGB(192, 192, 192); // Default: &H00C0C0C0& (Silver)
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None; // VB6 default
        public TextDirection RightToLeft { get; set; } = Controls.TextDirection.LeftToRight;
        public Style Style { get; set; } = Controls.Style.Standard; // Standard or Graphical
        public int TabIndex { get; set; } = 0; // Placeholder
        public TabStop TabStop { get; set; } = Controls.TabStop.Included;
        public string ToolTipText { get; set; } = "";
        public int Top { get; set; } = 120;    // Typical default Top in Twips
        public UseMaskColor UseMaskColor { get; set; } = Controls.UseMaskColor.DoNotUseMaskColor;
        public OptionButtonValue Value { get; set; } = OptionButtonValue.UnSelected; // True if selected, False if not
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1275;   // Typical default width in Twips

        // DataField, DataSource, DataMember, DataFormat are less common for OptionButton but possible.
        public string DataField { get; set; } = "";
        public string DataSource { get; set; } = "";
        public string DataMember { get; set; } = "";
        public string DataFormat { get; set; } = "";


        /// <summary>
        /// Default constructor initializing with common VB6 defaults for an OptionButton control.
        /// </summary>
        public OptionButtonProperties() { }

        /// <summary>
        /// Initializes OptionButtonProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public OptionButtonProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Alignment = rawProps.GetEnum("Alignment", Alignment);
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            Caption = rawProps.GetString("Caption", Caption);
            CausesValidation = rawProps.GetEnum("CausesValidation", CausesValidation);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            // Font handled by PropertyGroup
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor);
            Height = rawProps.GetInt32("Height", Height);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            Left = rawProps.GetInt32("Left", Left);
            MaskColor = rawProps.GetVB6Color("MaskColor", MaskColor);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            RightToLeft = rawProps.GetEnum("RightToLeft", RightToLeft);
            Style = rawProps.GetEnum("Style", Style);
            TabIndex = rawProps.GetInt32("TabIndex", TabIndex);
            TabStop = rawProps.GetEnum("TabStop", TabStop);
            ToolTipText = rawProps.GetString("ToolTipText", ToolTipText);
            Top = rawProps.GetInt32("Top", Top);
            UseMaskColor = rawProps.GetEnum("UseMaskColor", UseMaskColor);
            Value = rawProps.GetEnum("Value", Value); // Maps to bool True/False (-1/0) in VB6
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);
            Width = rawProps.GetInt32("Width", Width);

            // Data binding properties (less common but possible)
            DataField = rawProps.GetString("DataField", DataField);
            DataSource = rawProps.GetString("DataSource", DataSource);
            DataMember = rawProps.GetString("DataMember", DataMember);
            DataFormat = rawProps.GetString("DataFormat", DataFormat);
        }
    }
}