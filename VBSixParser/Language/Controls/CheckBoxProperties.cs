// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For common enums like Appearance, DragMode, etc. they are in this namespace or CommonControlEnums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents the state of a CheckBox or OptionButton control.
    /// </summary>
    public enum CheckBoxValue // VB6 Value property for CheckBox/OptionButton
    {
        /// <summary>The control is unchecked/deselected. (Default for CheckBox)</summary>
        Unchecked = 0, // vbUnchecked
        /// <summary>The control is checked/selected.</summary>
        Checked = 1,   // vbChecked
        /// <summary>The control is grayed (disabled but visually distinct). (Default for OptionButton if part of group and none selected initially)</summary>
        Grayed = 2     // vbGrayed (Applies to CheckBox, OptionButton doesn't usually show grayed by user action)
    }

    /// <summary>
    /// Properties specific to a VB6 CheckBox control.
    /// </summary>
    public class CheckBoxProperties
    {
        public JustifyAlignment Alignment { get; set; } // LeftJustify or RightJustify
        public Appearance Appearance { get; set; }
        public Vb6Color BackColor { get; set; }
        public string Caption { get; set; }
        public CausesValidation CausesValidation { get; set; }
        public string DataField { get; set; }
        // public string DataFormat { get; set; } // StdDataFormat object in VB6. String for simplicity.
        public string DataMember { get; set; }
        public string DataSource { get; set; }
        // public byte[] DisabledPicture { get; set; } // Placeholder for image data
        // public byte[] DownPicture { get; set; }     // Placeholder for image data
        // public byte[] DragIcon { get; set; }        // Placeholder for image data
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        public Vb6Color ForeColor { get; set; }
        public int Height { get; set; } // Typically in Twips
        public int HelpContextId { get; set; }
        public int Left { get; set; }   // Typically in Twips
        public Vb6Color MaskColor { get; set; } // Used with Picture property if Style is Graphical
        // public byte[] MouseIcon { get; set; }       // Placeholder for image data
        public MousePointer MousePointer { get; set; }
        public OLEDropMode OLEDropMode { get; set; }
        // public byte[] Picture { get; set; }         // Placeholder for image data
        public TextDirection RightToLeft { get; set; }
        public Style Style { get; set; } // Standard or Graphical
        public int TabIndex { get; set; }
        public TabStop TabStop { get; set; }
        public string ToolTipText { get; set; }
        public int Top { get; set; }     // Typically in Twips
        public UseMaskColor UseMaskColor { get; set; }
        public CheckBoxValue Value { get; set; } // Unchecked, Checked, or Grayed
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }   // Typically in Twips

        /// <summary>
        /// Initializes a new instance of the <see cref="CheckBoxProperties"/> class with default VB6 values.
        /// </summary>
        public CheckBoxProperties()
        {
            Alignment = JustifyAlignment.LeftJustify; // Default for CheckBox
            Appearance = Appearance.ThreeD; // Checkboxes are typically 3D unless Style is Graphical & Flat
            BackColor = Vb6Color.FromOleColor(0x8000000F); // System Button Face
            Caption = string.Empty; // Default is "CheckN"
            CausesValidation = CausesValidation.Yes;
            DataField = string.Empty;
            // DataFormat = string.Empty;
            DataMember = string.Empty;
            DataSource = string.Empty;
            // DisabledPicture = null;
            // DownPicture = null;
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            ForeColor = Vb6Color.FromOleColor(0x80000012); // System Button Text
            Height = 255; // Example default. Rust uses 30. VB default can be around 195-255 twips.
            HelpContextId = 0;
            Left = 0;     // Example default. Rust uses 30.
            MaskColor = Vb6Color.FromOleColor(0x00C0C0C0); // Silver
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            OLEDropMode = OLEDropMode.None;
            // Picture = null;
            RightToLeft = TextDirection.LeftToRight; // Default is False
            Style = Style.Standard; // Standard (checkbox) or Graphical (button-like)
            TabIndex = 0;
            TabStop = TabStop.Included; // Default is True
            ToolTipText = string.Empty;
            Top = 0;      // Example default. Rust uses 30.
            UseMaskColor = UseMaskColor.DoNotUseMaskColor; // Default is False
            Value = CheckBoxValue.Unchecked; // Default
            Visible = Visibility.Visible;
            WhatsThisHelpId = 0;
            Width = 1215; // Example default. Rust uses 100.
        }
    }
}