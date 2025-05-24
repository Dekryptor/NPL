// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For common enums like Appearance, DragMode, etc. they are in this namespace or CommonControlEnums.
using System.Collections.Generic;

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Defines the behavior and appearance style of a ComboBox control.
    /// </summary>
    public enum ComboBoxStyle
    {
        /// <summary>Dropdown Combo: Editable text box with a dropdown list. (Default)</summary>
        DropDownCombo = 0, // vbComboDropDown
        /// <summary>Simple Combo: Text box and list are both always visible. List is not dropdown.</summary>
        SimpleCombo = 1,   // vbComboSimple
        /// <summary>Dropdown List: Non-editable text box, user must choose from the list.</summary>
        DropDownList = 2 // vbComboDropDownList
    }

    /// <summary>
    /// Properties specific to a VB6 ComboBox control.
    /// </summary>
    public class ComboBoxProperties
    {
        public Appearance Appearance { get; set; }
        public Vb6Color BackColor { get; set; }
        public CausesValidation CausesValidation { get; set; }
        public string DataField { get; set; }
        // public string DataFormat { get; set; } // StdDataFormat object in VB6. String for simplicity.
        public string DataMember { get; set; }
        public string DataSource { get; set; }
        // public byte[] DragIcon { get; set; } // Placeholder for image data
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        public Vb6Color ForeColor { get; set; }
        public int Height { get; set; } // Typically in Twips. For SimpleCombo, includes list height.
        public int HelpContextId { get; set; }
        public bool IntegralHeight { get; set; } // True to resize to show full items only (applies to list part)
        public List<string> ListItems { get; set; } // Represents the List property (items in the list)
        public List<long> ItemData { get; set; } // Represents the ItemData property (long value per item)
        public int Left { get; set; }   // Typically in Twips
        public bool Locked { get; set; } // True if the text portion cannot be edited (Style=DropDownCombo or SimpleCombo)
        // public byte[] MouseIcon { get; set; } // Placeholder for image data
        public MousePointer MousePointer { get; set; }
        public OLEDragMode OLEDragMode { get; set; }
        public OLEDropMode OLEDropMode { get; set; }
        public TextDirection RightToLeft { get; set; }
        public bool Sorted { get; set; } // True if items are automatically sorted alphabetically
        public ComboBoxStyle Style { get; set; } // DropDownCombo, SimpleCombo, DropDownList
        public int TabIndex { get; set; }
        public TabStop TabStop { get; set; }
        public string Text { get; set; } // The text in the editable portion, or the selected item's text.
        public string ToolTipText { get; set; }
        public int Top { get; set; }     // Typically in Twips
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }   // Typically in Twips


        /// <summary>
        /// Initializes a new instance of the <see cref="ComboBoxProperties"/> class with default VB6 values.
        /// </summary>
        public ComboBoxProperties()
        {
            Appearance = Appearance.ThreeD;
            BackColor = Vb6Color.FromOleColor(0x80000005); // System Window Background
            CausesValidation = CausesValidation.Yes;
            DataField = string.Empty;
            // DataFormat = string.Empty;
            DataMember = string.Empty;
            DataSource = string.Empty;
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            ForeColor = Vb6Color.FromOleColor(0x80000008); // System Window Text
            Height = 315; // Example default (Rust: 30). VB default for DropDown style is for the text box part.
                          // For SimpleCombo, height includes the list.
            HelpContextId = 0;
            IntegralHeight = true; // Default
            ListItems = new List<string>();
            ItemData = new List<long>();
            Left = 0;     // Example default. Rust: 30.
            Locked = false; // Default
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            OLEDragMode = OLEDragMode.Manual;
            OLEDropMode = OLEDropMode.None; // Default
            RightToLeft = TextDirection.LeftToRight; // Default is False
            Sorted = false; // Default
            Style = ComboBoxStyle.DropDownCombo; // Default
            TabIndex = 0;
            TabStop = TabStop.Included; // Default is True
            Text = string.Empty; // Default (often set to control's name initially by IDE)
            ToolTipText = string.Empty;
            Top = 0;      // Example default. Rust: 30.
            Visible = Visibility.Visible;
            WhatsThisHelpId = 0;
            Width = 1800; // Example default. Rust: 100.
        }
    }
}