// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For common enums like Appearance, DragMode, etc. they are in this namespace or CommonControlEnums.
using System.Collections.Generic;

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Defines the appearance style of a ListBox control.
    /// </summary>
    public enum ListBoxStyle
    {
        /// <summary>Standard ListBox. (Default)</summary>
        Standard = 0, // vbListBoxStandard
        /// <summary>ListBox with a checkbox next to each item. Requires MultiSelect to be None.</summary>
        Checkbox = 1  // vbListBoxCheckbox
    }

    /// <summary>
    /// Properties specific to a VB6 ListBox control.
    /// </summary>
    public class ListBoxProperties
    {
        public Appearance Appearance { get; set; }
        public Vb6Color BackColor { get; set; }
        public CausesValidation CausesValidation { get; set; }
        public int Columns { get; set; } // 0 for single column, >0 for multi-column (snaking)
        public string DataField { get; set; }
        // public string DataFormat { get; set; } // StdDataFormat object in VB6. String for simplicity.
        public string DataMember { get; set; }
        public string DataSource { get; set; }
        // public byte[] DragIcon { get; set; } // Placeholder for image data
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        public Vb6Color ForeColor { get; set; }
        public int Height { get; set; } // Typically in Twips
        public int HelpContextId { get; set; }
        public bool IntegralHeight { get; set; } // True to resize to show full items only
        public List<string> ListItems { get; set; } // Represents the List property (items in the list)
        public List<long> ItemData { get; set; } // Represents the ItemData property (long value per item)
        public int Left { get; set; }   // Typically in Twips
        // public byte[] MouseIcon { get; set; } // Placeholder for image data
        public MousePointer MousePointer { get; set; }
        public MultiSelect MultiSelect { get; set; } // None, Simple, Extended
        public OLEDragMode OLEDragMode { get; set; }
        public OLEDropMode OLEDropMode { get; set; }
        public TextDirection RightToLeft { get; set; }
        public bool Sorted { get; set; } // True if items are automatically sorted alphabetically
        public ListBoxStyle Style { get; set; } // Standard or Checkbox
        public int TabIndex { get; set; }
        public TabStop TabStop { get; set; }
        public string ToolTipText { get; set; }
        public int Top { get; set; }     // Typically in Twips
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }   // Typically in Twips

        // Runtime/Design-time selected state (not always in .frm as a direct property like this)
        // For Style=Checkbox, Selected(index) is boolean.
        // For Style=Standard, Selected(index) is boolean, and ListIndex is the single selected index.
        // This might be too complex for a direct properties class if it reflects runtime state.
        // However, initial selected items can be set at design time.
        // For simplicity, we'll omit direct representation of selected items here,
        // as it's often managed via ListIndex or Selected(index) at runtime or through initial List values.

        /// <summary>
        /// Initializes a new instance of the <see cref="ListBoxProperties"/> class with default VB6 values.
        /// </summary>
        public ListBoxProperties()
        {
            Appearance = Appearance.ThreeD;
            BackColor = Vb6Color.FromOleColor(0x80000005); // System Window Background
            CausesValidation = CausesValidation.Yes;
            Columns = 0; // Default
            DataField = string.Empty;
            // DataFormat = string.Empty;
            DataMember = string.Empty;
            DataSource = string.Empty;
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            ForeColor = Vb6Color.FromOleColor(0x80000008); // System Window Text
            Height = 690; // Example default. Rust: 30. VB default is ~4-5 lines.
            HelpContextId = 0;
            IntegralHeight = true; // Default
            ListItems = new List<string>();
            ItemData = new List<long>();
            Left = 0;     // Example default. Rust: 30.
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            MultiSelect = MultiSelect.None; // Default
            OLEDragMode = OLEDragMode.Manual;
            OLEDropMode = OLEDropMode.None; // Default
            RightToLeft = TextDirection.LeftToRight; // Default is False
            Sorted = false; // Default
            Style = ListBoxStyle.Standard; // Default
            TabIndex = 0;
            TabStop = TabStop.Included; // Default is True
            ToolTipText = string.Empty;
            Top = 0;      // Example default. Rust: 30.
            Visible = Visibility.Visible;
            WhatsThisHelpId = 0;
            Width = 1800; // Example default. Rust: 100.
        }
    }
}