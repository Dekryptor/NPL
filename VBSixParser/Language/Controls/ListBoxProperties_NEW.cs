// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class ListBoxProperties : ControlSpecificPropertiesBase
    {
        public Appearance Appearance { get; set; } = Appearance.ThreeD;
        public Vb6Color BackColor { get; set; } = Vb6Color.DefaultSystemColors.WindowBackground;
        public Vb6Color ForeColor { get; set; } = Vb6Color.DefaultSystemColors.WindowText;
        public int Columns { get; set; } = 0; // 0 for single column, >0 for multi-column (snaking)
        public DataFormat DataFormat { get; set; } // Placeholder
        public string DataField { get; set; } = string.Empty;
        public object DataSource { get; set; } // Placeholder
        public DragMode DragMode { get; set; } = DragMode.Manual;
        public object DragIcon { get; set; } // Placeholder
        public Activation Enabled { get; set; } = Activation.Enabled;
        // Font is inherited
        public int Height { get; set; } = 1200; // Default in Twips (approx)
        public int Index { get; set; } = -1; // Contextual
        public bool IntegralHeight { get; set; } = true;
        public int Left { get; set; } = 0;
        public LockedConstants Locked { get; set; } = LockedConstants.False; // If items can be changed by user
        public MousePointer MousePointer { get; set; } = MousePointer.Default;
        public object MouseIcon { get; set; } // Placeholder
        public ListBoxMultiSelectConstants MultiSelect { get; set; } = ListBoxMultiSelectConstants.None;
        public OLEDropMode OLEDropMode { get; set; } = OLEDropMode.None;
        public RightToLeftConstants RightToLeft { get; set; } = RightToLeftConstants.False;
        public bool Sorted { get; set; } = false;
        public ListBoxStyleConstants Style { get; set; } = ListBoxStyleConstants.Standard; // Standard or Checkbox
        public int TabIndex { get; set; } = 0;
        public bool TabStop { get; set; } = true;
        public string Tag { get; set; } = string.Empty;
        public string ToolTipText { get; set; } = string.Empty;
        public int Top { get; set; } = 0;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1800; // Default in Twips (approx)

        // Properties for list items and associated data
        // These might be populated from textual properties or binary (FRX) data.
        public List<string> ListItems { get; set; } = new List<string>();
        public List<long> ItemData { get; set; } = new List<long>();
        
        // Raw string value of the 'List' property if it's textual and not an FRX reference.
        // This might need further parsing by the consumer if it's a complex format.
        public string RawListPropertyText { get; set; }

        public ListBoxProperties()
        {
            Font = new Font();
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            Appearance = PropertyParsingHelpers.GetEnum(textualProps, "Appearance", this.Appearance);
            BackColor = PropertyParsingHelpers.GetColor(textualProps, "BackColor", this.BackColor);
            ForeColor = PropertyParsingHelpers.GetColor(textualProps, "ForeColor", this.ForeColor);
            Columns = PropertyParsingHelpers.GetInt32(textualProps, "Columns", this.Columns);
            DataField = PropertyParsingHelpers.GetString(textualProps, "DataField", this.DataField);
            DragMode = PropertyParsingHelpers.GetEnum(textualProps, "DragMode", this.DragMode);
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", this.Enabled);

            PopulateFontProperty(textualProps, propertyGroups);

            Height = PropertyParsingHelpers.GetInt32(textualProps, "Height", this.Height);
            IntegralHeight = PropertyParsingHelpers.GetBool(textualProps, "IntegralHeight", this.IntegralHeight);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", this.Left);
            Locked = PropertyParsingHelpers.GetEnum(textualProps, "Locked", this.Locked);
            MousePointer = PropertyParsingHelpers.GetEnum(textualProps, "MousePointer", this.MousePointer);
            MultiSelect = PropertyParsingHelpers.GetEnum(textualProps, "MultiSelect", this.MultiSelect);
            OLEDropMode = PropertyParsingHelpers.GetEnum(textualProps, "OLEDropMode", this.OLEDropMode);
            RightToLeft = PropertyParsingHelpers.GetEnum(textualProps, "RightToLeft", this.RightToLeft);
            Sorted = PropertyParsingHelpers.GetBool(textualProps, "Sorted", this.Sorted);
            Style = PropertyParsingHelpers.GetEnum(textualProps, "Style", this.Style);
            TabIndex = PropertyParsingHelpers.GetInt32(textualProps, "TabIndex", this.TabIndex);
            TabStop = PropertyParsingHelpers.GetBool(textualProps, "TabStop", this.TabStop);
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            ToolTipText = PropertyParsingHelpers.GetString(textualProps, "ToolTipText", this.ToolTipText);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", this.Top);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", this.Visible);
            WhatsThisHelpID = PropertyParsingHelpers.GetInt32(textualProps, "WhatsThisHelpID", this.WhatsThisHelpID);
            Width = PropertyParsingHelpers.GetInt32(textualProps, "Width", this.Width);
            
            // Handle List and ItemData properties
            // These can be complex and might be in textualProps (as a string needing parsing)
            // or in binaryProps (if they are FRX references).

            if (textualProps.TryGetValue("List", out string listText))
            {
                // This is a simple textual representation. It might be semicolon-delimited
                // or a "property page string". For now, store the raw text.
                // A more sophisticated parser might split this into ListItems.
                // Example: "Item1;Item2;Item3"
                RawListPropertyText = listText;
                // Simple parsing for semicolon delimited, common for basic lists:
                if (!string.IsNullOrEmpty(listText) && listText.Contains(";"))
                {
                    ListItems.AddRange(listText.Split(';').Select(s => s.Trim('"'))); 
                } else if (!string.IsNullOrEmpty(listText)) {
                    // If not semicolon delimited and not an FRX ref, it might be a single item or malformed.
                    // Or it could be a property page string "Prop=""VB6...""" which needs specific parsing.
                    // For now, if not obviously delimited, add as single item or keep in RawListPropertyText.
                    // ListItems.Add(listText.Trim('"')); // Cautious: could be FRX ref string not caught by binary.
                }
            }
            else if (binaryProps.TryGetValue("List", out byte[] listBinaryData))
            {
                // Data is in FRX. This byte array needs to be parsed according
                // to FRX format for list items. This is complex.
                // For now, we acknowledge it's binary. A full FRX parser is needed.
                // Placeholder: could try to decode as simple strings if encoding known.
                // ListItems.Add($"[Binary List Data: {listBinaryData.Length} bytes]");
            }

            if (binaryProps.TryGetValue("ItemData", out byte[] itemDataBinary))
            {
                // ItemData is almost always binary from FRX.
                // This byte array contains an array of Longs (4 bytes each).
                // A full FRX parser is needed.
                // Placeholder:
                // for (int i = 0; i + 3 < itemDataBinary.Length; i += 4)
                // {
                //    ItemData.Add(BitConverter.ToInt32(itemDataBinary, i));
                // }
            }
            // Note: If ItemData is textual (rare), it would need specific parsing logic.

            if (textualProps.TryGetValue("DataSource", out string dsName))
            {
                DataSource = dsName;
            }
        }
    }

    // Supporting Enums
    public enum ListBoxMultiSelectConstants
    {
        None = 0,
        Simple = 1, // Shift+Click or Space toggles, Click selects
        Extended = 2 // Ctrl+Click, Shift+Click for ranges
    }

    public enum ListBoxStyleConstants
    {
        Standard = 0,
        CheckBox = 1 // ListBox with checkboxes next to items
    }
}