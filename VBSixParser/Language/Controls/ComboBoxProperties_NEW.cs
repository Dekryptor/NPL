// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class ComboBoxProperties : ControlSpecificPropertiesBase
    {
        public Appearance Appearance { get; set; } = Appearance.ThreeD;
        public Vb6Color BackColor { get; set; } = Vb6Color.DefaultSystemColors.WindowBackground;
        public Vb6Color ForeColor { get; set; } = Vb6Color.DefaultSystemColors.WindowText;
        public DataFormat DataFormat { get; set; } // Placeholder
        public string DataField { get; set; } = string.Empty;
        public object DataSource { get; set; } // Placeholder
        public DragMode DragMode { get; set; } = DragMode.Manual;
        public object DragIcon { get; set; } // Placeholder
        public Activation Enabled { get; set; } = Activation.Enabled;
        // Font is inherited
        public int Height { get; set; } = 315; // Default in Twips (approx)
        public int Index { get; set; } = -1; // Contextual
        public bool IntegralHeight { get; set; } = true; // Usually True for ComboBox
        public int Left { get; set; } = 0;
        public LockedConstants Locked { get; set; } = LockedConstants.False; // If text portion can be edited
        public MousePointer MousePointer { get; set; } = MousePointer.Default;
        public object MouseIcon { get; set; } // Placeholder
        public OLEDropMode OLEDropMode { get; set; } = OLEDropMode.None;
        public RightToLeftConstants RightToLeft { get; set; } = RightToLeftConstants.False;
        public bool Sorted { get; set; } = false;
        public ComboBoxStyleConstants Style { get; set; } = ComboBoxStyleConstants.DropDownCombo;
        public int TabIndex { get; set; } = 0;
        public bool TabStop { get; set; } = true;
        public string Tag { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty; // The current text in the ComboBox
        public string ToolTipText { get; set; } = string.Empty;
        public int Top { get; set; } = 0;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1800; // Default in Twips (approx)

        // Properties for list items and associated data
        public List<string> ListItems { get; set; } = new List<string>();
        public List<long> ItemData { get; set; } = new List<long>();
        public string RawListPropertyText { get; set; }


        public ComboBoxProperties()
        {
            Font = new Font();
            Text = string.Empty; // Default Text is usually the control name like "Combo1" or empty
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            Appearance = PropertyParsingHelpers.GetEnum(textualProps, "Appearance", this.Appearance);
            BackColor = PropertyParsingHelpers.GetColor(textualProps, "BackColor", this.BackColor);
            ForeColor = PropertyParsingHelpers.GetColor(textualProps, "ForeColor", this.ForeColor);
            DataField = PropertyParsingHelpers.GetString(textualProps, "DataField", this.DataField);
            DragMode = PropertyParsingHelpers.GetEnum(textualProps, "DragMode", this.DragMode);
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", this.Enabled);

            PopulateFontProperty(textualProps, propertyGroups);

            Height = PropertyParsingHelpers.GetInt32(textualProps, "Height", this.Height);
            IntegralHeight = PropertyParsingHelpers.GetBool(textualProps, "IntegralHeight", this.IntegralHeight);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", this.Left);
            Locked = PropertyParsingHelpers.GetEnum(textualProps, "Locked", this.Locked);
            MousePointer = PropertyParsingHelpers.GetEnum(textualProps, "MousePointer", this.MousePointer);
            OLEDropMode = PropertyParsingHelpers.GetEnum(textualProps, "OLEDropMode", this.OLEDropMode);
            RightToLeft = PropertyParsingHelpers.GetEnum(textualProps, "RightToLeft", this.RightToLeft);
            Sorted = PropertyParsingHelpers.GetBool(textualProps, "Sorted", this.Sorted);
            Style = PropertyParsingHelpers.GetEnum(textualProps, "Style", this.Style);
            TabIndex = PropertyParsingHelpers.GetInt32(textualProps, "TabIndex", this.TabIndex);
            TabStop = PropertyParsingHelpers.GetBool(textualProps, "TabStop", this.TabStop);
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            Text = PropertyParsingHelpers.GetString(textualProps, "Text", this.Text);
            ToolTipText = PropertyParsingHelpers.GetString(textualProps, "ToolTipText", this.ToolTipText);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", this.Top);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", this.Visible);
            WhatsThisHelpID = PropertyParsingHelpers.GetInt32(textualProps, "WhatsThisHelpID", this.WhatsThisHelpID);
            Width = PropertyParsingHelpers.GetInt32(textualProps, "Width", this.Width);

            // Handle List and ItemData properties (similar to ListBox)
            if (textualProps.TryGetValue("List", out string listText))
            {
                RawListPropertyText = listText;
                if (!string.IsNullOrEmpty(listText) && listText.Contains(";"))
                {
                    ListItems.AddRange(listText.Split(';').Select(s => s.Trim('"')));
                }
                else if (!string.IsNullOrEmpty(listText))
                {
                    // Could be single item or property page string. Cautious add.
                    // ListItems.Add(listText.Trim('"'));
                }
            }
            else if (binaryProps.TryGetValue("List", out byte[] listBinaryData))
            {
                // FRX data. Needs proper FRX parsing.
                // Placeholder: ListItems.Add($"[Binary List Data: {listBinaryData.Length} bytes]");
            }

            if (binaryProps.TryGetValue("ItemData", out byte[] itemDataBinary))
            {
                // FRX data for ItemData (array of Longs). Needs FRX parsing.
                // Placeholder:
                // for (int i = 0; i + 3 < itemDataBinary.Length; i += 4)
                // {
                //    ItemData.Add(BitConverter.ToInt32(itemDataBinary, i));
                // }
            }
            
            if (textualProps.TryGetValue("DataSource", out string dsName))
            {
                DataSource = dsName;
            }
        }
    }

    // Supporting Enums
    public enum ComboBoxStyleConstants
    {
        DropDownCombo = 0,  // Editable text box with dropdown list
        SimpleCombo = 1,    // Editable text box with list always visible below
        DropDownList = 2    // Non-editable, user must select from dropdown list
    }
}