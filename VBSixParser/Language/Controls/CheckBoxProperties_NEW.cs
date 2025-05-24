// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class CheckBoxProperties : ControlSpecificPropertiesBase
    {
        public AlignmentConstants Alignment { get; set; } = AlignmentConstants.LeftJustify; // Left/Right of the checkbox itself
        public Appearance Appearance { get; set; } = Appearance.ThreeD;
        public Vb6Color BackColor { get; set; } = Vb6Color.DefaultSystemColors.ButtonFace;
        public string Caption { get; set; } = "CheckBox"; // Default "Check1", "Check2"
        public DataFormat DataFormat { get; set; } // Placeholder
        public string DataField { get; set; } = string.Empty;
        public object DataSource { get; set; } // Placeholder
        public bool DownPicture { get; set; } // Actually PictureDisabled in some contexts, or specific to Style=Graphical
        public DragMode DragMode { get; set; } = DragMode.Manual;
        public object DragIcon { get; set; } // Placeholder
        public Activation Enabled { get; set; } = Activation.Enabled;
        // Font is inherited
        public Vb6Color ForeColor { get; set; } = Vb6Color.DefaultSystemColors.WindowText;
        public int Height { get; set; } = 255; // Default in Twips (approx)
        public int Index { get; set; } = -1; // Contextual
        public int Left { get; set; } = 0;
        public MaskColorMode MaskColor { get; set; } // For graphical style
        public MousePointer MousePointer { get; set; } = MousePointer.Default;
        public object MouseIcon { get; set; } // Placeholder
        public OLEDropMode OLEDropMode { get; set; } = OLEDropMode.None;
        public object Picture { get; set; } // For graphical style
        public object PictureDisabled { get; set; } // For graphical style (when Enabled=False)
        public RightToLeftConstants RightToLeft { get; set; } = RightToLeftConstants.False;
        public CheckBoxStyleConstants Style { get; set; } = CheckBoxStyleConstants.Standard; // Standard or Graphical
        public int TabIndex { get; set; } = 0;
        public bool TabStop { get; set; } = true;
        public string Tag { get; set; } = string.Empty;
        public string ToolTipText { get; set; } = string.Empty;
        public int Top { get; set; } = 0;
        public UseMaskColorConstants UseMaskColor { get; set; } = UseMaskColorConstants.False; // For graphical style
        public CheckStateConstants Value { get; set; } = CheckStateConstants.Unchecked; // The state of the checkbox
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1215; // Default in Twips


        public CheckBoxProperties()
        {
            Font = new Font();
            Caption = "Check1"; // Default caption often "Check1"
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            Alignment = PropertyParsingHelpers.GetEnum(textualProps, "Alignment", this.Alignment);
            Appearance = PropertyParsingHelpers.GetEnum(textualProps, "Appearance", this.Appearance);
            BackColor = PropertyParsingHelpers.GetColor(textualProps, "BackColor", this.BackColor);
            Caption = PropertyParsingHelpers.GetString(textualProps, "Caption", this.Caption);
            DataField = PropertyParsingHelpers.GetString(textualProps, "DataField", this.DataField);
            DragMode = PropertyParsingHelpers.GetEnum(textualProps, "DragMode", this.DragMode);
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", this.Enabled);

            PopulateFontProperty(textualProps, propertyGroups);

            ForeColor = PropertyParsingHelpers.GetColor(textualProps, "ForeColor", this.ForeColor);
            Height = PropertyParsingHelpers.GetInt32(textualProps, "Height", this.Height);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", this.Left);
            // MaskColor is a Vb6Color, not an enum directly from text. If Style=Graphical.
            MousePointer = PropertyParsingHelpers.GetEnum(textualProps, "MousePointer", this.MousePointer);
            OLEDropMode = PropertyParsingHelpers.GetEnum(textualProps, "OLEDropMode", this.OLEDropMode);
            RightToLeft = PropertyParsingHelpers.GetEnum(textualProps, "RightToLeft", this.RightToLeft);
            Style = PropertyParsingHelpers.GetEnum(textualProps, "Style", this.Style);
            TabIndex = PropertyParsingHelpers.GetInt32(textualProps, "TabIndex", this.TabIndex);
            TabStop = PropertyParsingHelpers.GetBool(textualProps, "TabStop", this.TabStop);
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            ToolTipText = PropertyParsingHelpers.GetString(textualProps, "ToolTipText", this.ToolTipText);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", this.Top);
            UseMaskColor = PropertyParsingHelpers.GetEnum(textualProps, "UseMaskColor", this.UseMaskColor);
            Value = PropertyParsingHelpers.GetEnum(textualProps, "Value", this.Value);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", this.Visible);
            WhatsThisHelpID = PropertyParsingHelpers.GetInt32(textualProps, "WhatsThisHelpID", this.WhatsThisHelpID);
            Width = PropertyParsingHelpers.GetInt32(textualProps, "Width", this.Width);

            // Binary properties (for Style = Graphical)
            if (binaryProps.TryGetValue("DownPicture", out byte[] dpData)) DownPicture = true; // Or store dpData
            if (binaryProps.TryGetValue("Picture", out byte[] picData)) Picture = picData;
            if (binaryProps.TryGetValue("DisabledPicture", out byte[] disPicData)) PictureDisabled = disPicData; // VB uses "DisabledPicture"
            if (binaryProps.TryGetValue("DragIcon", out byte[] dragIconData)) DragIcon = dragIconData;
            if (binaryProps.TryGetValue("MouseIcon", out byte[] mouseIconData)) MouseIcon = mouseIconData;

            // DataFormat and DataSource
            if (textualProps.TryGetValue("DataSource", out string dsName))
            {
                DataSource = dsName;
            }
            
            // MaskColor for graphical style (OLE_COLOR)
            if (textualProps.TryGetValue("MaskColor", out string maskColorString))
            {
                if (Vb6Color.TryParseHex(maskColorString, out Vb6Color mc))
                {
                    // This would be assigned to a Vb6Color MaskColor property if defined as such.
                    // Currently MaskColor is an enum MaskColorMode.
                    // This indicates the `MaskColor` property on the control is an OLE_COLOR (Vb6Color).
                    // The `MaskColorMode` enum is likely for different context or not directly the property name.
                }
            }
        }
    }

    // Supporting Enums (if not already defined elsewhere)
    public enum CheckBoxStyleConstants
    {
        Standard = 0,
        Graphical = 1
    }

    public enum CheckStateConstants // Also known as Value property for CheckBox/OptionButton
    {
        Unchecked = 0,
        Checked = 1,
        Grayed = 2 // (Dimmed / Indeterminate)
    }
}