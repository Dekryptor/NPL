// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class LabelProperties : ControlSpecificPropertiesBase
    {
        public AlignmentConstants Alignment { get; set; } = AlignmentConstants.LeftJustify;
        public Appearance Appearance { get; set; } = Appearance.ThreeD; // Labels are often Flat in newer styles though
        public AutoSizeConstants AutoSize { get; set; } = AutoSizeConstants.False;
        public Vb6Color BackColor { get; set; } = Vb6Color.DefaultSystemColors.ButtonFace; // Or control panel default
        public BackStyleConstants BackStyle { get; set; } = BackStyleConstants.Transparent;
        public BorderStyleConstants BorderStyle { get; set; } = BorderStyleConstants.None;
        public string Caption { get; set; } = "Label"; // Default "Label1", "Label2"
        public DataFormat DataFormat { get; set; } // Complex object, placeholder
        public string DataField { get; set; } = string.Empty;
        public object DataSource { get; set; } // Placeholder, typically a control name string
        public DragMode DragMode { get; set; } = DragMode.Manual;
        public object DragIcon { get; set; } // Placeholder for image data
        public Activation Enabled { get; set; } = Activation.Enabled;
        // Font is inherited
        public Vb6Color ForeColor { get; set; } = Vb6Color.DefaultSystemColors.WindowText;
        public int Height { get; set; } = 255; // Default in Twips
        public int Index { get; set; } = -1; // Contextual
        public int Left { get; set; } = 0;
        public LinkModeConstants LinkMode { get; set; } = LinkModeConstants.None; // For DDE
        public string LinkItem { get; set; } = string.Empty;
        public string LinkTopic { get; set; } = string.Empty;
        public int LinkTimeout { get; set; } = 50; // Tenths of a second
        public MousePointer MousePointer { get; set; } = MousePointer.Default;
        public object MouseIcon { get; set; } // Placeholder
        public OLEDropMode OLEDropMode { get; set; } = OLEDropMode.None;
        public RightToLeftConstants RightToLeft { get; set; } = RightToLeftConstants.False;
        public string Tag { get; set; } = string.Empty;
        public string ToolTipText { get; set; } = string.Empty;
        public int Top { get; set; } = 0;
        public UseMnemonicConstants UseMnemonic { get; set; } = UseMnemonicConstants.True;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1215; // Default in Twips
        public WordWrapConstants WordWrap { get; set; } = WordWrapConstants.False;

        public LabelProperties()
        {
            Font = new Font(); // Default Font
            Caption = "Label"; // Default caption often includes a number like "Label1"
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            Alignment = PropertyParsingHelpers.GetEnum(textualProps, "Alignment", this.Alignment);
            Appearance = PropertyParsingHelpers.GetEnum(textualProps, "Appearance", this.Appearance);
            AutoSize = PropertyParsingHelpers.GetEnum(textualProps, "AutoSize", this.AutoSize);
            BackColor = PropertyParsingHelpers.GetColor(textualProps, "BackColor", this.BackColor);
            BackStyle = PropertyParsingHelpers.GetEnum(textualProps, "BackStyle", this.BackStyle);
            BorderStyle = PropertyParsingHelpers.GetEnum(textualProps, "BorderStyle", this.BorderStyle);
            Caption = PropertyParsingHelpers.GetString(textualProps, "Caption", this.Caption);
            DataField = PropertyParsingHelpers.GetString(textualProps, "DataField", this.DataField);
            // DataSource would need to resolve to an object or control name string.
            DragMode = PropertyParsingHelpers.GetEnum(textualProps, "DragMode", this.DragMode);
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", this.Enabled);
            
            PopulateFontProperty(textualProps, propertyGroups);

            ForeColor = PropertyParsingHelpers.GetColor(textualProps, "ForeColor", this.ForeColor);
            Height = PropertyParsingHelpers.GetInt32(textualProps, "Height", this.Height);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", this.Left);
            LinkItem = PropertyParsingHelpers.GetString(textualProps, "LinkItem", this.LinkItem);
            LinkTopic = PropertyParsingHelpers.GetString(textualProps, "LinkTopic", this.LinkTopic);
            LinkTimeout = PropertyParsingHelpers.GetInt32(textualProps, "LinkTimeout", this.LinkTimeout);
            MousePointer = PropertyParsingHelpers.GetEnum(textualProps, "MousePointer", this.MousePointer);
            OLEDropMode = PropertyParsingHelpers.GetEnum(textualProps, "OLEDropMode", this.OLEDropMode);
            RightToLeft = PropertyParsingHelpers.GetEnum(textualProps, "RightToLeft", this.RightToLeft);
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            ToolTipText = PropertyParsingHelpers.GetString(textualProps, "ToolTipText", this.ToolTipText);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", this.Top);
            UseMnemonic = PropertyParsingHelpers.GetEnum(textualProps, "UseMnemonic", this.UseMnemonic);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", this.Visible); // Visible is an enum (0=False, -1=True)
            WhatsThisHelpID = PropertyParsingHelpers.GetInt32(textualProps, "WhatsThisHelpID", this.WhatsThisHelpID);
            Width = PropertyParsingHelpers.GetInt32(textualProps, "Width", this.Width);
            WordWrap = PropertyParsingHelpers.GetEnum(textualProps, "WordWrap", this.WordWrap);

            // Binary properties
            if (binaryProps.TryGetValue("DragIcon", out byte[] dragIconData)) DragIcon = dragIconData;
            if (binaryProps.TryGetValue("MouseIcon", out byte[] mouseIconData)) MouseIcon = mouseIconData;
            
            // DataFormat and DataSource are more complex.
            // DataFormat is often a nested property group or requires specific parsing.
            // DataSource is typically the name of another control.
            if (textualProps.TryGetValue("DataSource", out string dsName))
            {
                // Store the name, resolution to actual object happens later if needed.
                DataSource = dsName;
            }
        }
    }
}