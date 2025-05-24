// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class TextBoxProperties : ControlSpecificPropertiesBase
    {
        public AlignmentConstants Alignment { get; set; } = AlignmentConstants.LeftJustify;
        public Appearance Appearance { get; set; } = Appearance.ThreeD;
        public Vb6Color BackColor { get; set; } = Vb6Color.DefaultSystemColors.WindowBackground;
        public BorderStyleConstants BorderStyle { get; set; } = BorderStyleConstants.FixedSingle;
        public DataFormat DataFormat { get; set; } // Placeholder
        public string DataField { get; set; } = string.Empty;
        public object DataSource { get; set; } // Placeholder
        public DragMode DragMode { get; set; } = DragMode.Manual;
        public object DragIcon { get; set; } // Placeholder
        public Activation Enabled { get; set; } = Activation.Enabled;
        // Font is inherited
        public Vb6Color ForeColor { get; set; } = Vb6Color.DefaultSystemColors.WindowText;
        public int Height { get; set; } = 300; // Default in Twips (approx)
        public HideSelectionConstants HideSelection { get; set; } = HideSelectionConstants.True;
        public int Index { get; set; } = -1; // Contextual
        public int Left { get; set; } = 0;
        public LinkModeConstants LinkMode { get; set; } = LinkModeConstants.None;
        public string LinkItem { get; set; } = string.Empty;
        public string LinkTopic { get; set; } = string.Empty;
        public int LinkTimeout { get; set; } = 50;
        public LockedConstants Locked { get; set; } = LockedConstants.False;
        public int MaxLength { get; set; } = 0; // 0 means no limit (or max system limit)
        public MousePointer MousePointer { get; set; } = MousePointer.IBeam;
        public object MouseIcon { get; set; } // Placeholder
        public MultiLineConstants MultiLine { get; set; } = MultiLineConstants.False;
        public OLEDropMode OLEDropMode { get; set; } = OLEDropMode.None;
        public string PasswordChar { get; set; } = string.Empty;
        public RightToLeftConstants RightToLeft { get; set; } = RightToLeftConstants.False;
        public ScrollBarsConstants ScrollBars { get; set; } = ScrollBarsConstants.None;
        public int SelLength { get; set; } = 0; // Runtime
        public int SelStart { get; set; } = 0; // Runtime
        public string SelText { get; set; } = string.Empty; // Runtime
        public string Tag { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty; // Default often "Text1"
        public string ToolTipText { get; set; } = string.Empty;
        public int Top { get; set; } = 0;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1215; // Default in Twips

        public TextBoxProperties()
        {
            Font = new Font();
            Text = string.Empty; // Default Text property is empty, Caption is "Text1" in IDE
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            Alignment = PropertyParsingHelpers.GetEnum(textualProps, "Alignment", this.Alignment);
            Appearance = PropertyParsingHelpers.GetEnum(textualProps, "Appearance", this.Appearance);
            BackColor = PropertyParsingHelpers.GetColor(textualProps, "BackColor", this.BackColor);
            BorderStyle = PropertyParsingHelpers.GetEnum(textualProps, "BorderStyle", this.BorderStyle);
            DataField = PropertyParsingHelpers.GetString(textualProps, "DataField", this.DataField);
            DragMode = PropertyParsingHelpers.GetEnum(textualProps, "DragMode", this.DragMode);
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", this.Enabled);

            PopulateFontProperty(textualProps, propertyGroups);

            ForeColor = PropertyParsingHelpers.GetColor(textualProps, "ForeColor", this.ForeColor);
            Height = PropertyParsingHelpers.GetInt32(textualProps, "Height", this.Height);
            HideSelection = PropertyParsingHelpers.GetEnum(textualProps, "HideSelection", this.HideSelection);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", this.Left);
            LinkItem = PropertyParsingHelpers.GetString(textualProps, "LinkItem", this.LinkItem);
            LinkTopic = PropertyParsingHelpers.GetString(textualProps, "LinkTopic", this.LinkTopic);
            LinkTimeout = PropertyParsingHelpers.GetInt32(textualProps, "LinkTimeout", this.LinkTimeout);
            Locked = PropertyParsingHelpers.GetEnum(textualProps, "Locked", this.Locked);
            MaxLength = PropertyParsingHelpers.GetInt32(textualProps, "MaxLength", this.MaxLength);
            MousePointer = PropertyParsingHelpers.GetEnum(textualProps, "MousePointer", this.MousePointer);
            MultiLine = PropertyParsingHelpers.GetEnum(textualProps, "MultiLine", this.MultiLine);
            OLEDropMode = PropertyParsingHelpers.GetEnum(textualProps, "OLEDropMode", this.OLEDropMode);
            PasswordChar = PropertyParsingHelpers.GetString(textualProps, "PasswordChar", this.PasswordChar);
            RightToLeft = PropertyParsingHelpers.GetEnum(textualProps, "RightToLeft", this.RightToLeft);
            ScrollBars = PropertyParsingHelpers.GetEnum(textualProps, "ScrollBars", this.ScrollBars);
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            Text = PropertyParsingHelpers.GetString(textualProps, "Text", this.Text); // The initial text content
            ToolTipText = PropertyParsingHelpers.GetString(textualProps, "ToolTipText", this.ToolTipText);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", this.Top);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", this.Visible);
            WhatsThisHelpID = PropertyParsingHelpers.GetInt32(textualProps, "WhatsThisHelpID", this.WhatsThisHelpID);
            Width = PropertyParsingHelpers.GetInt32(textualProps, "Width", this.Width);

            // Binary properties
            if (binaryProps.TryGetValue("DragIcon", out byte[] dragIconData)) DragIcon = dragIconData;
            if (binaryProps.TryGetValue("MouseIcon", out byte[] mouseIconData)) MouseIcon = mouseIconData;
            
            // DataFormat and DataSource
            if (textualProps.TryGetValue("DataSource", out string dsName))
            {
                DataSource = dsName;
            }
            // DataFormat would require parsing its own group or set of properties.
        }
    }

    // Supporting Enums (if not already defined elsewhere)
    public enum HideSelectionConstants
    {
        True = -1, // VB True
        False = 0
    }

    public enum LockedConstants
    {
        True = -1,
        False = 0
    }

    public enum MultiLineConstants
    {
        True = -1,
        False = 0
    }

    public enum ScrollBarsConstants
    {
        None = 0,
        Horizontal = 1,
        Vertical = 2,
        Both = 3
    }
}