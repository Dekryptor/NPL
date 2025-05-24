// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class FrameProperties : ControlSpecificPropertiesBase
    {
        public Appearance Appearance { get; set; } = Appearance.ThreeD;
        public Vb6Color BackColor { get; set; } = Vb6Color.DefaultSystemColors.ButtonFace;
        public BorderStyleConstants BorderStyle { get; set; } = BorderStyleConstants.FixedSingle; // Or None
        public string Caption { get; set; } = "Frame"; // Default "Frame1", "Frame2"
        public bool ClipControls { get; set; } = true;
        public DragMode DragMode { get; set; } = DragMode.Manual;
        public object DragIcon { get; set; } // Placeholder
        public Activation Enabled { get; set; } = Activation.Enabled;
        // Font is inherited (applies to the caption of the Frame)
        public Vb6Color ForeColor { get; set; } = Vb6Color.DefaultSystemColors.WindowText; // For caption text
        public int Height { get; set; } = 1500; // Default in Twips (approx)
        public int Index { get; set; } = -1; // Contextual
        public int Left { get; set; } = 0;
        public MousePointer MousePointer { get; set; } = MousePointer.Default;
        public object MouseIcon { get; set; } // Placeholder
        public OLEDropMode OLEDropMode { get; set; } = OLEDropMode.None;
        public RightToLeftConstants RightToLeft { get; set; } = RightToLeftConstants.False;
        public int TabIndex { get; set; } = 0;
        // TabStop for Frame itself is usually False, as it's a container.
        // However, the property exists. Controls inside the frame have their own TabStop.
        public bool TabStop { get; set; } = false; 
        public string Tag { get; set; } = string.Empty;
        public string ToolTipText { get; set; } = string.Empty;
        public int Top { get; set; } = 0;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 3000; // Default in Twips (approx)

        public FrameProperties()
        {
            Font = new Font(); // Font for the Frame's caption
            Caption = "Frame1"; // Default caption often "Frame1"
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            Appearance = PropertyParsingHelpers.GetEnum(textualProps, "Appearance", this.Appearance);
            BackColor = PropertyParsingHelpers.GetColor(textualProps, "BackColor", this.BackColor);
            BorderStyle = PropertyParsingHelpers.GetEnum(textualProps, "BorderStyle", this.BorderStyle);
            Caption = PropertyParsingHelpers.GetString(textualProps, "Caption", this.Caption);
            ClipControls = PropertyParsingHelpers.GetBool(textualProps, "ClipControls", this.ClipControls);
            DragMode = PropertyParsingHelpers.GetEnum(textualProps, "DragMode", this.DragMode);
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", this.Enabled);

            PopulateFontProperty(textualProps, propertyGroups); // For the Frame's caption

            ForeColor = PropertyParsingHelpers.GetColor(textualProps, "ForeColor", this.ForeColor);
            Height = PropertyParsingHelpers.GetInt32(textualProps, "Height", this.Height);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", this.Left);
            MousePointer = PropertyParsingHelpers.GetEnum(textualProps, "MousePointer", this.MousePointer);
            OLEDropMode = PropertyParsingHelpers.GetEnum(textualProps, "OLEDropMode", this.OLEDropMode);
            RightToLeft = PropertyParsingHelpers.GetEnum(textualProps, "RightToLeft", this.RightToLeft);
            TabIndex = PropertyParsingHelpers.GetInt32(textualProps, "TabIndex", this.TabIndex);
            TabStop = PropertyParsingHelpers.GetBool(textualProps, "TabStop", this.TabStop);
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            ToolTipText = PropertyParsingHelpers.GetString(textualProps, "ToolTipText", this.ToolTipText);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", this.Top);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", this.Visible);
            WhatsThisHelpID = PropertyParsingHelpers.GetInt32(textualProps, "WhatsThisHelpID", this.WhatsThisHelpID);
            Width = PropertyParsingHelpers.GetInt32(textualProps, "Width", this.Width);

            // Binary properties
            if (binaryProps.TryGetValue("DragIcon", out byte[] dragIconData)) DragIcon = dragIconData;
            if (binaryProps.TryGetValue("MouseIcon", out byte[] mouseIconData)) MouseIcon = mouseIconData;
        }
    }
}