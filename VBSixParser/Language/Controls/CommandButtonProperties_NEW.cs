// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class CommandButtonProperties : ControlSpecificPropertiesBase
    {
        public Appearance Appearance { get; set; } = Appearance.ThreeD;
        public Vb6Color BackColor { get; set; } = Vb6Color.DefaultButtonFace;
        public string Caption { get; set; } = "CommandButton";
        public bool Cancel { get; set; } = false;
        public bool Default { get; set; } = false;
        public DragMode DragMode { get; set; } = DragMode.Manual;
        public object DragIcon { get; set; } // Placeholder for image data
        public Activation Enabled { get; set; } = Activation.Enabled;
        // Font is inherited from ControlSpecificPropertiesBase
        public int Height { get; set; } = 375; // Default in Twips
        public int HelpContextID { get; set; } = 0;
        public int Index { get; set; } = -1; // Not a property, but often part of control context
        public int Left { get; set; } = 0;
        public MaskColorMode MaskColor { get; set; } // Used with Style = Graphical
        public object MouseIcon { get; set; } // Placeholder
        public MousePointer MousePointer { get; set; } = MousePointer.Default;
        public OLEDropMode OLEDropMode { get; set; } = OLEDropMode.None;
        public object Picture { get; set; } // Placeholder
        public CommandButtonStyleConstants Style { get; set; } = CommandButtonStyleConstants.Standard;
        public int TabIndex { get; set; } = 0;
        public bool TabStop { get; set; } = true;
        public string Tag { get; set; } = string.Empty;
        public string ToolTipText { get; set; } = string.Empty;
        public int Top { get; set; } = 0;
        public UseMaskColorConstants UseMaskColor { get; set; } = UseMaskColorConstants.False;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1215; // Default in Twips

        public CommandButtonProperties()
        {
            // Initialize with common VB6 defaults if not done by field initializers
            Font = new Font(); // Default font
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            Appearance = PropertyParsingHelpers.GetEnum(textualProps, "Appearance", Appearance.ThreeD);
            BackColor = PropertyParsingHelpers.GetColor(textualProps, "BackColor", Vb6Color.DefaultButtonFace);
            Caption = PropertyParsingHelpers.GetString(textualProps, "Caption", this.Caption); // Use field default if not found
            Cancel = PropertyParsingHelpers.GetBool(textualProps, "Cancel", false);
            Default = PropertyParsingHelpers.GetBool(textualProps, "Default", false);
            DragMode = PropertyParsingHelpers.GetEnum(textualProps, "DragMode", DragMode.Manual);
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", Activation.Enabled);
            
            PopulateFontProperty(textualProps, propertyGroups); // Use helper from base class

            Height = PropertyParsingHelpers.GetInt32(textualProps, "Height", this.Height);
            HelpContextID = PropertyParsingHelpers.GetInt32(textualProps, "HelpContextID", 0);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", 0);
            
            // MaskColor needs special handling if it's an optional color
            // MousePointer etc.
            MousePointer = PropertyParsingHelpers.GetEnum(textualProps, "MousePointer", MousePointer.Default);
            OLEDropMode = PropertyParsingHelpers.GetEnum(textualProps, "OLEDropMode", OLEDropMode.None);
            Style = PropertyParsingHelpers.GetEnum(textualProps, "Style", CommandButtonStyleConstants.Standard);
            TabIndex = PropertyParsingHelpers.GetInt32(textualProps, "TabIndex", 0);
            TabStop = PropertyParsingHelpers.GetBool(textualProps, "TabStop", true);
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", string.Empty); // Tag is also on Vb6Control directly
            ToolTipText = PropertyParsingHelpers.GetString(textualProps, "ToolTipText", string.Empty);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", 0);
            UseMaskColor = PropertyParsingHelpers.GetEnum(textualProps, "UseMaskColor", UseMaskColorConstants.False);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", Visibility.Visible);
            WhatsThisHelpID = PropertyParsingHelpers.GetInt32(textualProps, "WhatsThisHelpID", 0);
            Width = PropertyParsingHelpers.GetInt32(textualProps, "Width", this.Width);

            // Handle binary properties like Picture, DragIcon, MouseIcon
            if (binaryProps.TryGetValue("Picture", out byte[] picData)) Picture = picData; // Or convert to Image object
            if (binaryProps.TryGetValue("DragIcon", out byte[] dragIconData)) DragIcon = dragIconData;
            if (binaryProps.TryGetValue("MouseIcon", out byte[] mouseIconData)) MouseIcon = mouseIconData;
            
            // If MaskColor is stored as text (e.g. from a resource string in FRX)
            if (textualProps.TryGetValue("MaskColor", out string maskColorString))
            {
                 if(Vb6Color.TryParseHex(maskColorString, out Vb6Color mc))
                 {
                    // This is tricky, MaskColor property in VB is a color but enum is about mode.
                    // This property might be named differently or not exist for CommandButton.
                    // For now, let's assume it's not directly on CommandButton this way.
                 }
            }

        }
    }
}