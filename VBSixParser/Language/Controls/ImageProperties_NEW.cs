// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class ImageProperties : ControlSpecificPropertiesBase
    {
        public Appearance Appearance { get; set; } = Appearance.Flat; // Image controls are always flat
        public BorderStyleConstants BorderStyle { get; set; } = BorderStyleConstants.None; // Or FixedSingle
        public DataFormat DataFormat { get; set; } // Placeholder
        public string DataField { get; set; } = string.Empty;
        public object DataSource { get; set; } // Placeholder
        public DragMode DragMode { get; set; } = DragMode.Manual;
        public object DragIcon { get; set; } // Placeholder
        public Activation Enabled { get; set; } = Activation.Enabled;
        public int Height { get; set; } = 735; // Default in Twips (approx)
        public int Index { get; set; } = -1; // Contextual
        public int Left { get; set; } = 0;
        public MousePointer MousePointer { get; set; } = MousePointer.Default;
        public object MouseIcon { get; set; } // Placeholder
        public OLEDropMode OLEDropMode { get; set; } = OLEDropMode.None;
        public object Picture { get; set; } // The image displayed (can be FRX)
        public bool Stretch { get; set; } = false; // If True, image stretches to fit control bounds
        public string Tag { get; set; } = string.Empty;
        public string ToolTipText { get; set; } = string.Empty;
        public int Top { get; set; } = 0;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 975; // Default in Twips (approx)
        
        // Image controls do not have Font, BackColor, ForeColor, Caption properties.
        // The base Font property will be nullified.

        public ImageProperties()
        {
            this.Font = null; // Image controls don't have a Font property.
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            // Appearance for Image is always Flat, so we might not even need to parse it.
            // It's not typically shown in the property sheet for an Image control.
            // If it exists in textualProps, it's likely 0 (Flat).
            Appearance = PropertyParsingHelpers.GetEnum(textualProps, "Appearance", Appearance.Flat);
            BorderStyle = PropertyParsingHelpers.GetEnum(textualProps, "BorderStyle", this.BorderStyle);
            DataField = PropertyParsingHelpers.GetString(textualProps, "DataField", this.DataField);
            DragMode = PropertyParsingHelpers.GetEnum(textualProps, "DragMode", this.DragMode);
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", this.Enabled);
            Height = PropertyParsingHelpers.GetInt32(textualProps, "Height", this.Height);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", this.Left);
            MousePointer = PropertyParsingHelpers.GetEnum(textualProps, "MousePointer", this.MousePointer);
            OLEDropMode = PropertyParsingHelpers.GetEnum(textualProps, "OLEDropMode", this.OLEDropMode);
            Stretch = PropertyParsingHelpers.GetBool(textualProps, "Stretch", this.Stretch);
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            ToolTipText = PropertyParsingHelpers.GetString(textualProps, "ToolTipText", this.ToolTipText);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", this.Top);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", this.Visible);
            WhatsThisHelpID = PropertyParsingHelpers.GetInt32(textualProps, "WhatsThisHelpID", this.WhatsThisHelpID);
            Width = PropertyParsingHelpers.GetInt32(textualProps, "Width", this.Width);

            // Binary properties
            if (binaryProps.TryGetValue("Picture", out byte[] picData)) Picture = picData; // FRX data
            if (binaryProps.TryGetValue("DragIcon", out byte[] dragIconData)) DragIcon = dragIconData;
            if (binaryProps.TryGetValue("MouseIcon", out byte[] mouseIconData)) MouseIcon = mouseIconData;
            
            if (textualProps.TryGetValue("DataSource", out string dsName))
            {
                DataSource = dsName;
            }

            // Image controls don't have Font property group.
            // Call to base.PopulateFontProperty is not needed or should be overridden.
        }
        
        protected override void PopulateFontProperty(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            // Image controls do not have a Font property.
            this.Font = null;
        }
    }
}