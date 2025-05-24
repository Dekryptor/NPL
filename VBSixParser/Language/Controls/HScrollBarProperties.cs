// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class HScrollBarProperties : ControlSpecificPropertiesBase
    {
        public DragMode DragMode { get; set; } = DragMode.Manual;
        public object DragIcon { get; set; } // Placeholder
        public Activation Enabled { get; set; } = Activation.Enabled;
        public int Height { get; set; } = 255; // Default height for HScrollBar (fixed)
        public int Index { get; set; } = -1; // Contextual
        public int LargeChange { get; set; } = 1;
        public int Left { get; set; } = 0;
        public int Max { get; set; } = 32767;
        public int Min { get; set; } = 0;
        public MousePointer MousePointer { get; set; } = MousePointer.Default;
        public object MouseIcon { get; set; } // Placeholder
        public OLEDropMode OLEDropMode { get; set; } = OLEDropMode.None;
        public RightToLeftConstants RightToLeft { get; set; } = RightToLeftConstants.False; // Affects visual layout
        public int SmallChange { get; set; } = 1;
        public int TabIndex { get; set; } = 0;
        public bool TabStop { get; set; } = true;
        public string Tag { get; set; } = string.Empty;
        public int Top { get; set; } = 0;
        public int Value { get; set; } = 0;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1215; // Default width for HScrollBar

        // HScrollBars do not have Font, BackColor, ForeColor, Caption properties.
        // The base Font property will be nullified.

        public HScrollBarProperties()
        {
            this.Font = null; // ScrollBars don't have a Font property.
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            DragMode = PropertyParsingHelpers.GetEnum(textualProps, "DragMode", this.DragMode);
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", this.Enabled);
            Height = PropertyParsingHelpers.GetInt32(textualProps, "Height", this.Height);
            LargeChange = PropertyParsingHelpers.GetInt32(textualProps, "LargeChange", this.LargeChange);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", this.Left);
            Max = PropertyParsingHelpers.GetInt32(textualProps, "Max", this.Max);
            Min = PropertyParsingHelpers.GetInt32(textualProps, "Min", this.Min);
            MousePointer = PropertyParsingHelpers.GetEnum(textualProps, "MousePointer", this.MousePointer);
            OLEDropMode = PropertyParsingHelpers.GetEnum(textualProps, "OLEDropMode", this.OLEDropMode);
            RightToLeft = PropertyParsingHelpers.GetEnum(textualProps, "RightToLeft", this.RightToLeft);
            SmallChange = PropertyParsingHelpers.GetInt32(textualProps, "SmallChange", this.SmallChange);
            TabIndex = PropertyParsingHelpers.GetInt32(textualProps, "TabIndex", this.TabIndex);
            TabStop = PropertyParsingHelpers.GetBool(textualProps, "TabStop", this.TabStop);
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", this.Top);
            Value = PropertyParsingHelpers.GetInt32(textualProps, "Value", this.Value);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", this.Visible);
            WhatsThisHelpID = PropertyParsingHelpers.GetInt32(textualProps, "WhatsThisHelpID", this.WhatsThisHelpID);
            Width = PropertyParsingHelpers.GetInt32(textualProps, "Width", this.Width);

            // Binary properties
            if (binaryProps.TryGetValue("DragIcon", out byte[] dragIconData)) DragIcon = dragIconData;
            if (binaryProps.TryGetValue("MouseIcon", out byte[] mouseIconData)) MouseIcon = mouseIconData;
            
            // No Font property group for scrollbars.
        }
        
        protected override void PopulateFontProperty(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            // ScrollBars do not have a Font property.
            this.Font = null;
        }
    }
}