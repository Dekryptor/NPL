// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class PictureBoxProperties : ControlSpecificPropertiesBase
    {
        public Appearance Appearance { get; set; } = Appearance.Flat; // Typically Flat for PictureBox, but can be 3D
        public bool AutoRedraw { get; set; } = false;
        public bool AutoSize { get; set; } = false;
        public Vb6Color BackColor { get; set; } = Vb6Color.DefaultSystemColors.ButtonFace; // Or WindowBackground
        public BorderStyleConstants BorderStyle { get; set; } = BorderStyleConstants.FixedSingle;
        public bool ClipControls { get; set; } = true; // If it acts as a container
        public DataFormat DataFormat { get; set; } // Placeholder
        public string DataField { get; set; } = string.Empty;
        public object DataSource { get; set; } // Placeholder
        public DragMode DragMode { get; set; } = DragMode.Manual;
        public object DragIcon { get; set; } // Placeholder
        public DrawModeConstants DrawMode { get; set; } = DrawModeConstants.CopyPen;
        public DrawStyleConstants DrawStyle { get; set; } = DrawStyleConstants.Solid;
        public int DrawWidth { get; set; } = 1; // In pixels
        public Activation Enabled { get; set; } = Activation.Enabled;
        public Vb6Color FillColor { get; set; } = Vb6Color.Black;
        public FillStyleConstants FillStyle { get; set; } = FillStyleConstants.Transparent;
        // Font is inherited (for text drawn on the PictureBox or if it contains other controls)
        public Vb6Color ForeColor { get; set; } = Vb6Color.DefaultSystemColors.WindowText;
        public int Height { get; set; } = 1200; // Default in Twips
        public int Index { get; set; } = -1; // Contextual
        public object Image { get; set; } // Persistent bitmap, not cleared by graphics methods
        public int Left { get; set; } = 0;
        public LinkModeConstants LinkMode { get; set; } = LinkModeConstants.None; // DDE
        public string LinkItem { get; set; } = string.Empty; // DDE
        public string LinkTopic { get; set; } = string.Empty; // DDE
        public int LinkTimeout { get; set; } = 50; // DDE
        public MousePointer MousePointer { get; set; } = MousePointer.Default;
        public object MouseIcon { get; set; } // Placeholder
        public OLEDropMode OLEDropMode { get; set; } = OLEDropMode.None;
        public object Picture { get; set; } // The image displayed in the control (can be FRX)
        public ScaleModeConstants ScaleMode { get; set; } = ScaleModeConstants.Twips;
        public int ScaleHeight { get; set; } // User-defined scale height
        public int ScaleLeft { get; set; }   // User-defined scale X-coordinate
        public int ScaleTop { get; set; }    // User-defined scale Y-coordinate
        public int ScaleWidth { get; set; }  // User-defined scale width
        public int TabIndex { get; set; } = 0;
        public bool TabStop { get; set; } = false; // Usually false unless it's interactive
        public string Tag { get; set; } = string.Empty;
        public string ToolTipText { get; set; } = string.Empty;
        public int Top { get; set; } = 0;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1800; // Default in Twips

        public PictureBoxProperties()
        {
            Font = new Font();
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            Appearance = PropertyParsingHelpers.GetEnum(textualProps, "Appearance", this.Appearance);
            AutoRedraw = PropertyParsingHelpers.GetBool(textualProps, "AutoRedraw", this.AutoRedraw);
            AutoSize = PropertyParsingHelpers.GetBool(textualProps, "AutoSize", this.AutoSize);
            BackColor = PropertyParsingHelpers.GetColor(textualProps, "BackColor", this.BackColor);
            BorderStyle = PropertyParsingHelpers.GetEnum(textualProps, "BorderStyle", this.BorderStyle);
            ClipControls = PropertyParsingHelpers.GetBool(textualProps, "ClipControls", this.ClipControls);
            DataField = PropertyParsingHelpers.GetString(textualProps, "DataField", this.DataField);
            DragMode = PropertyParsingHelpers.GetEnum(textualProps, "DragMode", this.DragMode);
            DrawMode = PropertyParsingHelpers.GetEnum(textualProps, "DrawMode", this.DrawMode);
            DrawStyle = PropertyParsingHelpers.GetEnum(textualProps, "DrawStyle", this.DrawStyle);
            DrawWidth = PropertyParsingHelpers.GetInt32(textualProps, "DrawWidth", this.DrawWidth);
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", this.Enabled);
            FillColor = PropertyParsingHelpers.GetColor(textualProps, "FillColor", this.FillColor);
            FillStyle = PropertyParsingHelpers.GetEnum(textualProps, "FillStyle", this.FillStyle);

            PopulateFontProperty(textualProps, propertyGroups);

            ForeColor = PropertyParsingHelpers.GetColor(textualProps, "ForeColor", this.ForeColor);
            Height = PropertyParsingHelpers.GetInt32(textualProps, "Height", this.Height);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", this.Left);
            LinkItem = PropertyParsingHelpers.GetString(textualProps, "LinkItem", this.LinkItem);
            LinkTopic = PropertyParsingHelpers.GetString(textualProps, "LinkTopic", this.LinkTopic);
            LinkTimeout = PropertyParsingHelpers.GetInt32(textualProps, "LinkTimeout", this.LinkTimeout);
            MousePointer = PropertyParsingHelpers.GetEnum(textualProps, "MousePointer", this.MousePointer);
            OLEDropMode = PropertyParsingHelpers.GetEnum(textualProps, "OLEDropMode", this.OLEDropMode);
            ScaleMode = PropertyParsingHelpers.GetEnum(textualProps, "ScaleMode", this.ScaleMode);
            ScaleHeight = PropertyParsingHelpers.GetInt32(textualProps, "ScaleHeight", this.ScaleHeight); // Might be float
            ScaleLeft = PropertyParsingHelpers.GetInt32(textualProps, "ScaleLeft", this.ScaleLeft);     // Might be float
            ScaleTop = PropertyParsingHelpers.GetInt32(textualProps, "ScaleTop", this.ScaleTop);       // Might be float
            ScaleWidth = PropertyParsingHelpers.GetInt32(textualProps, "ScaleWidth", this.ScaleWidth);   // Might be float
            TabIndex = PropertyParsingHelpers.GetInt32(textualProps, "TabIndex", this.TabIndex);
            TabStop = PropertyParsingHelpers.GetBool(textualProps, "TabStop", this.TabStop);
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            ToolTipText = PropertyParsingHelpers.GetString(textualProps, "ToolTipText", this.ToolTipText);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", this.Top);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", this.Visible);
            WhatsThisHelpID = PropertyParsingHelpers.GetInt32(textualProps, "WhatsThisHelpID", this.WhatsThisHelpID);
            Width = PropertyParsingHelpers.GetInt32(textualProps, "Width", this.Width);

            // Binary properties
            if (binaryProps.TryGetValue("Picture", out byte[] picData)) Picture = picData; // FRX data
            if (binaryProps.TryGetValue("Image", out byte[] imgData)) Image = imgData; // FRX data
            if (binaryProps.TryGetValue("DragIcon", out byte[] dragIconData)) DragIcon = dragIconData;
            if (binaryProps.TryGetValue("MouseIcon", out byte[] mouseIconData)) MouseIcon = mouseIconData;
            
            if (textualProps.TryGetValue("DataSource", out string dsName))
            {
                DataSource = dsName;
            }
        }
    }

    // Note: DrawModeConstants, DrawStyleConstants, FillStyleConstants, ScaleModeConstants
    // would need to be defined, likely in a shared Enums.cs or within VB6Parse.Language.
    // For brevity, I'm assuming they exist. If not, they'd look like:
    /*
    public enum DrawModeConstants { CopyPen = 13, NotCopyPen = 4, ... }
    public enum DrawStyleConstants { Solid = 0, Dash = 1, ... }
    public enum FillStyleConstants { Solid = 0, Transparent = 1, ... }
    public enum ScaleModeConstants { User = 0, Twips = 1, Points = 2, Pixels = 3, Characters = 4, Inches = 5, Millimeters = 6, Centimeters = 7 }
    */
}