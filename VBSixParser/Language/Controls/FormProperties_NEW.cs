// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model; // For Vb6PropertyGroup
using VB6Parse.Utilities; // For PropertyParsingHelpers

namespace VB6Parse.Language.Controls
{
    public class FormProperties : ControlSpecificPropertiesBase
    {
        public bool ActiveControl { get; set; } // Runtime, not saved
        public Appearance Appearance { get; set; } = Appearance.ThreeD;
        public AutoRedrawConstants AutoRedraw { get; set; } = AutoRedrawConstants.False;
        public Vb6Color BackColor { get; set; } = Vb6Color.DefaultSystemColors.WindowBackground; // System Window Background
        public BorderStyleConstants BorderStyle { get; set; } = BorderStyleConstants.Sizable;
        public string Caption { get; set; } = "Form"; // Default typically "Form1", "Form2" etc.
        public bool ClipControls { get; set; } = true;
        public bool ControlBox { get; set; } = true;
        public DrawModeConstants DrawMode { get; set; } = DrawModeConstants.CopyPen;
        public DrawStyleConstants DrawStyle { get; set; } = DrawStyleConstants.Solid;
        public int DrawWidth { get; set; } = 1;
        public Activation Enabled { get; set; } = Activation.Enabled;
        public FillColorConstants FillColor { get; set; } = FillColorConstants.Black; // Default 0
        public FillStyleConstants FillStyle { get; set; } = FillStyleConstants.Transparent;
        // Font is inherited
        public Vb6Color ForeColor { get; set; } = Vb6Color.DefaultSystemColors.WindowText; // System Window Text
        public bool HasDC { get; set; } = true; // Runtime, not saved
        public int Height { get; set; } = 4000; // Default in Twips, example value
        public string HelpContextID { get; set; } = "0";
        public object Icon { get; set; } // Placeholder for icon data (binary from FRX)
        public bool KeyPreview { get; set; } = false;
        public int Left { get; set; } = 1620; // Default in Twips, example value (often screen dependent)
        public LinkModeConstants LinkMode { get; set; } = LinkModeConstants.None;
        public string LinkTopic { get; set; } = string.Empty; // Typically "Form1"
        public bool MaxButton { get; set; } = true;
        public bool MDIChild { get; set; } = false;
        public bool MinButton { get; set; } = true;
        public MousePointer MousePointer { get; set; } = MousePointer.Default;
        public object MouseIcon { get; set; } // Placeholder for icon data
        public bool Moveable { get; set; } = true;
        public bool NegotiateMenus { get; set; } = true; // Only if MDIChild = True
        public OLEDropMode OLEDropMode { get; set; } = OLEDropMode.None;
        public PaletteModeConstants PaletteMode { get; set; } = PaletteModeConstants.Halftone;
        public object Palette { get; set; } // Placeholder for palette data
        public object Picture { get; set; } // Placeholder for picture data
        public RightToLeftConstants RightToLeft { get; set; } = RightToLeftConstants.False; // For BiDi versions of Windows
        public ScaleModeConstants ScaleMode { get; set; } = ScaleModeConstants.Twip;
        public int ScaleHeight { get; set; } // Depends on ClientHeight and ScaleMode
        public int ScaleLeft { get; set; } = 0;
        public int ScaleTop { get; set; } = 0;
        public int ScaleWidth { get; set; } // Depends on ClientWidth and ScaleMode
        public bool ShowInTaskbar { get; set; } = true;
        public StartUpPositionMode StartUpPosition { get; set; } = StartUpPositionMode.WindowsDefault;
        public string Tag { get; set; } = string.Empty;
        public int Top { get; set; } = 1665; // Default in Twips, example value
        public bool Visible { get; set; } = true;
        public string WhatsThisButton { get; set; } // Property name seems off, usually WhatsThisHelp + WhatsThisButton on control
        public bool WhatsThisHelp { get; set; } = false;
        public int Width { get; set; } = 7000; // Default in Twips, example value
        public WindowStateConstants WindowState { get; set; } = WindowStateConstants.Normal;

        // Client area coordinates (usually derived or set by IDE based on Form size and borders)
        public int ClientHeight { get; set; }
        public int ClientLeft { get; set; }
        public int ClientTop { get; set; }
        public int ClientWidth { get; set; }


        public FormProperties()
        {
            Font = new Font(); // Initialize with default Font
            // Set default caption based on common pattern if possible, or leave as "Form"
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            Appearance = PropertyParsingHelpers.GetEnum(textualProps, "Appearance", this.Appearance);
            AutoRedraw = PropertyParsingHelpers.GetEnum(textualProps, "AutoRedraw", this.AutoRedraw);
            BackColor = PropertyParsingHelpers.GetColor(textualProps, "BackColor", this.BackColor);
            BorderStyle = PropertyParsingHelpers.GetEnum(textualProps, "BorderStyle", this.BorderStyle);
            Caption = PropertyParsingHelpers.GetString(textualProps, "Caption", this.Caption);
            ClipControls = PropertyParsingHelpers.GetBool(textualProps, "ClipControls", this.ClipControls);
            ControlBox = PropertyParsingHelpers.GetBool(textualProps, "ControlBox", this.ControlBox);
            DrawMode = PropertyParsingHelpers.GetEnum(textualProps, "DrawMode", this.DrawMode);
            DrawStyle = PropertyParsingHelpers.GetEnum(textualProps, "DrawStyle", this.DrawStyle);
            DrawWidth = PropertyParsingHelpers.GetInt32(textualProps, "DrawWidth", this.DrawWidth);
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", this.Enabled);
            FillColor = PropertyParsingHelpers.GetEnum(textualProps, "FillColor", this.FillColor); // This is a Vb6Color, needs GetColor
            FillStyle = PropertyParsingHelpers.GetEnum(textualProps, "FillStyle", this.FillStyle);
            
            PopulateFontProperty(textualProps, propertyGroups); // Base class helper for Font

            ForeColor = PropertyParsingHelpers.GetColor(textualProps, "ForeColor", this.ForeColor);
            Height = PropertyParsingHelpers.GetInt32(textualProps, "Height", this.Height);
            HelpContextID = PropertyParsingHelpers.GetString(textualProps, "HelpContextID", this.HelpContextID); // String in VB6 docs
            KeyPreview = PropertyParsingHelpers.GetBool(textualProps, "KeyPreview", this.KeyPreview);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", this.Left);
            LinkMode = PropertyParsingHelpers.GetEnum(textualProps, "LinkMode", this.LinkMode);
            LinkTopic = PropertyParsingHelpers.GetString(textualProps, "LinkTopic", this.LinkTopic);
            MaxButton = PropertyParsingHelpers.GetBool(textualProps, "MaxButton", this.MaxButton);
            MDIChild = PropertyParsingHelpers.GetBool(textualProps, "MDIChild", this.MDIChild);
            MinButton = PropertyParsingHelpers.GetBool(textualProps, "MinButton", this.MinButton);
            MousePointer = PropertyParsingHelpers.GetEnum(textualProps, "MousePointer", this.MousePointer);
            Moveable = PropertyParsingHelpers.GetBool(textualProps, "Moveable", this.Moveable);
            NegotiateMenus = PropertyParsingHelpers.GetBool(textualProps, "NegotiateMenus", this.NegotiateMenus);
            OLEDropMode = PropertyParsingHelpers.GetEnum(textualProps, "OLEDropMode", this.OLEDropMode);
            PaletteMode = PropertyParsingHelpers.GetEnum(textualProps, "PaletteMode", this.PaletteMode);
            RightToLeft = PropertyParsingHelpers.GetEnum(textualProps, "RightToLeft", this.RightToLeft);
            ScaleMode = PropertyParsingHelpers.GetEnum(textualProps, "ScaleMode", this.ScaleMode);
            
            // ScaleHeight, ScaleLeft, ScaleTop, ScaleWidth are often explicitly set
            ScaleHeight = PropertyParsingHelpers.GetInt32(textualProps, "ScaleHeight", this.ScaleHeight);
            ScaleLeft = PropertyParsingHelpers.GetInt32(textualProps, "ScaleLeft", this.ScaleLeft);
            ScaleTop = PropertyParsingHelpers.GetInt32(textualProps, "ScaleTop", this.ScaleTop);
            ScaleWidth = PropertyParsingHelpers.GetInt32(textualProps, "ScaleWidth", this.ScaleWidth);

            ShowInTaskbar = PropertyParsingHelpers.GetBool(textualProps, "ShowInTaskbar", this.ShowInTaskbar);
            
            // StartUpPosition and related Client coordinates
            int clientL, clientT, clientW, clientH;
            StartUpPosition = PropertyParsingHelpers.GetStartUpPosition(textualProps, "StartUpPosition", this.StartUpPosition,
                                                                  out clientL, out clientT, out clientW, out clientH);
            ClientLeft = PropertyParsingHelpers.GetInt32(textualProps, "ClientLeft", clientL); // Use value from GetStartUpPosition if not explicitly set
            ClientTop = PropertyParsingHelpers.GetInt32(textualProps, "ClientTop", clientT);
            ClientWidth = PropertyParsingHelpers.GetInt32(textualProps, "ClientWidth", clientW);
            ClientHeight = PropertyParsingHelpers.GetInt32(textualProps, "ClientHeight", clientH);


            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", this.Top);
            Visible = PropertyParsingHelpers.GetBool(textualProps, "Visible", this.Visible); // Visible is boolean
            WhatsThisHelp = PropertyParsingHelpers.GetBool(textualProps, "WhatsThisHelp", this.WhatsThisHelp);
            Width = PropertyParsingHelpers.GetInt32(textualProps, "Width", this.Width);
            WindowState = PropertyParsingHelpers.GetEnum(textualProps, "WindowState", this.WindowState);

            // Binary properties
            if (binaryProps.TryGetValue("Icon", out byte[] iconData)) Icon = iconData;
            if (binaryProps.TryGetValue("MouseIcon", out byte[] mouseIconData)) MouseIcon = mouseIconData;
            if (binaryProps.TryGetValue("Picture", out byte[] pictureData)) Picture = pictureData;
            if (binaryProps.TryGetValue("Palette", out byte[] paletteData)) Palette = paletteData;
            
            // FillColor is a color, not an enum from textual properties directly
            if (textualProps.TryGetValue("FillColor", out string fcStr))
            {
                if (Vb6Color.TryParseHex(fcStr, out Vb6Color fc))
                {
                     // This specific property "FillColor" on Form is actually a color.
                     // The FillColorConstants enum is for other contexts.
                     // Let's assume we need a Vb6Color field for this.
                     // For now, I'll map it to a new Vb6Color field if we had one, or ignore if FillColorConstants is strict.
                     // The original `FillColorConstants FillColor` field should be `Vb6Color FillColor`
                }
            }


        }
    }
}