using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary
using System; // For Enum.IsDefined

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 Form or MDIForm control.
    /// These properties are typically found in a .frm file.
    /// </summary>
    public class FormProperties
    {
        public Appearance Appearance { get; set; } = Controls.Appearance.ThreeD;
        public AutoRedraw AutoRedraw { get; set; } = Controls.AutoRedraw.Manual;
        public VB6Color BackColor { get; set; } = VB6Color.VbButtonFace; // Default: &H8000000F&
        public FormBorderStyle BorderStyle { get; set; } = FormBorderStyle.Sizable;
        public string Caption { get; set; } = "Form1"; // Default caption can vary
        public int ClientHeight { get; set; } = 3000; // Typical default in Twips
        public int ClientLeft { get; set; } = 1020;  // Typical default in Twips
        public int ClientTop { get; set; } = 1590;   // Typical default in Twips
        public int ClientWidth { get; set; } = 4680;  // Typical default in Twips
        public ClipControls ClipControls { get; set; } = Controls.ClipControls.Clipped;
        public ControlBox ControlBox { get; set; } = Controls.ControlBox.Included;
        public DrawMode DrawMode { get; set; } = Controls.DrawMode.CopyPen;
        public DrawStyle DrawStyle { get; set; } = Controls.DrawStyle.Solid;
        public int DrawWidth { get; set; } = 1;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        public VB6Color FillColor { get; set; } = VB6Color.VbBlack; // Default: &H00000000&
        public FillStyle FillStyle { get; set; } = Controls.FillStyle.Transparent;
        // Font is a group property, handled by VB6PropertyGroup
        public FontTransparency FontTransparent { get; set; } = Controls.FontTransparency.Transparent;
        public VB6Color ForeColor { get; set; } = VB6Color.VbWindowText; // Default: &H80000012&
        public HasDeviceContext HasDC { get; set; } = Controls.HasDeviceContext.Yes;
        public int Height { get; set; } = 3885; // Overall height, sum of ClientHeight + borders/title bar
        public int HelpContextID { get; set; } = 0;
        // Icon, MouseIcon, Palette, Picture are FRX-based.
        // Their resolved byte[] data will be stored in the VB6FormKind/VB6MDIFormKind class.
        public bool KeyPreview { get; set; } = false;
        public int Left { get; set; } = 975; // Screen Left
        public FormLinkMode LinkMode { get; set; } = FormLinkMode.None; // For DDE
        public string LinkTopic { get; set; } = ""; // For DDE
        public MaxButton MaxButton { get; set; } = Controls.MaxButton.Included;
        public bool MDIChild { get; set; } = false; // True if this form is an MDI child
        public MinButton MinButton { get; set; } = Controls.MinButton.Included;
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public Movability Moveable { get; set; } = Controls.Movability.Moveable;
        public bool NegotiateMenus { get; set; } = true; // For MDI child forms and OLE
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None;
        public PaletteMode PaletteMode { get; set; } = Controls.PaletteMode.HalfTone;
        public TextDirection RightToLeft { get; set; } = Controls.TextDirection.LeftToRight;
        public int ScaleHeight { get; set; } = 3000; // Typically matches ClientHeight initially
        public int ScaleLeft { get; set; } = 0;
        public ScaleMode ScaleMode { get; set; } = Controls.ScaleMode.Twip;
        public int ScaleTop { get; set; } = 0;
        public int ScaleWidth { get; set; } = 4680; // Typically matches ClientWidth initially
        public ShowInTaskbar ShowInTaskbar { get; set; } = Controls.ShowInTaskbar.Show;
        public StartUpPositionType StartUpPosition { get; set; } = StartUpPositionType.WindowsDefault;
        public int Top { get; set; } = 1200; // Screen Top
        public Visibility Visible { get; set; } = Controls.Visibility.Visible; // Default for new forms at design time is True
        public WhatsThisButton WhatsThisButton { get; set; } = Controls.WhatsThisButton.Excluded; // VB6 default is False
        public WhatsThisHelp WhatsThisHelp { get; set; } = Controls.WhatsThisHelp.F1Help;
        public int Width { get; set; } = 4800; // Overall width
        public WindowState WindowState { get; set; } = Controls.WindowState.Normal;

        /// <summary>
        /// Default constructor initializing with common VB6 defaults.
        /// The caption will be set by the parser based on Attribute VB_Name or Form.Caption.
        /// </summary>
        public FormProperties() { }

        /// <summary>
        /// Initializes FormProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public FormProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            AutoRedraw = rawProps.GetEnum("AutoRedraw", AutoRedraw);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            BorderStyle = rawProps.GetEnum("BorderStyle", BorderStyle);
            Caption = rawProps.GetString("Caption", Caption); // Often from Attribute VB_Name
            ClientHeight = rawProps.GetInt32("ClientHeight", ClientHeight);
            ClientLeft = rawProps.GetInt32("ClientLeft", ClientLeft);
            ClientTop = rawProps.GetInt32("ClientTop", ClientTop);
            ClientWidth = rawProps.GetInt32("ClientWidth", ClientWidth);
            ClipControls = rawProps.GetEnum("ClipControls", ClipControls);
            ControlBox = rawProps.GetEnum("ControlBox", ControlBox);
            DrawMode = rawProps.GetEnum("DrawMode", DrawMode);
            DrawStyle = rawProps.GetEnum("DrawStyle", DrawStyle);
            DrawWidth = rawProps.GetInt32("DrawWidth", DrawWidth);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            FillColor = rawProps.GetVB6Color("FillColor", FillColor);
            FillStyle = rawProps.GetEnum("FillStyle", FillStyle);
            FontTransparent = rawProps.GetEnum("FontTransparent", FontTransparent);
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor);
            HasDC = rawProps.GetEnum("HasDC", HasDC);
            Height = rawProps.GetInt32("Height", Height); // This is the Form.Height, not ClientHeight
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            KeyPreview = rawProps.GetBoolean("KeyPreview", KeyPreview);
            Left = rawProps.GetInt32("Left", Left); // This is the Form.Left (screen coordinate)
            LinkMode = rawProps.GetEnum("LinkMode", LinkMode);
            LinkTopic = rawProps.GetString("LinkTopic", LinkTopic);
            MaxButton = rawProps.GetEnum("MaxButton", MaxButton);
            MDIChild = rawProps.GetBoolean("MDIChild", MDIChild);
            MinButton = rawProps.GetEnum("MinButton", MinButton);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            Moveable = rawProps.GetEnum("Moveable", Moveable);
            NegotiateMenus = rawProps.GetBoolean("NegotiateMenus", NegotiateMenus);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            PaletteMode = rawProps.GetEnum("PaletteMode", PaletteMode);
            RightToLeft = rawProps.GetEnum("RightToLeft", RightToLeft);
            ScaleHeight = rawProps.GetInt32("ScaleHeight", ScaleHeight);
            ScaleLeft = rawProps.GetInt32("ScaleLeft", ScaleLeft);
            ScaleMode = rawProps.GetEnum("ScaleMode", ScaleMode);
            ScaleTop = rawProps.GetInt32("ScaleTop", ScaleTop);
            ScaleWidth = rawProps.GetInt32("ScaleWidth", ScaleWidth);
            ShowInTaskbar = rawProps.GetEnum("ShowInTaskbar", ShowInTaskbar);
            
            // StartUpPosition property needs careful handling as its value depends on other properties if Manual(0)
            int startUpPosIntValue = rawProps.GetInt32("StartUpPosition", (int)StartUpPositionType.WindowsDefault);
            if (Enum.IsDefined(typeof(StartUpPositionType), startUpPosIntValue))
            {
                StartUpPosition = (StartUpPositionType)startUpPosIntValue;
            }
            else
            {
                StartUpPosition = StartUpPositionType.WindowsDefault; // Fallback
            }
            // If StartUpPosition is Manual, ClientLeft/Top/Width/Height would have been set above.

            Top = rawProps.GetInt32("Top", Top); // This is the Form.Top (screen coordinate)
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisButton = rawProps.GetEnum("WhatsThisButton", WhatsThisButton);
            WhatsThisHelp = rawProps.GetEnum("WhatsThisHelp", WhatsThisHelp);
            Width = rawProps.GetInt32("Width", Width); // This is the Form.Width
            WindowState = rawProps.GetEnum("WindowState", WindowState);
        }
    }
}