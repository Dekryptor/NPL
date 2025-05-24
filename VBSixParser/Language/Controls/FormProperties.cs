// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For common enums like Appearance, AutoRedraw, etc. they are in this namespace or CommonControlEnums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 Form or MDIForm control.
    /// </summary>
    public class FormProperties // Could be a base for MDIFormProperties if they share much
    {
        public Appearance Appearance { get; set; }
        public AutoRedraw AutoRedraw { get; set; }
        public Vb6Color BackColor { get; set; }
        public FormBorderStyle BorderStyle { get; set; } // Form-specific border style
		
        public string Caption { get; set; }
        public int ClientHeight { get; set; } // Internal height in ScaleMode units
        public int ClientLeft { get; set; }   // Internal left in ScaleMode units
        public int ClientTop { get; set; }    // Internal top in ScaleMode units
        public int ClientWidth { get; set; }  // Internal width in ScaleMode units
		
        public ClipControls ClipControls { get; set; }
        public ControlBox ControlBox { get; set; }
        public DrawMode DrawMode { get; set; }
        public DrawStyle DrawStyle { get; set; }
        public int DrawWidth { get; set; }
        public Activation Enabled { get; set; }
        public Vb6Color FillColor { get; set; }
        public FillStyle FillStyle { get; set; }
        public FontTransparency FontTransparent { get; set; } // Property name is FontTransparent in VB6
        public Vb6Color ForeColor { get; set; }
        public HasDeviceContext HasDC { get; set; } // If True, form has persistent hDC
        public int Height { get; set; } // Outer height in Twips
        public int HelpContextId { get; set; }
        public byte[] Icon { get; set; } // Placeholder for image data (form icon)
        public bool KeyPreview { get; set; } // True if form previews keyboard events
        public int Left { get; set; }   // Outer left in Twips
        public FormLinkMode LinkMode { get; set; } // Form-specific DDE link mode
        public string LinkTopic { get; set; }
        public MaxButton MaxButton { get; set; }
        public bool MDIChild { get; set; } // True if this is an MDI child form
        public MinButton MinButton { get; set; }
        public byte[] MouseIcon { get; set; } // Placeholder for image data (custom mouse cursor)
        public MousePointer MousePointer { get; set; }
        public Movability Moveable { get; set; } // Property name is Moveable in VB6
        public bool NegotiateMenus { get; set; } // For MDI forms primarily
        public OLEDropMode OLEDropMode { get; set; } // If form accepts OLE drop operations
        public byte[] Palette { get; set; } // Placeholder for image data (custom palette image)
        public PaletteMode PaletteMode { get; set; }
        public byte[] Picture { get; set; } // Placeholder for image data (form background picture)
        public TextDirection RightToLeft { get; set; }
        
		public float ScaleHeight { get; set; } // User-defined height for scaling
        public float ScaleLeft { get; set; }   // User-defined left for scaling
        public ScaleMode ScaleMode { get; set; }
        public float ScaleTop { get; set; }    // User-defined top for scaling
        public float ScaleWidth { get; set; }  // User-defined width for scaling
        
		public ShowInTaskbar ShowInTaskbar { get; set; }
        public StartUpPosition StartUpPosition { get; set; }
        public int Top { get; set; }     // Outer top in Twips
        public Visibility Visible { get; set; } // Property name is Visible in VB6
        public WhatsThisButton WhatsThisButton { get; set; }
        public WhatsThisHelp WhatsThisHelp { get; set; }
        public int Width { get; set; }   // Outer width in Twips
        public WindowState WindowState { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormProperties"/> class with default VB6 values.
        /// </summary>
        public FormProperties()
        {
            Appearance = Appearance.ThreeD;
            AutoRedraw = AutoRedraw.Manual; // VB False
            BackColor = Vb6Color.FromOleColor(0x8000000F); // System Button Face (often overridden by user)
            BorderStyle = FormBorderStyle.Sizable;
            Caption = "Form1"; // Default
            ClientHeight = 3150; // Example, depends on ScaleMode & initial size. Rust: 200
            ClientLeft = 0;      // Rust: 0
            ClientTop = 0;       // Rust: 0
            ClientWidth = 4680;  // Example. Rust: 300
            ClipControls = ClipControls.Included; // VB True
            ControlBox = ControlBox.Included; // VB True
            DrawMode = DrawMode.CopyPen;
            DrawStyle = DrawStyle.Solid;
            DrawWidth = 1;
            Enabled = Activation.Enabled; // VB True
            FillColor = Vb6Color.FromOleColor(0x00000000); // Black
            FillStyle = FillStyle.Transparent;
            FontTransparent = FontTransparency.Transparent; // VB True
            ForeColor = Vb6Color.FromOleColor(0x80000012); // System Button Text
            HasDC = HasDeviceContext.Yes; // VB True
            Height = 3600; // Example outer height. Rust: 240
            HelpContextId = 0;
            Icon = null;
            KeyPreview = false;
            Left = 0; // Example. Rust: 0
            LinkMode = FormLinkMode.None;
            LinkTopic = string.Empty;
            MaxButton = MaxButton.Included; // VB True
            MDIChild = false;
            MinButton = MinButton.Included; // VB True
            MouseIcon = null;
            MousePointer = MousePointer.Default;
            Moveable = Movability.Moveable; // VB True
            NegotiateMenus = true; // VB True
            OLEDropMode = OLEDropMode.None;
            Palette = null;
            PaletteMode = PaletteMode.HalfTone;
            Picture = null;
            RightToLeft = TextDirection.LeftToRight; // VB False
            ScaleHeight = 3150; // Typically matches ClientHeight initially if ScaleMode is Twips. Rust: 240
            ScaleLeft = 0;      // Rust: 0
            ScaleMode = ScaleMode.Twip;
            ScaleTop = 0;       // Rust: 0
            ScaleWidth = 4680;  // Typically matches ClientWidth initially if ScaleMode is Twips. Rust: 240
            ShowInTaskbar = ShowInTaskbar.Show; // VB True
            StartUpPosition = StartUpPosition.WindowsDefault;
            Top = 0;        // Example. Rust: 0
            Visible = Visibility.Visible; // VB True (usually set False at design then True in Form_Load)
            WhatsThisButton = WhatsThisButton.Excluded; // VB False
            WhatsThisHelp = WhatsThisHelp.F1Help; // VB False (property is WhatsThisHelp, enum is different)
                                                  // Rust used WhatsThisHelp::F1Help for WhatsThisHelp property, which is boolean in VB6.
                                                  // VB6 WhatsThisHelp is a boolean. If True, context-sensitive help is enabled.
                                                  // The WhatsThisHelpID is a separate numeric property.
                                                  // The enum WhatsThisHelp in Rust might be a misinterpretation or a combined concept.
                                                  // Let's assume WhatsThisHelp property is boolean.
                                                  // The Rust WhatsThisHelp enum was: { F1Help = 0, ContextSensitive = -1 }
                                                  // This looks like a boolean mapping. Let's make it bool.
            // Re-evaluating WhatsThisHelp:
            // VB6 Form has:
            // WhatsThisButton (Boolean) - Show the "?" button
            // WhatsThisHelp (Boolean) - Enable context-sensitive help mode (Shift+F1)
            // WhatsThisHelpID (Long) - For specific context ID.
            // The Rust `WhatsThisHelp` enum seems to map to the boolean `WhatsThisHelp` property.
            // Let's keep the WhatsThisHelp enum from CommonControlEnums for now if it represents the boolean state.
            // The Rust 'form.rs' uses the 'WhatsThisHelp' enum for the 'whats_this_help' field.
            // Our CommonControlEnums.WhatsThisHelp is: { F1Help = 0, ContextSensitive = -1 }
            // This matches a boolean: F1Help (False), ContextSensitive (True).
            this.WhatsThisHelp = Controls.WhatsThisHelp.F1Help; // Default is False in VB6

            Width = 4800;   // Example outer width. Rust: 240
            WindowState = WindowState.Normal;
        }
    }
	
	// Placeholder for Font class - this would be more complex
    public class Font
    {
        public string Name { get; set; }
        public float Size { get; set; } // VB6 uses Single
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikethrough { get; set; }
        public short Charset { get; set; } // VB6 uses Integer
        public short Weight { get; set; } // VB6 uses Integer

        public Font()
        {
            Name = "MS Sans Serif"; // Common VB6 default
            Size = 8.25f;
            Bold = false;
            Italic = false;
            Underline = false;
            Strikethrough = false;
            Charset = 0; // ANSI_CHARSET
            Weight = 400; // FW_NORMAL
        }
    }
}