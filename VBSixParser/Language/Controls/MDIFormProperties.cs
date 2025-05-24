// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For common enums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 MDIForm (Multiple Document Interface Form).
    /// An MDIForm acts as a container for MDI child forms.
    /// </summary>
    public class MDIFormProperties
    {
        public Appearance Appearance { get; set; }
        public bool AutoShowChildren { get; set; } // If True, MDI child forms are shown when loaded.
        public Vb6Color BackColor { get; set; } // Background color of the MDI client area.
        public string Caption { get; set; } // Text in the title bar.
        public Activation Enabled { get; set; } // If False, user interaction is disabled.
        // Note: Font object properties (Name, Size, Bold, etc.) for child form captions if not overridden.
        public int Height { get; set; } // Outer height in Twips.
        public int HelpContextId { get; set; }
        public byte[] Icon { get; set; } // Icon for the MDIForm title bar and taskbar. (FRX data)
        public int Left { get; set; }   // ScreenLeft in Twips.
        public FormLinkMode LinkMode { get; set; } // DDE link mode (usually None or Source).
        public string LinkTopic { get; set; } // DDE topic.
        // public byte[] MouseIcon { get; set; } // Custom mouse cursor (FRX data).
        public MousePointer MousePointer { get; set; }
        public Movability Moveable { get; set; } // If True, user can move the MDIForm.
        public bool NegotiateMenus { get; set; } // VB6 "NegotiateMenus" property (always True for MDIForm, not in Rust struct)
        public bool NegotiateToolbars { get; set; } // If True, child form toolbars can be on MDIForm.
        public OLEDropMode OLEDropMode { get; set; }
        public byte[] Picture { get; set; } // Background picture for the MDI client area (FRX data).
        public TextDirection RightToLeft { get; set; }
        public bool ScrollBars { get; set; } // If True, MDIForm has scrollbars for child windows.
        public StartUpPosition StartUpPosition { get; set; } // Initial position on screen.
        public int Top { get; set; }     // ScreenTop in Twips.
        public Visibility Visible { get; set; } // If True, MDIForm is visible at startup (if it's the startup object).
        public WhatsThisHelpMode WhatsThisHelp { get; set; } // Enables What's This Help button.
        public int Width { get; set; }   // Outer width in Twips.
        public WindowState WindowState { get; set; } // Initial state (Normal, Minimized, Maximized).

        public MDIFormProperties()
        {
            Appearance = Appearance.ThreeD; // VB6 default. Rust: Same.
            AutoShowChildren = true; // VB6 default. Rust: Same.
            BackColor = Vb6Color.FromOleColor(0x8000000C); // System Application Workspace color. Rust: Same.
            Caption = "MDIForm1"; // Default, often set by user. Rust default is empty.
            Enabled = Activation.Enabled; // VB6 default. Rust: Same.
            // Font typically MS Sans Serif, 8pt.
            Height = 3600; // Example default. Rust: Same.
            HelpContextId = 0;
            Icon = null;
            Left = 0; // Depends on StartUpPosition. Rust: Same.
            LinkMode = FormLinkMode.None; // VB6 default. Rust: Same.
            LinkTopic = string.Empty; // Typically "FormName" if LinkMode is Source. Rust: Same.
            // MouseIcon = null;
            MousePointer = MousePointer.Default; // VB6 default. Rust: Same.
            Moveable = Movability.Moveable; // VB6 default is True. Rust: Moveable (1).
            NegotiateMenus = true; // Always true for MDIForm, cannot be set to False.
            NegotiateToolbars = true; // VB6 default. Rust: Same.
            OLEDropMode = OLEDropMode.None; // VB6 default. Rust: Same.
            Picture = null;
            RightToLeft = TextDirection.LeftToRight; // Default is False. Rust: Same.
            ScrollBars = false; // VB6 MDIForm default is False. Rust is True. Let's use VB6 default.
            StartUpPosition = StartUpPosition.WindowsDefault; // VB6 default. Rust: Same.
            Top = 0; // Depends on StartUpPosition. Rust: Same.
            Visible = Visibility.Visible; // VB6 default is True (but only shown if startup object or explicitly). Rust: Same.
            WhatsThisHelp = WhatsThisHelpMode.HelpEnabled; // VB6 default is True. Rust: F1Help (1).
            Width = 4800;  // Example default. Rust: Same.
            WindowState = WindowState.Normal; // VB6 default. Rust: Same.
        }
    }
}