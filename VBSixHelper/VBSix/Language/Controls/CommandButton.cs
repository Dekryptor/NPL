using VBSix.Language;
using VBSix.Parsers;
using VBSix.Utilities;

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 CommandButton control.
    /// These properties are typically found in a .frm file.
    /// </summary>
    public class CommandButtonProperties
    {
        public Appearance Appearance { get; set; } = Controls.Appearance.ThreeD;
        public VB6Color BackColor { get; set; } = VB6Color.VbButtonFace; // Default: &H8000000F&
        public bool Cancel { get; set; } = false; // True if this button acts as a Cancel button (Esc key)
        public string Caption { get; set; } = "Command1"; // Default caption can vary
        public CausesValidation CausesValidation { get; set; } = Controls.CausesValidation.Yes;
        public bool Default { get; set; } = false; // True if this button is the Default button (Enter key)
        // DisabledPicture, DownPicture, DragIcon, MouseIcon, Picture are FRX-based.
        // Their resolved byte[] data will be stored in the VB6CommandButtonKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        // Font is a group property, handled by VB6PropertyGroup
        public VB6Color ForeColor { get; set; } = VB6Color.VbButtonText; // Typically not used if BackColor is ButtonFace
        public int Height { get; set; } = 375; // Typical default height in Twips
        public int HelpContextID { get; set; } = 0;
        public int Left { get; set; } = 120; // Typical default Left in Twips
        public VB6Color MaskColor { get; set; } = VB6Color.FromRGB(192, 192, 192); // Default: &H00C0C0C0& (Silver)
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None; // CommandButtons are not typical OLE drop targets
        public TextDirection RightToLeft { get; set; } = Controls.TextDirection.LeftToRight;
        public Style Style { get; set; } = Controls.Style.Standard; // Standard or Graphical (if Picture is used)
        public int TabIndex { get; set; } = 0; // Placeholder
        public TabStop TabStop { get; set; } = Controls.TabStop.Included;
        public string ToolTipText { get; set; } = "";
        public int Top { get; set; } = 120; // Typical default Top in Twips
        public UseMaskColor UseMaskColor { get; set; } = Controls.UseMaskColor.DoNotUseMaskColor;
        // Value property is runtime for CommandButton (True when clicked, then False) - not stored in .frm usually
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 1215; // Typical default width in Twips

        /// <summary>
        /// Default constructor initializing with common VB6 defaults.
        /// </summary>
        public CommandButtonProperties() { }

        /// <summary>
        /// Initializes CommandButtonProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public CommandButtonProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            Cancel = rawProps.GetBoolean("Cancel", Cancel);
            Caption = rawProps.GetString("Caption", Caption);
            CausesValidation = rawProps.GetEnum("CausesValidation", CausesValidation);
            Default = rawProps.GetBoolean("Default", Default);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            // Font handled by PropertyGroup
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor); // Often ignored visually if BackColor is system button face
            Height = rawProps.GetInt32("Height", Height);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            Left = rawProps.GetInt32("Left", Left);
            MaskColor = rawProps.GetVB6Color("MaskColor", MaskColor);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            RightToLeft = rawProps.GetEnum("RightToLeft", RightToLeft);
            Style = rawProps.GetEnum("Style", Style);
            TabIndex = rawProps.GetInt32("TabIndex", TabIndex);
            TabStop = rawProps.GetEnum("TabStop", TabStop);
            ToolTipText = rawProps.GetString("ToolTipText", ToolTipText);
            Top = rawProps.GetInt32("Top", Top);
            UseMaskColor = rawProps.GetEnum("UseMaskColor", UseMaskColor);
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);
            Width = rawProps.GetInt32("Width", Width);
        }
    }
}