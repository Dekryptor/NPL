using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary
using System; // For Enum.IsDefined

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 MDIForm control.
    /// An MDIForm is a window that acts as a container for MDI child forms.
    /// </summary>
    public class MDIFormProperties
    {
        public Appearance Appearance { get; set; } = Controls.Appearance.ThreeD;
        public bool AutoShowChildren { get; set; } = true; // If True, MDI child forms are shown when loaded
        public VB6Color BackColor { get; set; } = VB6Color.VbApplicationWorkspace; // Default: &H8000000C&
        public string Caption { get; set; } = "MDIForm1"; // Default caption can vary
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        // Font is a group property.
        public int Height { get; set; } = 7200; // Typical default height in Twips
        public int HelpContextID { get; set; } = 0;
        // Icon, MouseIcon, Picture are FRX-based.
        // Their resolved byte[] data will be stored in the VB6MDIFormKind class.
        public int Left { get; set; } = 0;   // Screen Left, often 0 or a small value
        public FormLinkMode LinkMode { get; set; } = FormLinkMode.None; // For DDE (less common on MDI forms)
        public string LinkTopic { get; set; } = ""; // For DDE
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public Movability Moveable { get; set; } = Controls.Movability.Moveable; // MDI Forms are typically moveable
        public bool NegotiateToolbars { get; set; } = true; // If True, child form toolbars can merge
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None; // MDI forms can be OLE drop targets
        public TextDirection RightToLeft { get; set; } = Controls.TextDirection.LeftToRight;
        public bool ScrollBars { get; set; } = true; // If True, MDI parent shows scrollbars if children exceed view
        public StartUpPositionType StartUpPosition { get; set; } = StartUpPositionType.WindowsDefault;
        public int Top { get; set; } = 0;    // Screen Top, often 0 or a small value
        public Visibility Visible { get; set; } = Controls.Visibility.Visible; // Default for new MDI forms
        public WhatsThisHelp WhatsThisHelp { get; set; } = Controls.WhatsThisHelp.F1Help; // MDI Forms don't have What's This button
        public int Width { get; set; } = 9600;   // Typical default width in Twips
        public WindowState WindowState { get; set; } = Controls.WindowState.Normal; // Can be Maximized by default

        // Runtime properties not stored in .frm typically:
        // ActiveControl, ActiveForm, MDIChildForms (collection)

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for an MDIForm control.
        /// </summary>
        public MDIFormProperties() { }

        /// <summary>
        /// Initializes MDIFormProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public MDIFormProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            AutoShowChildren = rawProps.GetBoolean("AutoShowChildren", AutoShowChildren);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            Caption = rawProps.GetString("Caption", Caption); // Often from Attribute VB_Name
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            // Font handled by PropertyGroup
            Height = rawProps.GetInt32("Height", Height);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            Left = rawProps.GetInt32("Left", Left);
            LinkMode = rawProps.GetEnum("LinkMode", LinkMode);
            LinkTopic = rawProps.GetString("LinkTopic", LinkTopic);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            Moveable = rawProps.GetEnum("Moveable", Moveable);
            NegotiateToolbars = rawProps.GetBoolean("NegotiateToolbars", NegotiateToolbars);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            RightToLeft = rawProps.GetEnum("RightToLeft", RightToLeft);
            ScrollBars = rawProps.GetBoolean("ScrollBars", ScrollBars); // VB6 property is "ScrollBars" (plural)

            int startUpPosIntValue = rawProps.GetInt32("StartUpPosition", (int)StartUpPositionType.WindowsDefault);
            if (Enum.IsDefined(typeof(StartUpPositionType), startUpPosIntValue))
            {
                StartUpPosition = (StartUpPositionType)startUpPosIntValue;
            }
            else
            {
                StartUpPosition = StartUpPositionType.WindowsDefault; // Fallback
            }
            // ClientLeft/Top/Width/Height are not directly set for MDIForm as they are for Form (Manual StartUpPosition)

            Top = rawProps.GetInt32("Top", Top);
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisHelp = rawProps.GetEnum("WhatsThisHelp", WhatsThisHelp);
            Width = rawProps.GetInt32("Width", Width);
            WindowState = rawProps.GetEnum("WindowState", WindowState);
        }
    }
}