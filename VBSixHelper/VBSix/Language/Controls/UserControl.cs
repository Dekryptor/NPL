using System.Text;
using VBSix.Parsers;
using VBSix.Utilities;

namespace VBSix.Language.Controls
{
    public class UserControlProperties
    {
        // Visual Properties for the UserControl surface
        public VB6Color BackColor { get; set; } = VB6Color.VbButtonFace;
        public Appearance Appearance { get; set; } = Appearance.ThreeD;
        public BorderStyle BorderStyle { get; set; } = BorderStyle.None;
        // Font is a group property

        // Sizing and Scaling
        public int Height { get; set; } = 1500; // Design-time surface height
        public int Width { get; set; } = 2000;  // Design-time surface width
        public ScaleMode ScaleMode { get; set; } = ScaleMode.Twip;
        public int ScaleHeight { get; set; } = 1500;
        public int ScaleWidth { get; set; } = 2000;

        // Behavior Properties
        public bool Alignable { get; set; } = false;
        public bool CanGetFocus { get; set; } = true;
        public bool ControlContainer { get; set; } = false; // Can this UC host other controls?
        public bool DefaultCancel { get; set; } = false; // Does it have Default/Cancel button behavior?
        public bool ForwardFocus { get; set; } = false;
        public bool InvisibleAtRuntime { get; set; } = false;
        public bool Public { get; set; } = true; // Whether it's creatable by other projects
        public string ToolboxBitmap_SourceString { get; set; } = ""; // FRX link to .ctx file usually
        public bool WindowLess { get; set; } = false;

        // Standard Control properties
        public Activation Enabled { get; set; } = Activation.Enabled;
        public TabStop TabStop { get; set; } = TabStop.Included; // For the UC itself
        public Visibility Visible { get; set; } = Visibility.Visible; // For the UC itself
        public DragMode DragMode { get; set; } = DragMode.Manual;
        public MousePointer MousePointer { get; set; } = MousePointer.Default;
        public OLEDropMode OleDropMode { get; set; } = OLEDropMode.None;
        public string ToolTipText { get; set; } = "";
        public int HelpContextID { get; set; } = 0;
        public int WhatsThisHelpID { get; set; } = 0;  // UserControls don't show What's This button

        // ActiveX Internal Properties
        public int _ExtentX { get; set; }
        public int _ExtentY { get; set; }
        public int _Version { get; set; }

        // Client area dimensions (often saved for UserControls, similar to Forms)
        public int ClientHeight { get; set; } = 1500; // Should ideally match ScaleHeight if no border/scroll
        public int ClientLeft { get; set; } = 0;      // Usually 0
        public int ClientTop { get; set; } = 0;       // Usually 0
        public int ClientWidth { get; set; } = 2000;  // Should ideally match ScaleWidth

        public UserControlProperties() { }

        public UserControlProperties(IDictionary<string, PropertyValue> rawProps)
        {
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            BorderStyle = rawProps.GetEnum("BorderStyle", BorderStyle);
            Height = rawProps.GetInt32("Height", Height);
            Width = rawProps.GetInt32("Width", Width);
            ScaleMode = rawProps.GetEnum("ScaleMode", ScaleMode);
            ScaleHeight = rawProps.GetInt32("ScaleHeight", ScaleHeight); // Often matches Height
            ScaleWidth = rawProps.GetInt32("ScaleWidth", ScaleWidth);   // Often matches Width

            Alignable = rawProps.GetBoolean("Alignable", Alignable);
            CanGetFocus = rawProps.GetBoolean("CanGetFocus", CanGetFocus);
            ControlContainer = rawProps.GetBoolean("ControlContainer", ControlContainer);
            DefaultCancel = rawProps.GetBoolean("DefaultCancel", DefaultCancel);
            ForwardFocus = rawProps.GetBoolean("ForwardFocus", ForwardFocus);
            InvisibleAtRuntime = rawProps.GetBoolean("InvisibleAtRuntime", InvisibleAtRuntime);
            Public = rawProps.GetBoolean("Public", Public); // Important for how it's used
            WindowLess = rawProps.GetBoolean("WindowLess", WindowLess);

            if (rawProps.TryGetValue("ToolboxBitmap", out var tbPropVal) && !tbPropVal.IsResource)
            {
                ToolboxBitmap_SourceString = tbPropVal.AsString() ?? "";
            }

            Enabled = rawProps.GetEnum("Enabled", Enabled);
            TabStop = rawProps.GetEnum("TabStop", TabStop);
            Visible = rawProps.GetEnum("Visible", Visible); // UserControl's own visibility
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            ToolTipText = rawProps.GetString("ToolTipText", ToolTipText);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);

            _ExtentX = rawProps.GetInt32("_ExtentX", _ExtentX);
            _ExtentY = rawProps.GetInt32("_ExtentY", _ExtentY);
            _Version = rawProps.GetInt32("_Version", _Version);
            // Font properties are handled via PropertyGroups

            ClientHeight = rawProps.GetInt32("ClientHeight", ClientHeight);
            ClientLeft = rawProps.GetInt32("ClientLeft", ClientLeft);
            ClientTop = rawProps.GetInt32("ClientTop", ClientTop);
            ClientWidth = rawProps.GetInt32("ClientWidth", ClientWidth);

            // Ensure ScaleHeight/Width are also read AFTER ClientHeight/Width if they can differ
            ScaleHeight = rawProps.GetInt32("ScaleHeight", ClientHeight); // Default ScaleHeight to ClientHeight
            ScaleWidth = rawProps.GetInt32("ScaleWidth", ClientWidth);   // Default ScaleWidth to ClientWidth
        }
    }
}