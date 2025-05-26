using VBSix.Language; // For VB6Color
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using VBSix.Errors;    // For VB6ErrorKind
using System.Collections.Generic; // For IDictionary
using System; // For StringComparison

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 Data control.
    /// The Data control provides access to databases.
    /// </summary>
    public class DataProperties
    {
        public Align Align { get; set; } = Controls.Align.None;
        public Appearance Appearance { get; set; } = Controls.Appearance.ThreeD;
        public VB6Color BackColor { get; set; } = VB6Color.VbWindowBackground; // Default: &H80000005& (System Color: Window Background)
        public BOFAction BofAction { get; set; } = BOFAction.MoveFirst;
        public string Caption { get; set; } = "Data1"; // Default caption can vary (e.g., "Data1")
        public Connection ConnectionType { get; set; } = Controls.Connection.Access; // VB6 Property: Connect
        public string DatabaseName { get; set; } = ""; // Path to the .mdb or other database file
        public DefaultCursorType DefaultCursorType { get; set; } = DefaultCursorType.Default;
        public DefaultType DefaultType { get; set; } = Controls.DefaultType.UseJet; // Jet or ODBCDirect
        // DragIcon, MouseIcon are FRX-based.
        // Their resolved byte[] data will be stored in the VB6DataKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        public EOFAction EofAction { get; set; } = EOFAction.MoveLast;
        public bool Exclusive { get; set; } = false; // Open database in exclusive mode
        // Font is a group property.
        public VB6Color ForeColor { get; set; } = VB6Color.VbWindowText; // Default: &H80000008& (System Color: Window Text)
        public int Height { get; set; } = 345; // Typical default height in Twips
        public int Left { get; set; } = 240;  // Typical default Left in Twips
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public bool Negotiate { get; set; } = false; // If control participates in OLE container negotiation
        public OLEDropMode OleDropMode { get; set; } = Controls.OLEDropMode.None;
        public int Options { get; set; } = 0; // Bitmask for various options (e.g., dbDenyWrite, dbDenyRead)
        public bool ReadOnly { get; set; } = false; // Open recordset as read-only
        public RecordSetType RecordSetType { get; set; } = Controls.RecordSetType.Dynaset; // VB6 Property: RecordsetType
        public string RecordSource { get; set; } = ""; // SQL query, table name, or query name
        public TextDirection RightToLeft { get; set; } = Controls.TextDirection.LeftToRight;
        public string ToolTipText { get; set; } = "";
        public int Top { get; set; } = 240; // Typical default Top in Twips
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 2295; // Typical default width in Twips

        /// <summary>
        /// Default constructor initializing with common VB6 defaults.
        /// </summary>
        public DataProperties() { }

        /// <summary>
        /// Initializes DataProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public DataProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Align = rawProps.GetEnum("Align", Align);
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            BofAction = rawProps.GetEnum("BOFAction", BofAction); // VB6 Property name
            Caption = rawProps.GetString("Caption", Caption);
            
            // The "Connect" property in .frm maps to ConnectionType.
            // Its value is a string like "Access", "dBase III;", etc.
            string connectString = rawProps.GetString("Connect", "Access"); // Default to "Access" if not found
            ConnectionType = PropertyParserHelpers.ParseConnectionString(connectString) ?? Controls.Connection.Access;

            DatabaseName = rawProps.GetString("DatabaseName", DatabaseName);
            DefaultCursorType = rawProps.GetEnum("DefaultCursorType", DefaultCursorType);
            DefaultType = rawProps.GetEnum("DefaultType", DefaultType);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            EofAction = rawProps.GetEnum("EOFAction", EofAction); // VB6 Property name
            Exclusive = rawProps.GetBoolean("Exclusive", Exclusive);
            // Font handled by PropertyGroup
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor);
            Height = rawProps.GetInt32("Height", Height);
            Left = rawProps.GetInt32("Left", Left);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            Negotiate = rawProps.GetBoolean("Negotiate", Negotiate);
            OleDropMode = rawProps.GetEnum("OLEDropMode", OleDropMode);
            Options = rawProps.GetInt32("Options", Options);
            ReadOnly = rawProps.GetBoolean("ReadOnly", ReadOnly);
            RecordSetType = rawProps.GetEnum("RecordsetType", RecordSetType); // VB6 Property: RecordsetType
            RecordSource = rawProps.GetString("RecordSource", RecordSource);
            RightToLeft = rawProps.GetEnum("RightToLeft", RightToLeft);
            ToolTipText = rawProps.GetString("ToolTipText", ToolTipText);
            Top = rawProps.GetInt32("Top", Top);
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);
            Width = rawProps.GetInt32("Width", Width);
        }
    }
}