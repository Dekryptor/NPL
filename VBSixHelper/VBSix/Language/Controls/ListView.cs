using System.Text;
using VBSix.Parsers;
using VBSix.Utilities;

namespace VBSix.Language.Controls
{
    // ListView-specific Enums (many of these will map to MSComctlLib constants)
    public enum ListViewView
    {
        lvwIcon = 0,
        lvwSmallIcon = 1,
        lvwList = 2,
        lvwReport = 3 // Default
    }

    public enum ListViewLabelEdit
    {
        lvwAutomatic = 0, // Default
        lvwManual = 1
    }

    public enum ListViewArrange
    {
        lvwNone = 0, // Default
        lvwAutoLeft = 1,
        lvwAutoTop = 2
    }

    // Other enums like SortOrder (lvwAscending/lvwDescending), ListSortOrder, ColumnHeaderIcon, etc.
    // would be added as needed if those properties are commonly found/parsed from .frm.

    public class ListViewProperties
    {
        // Standard Properties
        public int Height { get; set; } = 2535;
        public int Width { get; set; } = 3735;
        public int Left { get; set; } = 240;
        public int Top { get; set; } = 240;
        public Appearance Appearance { get; set; } = Appearance.ThreeD;
        public BorderStyle BorderStyle { get; set; } = BorderStyle.FixedSingle;
        public Activation Enabled { get; set; } = Activation.Enabled;
        public TabStop TabStop { get; set; } = TabStop.Included;
        public int TabIndex { get; set; } = 0;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public DragMode DragMode { get; set; } = DragMode.Manual;
        public MousePointer MousePointer { get; set; } = MousePointer.Default;
        public string ToolTipText { get; set; } = "";
        public int HelpContextID { get; set; } = 0;
        public int WhatsThisHelpID { get; set; } = 0;
        public VB6Color BackColor { get; set; } = VB6Color.VbWindowBackground;
        public VB6Color ForeColor { get; set; } = VB6Color.VbWindowText; // For text of items

        // ListView Specific Properties
        public ListViewView View { get; set; } = ListViewView.lvwReport;
        public ListViewLabelEdit LabelEdit { get; set; } = ListViewLabelEdit.lvwAutomatic;
        public bool MultiSelect { get; set; } = false;
        public bool HideColumnHeaders { get; set; } = false;
        public bool FullRowSelect { get; set; } = false;
        public bool GridLines { get; set; } = false;
        public bool HotTracking { get; set; } = false;
        public bool HoverSelection { get; set; } = false;
        public bool CheckBoxes { get; set; } = false;
        public bool FlatScrollBar { get; set; } = false;
        public ListViewArrange Arrange { get; set; } = ListViewArrange.lvwNone;
        public bool LabelWrap { get; set; } = true; // Default for Icon views
        public bool AllowColumnReorder { get; set; } = false;

        // ImageList properties - these store the NAME of the ImageList control on the same form.
        public string Icons_SourceString { get; set; } = ""; // Property name in .frm is "Icons"
        public string SmallIcons_SourceString { get; set; } = ""; // Property name in .frm is "SmallIcons"
        // ColumnHeaders are handled via a BeginProperty ColumnHeaders block

        // OLE Drag/Drop
        public OLEDragMode OLEDragMode { get; set; } = OLEDragMode.Manual;
        public OLEDropMode OLEDropMode { get; set; } = OLEDropMode.None;

        // Data Binding (less common for ListView items at design time)
        public string DataSource { get; set; } = "";
        public string DataMember { get; set; } = "";
        // DataField is usually per-column, not a direct ListView property for items.

        // ActiveX Internal Properties
        public int _ExtentX { get; set; }
        public int _ExtentY { get; set; }
        public int _Version { get; set; } // e.g., 393217 for MSComctlLib ListView v6

        public ListViewProperties() { }

        public ListViewProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Height = rawProps.GetInt32("Height", Height);
            Width = rawProps.GetInt32("Width", Width);
            Left = rawProps.GetInt32("Left", Left);
            Top = rawProps.GetInt32("Top", Top);
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            BorderStyle = rawProps.GetEnum("BorderStyle", BorderStyle);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            TabStop = rawProps.GetEnum("TabStop", TabStop);
            TabIndex = rawProps.GetInt32("TabIndex", TabIndex);
            Visible = rawProps.GetEnum("Visible", Visible);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            ToolTipText = rawProps.GetString("ToolTipText", ToolTipText);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor);

            View = rawProps.GetEnum("View", View);
            LabelEdit = rawProps.GetEnum("LabelEdit", LabelEdit);
            MultiSelect = rawProps.GetBoolean("MultiSelect", MultiSelect);
            HideColumnHeaders = rawProps.GetBoolean("HideColumnHeaders", HideColumnHeaders);
            FullRowSelect = rawProps.GetBoolean("FullRowSelect", FullRowSelect);
            GridLines = rawProps.GetBoolean("GridLines", GridLines);
            HotTracking = rawProps.GetBoolean("HotTracking", HotTracking);
            HoverSelection = rawProps.GetBoolean("HoverSelection", HoverSelection);
            CheckBoxes = rawProps.GetBoolean("CheckBoxes", CheckBoxes);
            FlatScrollBar = rawProps.GetBoolean("FlatScrollBar", FlatScrollBar);
            Arrange = rawProps.GetEnum("Arrange", Arrange);
            LabelWrap = rawProps.GetBoolean("LabelWrap", LabelWrap);
            AllowColumnReorder = rawProps.GetBoolean("AllowColumnReorder", AllowColumnReorder);

            Icons_SourceString = rawProps.GetString("Icons", Icons_SourceString);
            SmallIcons_SourceString = rawProps.GetString("SmallIcons", SmallIcons_SourceString);

            OLEDragMode = rawProps.GetEnum("OLEDragMode", OLEDragMode);
            OLEDropMode = rawProps.GetEnum("OLEDropMode", OLEDropMode);

            DataSource = rawProps.GetString("DataSource", DataSource);
            DataMember = rawProps.GetString("DataMember", DataMember);

            _ExtentX = rawProps.GetInt32("_ExtentX", _ExtentX);
            _ExtentY = rawProps.GetInt32("_ExtentY", _ExtentY);
            _Version = rawProps.GetInt32("_Version", _Version);

            // Font is handled as a PropertyGroup
            // ColumnHeaders is a PropertyGroup
            // ListItems are typically added at runtime or via specific data binding; design-time items are rare in .frm text
        }
    }
}