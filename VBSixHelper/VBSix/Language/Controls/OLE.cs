using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Language; // For VB6Color
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary
using System; // For StringComparison

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 OLE Container control.
    /// The OLE control is used to embed or link objects from other applications.
    /// </summary>
    public class OLEProperties
    {
        public Appearance Appearance { get; set; } = Controls.Appearance.Flat; // OLE control default is Flat (0)
        public AutoActivate AutoActivate { get; set; } = Controls.AutoActivate.DoubleClick;
        public bool AutoVerbMenu { get; set; } = true; // If True, context menu shows object's verbs
        public VB6Color BackColor { get; set; } = VB6Color.VbButtonFace; // Default: &H8000000F&
        public BackStyle BackStyle { get; set; } = Controls.BackStyle.Opaque;
        public BorderStyle BorderStyle { get; set; } = Controls.BorderStyle.FixedSingle;
        public CausesValidation CausesValidation { get; set; } = Controls.CausesValidation.Yes;
        
        /// <summary>
        /// The programmatic identifier (ProgID or ClassID) of the OLE object (e.g., "Excel.Sheet.8").
        /// This is set when an object is embedded or linked.
        /// </summary>
        public string? Class { get; set; } = null; 
        
        public string DataField { get; set; } = ""; // For data binding
        public string DataSource { get; set; } = ""; // For data binding
        public DisplayType DisplayType { get; set; } = Controls.DisplayType.Content; // Content or Icon
        // DragIcon, MouseIcon are FRX-based.
        // SourceDoc (if it's a file path for linking), ObjectData (for embedded objects),
        // ObjectVerbs, ObjectVerbFlags are typically stored in or referenced via .frx.
        // Their resolved data will be stored in the VB6OleKind class.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        // Font can be a group property.
        public int Height { get; set; } = 2055; // Typical default height in Twips
        public int HelpContextID { get; set; } = 0;
        public string HostName { get; set; } = ""; // Name of the application hosting the OLE control
        public int Left { get; set; } = 120;   // Typical default Left in Twips
        public int MiscFlags { get; set; } = 0; // OLE miscellaneous flags
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public bool OleDropAllowed { get; set; } = false; // If True, control can act as an OLE drop target itself
        public OLETypeAllowed OleTypeAllowed { get; set; } = Controls.OLETypeAllowed.Either; // Linked, Embedded, or Either
        public SizeMode SizeMode { get; set; } = Controls.SizeMode.Clip; // Clip, Stretch, AutoSize, Zoom
        
        /// <summary>
        /// Path to the source document for linked objects.
        /// </summary>
        public string SourceDoc { get; set; } = ""; 
        
        /// <summary>
        /// Specific item within the source document for linked objects (e.g., a range in Excel).
        /// </summary>
        public string SourceItem { get; set; } = ""; 
        
        public int TabIndex { get; set; } = 0; // Placeholder
        public TabStop TabStop { get; set; } = Controls.TabStop.Included;
        public int Top { get; set; } = 120;    // Typical default Top in Twips
        public UpdateOptions UpdateOptions { get; set; } = Controls.UpdateOptions.Automatic; // For linked objects
        
        /// <summary>
        /// The default verb to execute on AutoActivate. (0 is usually Primary, -1 Open, -2 Show, etc.)
        /// </summary>
        public int Verb { get; set; } = 0; 
        
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 2775;   // Typical default width in Twips

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for an OLE control.
        /// </summary>
        public OLEProperties() { }

        /// <summary>
        /// Initializes OLEProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public OLEProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            AutoActivate = rawProps.GetEnum("AutoActivate", AutoActivate);
            AutoVerbMenu = rawProps.GetBoolean("AutoVerbMenu", AutoVerbMenu);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            BackStyle = rawProps.GetEnum("BackStyle", BackStyle);
            BorderStyle = rawProps.GetEnum("BorderStyle", BorderStyle);
            CausesValidation = rawProps.GetEnum("CausesValidation", CausesValidation);
            Class = rawProps.GetOptionalString("Class"); // VB6 stores Class property directly
            DataField = rawProps.GetString("DataField", DataField);
            DataSource = rawProps.GetString("DataSource", DataSource);
            DisplayType = rawProps.GetEnum("DisplayType", DisplayType);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            // Font handled by PropertyGroup
            Height = rawProps.GetInt32("Height", Height);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            HostName = rawProps.GetString("HostName", HostName);
            Left = rawProps.GetInt32("Left", Left);
            MiscFlags = rawProps.GetInt32("MiscFlags", MiscFlags);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            OleDropAllowed = rawProps.GetBoolean("OLEDropAllowed", OleDropAllowed);
            OleTypeAllowed = rawProps.GetEnum("OLETypeAllowed", OleTypeAllowed);
            SizeMode = rawProps.GetEnum("SizeMode", SizeMode);
            
            // SourceDoc and SourceItem are tricky. They can be direct strings or FRX links.
            // The BuildControl logic needs to check if the PropertyValue for these is a resource.
            // For typed properties here, we assume they are the string values if not FRX.
            // The VB6OleKind will hold the resolved byte[] if they are FRX.
            SourceDoc = rawProps.GetString("SourceDoc", SourceDoc);
            SourceItem = rawProps.GetString("SourceItem", SourceItem);

            TabIndex = rawProps.GetInt32("TabIndex", TabIndex);
            TabStop = rawProps.GetEnum("TabStop", TabStop);
            Top = rawProps.GetInt32("Top", Top);
            UpdateOptions = rawProps.GetEnum("UpdateOptions", UpdateOptions);
            Verb = rawProps.GetInt32("Verb", Verb);
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);
            Width = rawProps.GetInt32("Width", Width);
        }
    }

    /// <summary>
    /// Represents information about a single OLE verb for an embedded or linked object.
    /// </summary>
    public class OleVerbInfo
    {
        /// <summary>
        /// The numeric identifier for the verb.
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// The display name of the verb (e.g., "Edit", "Open", "Play").
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Flags associated with the verb (e.g., OLEVERBATTRIB_NEVERDIRTIES, OLEVERBATTRIB_ONCONTAINERMENU).
        /// From OLEVERBATTRIB enum in OLE2.H.
        /// </summary>
        public int Flags { get; set; } 

        public override string ToString()
        {
            return $"Verb ID: {Id}, Name: \"{Name}\", Flags: 0x{Flags:X}";
        }
    }
}