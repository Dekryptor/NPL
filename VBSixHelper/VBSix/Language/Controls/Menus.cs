using VBSix.Language; // For VB6Color (though not directly used by MenuProperties itself, Font group would)
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary, List
using System; // For StringComparison

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a single VB6 Menu item.
    /// These properties are typically found within a `Begin VB.Menu ... End` block in a .frm file.
    /// </summary>
    public class MenuProperties
    {
        public string Caption { get; set; } = "MenuItem"; // Default can vary (e.g., "mnuFile")
        public bool Checked { get; set; } = false; // If True, a checkmark appears next to the menu item
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        public int HelpContextID { get; set; } = 0;
        public NegotiatePosition NegotiatePosition { get; set; } = Controls.NegotiatePosition.None;
        public ShortCut? Shortcut { get; set; } = null; // Keyboard shortcut (e.g., Ctrl+O)
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public bool WindowList { get; set; } = false; // If True, MDI child window list is appended here

        // Name, Index, and Tag are properties of the VB6MenuControl itself, not within this properties bag.
        // Parent property (runtime) is also not stored here.

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for a Menu item.
        /// </summary>
        public MenuProperties() { }

        /// <summary>
        /// Initializes MenuProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public MenuProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Caption = rawProps.GetString("Caption", Caption);
            Checked = rawProps.GetBoolean("Checked", Checked);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            // VB6 .frm uses "NegotiatePosition"
            NegotiatePosition = rawProps.GetEnum("NegotiatePosition", NegotiatePosition); 
            Shortcut = rawProps.GetShortCut("Shortcut", Shortcut); // Uses helper for string->enum parsing
            Visible = rawProps.GetEnum("Visible", Visible);
            WindowList = rawProps.GetBoolean("WindowList", WindowList);
        }
    }

    /// <summary>
    /// Represents a VB6 menu control item.
    /// A menu item can have its own properties (like Caption, Checked, Enabled)
    /// and can also contain a list of sub-menu items, forming a hierarchical menu structure.
    /// which is distinct from the general `VB6Control` used for visual controls on a form.
    /// </summary>
    public class VB6MenuControl
    {
        /// <summary>
        /// The programmatic name of the menu item (e.g., "mnuFileOpen").
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// User-defined data associated with the menu item.
        /// </summary>
        public string Tag { get; set; } = string.Empty;

        /// <summary>
        /// If the menu item is part of a control array, this is its index. Otherwise, typically 0 or not set.
        /// Menu items in VB6 can form control arrays.
        /// </summary>
        public int Index { get; set; } // Default to 0 if not specified

        /// <summary>
        /// The specific properties of this menu item (Caption, Checked, Enabled, etc.).
        /// </summary>
        public MenuProperties TypedProperties { get; set; } = new MenuProperties();

        /// <summary>
        /// A list of sub-menu items parented by this menu item.
        /// </summary>
        public List<VB6MenuControl> SubMenus { get; set; } = new List<VB6MenuControl>();
        
        /// <summary>
        /// Property groups associated with this menu, like Font.
        /// While less common than for visual controls, menus can have Font properties.
        /// </summary>
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = new List<VB6PropertyGroup>();


        public VB6MenuControl() {}

        public VB6MenuControl(string name, IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propertyGroups)
        {
            Name = name;
            Tag = rawProps.GetString("Tag", string.Empty);
            Index = rawProps.GetInt32("Index", 0); // Default to 0 if not found
            TypedProperties = new MenuProperties(rawProps);
            PropertyGroups = propertyGroups ?? new List<VB6PropertyGroup>();
            // SubMenus are added recursively during parsing.
        }


        public override string ToString()
        {
            return $"Menu: {Name} (Caption: \"{TypedProperties.Caption}\", SubMenus: {SubMenus.Count})";
        }
    }
}