// Namespace: VB6Parse.Language.Controls
using System.Collections.Generic;

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 menu item, which can have its own properties and a list of sub-menu items.
    /// This corresponds to the VB6MenuControl in Rust, and is used for the main menu structure.
    /// A top-level menu item in a Form's menu bar would be a Vb6MenuControl,
    /// and its dropdown items would be in its SubMenus list.
    /// </summary>
    public class Vb6MenuControl // Not deriving from Vb6ControlBase directly, as it's a distinct structure for menus
    {
        /// <summary>
        /// Gets or sets the name of the menu control (e.g., "mnuFileOpen").
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the tag string associated with the menu control.
        /// </summary>
        public string Tag { get; set; }

        /// <summary>
        /// Gets or sets the index if this menu item is part of a control array.
        /// </summary>
        public int Index { get; set; } // VB6 allows menu control arrays

        /// <summary>
        /// Gets or sets the specific properties of this menu item.
        /// </summary>
        public MenuProperties Properties { get; set; }

        /// <summary>
        /// Gets or sets the list of sub-menu items under this menu item.
        /// </summary>
        public List<Vb6MenuControl> SubMenus { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Vb6MenuControl"/> class.
        /// </summary>
        public Vb6MenuControl(string name, MenuProperties properties, int index = -1, string tag = "") // Index -1 if not in array
        {
            Name = name;
            Properties = properties ?? new MenuProperties();
            Index = index;
            Tag = tag ?? string.Empty;
            SubMenus = new List<Vb6MenuControl>();
        }
    }
}