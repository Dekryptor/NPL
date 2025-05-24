// Namespace: VB6Parse.Language.Controls
using System.Collections.Generic;

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 Menu item when defined as a control within a Form's control collection.
    /// It wraps the menu's properties and its hierarchy of sub-menus.
    /// </summary>
    public class MenuControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties of this menu item (like Caption, Checked, etc.).
        /// </summary>
        public MenuProperties MenuProperties { get; set; }

        /// <summary>
        /// Gets or sets the list of direct sub-menu items under this menu item.
        /// </summary>
        public List<Vb6MenuControl> SubMenus { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="MenuControl"/> class.
        /// </summary>
        /// <param name="name">The name of the menu control (e.g., "mnuFileOpen").</param>
        /// <param name="index">The index if part of a menu control array.</param>
        /// <param name="tag">The tag associated with the menu control.</param>
        /// <param name="properties">The properties of the menu item.</param>
        /// <param name="subMenus">The list of sub-menus.</param>
        public MenuControl(
            string name, 
            int index, 
            string tag,
            MenuProperties properties, 
            List<Vb6MenuControl> subMenus)
            : base(name, index, tag)
        {
            MenuProperties = properties ?? new MenuProperties();
            SubMenus = subMenus ?? new List<Vb6MenuControl>();
        }
    }
}