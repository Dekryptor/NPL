// Namespace: VB6Parse.Language.Controls
using System.Collections.Generic; // For List if we were to add sub-items here, but they are on Vb6MenuControl

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Contains properties specific to a VB6 Menu item.
    /// </summary>
    public class MenuProperties
    {
        /// <summary>
        /// Gets or sets the caption displayed for the menu item.
        /// Use an ampersand (&) to define an access key.
        /// </summary>
        public string Caption { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the menu item is checked.
        /// </summary>
        public bool Checked { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the menu item is enabled.
        /// </summary>
        public Activation Enabled { get; set; } // Uses common enum

        /// <summary>
        /// Gets or sets the Help context ID associated with the menu item.
        /// </summary>
        public int HelpContextId { get; set; }

        /// <summary>
        /// Gets or sets the position of the menu item on the menu bar when an
        /// OLE object on a form is active and displaying its menus.
        /// </summary>
        public NegotiatePosition NegotiatePosition { get; set; }

        /// <summary>
        /// Gets or sets the shortcut key for the menu item.
        /// Null if no shortcut is assigned.
        /// </summary>
        public ShortCut? Shortcut { get; set; } // Nullable enum

        /// <summary>
        /// Gets or sets a value indicating whether the menu item is visible.
        /// </summary>
        public Visibility Visible { get; set; } // Uses common enum

        /// <summary>
        /// Gets or sets a value indicating whether the menu item is a Windows List menu.
        /// This applies to MDI forms, where it displays a list of open MDI child windows.
        /// </summary>
        public bool WindowList { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="MenuProperties"/> class with default values.
        /// </summary>
        public MenuProperties()
        {
            Caption = string.Empty;
            Checked = false;
            Enabled = Activation.Enabled; // Default from Rust
            HelpContextId = 0;
            NegotiatePosition = NegotiatePosition.None; // Default from Rust
            Shortcut = null; // Corresponds to Option<ShortCut> being None
            Visible = Visibility.Visible; // Default from Rust
            WindowList = false;
        }
    }
}