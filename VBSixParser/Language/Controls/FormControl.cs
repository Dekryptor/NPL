// Namespace: VB6Parse.Language.Controls
using System.Collections.Generic;

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 Form, which is a top-level control container.
    /// It holds its own properties, a collection of child controls, and its menu structure.
    /// </summary>
    public class FormControl : Vb6ControlBase // A Form has a Name, but Index and Tag are less common for Forms themselves.
    {
        /// <summary>
        /// Gets or sets the specific properties for this Form.
        /// </summary>
        public FormProperties FormProperties { get; set; }

        /// <summary>
        /// Gets or sets the list of child controls contained within this Form.
        /// The order reflects the Z-order and .frm file definition order.
        /// </summary>
        public List<Vb6ControlBase> ChildControls { get; set; }

        /// <summary>
        /// Gets or sets the list of top-level menu items for this Form.
        /// Each Vb6MenuControl can contain its own sub-menus.
        /// </summary>
        public List<Vb6MenuControl> Menus { get; set; }
        
        // VB6 Forms also have a "Code Behind" module, which is not represented here yet.
        // This would be a separate string or object property holding the VB code.
        public string CodeModule { get; set; }


        /// <summary>
        /// Initializes a new instance of the <see cref="FormControl"/> class.
        /// </summary>
        /// <param name="name">The name of the Form (its identifier).</param>
        /// <param name="properties">The Form-specific properties.</param>
        /// <param name="childControls">The list of controls on the Form.</param>
        /// <param name="menus">The list of menus on the Form.</param>
        public FormControl(
            string name, 
            FormProperties properties, 
            List<Vb6ControlBase> childControls = null, 
            List<Vb6MenuControl> menus = null,
            string codeModule = "")
            : base(name, -1, "") // Forms typically don't have an Index or Tag in the same way other controls do.
        {
            FormProperties = properties ?? new FormProperties();
            ChildControls = childControls ?? new List<Vb6ControlBase>();
            Menus = menus ?? new List<Vb6MenuControl>();
            CodeModule = codeModule ?? string.Empty;
        }
    }
}