// Namespace: VB6Parse.Language.Controls
using System.Collections.Generic; // Required for List<Vb6ControlBase> (for Menus, though not for child forms here)

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 MDIForm (Multiple Document Interface Form).
    /// This is the main container window in an MDI application.
    /// It does not directly contain other controls in the same way a standard Form does (apart from Menus).
    /// Child forms are associated with it programmatically or by setting their MDIChild property.
    /// </summary>
    public class MDIFormControl : Vb6ControlBase // MDIForm itself is a type of control/component
    {
        /// <summary>
        /// Gets or sets the specific properties for this MDIForm.
        /// </summary>
        public MDIFormProperties MDIFormProperties { get; set; }

        // MDIForms can have their own menus.
        // Child controls (like Buttons, TextBoxes) are NOT placed directly on an MDIForm's surface
        // in the same way as a standard Form. They are placed on MDIChild forms or on toolbars/statusbars
        // that might be associated with the MDIForm.
        // For simplicity here, we might only consider menus as "contained" if following the general pattern.
        // However, the primary role is a container for MDIChild FORMS, not standard controls.
        // The list of MDIChild forms is a runtime concept or managed at the application level.

        /// <summary>
        /// Initializes a new instance of the <see cref="MDIFormControl"/> class.
        /// </summary>
        /// <param name="name">The name of the MDIForm (e.g., "MDIForm1").</param>
        /// <param name="index">Index is not typically used for MDIForms as there's only one.</param>
        /// <param name="tag">The tag string associated with the MDIForm.</param>
        /// <param name="properties">The MDIForm-specific properties.</param>
        public MDIFormControl(string name, int index, string tag, MDIFormProperties properties)
            : base(name, index, tag) // Name is key, Index is usually irrelevant.
        {
            MDIFormProperties = properties ?? new MDIFormProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MDIFormControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the MDIForm.</param>
        public MDIFormControl(string name)
            : this(name, -1, string.Empty, new MDIFormProperties())
        {
        }
    }
}