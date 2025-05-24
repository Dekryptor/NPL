// Namespace: VB6Parse.Language.Controls
using System.Collections.Generic;

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a generic VB6 control, encompassing common identification
    /// properties and a reference to its kind-specific properties.
    /// This class mirrors the Rust VB6Control struct and uses Vb6ControlKindEnum
    /// to discriminate the type of SpecificProperties.
    /// </summary>
    public class Vb6Control
    {
        /// <summary>
        /// The name of the control instance (e.g., "Command1", "Text1").
        /// This is the `Name` property of the control.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// The user-defined string data associated with the control.
        /// This is the `Tag` property of the control.
        /// </summary>
        public string Tag { get; set; }

        /// <summary>
        /// The index of the control if it is part of a control array.
        /// A value like -1 (or another agreed-upon sentinel) indicates it's not in an array.
        /// This is the `Index` property of the control.
        /// </summary>
        public int Index { get; set; } // VB6 uses short, but int is safer here.

        /// <summary>
        /// The specific kind (type) of this VB6 control (e.g., TextBox, Form).
        /// </summary>
        public Vb6ControlKindEnum ControlKind { get; set; }

        /// <summary>
        /// Holds the properties specific to the control type defined by <see cref="ControlKind"/>.
        /// This object should be cast to the corresponding properties class
        /// (e.g., <see cref="TextBoxProperties"/>, <see cref="FormProperties"/>)
        /// to access detailed attributes.
        /// </summary>
        public object SpecificProperties { get; set; }

        // --- Container-Specific Properties ---

        /// <summary>
        /// For container controls (Form, Frame, MDIForm, PictureBox), this list
        /// holds their child controls. It will be null or empty if the control
        /// is not a container or has no children.
        /// </summary>
        public List<Vb6Control> ChildControls { get; set; }

        // --- Form/MDIForm-Specific Properties ---
        /// <summary>
        /// For Form and MDIForm controls, this list holds their menu structure.
        /// It will be null or empty if the form has no menus.
        /// (Vb6MenuControl will be defined based on menus.rs)
        /// </summary>
        public List<object> Menus { get; set; } // Placeholder: List<Vb6MenuControl>

        // --- Custom Control-Specific Properties ---
        /// <summary>
        /// For Custom controls (ActiveX/OCX), this may hold additional property groupings
        /// if identified by the parser (e.g., property bags).
        /// (Vb6PropertyGroup will be defined based on parser needs)
        /// </summary>
        public List<object> PropertyGroups { get; set; } // Placeholder: List<Vb6PropertyGroup>

        /// <summary>
        /// Initializes a new instance of the <see cref="Vb6Control"/> class.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="kind">The specific kind of the control.</param>
        /// <param name="index">The control's index if part of an array; defaults to -1.</param>
        /// <param name="tag">The control's tag string; defaults to empty.</param>
        public Vb6Control(string name, Vb6ControlKindEnum kind, int index = -1, string tag = "")
        {
            Name = name;
            ControlKind = kind;
            Index = index;
            Tag = tag ?? string.Empty; // Ensure tag is not null

            // Initialize collections for controls that might use them.
            // The parser will populate these as needed.
            ChildControls = new List<Vb6Control>();
            Menus = new List<object>(); // Will be List<Vb6MenuControl>
            PropertyGroups = new List<object>(); // Will be List<Vb6PropertyGroup>
        }

        /// <summary>
        /// Helper method to safely get the typed specific properties.
        /// </summary>
        /// <typeparam name="T">The expected type of the SpecificProperties (e.g., TextBoxProperties).</typeparam>
        /// <returns>The specific properties cast to type T, or null if the cast is invalid or SpecificProperties is null.</returns>
        public T GetProperties<T>() where T : class
        {
            return SpecificProperties as T;
        }
    }
}