// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Base class for all VB6 controls, containing common properties.
    /// </summary>
    public abstract class Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the name of the control.
        /// This is the identifier used to refer to the control in code.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the tag string associated with the control.
        /// This property can store extra data about the control.
        /// </summary>
        public string Tag { get; set; }

        /// <summary>
        /// Gets or sets the index of the control if it is part of a control array.
        /// For controls not in an array, this value might be non-applicable or a default (e.g., -1),
        /// but VB6 form files often list an index.
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Vb6ControlBase"/> class.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        protected Vb6ControlBase(string name, int index, string tag = "")
        {
            Name = name;
            Index = index;
            Tag = tag ?? string.Empty;
        }

        // We might add more common properties or abstract methods here later
        // if they are identified as universally applicable. For example,
        // methods related to serialization or common property access.

        /// <summary>
        /// Returns a string that represents the current control.
        /// </summary>
        /// <returns>A string representation of the control.</returns>
        public override string ToString()
        {
            return $"{GetType().Name} '{Name}' (Index: {Index})";
        }
    }
}