// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents properties for a custom control or a standard control not specifically handled.
    /// This class stores the raw properties as collected by the parser.
    /// </summary>
    public class CustomControlProperties : ControlSpecificPropertiesBase
    {
        /// <summary>
        /// The full type name of the custom control (e.g., "MSComctlLib.ListViewCtrl").
        /// </summary>
        public string ControlTypeFullName { get; }

        /// <summary>
        /// Dictionary of textual properties collected for this custom control.
        /// Keys are property names, values are their string representations.
        /// </summary>
        public IReadOnlyDictionary<string, string> TextualProperties { get; private set; }

        /// <summary>
        /// Dictionary of binary properties (e.g., from .frx files) collected for this custom control.
        /// Keys are property names, values are the raw byte arrays.
        /// </summary>
        public IReadOnlyDictionary<string, byte[]> BinaryProperties { get; private set; }

        /// <summary>
        /// List of Vb6PropertyGroup objects collected for this custom control.
        /// </summary>
        public IReadOnlyList<Vb6PropertyGroup> PropertyGroups { get; private set; }
        
        // Standard properties that custom controls might still commonly have.
        // These can be populated if found, otherwise they'll keep default values.
        public int Height { get; set; }
        public int Left { get; set; }
        public string Tag { get; set; } = string.Empty;
        public int TabIndex { get; set; } = 0; 
        public bool TabStop { get; set; } = true; 
        public string ToolTipText { get; set; } = string.Empty;
        public int Top { get; set; }
        public Visibility Visible { get; set; } = Visibility.Visible;
        public int Width { get; set; }

        public CustomControlProperties(string controlTypeFullName)
        {
            ControlTypeFullName = controlTypeFullName;
            TextualProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            BinaryProperties = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            PropertyGroups = new List<Vb6PropertyGroup>();
            Font = new Font(); // Initialize with a default Font object
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            // Store the raw collections
            this.TextualProperties = new Dictionary<string, string>(textualProps, StringComparer.OrdinalIgnoreCase);
            this.BinaryProperties = new Dictionary<string, byte[]>(binaryProps, StringComparer.OrdinalIgnoreCase);
            this.PropertyGroups = new List<Vb6PropertyGroup>(propertyGroups);

            // Attempt to populate common/standard properties that custom controls often share.
            // Use the stored TextualProperties for these lookups.
            PopulateFontProperty(this.TextualProperties, this.PropertyGroups);

            Height = PropertyParsingHelpers.GetInt32(this.TextualProperties, "Height", 0);
            Left = PropertyParsingHelpers.GetInt32(this.TextualProperties, "Left", 0);
            Tag = PropertyParsingHelpers.GetString(this.TextualProperties, "Tag", string.Empty);
            TabIndex = PropertyParsingHelpers.GetInt32(this.TextualProperties, "TabIndex", 0);
            TabStop = PropertyParsingHelpers.GetBool(this.TextualProperties, "TabStop", true);
            ToolTipText = PropertyParsingHelpers.GetString(this.TextualProperties, "ToolTipText", string.Empty);
            Top = PropertyParsingHelpers.GetInt32(this.TextualProperties, "Top", 0);
            Visible = PropertyParsingHelpers.GetEnum(this.TextualProperties, "Visible", Visibility.Visible);
            Width = PropertyParsingHelpers.GetInt32(this.TextualProperties, "Width", 0);
            
            // Users of CustomControlProperties will need to access TextualProperties, BinaryProperties,
            // and PropertyGroups directly to interpret control-specific values.
            // Example: string customProp = customControl.SpecificProperties.TextualProperties.GetValueOrDefault("MyCustomProp", "default");
        }
    }
}