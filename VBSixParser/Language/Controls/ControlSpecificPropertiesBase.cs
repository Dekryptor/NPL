// Namespace: VB6Parse.Language.Controls
using System.Collections.Generic;
using System.Linq; // Required for FirstOrDefault
using VB6Parse.Model; // For Vb6PropertyGroup
using VB6Parse.Utilities; // For PropertyParsingHelpers

namespace VB6Parse.Language.Controls
{
    public abstract class ControlSpecificPropertiesBase
    {
        /// <summary>
        /// The font used for text display by the control.
        /// </summary>
        public Font Font { get; set; }

        /// <summary>
        /// Populates the strongly-typed properties of this instance from
        /// the raw dictionaries and property groups collected by the parser.
        /// </summary>
        /// <param name="textualProps">Dictionary of property names to their string values.</param>
        /// <param name="binaryProps">Dictionary of property names to their binary data (e.g., from FRX).</param>
        /// <param name="propertyGroups">List of parsed Vb6PropertyGroup objects (e.g., for Font).</param>
        public abstract void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups);

        /// <summary>
        /// Helper method to be used by derived classes to populate the Font property.
        /// It checks for a "Font" Vb6PropertyGroup first, then falls back to
        /// flat properties like "FontName", "FontSize" if the group is not found.
        /// </summary>
        protected virtual void PopulateFontProperty(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            var fontGroup = propertyGroups.FirstOrDefault(pg => pg.Name.Equals("Font", StringComparison.OrdinalIgnoreCase));
            if (fontGroup != null)
            {
                this.Font = PropertyParsingHelpers.ParseFontFromGroup(fontGroup);
            }
            else
            {
                // Fallback if Font is not in a group (less common for standard controls but good for robustness)
                this.Font = PropertyParsingHelpers.ParseFontFromFlatProperties(textualProps, "Font");
            }
             // Ensure Font is never null after population
            if (this.Font == null)
            {
                this.Font = new Font(); // Default Font object
            }
        }
    }
}