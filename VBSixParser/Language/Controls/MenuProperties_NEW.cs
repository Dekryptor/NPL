// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class MenuProperties : ControlSpecificPropertiesBase // Menus don't usually have their own Font property in VB6
    {
        public string Caption { get; set; } = "Menu"; // Default "Menu1", or actual caption
        public CheckedConstants Checked { get; set; } = CheckedConstants.False;
        public Activation Enabled { get; set; } = Activation.Enabled;
        public string HelpContextID { get; set; } = "0";
        public int Index { get; set; } = -1; // Contextual, for menu control arrays
        public bool NegotiatePosition { get; set; } = false; // For MDI applications, 0=None, 1=Left, 2=Middle, 3=Right
                                                             // Stored as an integer, needs mapping or an enum
        public string Shortcut { get; set; } = string.Empty; // e.g., "Ctrl+A", stored as text or an enum for keys
        public string Tag { get; set; } = string.Empty;
        public Visibility Visible { get; set; } = Visibility.Visible;
        public bool WindowList { get; set; } = false; // For MDI parent forms

        // Note: Menus in VB6 do not have their own Font, BackColor, ForeColor properties.
        // They inherit appearance from the system or container.
        // So, PopulateFontProperty from base might not be strictly needed or could be overridden to do nothing.

        public MenuProperties()
        {
            // Default Caption is often like "mnuFileOpen" or user-defined.
            // The parser will set the actual name/caption.
            Caption = "Menu"; 
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            Caption = PropertyParsingHelpers.GetString(textualProps, "Caption", this.Caption);
            Checked = PropertyParsingHelpers.GetEnum(textualProps, "Checked", this.Checked);
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", this.Enabled);
            HelpContextID = PropertyParsingHelpers.GetString(textualProps, "HelpContextID", this.HelpContextID);
            
            // NegotiatePosition is an integer in .frm (0-3 typically)
            // We can store it as int or map to a more descriptive enum.
            // For now, let's assume it's stored as an int if we don't have a specific enum yet.
            // If an enum NegotiatePositionConstants exists (e.g., None=0, Left=1, Middle=2, Right=3):
            // NegotiatePosition = PropertyParsingHelpers.GetEnum(textualProps, "NegotiatePosition", DefaultNegotiatePositionEnum);
            // As a boolean, it's often just whether it *can* negotiate, actual position is different.
            // The VB6 IDE shows it as an integer value. Let's treat it as bool for simplicity of "does it participate".
            // The actual property "NegotiatePosition" is not a boolean in VB, it's an integer (0-3).
            // Let's define a field for it as int or an enum if we create one.
            // For now, I'll treat the boolean 'NegotiatePosition' as a simplification.
            // The actual property is an integer. Let's assume it's simplified to a boolean for now.
            NegotiatePosition = PropertyParsingHelpers.GetBool(textualProps, "NegotiatePosition", false); // This might be an oversimplification.
                                                                                                       // A more accurate representation would be an enum or int.

            // Shortcut is complex, it's an enum value (integer)
            // e.g. None=0, CtrlA=257. For now, storing as string from FRM is simpler if it's text.
            // If it's an integer: int shortcutVal = PropertyParsingHelpers.GetInt32(textualProps, "Shortcut", 0);
            // Then map shortcutVal to a Keys enum or string representation.
            // VB stores it as an integer. The text like "Ctrl+A" is IDE representation.
            Shortcut = PropertyParsingHelpers.GetString(textualProps, "Shortcut", this.Shortcut); // This will get the integer as a string
            // To make it more useful, we'd parse this integer string into a System.Windows.Forms.Keys or custom enum.
            // Example: if (int.TryParse(Shortcut, out int scVal)) { MappedShortcut = (Keys)scVal; }


            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            Visible = PropertyParsingHelpers.GetEnum(textualProps, "Visible", this.Visible);
            WindowList = PropertyParsingHelpers.GetBool(textualProps, "WindowList", this.WindowList);

            // Menus typically don't have their own Font property group or binary properties directly.
            // So, no call to PopulateFontProperty or handling of binaryProps here unless custom menus do.
        }

        // Override if base class Font property is not applicable or handled differently
        protected override void PopulateFontProperty(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            // Menus in VB6 generally do not have their own Font property.
            // They use system fonts or the container's font.
            // So, this method can be a no-op for MenuProperties.
            this.Font = null; // Or some system default representation if needed.
        }
    }

    // Supporting Enums (if not already defined elsewhere)
    public enum CheckedConstants
    {
        True = -1, // VB True
        False = 0
    }

    // NegotiatePosition enum (example, actual values might vary or be specific to MDI)
    public enum MenuNegotiatePositionConstants
    {
        None = 0,
        Left = 1,
        Middle = 2,
        Right = 3
    }
}