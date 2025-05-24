// Namespace: VB6Parse.Utilities (or VB6Parse.Parsers.Helpers)
using System;
using System.Collections.Generic;
using System.Globalization;
using VB6Parse.Language; // For Vb6Color
using VB6Parse.Language.Controls; // For Font, StartUpPositionMode, and other enums
using VB6Parse.Model; // For Vb6PropertyGroup (if Font is parsed from it here)

namespace VB6Parse.Utilities
{
    /// <summary>
    /// Provides static helper methods for parsing property values from their
    /// string representations in VB6 form files.
    /// </summary>
    public static class PropertyParsingHelpers
    {
        /// <summary>
        /// Gets a string value from the properties dictionary.
        /// </summary>
        public static string GetString(IReadOnlyDictionary<string, string> props, string key, string defaultValue = "") =>
            props.TryGetValue(key, out var val) ? val : defaultValue;

        /// <summary>
        /// Gets an integer (i32) value from the properties dictionary.
        /// </summary>
        public static int GetInt32(IReadOnlyDictionary<string, string> props, string key, int defaultValue)
        {
            if (props.TryGetValue(key, out string valStr) && 
                int.TryParse(valStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result))
            {
                return result;
            }
            return defaultValue;
        }
		
		public static int GetInt(IReadOnlyDictionary<string, string> props, string key, int defaultValue = 0) =>
            props.TryGetValue(key, out var val) && int.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out int result) ? result : defaultValue;

        /// <summary>
        /// Gets a single-precision float (float) value from the properties dictionary.
        /// VB6 uses '.' as the decimal separator.
        /// </summary>
        public static float GetSingle(IReadOnlyDictionary<string, string> props, string key, float defaultValue)
        {
            if (props.TryGetValue(key, out string valStr) && 
                float.TryParse(valStr, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out float result))
            {
                return result;
            }
            return defaultValue;
        }
		
		public static float GetFloat(IReadOnlyDictionary<string, string> props, string key, float defaultValue = 0f) =>
            props.TryGetValue(key, out var val) && float.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out float result) ? result : defaultValue;
        
        /// <summary>
        /// Gets a boolean value from the properties dictionary.
        /// VB6 typically uses 0 for False and -1 (or sometimes 1) for True.
        /// </summary>
        public static bool GetBool(IReadOnlyDictionary<string, string> props, string key, bool defaultValue = false)
        {
            if (props.TryGetValue(key, out var val))
            {
                if (int.TryParse(val, out int intVal)) return intVal != 0; // VB uses 0 for False, -1 (or any non-zero) for True
                if (bool.TryParse(val, out bool boolVal)) return boolVal; // For "True" / "False" strings (less common in .frm)
                // VB often uses "0 'False" or "-1 'True"
                if (val.StartsWith("0")) return false;
                if (val.StartsWith("-1")) return true;
            }
            return defaultValue;
        }

        /// <summary>
        /// Gets a Vb6Color value from the properties dictionary.
        /// Color values are typically hex strings like "&H00FFFFFF&".
        /// </summary>
        public static Color GetColor(IReadOnlyDictionary<string, string> props, string key, Color defaultValue)
        {
            if (props.TryGetValue(key, out var valStr))
            {
                if (string.IsNullOrWhiteSpace(valStr)) return defaultValue;

                if (valStr.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
                {
                    string hexVal = valStr.Substring(2);
                    if (hexVal.EndsWith("&"))
                    {
                        hexVal = hexVal.Substring(0, hexVal.Length - 1);
                    }

                    if (uint.TryParse(hexVal, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint oleColor))
                    {
                        if ((oleColor & 0x80000000) != 0)
                        {
                            KnownColor knownColor;
                            switch ((int)(oleColor & 0xFF)) 
                            {
                                case 0: knownColor = KnownColor.ScrollBar; break;
                                case 1: knownColor = KnownColor.Desktop; break; 
                                case 2: knownColor = KnownColor.ActiveCaption; break;
                                case 3: knownColor = KnownColor.InactiveCaption; break;
                                case 4: knownColor = KnownColor.Menu; break;
                                case 5: knownColor = KnownColor.Window; break;
                                case 6: knownColor = KnownColor.WindowFrame; break;
                                case 7: knownColor = KnownColor.MenuText; break;
                                case 8: knownColor = KnownColor.WindowText; break;
                                case 9: knownColor = KnownColor.ActiveCaptionText; break; 
                                case 10: knownColor = KnownColor.ActiveBorder; break;
                                case 11: knownColor = KnownColor.InactiveBorder; break;
                                case 12: knownColor = KnownColor.AppWorkspace; break;
                                case 13: knownColor = KnownColor.Highlight; break;
                                case 14: knownColor = KnownColor.HighlightText; break;
                                case 15: knownColor = KnownColor.Control; break; 
                                case 16: knownColor = KnownColor.ControlDark; break; 
                                case 17: knownColor = KnownColor.GrayText; break;
                                case 18: knownColor = KnownColor.ControlText; break; 
                                case 19: knownColor = KnownColor.InactiveCaptionText; break;
                                case 20: knownColor = KnownColor.ControlLightLight; break; 
                                case 21: knownColor = KnownColor.ControlDarkDark; break;
                                case 22: knownColor = KnownColor.ControlLight; break;
                                case 23: knownColor = KnownColor.InfoText; break;
                                case 24: knownColor = KnownColor.Info; break; 
                                case 25: knownColor = KnownColor.ButtonFace; break; // Not a standard system color index but sometimes used.
                                case 26: knownColor = KnownColor.HotTrack; break; 
                                case 27: knownColor = KnownColor.GradientActiveCaption; break;
                                case 28: knownColor = KnownColor.GradientInactiveCaption; break;
                                case 29: knownColor = KnownColor.MenuHighlight; break;
                                case 30: knownColor = KnownColor.MenuBar; break;
                                default: return defaultValue; 
                            }
                            return Color.FromKnownColor(knownColor);
                        }
                        else
                        {
                            int r = (int)(oleColor & 0xFF);
                            int g = (int)((oleColor >> 8) & 0xFF);
                            int b = (int)((oleColor >> 16) & 0xFF);
                            return Color.FromArgb(r, g, b);
                        }
                    }
                }
                // Handle named colors if necessary, though .frm usually uses OLE_COLOR
                try { return Color.FromName(valStr); } catch { /* Fall through */ }
            }
            return defaultValue;
        }
		
		// Placeholder for Font parsing. You'll need a FontProperties class.
        // public static FontProperties GetFont(IReadOnlyList<Vb6PropertyGroup> propertyGroups, string fontGroupName = "Font")
        // {
        //     var fontGroup = propertyGroups.FirstOrDefault(pg => pg.Name.Equals(fontGroupName, StringComparison.OrdinalIgnoreCase));
        //     if (fontGroup != null)
        //     {
        //         var fontProps = new FontProperties();
        //         // fontProps.Name = GetString(fontGroup.GetTextualProperties(), "Name", "MS Sans Serif");
        //         // fontProps.Size = GetFloat(fontGroup.GetTextualProperties(), "Size", 8.25f);
        //         // ... and so on for Charset, Weight, Underline, Italic, Strikethrough
        //         return fontProps;
        //     }
        //     return null; // Or a default font
        // }

        /// <summary>
        /// Gets an enum value from the properties dictionary.
        /// Enum values in VB6 .frm files are stored as their underlying integer values.
        /// </summary>
        public static TEnum GetEnum<TEnum>(IReadOnlyDictionary<string, string> props, string key, TEnum defaultValue) where TEnum : struct, Enum
        {
            if (props.TryGetValue(key, out string valStr) && 
                int.TryParse(valStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out int intVal))
            {
                // Check if intVal is a defined value for TEnum or if TEnum can represent intVal.
                // This handles enums that might have non-contiguous values or negative values (like -1 for True).
                try
                {
                    object enumValue = Enum.ToObject(typeof(TEnum), intVal);
                    if (Enum.IsDefined(typeof(TEnum), enumValue))
                    {
                        return (TEnum)enumValue;
                    }
                }
                catch (ArgumentException)
                {
                    // intVal is not a valid underlying value for the enum type.
                }
            }
            return defaultValue;
        }

        /// <summary>
        /// Gets the StartUpPositionMode and associated client coordinates.
        /// </summary>
        public static StartUpPositionMode GetStartUpPosition(
            IReadOnlyDictionary<string, string> props,
            string key, // Key for the StartUpPosition property itself
            StartUpPositionMode defaultMode,
            out int clientLeft, out int clientTop, out int clientWidth, out int clientHeight)
        {
            // Get client coordinates which are always present, used if mode is Manual.
            // Provide sensible defaults if they are missing, though they should usually be there.
            clientLeft = GetInt32(props, "ClientLeft", 100); 
            clientTop = GetInt32(props, "ClientTop", 100);
            clientWidth = GetInt32(props, "ClientWidth", ConvertTwipsToPixels(4500)); // Default form width in twips
            clientHeight = GetInt32(props, "ClientHeight", ConvertTwipsToPixels(3000)); // Default form height in twips

            StartUpPositionMode mode = GetEnum(props, key, defaultMode);
            return mode;
        }

        /// <summary>
        /// Parses a Font object from a Vb6PropertyGroup (e.g., a "Font" property group).
        /// </summary>
        public static Font ParseFontFromGroup(Vb6PropertyGroup fontGroup, Font defaultFont = null)
        {
            if (fontGroup == null) return defaultFont ?? new Font();

            var font = new Font
            {
                Name = GetString(fontGroup.Properties, "Name", "MS Sans Serif"),
                Size = GetSingle(fontGroup.Properties, "Size", 8.25f),
                Bold = GetBool(fontGroup.Properties, "Bold", false),
                Italic = GetBool(fontGroup.Properties, "Italic", false),
                Underline = GetBool(fontGroup.Properties, "Underline", false),
                Strikethrough = GetBool(fontGroup.Properties, "Strikethrough", false),
                Weight = (short)GetInt32(fontGroup.Properties, "Weight", 400), // FW_NORMAL
                Charset = (short)GetInt32(fontGroup.Properties, "Charset", 0)   // ANSI_CHARSET
            };
            return font;
        }

        /// <summary>
        /// Parses a Font object from flat properties (e.g., "FontName", "FontSize").
        /// This is a fallback or alternative if Font isn't a BeginProperty block.
        /// VB6 usually stores Font as a group, but some custom controls might do it differently.
        /// </summary>
        public static Font ParseFontFromFlatProperties(IReadOnlyDictionary<string, string> props, string baseKeyName = "Font", Font defaultFont = null)
        {
             // Standard VB6 stores Font properties as a group. This is a fallback.
             // If properties are FontName, FontSize, etc.
            if (!props.ContainsKey(baseKeyName + "Name") && !props.ContainsKey("FontName")) // Check for actual properties
                 return defaultFont ?? new Font();


            var font = new Font
            {
                Name = GetString(props, baseKeyName + "Name", GetString(props, "FontName", "MS Sans Serif")),
                Size = GetSingle(props, baseKeyName + "Size", GetSingle(props, "FontSize", 8.25f)),
                Bold = GetBool(props, baseKeyName + "Bold", GetBool(props, "FontBold", false)),
                Italic = GetBool(props, baseKeyName + "Italic", GetBool(props, "FontItalic", false)),
                Underline = GetBool(props, baseKeyName + "Underline", GetBool(props, "FontUnderline", false)),
                Strikethrough = GetBool(props, baseKeyName + "Strikethrough", GetBool(props, "FontStrikethrough", false)),
                Weight = (short)GetInt32(props, baseKeyName + "Weight", GetInt32(props, "FontWeight", 400)),
                Charset = (short)GetInt32(props, baseKeyName + "Charset", GetInt32(props, "FontCharset", 0))
            };
            return font;
        }
        
        // Helper for default client sizes, assuming 15 twips per pixel for conversion if needed.
        // This is a rough estimate. ScaleMode should ideally be used.
        private static int ConvertTwipsToPixels(int twips) => (int)(twips / 15.0);


        // Extension method for convenience on Dictionaries
        public static TValue GetValueOrDefault<TKey, TValue>(this IReadOnlyDictionary<TKey, TValue> dictionary, TKey key, TValue defaultValue)
        {
            return dictionary.TryGetValue(key, out var value) ? value : defaultValue;
        }
    }
}