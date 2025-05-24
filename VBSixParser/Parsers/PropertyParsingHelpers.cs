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
        public static string GetString(IReadOnlyDictionary<string, string> props, string key, string defaultValue = "")
        {
            return props.TryGetValue(key, out string valStr) ? valStr : defaultValue;
        }

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
        
        /// <summary>
        /// Gets a boolean value from the properties dictionary.
        /// VB6 typically uses 0 for False and -1 (or sometimes 1) for True.
        /// </summary>
        public static bool GetBool(IReadOnlyDictionary<string, string> props, string key, bool defaultValue)
        {
            if (props.TryGetValue(key, out string valStr))
            {
                if (valStr == "0") return false;
                // VB6 True is -1. Some controls might save True as 1.
                if (valStr == "-1" || valStr == "1") return true; 
            }
            return defaultValue;
        }

        /// <summary>
        /// Gets a Vb6Color value from the properties dictionary.
        /// Color values are typically hex strings like "&H00FFFFFF&".
        /// </summary>
        public static Vb6Color GetColor(IReadOnlyDictionary<string, string> props, string key, Vb6Color defaultColor)
        {
            if (props.TryGetValue(key, out string valStr))
            {
                if (Vb6Color.TryParseHex(valStr, out Vb6Color color))
                {
                    return color;
                }
            }
            return defaultColor;
        }

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