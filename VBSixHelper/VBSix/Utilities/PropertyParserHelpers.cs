using System.Globalization;
using System.Text;
using VBSix.Errors;
using VBSix.Language;
using VBSix.Language.Controls;
using VBSix.Parsers;

namespace VBSix.Utilities
{
    public static class PropertyParserHelpers
    {
        private static string? GetStringFromPropValue(PropertyValue? propValue)
        {
            if (propValue == null || propValue.IsResource) return null;
            return propValue.AsString();
        }

        /// <summary>
        /// Retrieves a string property value. If the key is not found or the value is a resource,
        /// returns the default value.
        /// </summary>
        public static string GetString(this IDictionary<string, PropertyValue> props, string key, string defaultValue)
        {
            props.TryGetValue(key, out var propValue);
            return GetStringFromPropValue(propValue) ?? defaultValue;
        }

        /// <summary>
        /// Retrieves an optional string property value. If the key is not found, the value is a resource,
        /// or the value is an empty string, returns the default value (which is null by default).
        /// </summary>
        public static string? GetOptionalString(this IDictionary<string, PropertyValue> props, string key, string? defaultValue = null)
        {
            props.TryGetValue(key, out var propValue);
            string? stringVal = GetStringFromPropValue(propValue);
            return string.IsNullOrEmpty(stringVal) ? defaultValue : stringVal;
        }

        /// <summary>
        /// Retrieves an integer property value.
        /// </summary>
        public static int GetInt32(this IDictionary<string, PropertyValue> props, string key, int defaultValue)
        {
            props.TryGetValue(key, out var propValue);
            string? stringValue = GetStringFromPropValue(propValue);
            if (stringValue != null && int.TryParse(stringValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result))
            {
                return result;
            }
            return defaultValue;
        }

        /// <summary>
        /// Retrieves a boolean property value. VB6 typically uses -1 for True and 0 for False.
        /// </summary>
        public static bool GetBoolean(this IDictionary<string, PropertyValue> props, string key, bool defaultValue)
        {
            props.TryGetValue(key, out var propValue);
            string? stringValue = GetStringFromPropValue(propValue);
            if (stringValue != null)
            {
                if (int.TryParse(stringValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int intResult))
                {
                    // In VB6: -1 is True, 0 is False. Also accept 1 as True.
                    if (intResult == -1 || intResult == 1) return true;
                    if (intResult == 0) return false;
                }
                // Fallback for "True" / "False" strings, though less common for .frm numeric properties
                if (bool.TryParse(stringValue, out bool boolResult))
                {
                    return boolResult;
                }
            }
            return defaultValue;
        }

        /// <summary>
        /// Retrieves a VB6Color property value.
        /// </summary>
        public static VB6Color GetVB6Color(this IDictionary<string, PropertyValue> props, string key, VB6Color defaultValue)
        {
            props.TryGetValue(key, out var propValue);
            string? stringValue = GetStringFromPropValue(propValue);
            if (stringValue != null)
            {
                try
                {
                    return VB6Color.FromHex(stringValue);
                }
                catch (VB6ParseException ex) when (ex.Kind == VB6ErrorKind.HexColorParseError)
                {
                    // Optionally log a warning here if parsing fails
                    //Console.Error.WriteLine($"Warning: Could not parse color '{stringValue}' for key '{key}'. Using default. Error: {ex.Message}");
                    return defaultValue;
                }
            }
            return defaultValue;
        }

        /// <summary>
        /// Retrieves an enum property value. Assumes the enum is stored as an integer in the .frm file.
        /// </summary>
        public static TEnum GetEnum<TEnum>(this IDictionary<string, PropertyValue> props, string key, TEnum defaultValue) where TEnum : struct, Enum
        {
            props.TryGetValue(key, out var propValue);
            string? stringValue = GetStringFromPropValue(propValue);
            if (stringValue != null)
            {
                if (int.TryParse(stringValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int intValue))
                {
                    // Check if the integer value is defined for the enum
                    if (Enum.IsDefined(typeof(TEnum), intValue))
                    {
                        return (TEnum)Enum.ToObject(typeof(TEnum), intValue);
                    }
                    // If not directly defined, but TEnum might be Flags or have non-contiguous values,
                    // a simple cast might be attempted, but IsDefined is safer for standard enums.
                    // For TryFromPrimitive behavior, if it's a direct cast:
                    try
                    {
                        return (TEnum)Convert.ChangeType(intValue, Enum.GetUnderlyingType(typeof(TEnum)));
                    }
                    catch (Exception)
                    {
                        // Conversion failed, fall through to default
                    }
                }
                // Fallback: Try parsing by enum member name (less common for .frm numeric properties but adds robustness)
                try
                {
                    return Enum.Parse<TEnum>(stringValue, ignoreCase: true);
                }
                catch (ArgumentException)
                {
                    // Value was not a valid member name
                }
            }
            return defaultValue;
        }

        /// <summary>
        /// Retrieves a StartUpPositionType enum value. This is specifically for the StartUpPosition property
        /// which is stored as an integer (0-3) in .frm files.
        /// </summary>
        public static StartUpPositionType GetStartUpPositionType(this IDictionary<string, PropertyValue> props, string key, StartUpPositionType defaultValue)
        {
            props.TryGetValue(key, out var propValue);
            string? stringValue = GetStringFromPropValue(propValue);
            if (stringValue != null && int.TryParse(stringValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int intValue))
            {
                if (Enum.IsDefined(typeof(StartUpPositionType), intValue))
                {
                    return (StartUpPositionType)intValue;
                }
            }
            return defaultValue;
        }
        
        /// <summary>
        /// Retrieves an optional char property value. Used for PasswordChar.
        /// </summary>
        public static char? GetOptionalChar(this IDictionary<string, PropertyValue> props, string key)
        {
            props.TryGetValue(key, out var propValue);
            string? stringValue = GetStringFromPropValue(propValue);
            if (!string.IsNullOrEmpty(stringValue))
            {
                // VB6 PasswordChar is stored as the character itself if set, or empty string if not.
                // The .frm file often stores it as "X" or "*".
                return stringValue[0];
            }
            return null;
        }

        /// <summary>
        /// Retrieves a ShortCut enum value from its string representation.
        /// </summary>
        public static ShortCut? GetShortCut(this IDictionary<string, PropertyValue> props, string key, ShortCut? defaultValue)
        {
            props.TryGetValue(key, out var propValue);
            string? stringValue = GetStringFromPropValue(propValue);
            if (stringValue != null)
            {
                return ParseShortCutString(stringValue) ?? defaultValue;
            }
            return defaultValue;
        }
        
        /// <summary>
        /// Retrieves a Connection enum value from its string representation.
        /// </summary>
        public static Connection GetConnection(this IDictionary<string, PropertyValue> props, string key, Connection defaultValue)
        {
            props.TryGetValue(key, out var propValue);
            string? stringValue = GetStringFromPropValue(propValue);
            if (stringValue != null)
            {
                return ParseConnectionString(stringValue) ?? defaultValue;
            }
            return defaultValue;
        }


        // --- Parsing helpers for specific enums that have string representations ---

        public static ShortCut? ParseShortCutString(string s)
        {
            if (string.IsNullOrEmpty(s)) return ShortCut.None; // Explicitly return None if empty, matches Option<ShortCut> default

            switch (s)
            {
                // Ctrl + Letter
                case "^A": return ShortCut.CtrlA; case "^B": return ShortCut.CtrlB; case "^C": return ShortCut.CtrlC;
                case "^D": return ShortCut.CtrlD; case "^E": return ShortCut.CtrlE; case "^F": return ShortCut.CtrlF;
                case "^G": return ShortCut.CtrlG; case "^H": return ShortCut.CtrlH; case "^I": return ShortCut.CtrlI;
                case "^J": return ShortCut.CtrlJ; case "^K": return ShortCut.CtrlK; case "^L": return ShortCut.CtrlL;
                case "^M": return ShortCut.CtrlM; case "^N": return ShortCut.CtrlN; case "^O": return ShortCut.CtrlO;
                case "^P": return ShortCut.CtrlP; case "^Q": return ShortCut.CtrlQ; case "^R": return ShortCut.CtrlR;
                case "^S": return ShortCut.CtrlS; case "^T": return ShortCut.CtrlT; case "^U": return ShortCut.CtrlU;
                case "^V": return ShortCut.CtrlV; case "^W": return ShortCut.CtrlW; case "^X": return ShortCut.CtrlX;
                case "^Y": return ShortCut.CtrlY; case "^Z": return ShortCut.CtrlZ;

                // Function Keys
                case "{F1}": return ShortCut.F1; case "{F2}": return ShortCut.F2; case "{F3}": return ShortCut.F3;
                case "{F4}": return ShortCut.F4; case "{F5}": return ShortCut.F5; case "{F6}": return ShortCut.F6;
                case "{F7}": return ShortCut.F7; case "{F8}": return ShortCut.F8; case "{F9}": return ShortCut.F9;
                // F10 is not a valid shortcut in VB6 menu editor
                case "{F11}": return ShortCut.F11; case "{F12}": return ShortCut.F12;

                // Ctrl + Function Keys
                case "^{F1}": return ShortCut.CtrlF1; case "^{F2}": return ShortCut.CtrlF2; case "^{F3}": return ShortCut.CtrlF3;
                case "^{F4}": return ShortCut.CtrlF4; case "^{F5}": return ShortCut.CtrlF5; case "^{F6}": return ShortCut.CtrlF6;
                case "^{F7}": return ShortCut.CtrlF7; case "^{F8}": return ShortCut.CtrlF8; case "^{F9}": return ShortCut.CtrlF9;
                case "^{F11}": return ShortCut.CtrlF11; case "^{F12}": return ShortCut.CtrlF12;

                // Shift + Function Keys
                case "+{F1}": return ShortCut.ShiftF1; case "+{F2}": return ShortCut.ShiftF2; case "+{F3}": return ShortCut.ShiftF3;
                case "+{F4}": return ShortCut.ShiftF4; case "+{F5}": return ShortCut.ShiftF5; case "+{F6}": return ShortCut.ShiftF6;
                case "+{F7}": return ShortCut.ShiftF7; case "+{F8}": return ShortCut.ShiftF8; case "+{F9}": return ShortCut.ShiftF9;
                case "+{F11}": return ShortCut.ShiftF11; case "+{F12}": return ShortCut.ShiftF12;

                // Shift + Ctrl + Function Keys
                case "+^{F1}": return ShortCut.ShiftCtrlF1; case "+^{F2}": return ShortCut.ShiftCtrlF2; case "+^{F3}": return ShortCut.ShiftCtrlF3;
                case "+^{F4}": return ShortCut.ShiftCtrlF4; case "+^{F5}": return ShortCut.ShiftCtrlF5; case "+^{F6}": return ShortCut.ShiftCtrlF6;
                case "+^{F7}": return ShortCut.ShiftCtrlF7; case "+^{F8}": return ShortCut.ShiftCtrlF8; case "+^{F9}": return ShortCut.ShiftCtrlF9;
                case "+^{F11}": return ShortCut.ShiftCtrlF11; case "+^{F12}": return ShortCut.ShiftCtrlF12;

                // Other Keys
                case "^{INSERT}": return ShortCut.CtrlIns; // Ctrl+Insert
                case "+{INSERT}": return ShortCut.ShiftIns; // Shift+Insert
                case "{DEL}": return ShortCut.Del;       // Delete
                case "+{DEL}": return ShortCut.ShiftDel;  // Shift+Delete
                case "%{BKSP}": return ShortCut.AltBKsp; // Alt+Backspace
                
                case "": return ShortCut.None; 
                default:
                    // Optionally log an unknown shortcut string
                    // Console.Error.WriteLine($"Warning: Unknown shortcut string '{s}' encountered.");
                    return null; // Or throw an exception if strict parsing is required.
            }
        }
        
        public static Connection? ParseConnectionString(string s)
        {
            switch (s)
            {
                case "Access": return Connection.Access;
                case "dBase III": return Connection.DBaseIII; 
                case "dBase IV": return Connection.DBaseIV; 
                case "dBase 5.0": return Connection.DBase5_0;
                case "Excel 3.0": return Connection.Excel3_0; 
                case "Excel 4.0": return Connection.Excel4_0; 
                case "Excel 5.0": return Connection.Excel5_0; 
                case "Excel 8.0": return Connection.Excel8_0;
                case "FoxPro 2.0": return Connection.FoxPro2_0; 
                case "FoxPro 2.5": return Connection.FoxPro2_5; 
                case "FoxPro 2.6": return Connection.FoxPro2_6; 
                case "FoxPro 3.0": return Connection.FoxPro3_0;
                case "Lotus WK1": return Connection.LotusWk1; // Or maybe "Lotus Works 1"
                case "Lotus WK3": return Connection.LotusWk3; // Or maybe "Lotus Works 3"
                case "Lotus WK4": return Connection.LotusWk4; // Or maybe "Lotus Works 4"
                case "Paradox 3.X": return Connection.Paradox3X; 
                case "Paradox 4.X": return Connection.Paradox4X; 
                case "Paradox 5.X": return Connection.Paradox5X;
                case "Text": return Connection.Text;
                default: 
                    // Optionally log unknown connection string
                    // Console.Error.WriteLine($"Warning: Unknown connection string '{s}' encountered.");
                    return null; // Or throw if strict
            }
        }

        /// <summary>
        /// Strips a trailing VB6 comment (starting with ') from a string value,
        /// respecting quotes. Then, unquotes the string if it's surrounded by double quotes
        /// and replaces internal "" with ".
        /// </summary>
        public static string UnquoteAndStripComment(string rawValue)
        {
            if (string.IsNullOrEmpty(rawValue)) return string.Empty;

            int commentStartIndex = -1;
            bool inDoubleQuotes = false;
            StringBuilder sbValuePart = new StringBuilder();

            for (int i = 0; i < rawValue.Length; i++)
            {
                char currentChar = rawValue[i];
                if (currentChar == '\"')
                {
                    // Check for escaped quote: if previous was quote AND current is quote
                    if (inDoubleQuotes && i + 1 < rawValue.Length && rawValue[i + 1] == '\"')
                    {
                        sbValuePart.Append('\"'); // Append one quote for ""
                        i++; // Skip next quote
                        continue;
                    }
                    inDoubleQuotes = !inDoubleQuotes;
                    sbValuePart.Append('\"'); // Append the quote itself
                }
                else if (currentChar == '\'' && !inDoubleQuotes)
                {
                    commentStartIndex = i;
                    break;
                }
                else
                {
                    sbValuePart.Append(currentChar);
                }
            }

            string valueWithoutComment = sbValuePart.ToString().Trim(); // Trim spaces from the part before comment

            // Now, unquote if it's fully quoted
            if (valueWithoutComment.Length >= 2 && valueWithoutComment.StartsWith("\"") && valueWithoutComment.EndsWith("\""))
            {
                // This simple unquoting might be too naive if strings can contain escaped quotes
                // that were not handled by the loop above. The loop tries to build the content correctly.
                // The `Vb6.ParseVB6String` is more robust for string content.
                // For property values, it's usually simpler.
                // Let's assume the loop built the correct inner content, and we just need to strip
                // *potential* outer quotes if they were not part of an escaped sequence.

                // If the loop correctly built the string content (handling internal "" as "),
                // then `valueWithoutComment` might already be the unquoted version if the original was like `"""quoted"""`.
                // Or, if it was just `"simple"`, `valueWithoutComment` would be `"simple"`.
                // The `UnquoteValue` method from Project.cs was more straightforward.
                // Let's adapt that logic here.

                // First, strip comment based on being outside quotes
                string valueBeforeComment;
                int cmtIdx = -1;
                bool inDblQt = false;
                for (int i = 0; i < rawValue.Length; ++i)
                {
                    if (rawValue[i] == '\"') inDblQt = !inDblQt;
                    else if (rawValue[i] == '\'' && !inDblQt) { cmtIdx = i; break; }
                }
                valueBeforeComment = (cmtIdx != -1) ? rawValue.Substring(0, cmtIdx) : rawValue;

                string trimmedVal = valueBeforeComment.Trim();

                if (trimmedVal.Length >= 2 && trimmedVal.StartsWith("\"") && trimmedVal.EndsWith("\""))
                {
                    // Remove outer quotes
                    string inner = trimmedVal.Substring(1, trimmedVal.Length - 2);
                    // Replace VB's escaped double quotes ("") with a single double quote
                    return inner.Replace("\"\"", "\"");
                }
                return trimmedVal; // Not a quoted string, or malformed
            }
            // If not clearly quoted, return the comment-stripped and trimmed value
            return valueWithoutComment;
        }
    }
}