// File: Language/VB6Color.cs
using System;
using System.Globalization;
using VBSix.Errors; // For VB6ErrorKind, VB6ParseException

namespace VBSix.Language
{
    /// <summary>
    /// Specifies whether a <see cref="VB6Color"/> represents an RGB color or a system-defined color.
    /// </summary>
    public enum VB6ColorType
    {
        /// <summary>The color is defined by Red, Green, and Blue components.</summary>
        RGB,
        /// <summary>The color is a system-defined color, identified by an index.</summary>
        System
    }

    /// <summary>
    /// Represents a color in Visual Basic 6.
    /// VB6 colors are typically stored as 32-bit integers but represented in .frm files
    /// as hexadecimal strings, either as `&H00BBGGRR&` for RGB or `&H800000II&` for system colors.
    /// </summary>
    public class VB6Color : IEquatable<VB6Color>
    {
        /// <summary>
        /// Gets the type of the color (RGB or System).
        /// </summary>
        public VB6ColorType Type { get; private set; }

        /// <summary>
        /// Gets the Red component of the color (0-255). Only valid if Type is RGB.
        /// </summary>
        public byte R { get; private set; }

        /// <summary>
        /// Gets the Green component of the color (0-255). Only valid if Type is RGB.
        /// </summary>
        public byte G { get; private set; }

        /// <summary>
        /// Gets the Blue component of the color (0-255). Only valid if Type is RGB.
        /// </summary>
        public byte B { get; private set; }

        /// <summary>
        /// Gets the index of the system color. Only valid if Type is System.
        /// </summary>
        public byte SystemIndex { get; private set; }

        // Private constructor to enforce factory method usage.
        private VB6Color() { }

        /// <summary>
        /// Creates a <see cref="VB6Color"/> from RGB components.
        /// </summary>
        /// <param name="r">The red component (0-255).</param>
        /// <param name="g">The green component (0-255).</param>
        /// <param name="b">The blue component (0-255).</param>
        /// <returns>A new <see cref="VB6Color"/> instance.</returns>
        public static VB6Color FromRGB(byte r, byte g, byte b)
        {
            return new VB6Color { Type = VB6ColorType.RGB, R = r, G = g, B = b };
        }

        /// <summary>
        /// Creates a <see cref="VB6Color"/> from a system color index.
        /// </summary>
        /// <param name="systemIndex">The index of the system color.</param>
        /// <returns>A new <see cref="VB6Color"/> instance.</returns>
        public static VB6Color FromSystemIndex(byte systemIndex)
        {
            // VB6 system color indices are typically 0-24, but can go higher for OleColors.
            // No specific validation here, assuming caller provides valid VB6 index.
            return new VB6Color { Type = VB6ColorType.System, SystemIndex = systemIndex };
        }

        /// <summary>
        /// Parses a VB6 color string (e.g., "&H00BBGGRR&" or "&H800000II&").
        /// </summary>
        /// <param name="hexColorString">The VB6 hex color string.</param>
        /// <returns>A <see cref="VB6Color"/> object.</returns>
        /// <exception cref="ArgumentNullException">If hexColorString is null.</exception>
        /// <exception cref="VB6ParseException">If the string is not a valid VB6 color format.</exception>
        public static VB6Color FromHex(string hexColorString)
        {
            if (hexColorString == null) throw new ArgumentNullException(nameof(hexColorString));

            if (string.IsNullOrWhiteSpace(hexColorString) ||
                !hexColorString.StartsWith("&H", StringComparison.OrdinalIgnoreCase) ||
                !hexColorString.EndsWith("&", StringComparison.OrdinalIgnoreCase) ||
                hexColorString.Length != 11) // Format is &H + 8 hex chars + &
            {
                throw new VB6ParseException(VB6ErrorKind.HexColorParseError, 
                    new FormatException($"Invalid VB6 color string format: '{hexColorString}'. Expected format like '&H00BBGGRR&' or '&H800000II&'."));
            }

            // Extract the 8 hex characters representing the color data.
            string colorDataHex = hexColorString.Substring(2, 8); 
            uint colorValue;
            try
            {
                colorValue = uint.Parse(colorDataHex, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            }
            catch (FormatException ex)
            {
                throw new VB6ParseException(VB6ErrorKind.HexColorParseError, 
                    new FormatException($"Could not parse color data '{colorDataHex}' as hexadecimal.", ex));
            }
            catch (OverflowException ex) // Should not happen with uint for 8 hex chars
            {
                 throw new VB6ParseException(VB6ErrorKind.HexColorParseError, 
                    new FormatException($"Color data '{colorDataHex}' resulted in overflow.", ex));
            }


            // Check the most significant byte (first two hex chars of colorDataHex) to determine type
            // colorValue = 0x S S B B G G R R (S=System flag byte, B=Blue, G=Green, R=Red, I=Index)
            if ((colorValue & 0xFF000000) == 0x80000000) // System color if MSB is 0x80
            {
                byte index = (byte)(colorValue & 0xFF); // System index is in the least significant byte
                return FromSystemIndex(index);
            }
            else if ((colorValue & 0xFF000000) == 0x00000000 || (colorValue & 0x02000000) != 0) // Standard RGB or OLE_COLOR
            {
                // Standard RGB: &H00BBGGRR&
                // OLE_COLOR can also be represented, often with the 0x02000000 bit set if it's not a system color.
                // For simplicity, if not 0x80 prefix, treat as RGB.
                byte r_val = (byte)(colorValue & 0xFF);
                byte g_val = (byte)((colorValue >> 8) & 0xFF);
                byte b_val = (byte)((colorValue >> 16) & 0xFF);
                return FromRGB(r_val, g_val, b_val);
            }
            else
            {
                // This case implies the first byte (e.g. from `&HXX......&`) was neither 0x00 nor 0x80.
                throw new VB6ParseException(VB6ErrorKind.HexColorParseError, 
                    new FormatException($"Unknown color prefix byte in '{colorDataHex}'. Expected '00' for RGB or '80' for System."));
            }
        }

        public override string ToString()
        {
            if (Type == VB6ColorType.RGB)
            {
                // Format as &H00BBGGRR&
                return $"&H{0:X2}{B:X2}{G:X2}{R:X2}&";
            }
            else // System
            {
                // Format as &H800000II&
                return $"&H800000{SystemIndex:X2}&";
            }
        }

        public bool Equals(VB6Color? other)
        {
            if (other is null) return false;
            if (ReferenceEquals(this, other)) return true;
            if (Type != other.Type) return false;
            
            if (Type == VB6ColorType.RGB)
            {
                return R == other.R && G == other.G && B == other.B;
            }
            // Type == VB6ColorType.System
            return SystemIndex == other.SystemIndex;
        }

        public override bool Equals(object? obj) => Equals(obj as VB6Color);

        public override int GetHashCode()
        {
            if (Type == VB6ColorType.RGB)
            {
                return HashCode.Combine(Type, R, G, B);
            }
            // Type == VB6ColorType.System
            return HashCode.Combine(Type, SystemIndex);
        }

        // --- Predefined Colors ---
        public static readonly VB6Color VbBlack   = FromRGB(0x00, 0x00, 0x00);
        public static readonly VB6Color VbWhite   = FromRGB(0xFF, 0xFF, 0xFF);
        public static readonly VB6Color VbRed     = FromRGB(0xFF, 0x00, 0x00);
        public static readonly VB6Color VbGreen   = FromRGB(0x00, 0xFF, 0x00);
        public static readonly VB6Color VbBlue    = FromRGB(0x00, 0x00, 0xFF);
        public static readonly VB6Color VbYellow  = FromRGB(0xFF, 0xFF, 0x00);
        public static readonly VB6Color VbMagenta = FromRGB(0xFF, 0x00, 0xFF);
        public static readonly VB6Color VbCyan    = FromRGB(0x00, 0xFF, 0xFF);

        // System Colors (Indices match common VB system color constants)
        public static readonly VB6Color VbScrollBars           = FromSystemIndex(0);  // &H80000000&
        public static readonly VB6Color VbDesktop              = FromSystemIndex(1);  // &H80000001&
        public static readonly VB6Color VbActiveTitleBar       = FromSystemIndex(2);  // &H80000002&
        public static readonly VB6Color VbInactiveTitleBar     = FromSystemIndex(3);  // &H80000003&
        public static readonly VB6Color VbMenuBar              = FromSystemIndex(4);  // &H80000004&
        public static readonly VB6Color VbWindowBackground     = FromSystemIndex(5);  // &H80000005&
        public static readonly VB6Color VbWindowFrame          = FromSystemIndex(6);  // &H80000006&
        public static readonly VB6Color VbMenuText             = FromSystemIndex(7);  // &H80000007&
        public static readonly VB6Color VbWindowText           = FromSystemIndex(8);  // &H80000008&
        public static readonly VB6Color VbTitleBarText         = FromSystemIndex(9);  // &H80000009& (Active title bar text)
        public static readonly VB6Color VbActiveBorder         = FromSystemIndex(10); // &H8000000A&
        public static readonly VB6Color VbInactiveBorder       = FromSystemIndex(11); // &H8000000B&
        public static readonly VB6Color VbApplicationWorkspace = FromSystemIndex(12); // &H8000000C& (MDI background)
        public static readonly VB6Color VbHighlight            = FromSystemIndex(13); // &H8000000D&
        public static readonly VB6Color VbHighlightText        = FromSystemIndex(14); // &H8000000E&
        public static readonly VB6Color VbButtonFace           = FromSystemIndex(15); // &H8000000F& (Same as Vb3DFace)
        public static readonly VB6Color Vb3DFace               = VbButtonFace;
        public static readonly VB6Color VbButtonShadow         = FromSystemIndex(16); // &H80000010& (Same as Vb3DShadow)
        public static readonly VB6Color Vb3DShadow             = VbButtonShadow;
        public static readonly VB6Color VbGrayText             = FromSystemIndex(17); // &H80000011&
        public static readonly VB6Color VbButtonText           = FromSystemIndex(18); // &H80000012&
        public static readonly VB6Color VbInactiveCaptionText  = FromSystemIndex(19); // &H80000013& (Inactive title bar text)
        public static readonly VB6Color Vb3DHighlight          = FromSystemIndex(20); // &H80000014&
        public static readonly VB6Color Vb3DDkShadow           = FromSystemIndex(21); // &H80000015&
        public static readonly VB6Color Vb3DLight              = FromSystemIndex(22); // &H80000016&
        public static readonly VB6Color VbInfoText             = FromSystemIndex(23); // &H80000017&
        public static readonly VB6Color VbInfoBackground       = FromSystemIndex(24); // &H80000018&
        
        public static readonly VB6Color VB_MSG_BOX             = VbInfoText;
        public static readonly VB6Color VB_MSG_BOX_TEXT        = VbInfoBackground;
    }
}