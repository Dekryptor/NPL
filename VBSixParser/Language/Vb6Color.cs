using System;
using System.Globalization; // For NumberStyles.HexNumber
using VB6Parse.Errors;

namespace VB6Parse.Language
{
    /// <summary>
    /// Represents a VB6 color, which can be an RGB color or a system color.
    /// VB6Colors are stored and used within VB6 as text formatted as '&H00BBGGRR&' (RGB)
    /// or '&H800000II&' (System Color Index).
    /// </summary>
    public readonly struct Vb6Color : IEquatable<Vb6Color>
    {
        // --- Fields and Properties from previous step ---
        private readonly byte _r, _g, _b;
        private readonly byte _systemIndex;
        private readonly bool _isSystemColor;

        public byte R { get { if (_isSystemColor) throw new InvalidOperationException("Cannot get RGB components from a system color."); return _r; } }
        public byte G { get { if (_isSystemColor) throw new InvalidOperationException("Cannot get RGB components from a system color."); return _g; } }
        public byte B { get { if (_isSystemColor) throw new InvalidOperationException("Cannot get RGB components from a system color."); return _b; } }
        public byte SystemIndex { get { if (!_isSystemColor) throw new InvalidOperationException("Cannot get system index from an RGB color."); return _systemIndex; } }
        public bool IsSystemColor => _isSystemColor;
        public bool IsRgbColor => !_isSystemColor;

        // --- Constructors from previous step ---
        private Vb6Color(byte r, byte g, byte b)
        {
            _r = r;
            _g = g;
            _b = b;
            _systemIndex = 0;
            _isSystemColor = false;
        }
        private Vb6Color(byte systemIndex)
        {
            _r = 0;
            _g = 0;
            _b = 0;
            _systemIndex = systemIndex;
            _isSystemColor = true;
        }
        public static Vb6Color FromRgb(byte red, byte green, byte blue) => new Vb6Color(red, green, blue);
        public static Vb6Color FromSystemIndex(byte index) => new Vb6Color(index);

        // --- Predefined Color Constants ---
        public static readonly Vb6Color VbBlack = FromRgb(0x00, 0x00, 0x00);
        public static readonly Vb6Color VbWhite = FromRgb(0xFF, 0xFF, 0xFF);
        public static readonly Vb6Color VbRed = FromRgb(0xFF, 0x00, 0x00);
        public static readonly Vb6Color VbGreen = FromRgb(0x00, 0xFF, 0x00);
        public static readonly Vb6Color VbBlue = FromRgb(0x00, 0x00, 0xFF);
        public static readonly Vb6Color VbYellow = FromRgb(0xFF, 0xFF, 0x00);
        public static readonly Vb6Color VbMagenta = FromRgb(0xFF, 0x00, 0xFF);
        public static readonly Vb6Color VbCyan = FromRgb(0x00, 0xFF, 0xFF);

        // System Colors
        public static readonly Vb6Color Vb3dDkShadow = FromSystemIndex(0x15);         // Darkest shadow color for 3-D display elements.
        public static readonly Vb6Color Vb3dHighlight = FromSystemIndex(0x14);        // Highlight color for 3-D display elements
        public static readonly Vb6Color Vb3dLight = FromSystemIndex(0x16);            // Second lightest 3-D color after vb3DHighlight
        public static readonly Vb6Color Vb3dShadow = FromSystemIndex(0x10);           // Lightest shadow color for 3-D display elements (also VB_BUTTON_SHADOW)
        public static readonly Vb6Color VbActiveBorder = FromSystemIndex(0x0A);       // Border color of active window
        public static readonly Vb6Color VbActiveTitleBar = FromSystemIndex(0x02);     // Color of the title bar for the active window
        public static readonly Vb6Color VbApplicationWorkspace = FromSystemIndex(0x0C); // Background color of multiple document interface (MDI) applications
        public static readonly Vb6Color VbButtonFace = FromSystemIndex(0x0F);         // Color of shading on the face of command buttons (also VB_3D_FACE)
        public static readonly Vb6Color Vb3dFace = FromSystemIndex(0x0F);             // (Same as VB_BUTTON_FACE)
        public static readonly Vb6Color VbButtonShadow = FromSystemIndex(0x10);       // Color of shading on the edge of command buttons (Same as VB_3D_SHADOW)
        public static readonly Vb6Color VbButtonText = FromSystemIndex(0x12);         // Text color on push buttons
        public static readonly Vb6Color VbDesktop = FromSystemIndex(0x01);            // Desktop color
        public static readonly Vb6Color VbGrayText = FromSystemIndex(0x11);           // Grayed (disabled) text
        public static readonly Vb6Color VbHighlight = FromSystemIndex(0x0D);          // Background color of items selected in a control
        public static readonly Vb6Color VbHighlightText = FromSystemIndex(0x0E);      // Text color of items selected in a control
        public static readonly Vb6Color VbInactiveBorder = FromSystemIndex(0x0B);     // Border color of inactive window
        public static readonly Vb6Color VbInactiveCaptionText = FromSystemIndex(0x13); // Color of text in an inactive caption
        public static readonly Vb6Color VbInactiveTitleBar = FromSystemIndex(0x03);   // Color of the title bar for the inactive window
        public static readonly Vb6Color VbInfoBackground = FromSystemIndex(0x18);     // Background color of tool tips (also VB_MSG_BOX_TEXT)
        public static readonly Vb6Color VbMsgBoxText = FromSystemIndex(0x18);         // (Same as VB_INFO_BACKGROUND) - Note: Rust comments say "Background color of tool tips", VB name implies text. Assuming index from Rust is primary.
        public static readonly Vb6Color VbInfoText = FromSystemIndex(0x17);           // Color of text in tool tips (also VB_MSG_BOX)
        public static readonly Vb6Color VbMsgBox = FromSystemIndex(0x17);             // (Same as VB_INFO_TEXT) - Note: Rust comments say "Color of text in tool tips", VB name implies box (background). Assuming index from Rust is primary.
        public static readonly Vb6Color VbMenuBar = FromSystemIndex(0x04);            // Menu background color
        public static readonly Vb6Color VbMenuText = FromSystemIndex(0x07);           // Color of text on menus
        public static readonly Vb6Color VbScrollBars = FromSystemIndex(0x00);         // Scrollbar color
        public static readonly Vb6Color VbTitleBarText = FromSystemIndex(0x09);       // Color of text in caption, size box, and scroll arrow
        public static readonly Vb6Color VbWindowBackground = FromSystemIndex(0x05);   // Window background color
        public static readonly Vb6Color VbWindowFrame = FromSystemIndex(0x06);        // Window frame color
        public static readonly Vb6Color VbWindowText = FromSystemIndex(0x08);         // Color of text in windows
        
        // --- ParseHex method from previous step ---
        public static Vb6Color ParseHex(string input)
        {
            if (input == null)
                throw new ArgumentNullException(nameof(input));

            input = input.Trim();

            if (!input.StartsWith("&H", StringComparison.OrdinalIgnoreCase) || !input.EndsWith("&"))
                throw new Vb6ParseException(Vb6ErrorKind.HexColorParseError, $"Invalid VB6 color format: Must start with '&H' and end with '&'. Input: '{input}'");

            if (input.Length != 11)
                throw new Vb6ParseException(Vb6ErrorKind.HexColorParseError, $"Invalid VB6 color format: Incorrect length. Expected 11 characters. Input: '{input}'");

            string hexContent = input.Substring(2, 8);

            try
            {
                if (hexContent.StartsWith("80", StringComparison.OrdinalIgnoreCase))
                {
                    if (!hexContent.Substring(2, 4).Equals("0000", StringComparison.OrdinalIgnoreCase))
                         throw new Vb6ParseException(Vb6ErrorKind.HexColorParseError, $"Invalid system color format: Expected '&H800000II&'. Input: '{input}'");

                    byte index = byte.Parse(hexContent.Substring(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    return FromSystemIndex(index);
                }
                else if (hexContent.StartsWith("00", StringComparison.OrdinalIgnoreCase))
                {
                    byte blue = byte.Parse(hexContent.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    byte green = byte.Parse(hexContent.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    byte red = byte.Parse(hexContent.Substring(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    return FromRgb(red, green, blue);
                }
                else
                {
                    throw new Vb6ParseException(Vb6ErrorKind.HexColorParseError, $"Invalid VB6 color format: Expected '&H00...' for RGB or '&H80...' for System. Input: '{input}'");
                }
            }
            catch (FormatException ex)
            {
                throw new Vb6ParseException(Vb6ErrorKind.HexColorParseError, $"Failed to parse hex values in VB6 color string: '{input}'", null, ex);
            }
            catch (OverflowException ex)
            {
                throw new Vb6ParseException(Vb6ErrorKind.HexColorParseError, $"Hex value out of range in VB6 color string: '{input}'", null, ex);
            }
        }

        // --- Equality and ToString methods from previous step ---
        public override bool Equals(object obj) => obj is Vb6Color other && Equals(other);
        public bool Equals(Vb6Color other)
        {
            if (_isSystemColor != other._isSystemColor) return false;
            if (_isSystemColor) return _systemIndex == other._systemIndex;
            return _r == other._r && _g == other._g && _b == other._b;
        }
        public override int GetHashCode()
        {
            if (_isSystemColor) return HashCode.Combine(_isSystemColor, _systemIndex);
            return HashCode.Combine(_isSystemColor, _r, _g, _b);
        }
        public static bool operator ==(Vb6Color left, Vb6Color right) => left.Equals(right);
        public static bool operator !=(Vb6Color left, Vb6Color right) => !(left == right);
        public override string ToString()
        {
            if (_isSystemColor) return $"&H800000{_systemIndex:X2}&";
            return $"&H00{_b:X2}{_g:X2}{_r:X2}&";
        }
    }
}