using System;

namespace VBSix.Utilities
{
    /// <summary>
    /// Provides utility methods for converting between Twips and other units of measurement,
    /// particularly pixels, which are common in display contexts.
    /// VB6 uses Twips (Twentieths of an Inch Point) for many of its coordinate and size properties.
    /// </summary>
    /// <example>
    /// <code>
    /// // Assuming 'formProps' is an instance of a class like C# FormProperties
    /// // int heightInTwips = formProps.Height; // If Height is stored in Twips
    /// 
    /// // For demonstration:
    /// int heightInTwips = 1440; // Equivalent to 1 inch or 96 pixels at 96 DPI
    /// 
    /// double heightInPixels96Dpi = TwipsConverter.ToPixels(heightInTwips);
    /// double heightInPixels120Dpi = TwipsConverter.ToPixels(heightInTwips, 120.0);
    /// 
    /// Console.WriteLine($"Form height: {heightInTwips} twips");
    /// Console.WriteLine($"  -> {heightInPixels96Dpi} pixels @ 96 DPI");
    /// Console.WriteLine($"  -> {heightInPixels120Dpi} pixels @ 120 DPI");
    /// 
    /// double widthInPixels = 200;
    /// int widthInTwips = TwipsConverter.ToTwips(widthInPixels);
    /// Console.WriteLine($"{widthInPixels} pixels @ 96 DPI -> {widthInTwips} twips");
    /// </code>
    /// </example>
    public static class TwipsConverter
    {
        /// <summary>
        /// The number of twips in one logical inch.
        /// </summary>
        public const int TwipsPerInch = 1440;

        /// <summary>
        /// A common default screen DPI (Dots Per Inch) used for calculations
        /// when a specific DPI is not provided.
        /// </summary>
        public const double DefaultScreenDpi = 96.0;

        /// <summary>
        /// Converts a value from Twips to Pixels based on a given screen DPI.
        /// </summary>
        /// <param name="twips">The value in twips.</param>
        /// <param name="dpi">The Dots Per Inch (DPI) of the screen or rendering context.
        /// Defaults to <see cref="DefaultScreenDpi"/> if not provided.</param>
        /// <returns>The equivalent value in pixels.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if DPI is not positive.</exception>
        public static double ToPixels(int twips, double dpi = DefaultScreenDpi)
        {
            if (dpi <= 0) throw new ArgumentOutOfRangeException(nameof(dpi), "DPI must be positive.");
            return (double)twips * dpi / TwipsPerInch;
        }

        /// <summary>
        /// Converts a value from Pixels to Twips based on a given screen DPI.
        /// </summary>
        /// <param name="pixels">The value in pixels.</param>
        /// <param name="dpi">The Dots Per Inch (DPI) of the screen or rendering context.
        /// Defaults to <see cref="DefaultScreenDpi"/> if not provided.</param>
        /// <returns>The equivalent value in twips, rounded to the nearest integer.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if DPI is not positive.</exception>
        public static int ToTwips(double pixels, double dpi = DefaultScreenDpi)
        {
            if (dpi <= 0) throw new ArgumentOutOfRangeException(nameof(dpi), "DPI must be positive.");
            return (int)Math.Round(pixels * TwipsPerInch / dpi);
        }

        /// <summary>
        /// Converts Twips to Points. (1 Point = 20 Twips)
        /// </summary>
        /// <param name="twips">The value in twips.</param>
        /// <returns>The equivalent value in points.</returns>
        public static double ToPoints(int twips)
        {
            return (double)twips / 20.0;
        }

        /// <summary>
        /// Converts Points to Twips. (1 Point = 20 Twips)
        /// </summary>
        /// <param name="points">The value in points.</param>
        /// <returns>The equivalent value in twips, rounded to the nearest integer.</returns>
        public static int FromPoints(double points)
        {
            return (int)Math.Round(points * 20.0);
        }

        /// <summary>
        /// Converts Twips to Inches. (1 Inch = 1440 Twips)
        /// </summary>
        /// <param name="twips">The value in twips.</param>
        /// <returns>The equivalent value in inches.</returns>
        public static double ToInches(int twips)
        {
            return (double)twips / TwipsPerInch;
        }

        /// <summary>
        /// Converts Inches to Twips. (1 Inch = 1440 Twips)
        /// </summary>
        /// <param name="inches">The value in inches.</param>
        /// <returns>The equivalent value in twips, rounded to the nearest integer.</returns>
        public static int FromInches(double inches)
        {
            return (int)Math.Round(inches * TwipsPerInch);
        }

        /// <summary>
        /// Converts Twips to Centimeters. (1 Inch = 2.54 cm, 1 Inch = 1440 Twips)
        /// So, 1 cm = 1440 / 2.54 Twips approx 566.929
        /// Using a more precise constant: 1 cm = 567 Twips is a common VB6 approximation.
        /// </summary>
        public const double TwipsPerCentimeter = TwipsPerInch / 2.54; // More precise: 566.929133858...

        /// <summary>
        /// Converts Twips to Centimeters.
        /// </summary>
        /// <param name="twips">The value in twips.</param>
        /// <returns>The equivalent value in centimeters.</returns>
        public static double ToCentimeters(int twips)
        {
            return (double)twips / TwipsPerCentimeter;
        }

        /// <summary>
        /// Converts Centimeters to Twips.
        /// </summary>
        /// <param name="cm">The value in centimeters.</param>
        /// <returns>The equivalent value in twips, rounded to the nearest integer.</returns>
        public static int FromCentimeters(double cm)
        {
            return (int)Math.Round(cm * TwipsPerCentimeter);
        }

        /// <summary>
        /// Converts Twips to Millimeters. (1 cm = 10 mm)
        /// </summary>
        public const double TwipsPerMillimeter = TwipsPerCentimeter / 10.0; // Approx 56.6929

        /// <summary>
        /// Converts Twips to Millimeters.
        /// </summary>
        /// <param name="twips">The value in twips.</param>
        /// <returns>The equivalent value in millimeters.</returns>
        public static double ToMillimeters(int twips)
        {
            return (double)twips / TwipsPerMillimeter;
        }

        /// <summary>
        /// Converts Millimeters to Twips.
        /// </summary>
        /// <param name="mm">The value in millimeters.</param>
        /// <returns>The equivalent value in twips, rounded to the nearest integer.</returns>
        public static int FromMillimeters(double mm)
        {
            return (int)Math.Round(mm * TwipsPerMillimeter);
        }
    }
}