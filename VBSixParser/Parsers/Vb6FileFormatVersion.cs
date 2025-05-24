// Namespace: VB6Parse.Parsers (or VB6Parse.Model)
namespace VB6Parse.Model // Or VB6Parse.Parsers
{
    /// <summary>
    /// Represents the file format version of a VB6 file (e.g., "VERSION 5.00").
    /// </summary>
    public class Vb6FileFormatVersion
    {
        public int Major { get; set; }
        public int Minor { get; set; }

        public Vb6FileFormatVersion(int major = 0, int minor = 0)
        {
            Major = major;
            Minor = minor;
        }

        public override string ToString()
        {
            return $"Version {Major}.{Minor:D2}";
        }
    }
}