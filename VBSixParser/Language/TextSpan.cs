namespace VB6Parse.Language
{
    /// <summary>
    /// Represents a span of text in a source file, defined by a starting position and length.
    /// </summary>
    public readonly struct TextSpan : IEquatable<TextSpan>
    {
        /// <summary>
        /// Gets the starting character index of the span.
        /// </summary>
        public int Start { get; }

        /// <summary>
        /// Gets the length of the span in characters.
        /// </summary>
        public int Length { get; }

        /// <summary>
        /// Gets the ending character index of the span (exclusive).
        /// </summary>
        public int End => Start + Length;

        /// <summary>
        /// Initializes a new instance of the <see cref="TextSpan"/> struct.
        /// </summary>
        /// <param name="start">The starting character index.</param>
        /// <param name="length">The length of the span.</param>
        public TextSpan(int start, int length)
        {
            if (start < 0)
                throw new ArgumentOutOfRangeException(nameof(start), "Start index must be non-negative.");
            if (length < 0)
                throw new ArgumentOutOfRangeException(nameof(length), "Length must be non-negative.");
            
            Start = start;
            Length = length;
        }

        public override bool Equals(object obj) => obj is TextSpan other && Equals(other);

        public bool Equals(TextSpan other) => Start == other.Start && Length == other.Length;

        public override int GetHashCode() => HashCode.Combine(Start, Length);

        public static bool operator ==(TextSpan left, TextSpan right) => left.Equals(right);
        public static bool operator !=(TextSpan left, TextSpan right) => !(left == right);

        public override string ToString() => $"[{Start}..{End})";
    }
}