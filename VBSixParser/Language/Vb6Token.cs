using System;

namespace VB6Parse.Language
{
    /// <summary>
    /// Represents a lexical token in VB6 source code.
    /// </summary>
    public readonly struct Vb6Token : IEquatable<Vb6Token>
    {
        /// <summary>
        /// Gets the type of the token.
        /// </summary>
        public Vb6TokenType Type { get; }

        /// <summary>
        /// Gets the actual text of the token from the source code.
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// Gets the span of the token in the source text.
        /// </summary>
        public TextSpan Span { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Vb6Token"/> struct.
        /// </summary>
        /// <param name="type">The type of the token.</param>
        /// <param name="text">The text of the token.</param>
        /// <param name="span">The span of the token in the source text.</param>
        public Vb6Token(Vb6TokenType type, string text, TextSpan span)
        {
            Type = type;
            Text = text ?? throw new ArgumentNullException(nameof(text));
            Span = span;
        }

        public override bool Equals(object obj) => obj is Vb6Token other && Equals(other);

        public bool Equals(Vb6Token other) =>
            Type == other.Type &&
            Text == other.Text &&
            Span == other.Span;

        public override int GetHashCode() => HashCode.Combine(Type, Text, Span);

        public static bool operator ==(Vb6Token left, Vb6Token right) => left.Equals(right);
        public static bool operator !=(Vb6Token left, Vb6Token right) => !(left == right);

        public override string ToString() => $"{Type}: \"{Text}\" @ {Span}";
    }
}