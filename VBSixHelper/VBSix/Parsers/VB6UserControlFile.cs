using VBSix.Errors;
using VBSix.Language;
using VBSix.Language.Controls;

namespace VBSix.Parsers
{
    public class VB6UserControlFile
    {
        public VB6FileFormatVersion? FormatVersion { get; set; }
        public List<VB6ObjectReference> ObjectReferences { get; set; } = [];
        public VB6Control? UserControl { get; set; } // The main UserControl object
        public VB6FileAttributes Attributes { get; set; } = new VB6FileAttributes();
        public List<VB6Token> Tokens { get; set; } = [];

        public static VB6UserControlFile Parse(string ctlPath, byte[] inputBytes)
        {
            // For .ctl files, the associated resource file for ToolboxBitmap is typically .ctx
            // The DefaultResourceFileResolver should work fine if it can handle various extensions.
            return ParseWithResolver(ctlPath, inputBytes, VB6FormFile.DefaultResourceFileResolver);
        }

        public static VB6UserControlFile ParseWithResolver(string ctlPath, byte[] inputBytes, ResourceFileResolver resourceResolver)
        {
            var ucFile = new VB6UserControlFile();
            var stream = new VB6Stream(ctlPath, inputBytes);
            var initialStreamForError = stream.Checkpoint();

            try
            {
                // 1. Parse Version Line (e.g., "VERSION 5.00") - UserControls don't have "USERCONTROL" on this line
                var (sAfterVersion, version, _) = HeaderParser.ParseVersionLine(stream, HeaderKind.Form); // Use Form kind as it's similar
                ucFile.FormatVersion = version;
                stream = sAfterVersion;

                // 2. Parse Object References (Object=...) for constituent ActiveX controls
                (stream, ucFile.ObjectReferences) = VB6FormFile.ParseFormObjects(stream); // Re-use Form's object parser

                // Skip blank lines/comments before main BEGIN block
                stream = VB6FormFile.SkipBlankLinesAndComments(stream); // Re-use helper

                // 3. Parse Main Control Block (BEGIN VB.UserControl UserControlName ... END)
                VB6FullyQualifiedName mainFqn = ParseBeginUserControlLine(ref stream); // Specific to UserControl

                // Re-use ParseControlBlock from VB6FormFile, as the structure is identical
                var (streamAfterControlBlock, mainControl) = VB6FormFile.ParseControlBlock(stream, mainFqn, resourceResolver, 0, ctlPath);
                ucFile.UserControl = mainControl;
                stream = streamAfterControlBlock;

                stream = VB6FormFile.SkipBlankLinesAndComments(stream);

                // 4. Parse Attributes (Attribute VB_Name = ..., Attribute VB_UserControl = True, etc.)
                var (streamAfterAttrs, attributes) = HeaderParser.ParseAttributes(stream);
                ucFile.Attributes = attributes;
                stream = streamAfterAttrs;

                // Validate VB_UserControl attribute
                bool isUserControlAttributePresent = false;
                if (attributes.ExtKey.TryGetValue("VB_UserControl", out string? val) && (val.Equals("True", StringComparison.OrdinalIgnoreCase) || val == "-1"))
                {
                    isUserControlAttributePresent = true;
                }

                // Could also check attributes.ExtKey directly if VB_UserControl wasn't special cased in VB6FileAttributes
                // For example, if VB6FileAttributes had a specific field: `public bool IsUserControl {get;set;}`
                // that HeaderParser.ParseAttributes would populate.
                // For now, we assume it might be in ExtKey or a dedicated field if you add one.
                // If strict validation is needed:
                //if (!isUserControlAttributePresent) {
                //    throw stream.CreateException(VB6ErrorKind.AttributeParseError, innerException: new Exception("Missing 'Attribute VB_UserControl = True'"));
                //}

                // 5. Parse Code Tokens
                ucFile.Tokens = VB6.Vb6Parse(stream);

                return ucFile;
            }
            catch (VB6ParseException) { throw; }
            catch (Exception ex)
            {
                throw stream.CreateExceptionFromCheckpoint(initialStreamForError, VB6ErrorKind.Unparseable,
                    innerException: new Exception($"Unexpected error parsing UserControl file '{ctlPath}'. Details: {ex.Message}", ex));
            }
        }

        private static VB6FullyQualifiedName ParseBeginUserControlLine(ref VB6Stream stream)
        {
            var initialCheckpoint = stream.Checkpoint();
            VB6Stream s = HeaderParser.SkipSpace0(stream);

            if (!s.CompareSliceCaseless("BEGIN VB.UserControl ")) // Expect "VB.UserControl"
                throw s.CreateExceptionFromCheckpoint(initialCheckpoint, VB6ErrorKind.NoBegin, innerException: new Exception("Expected 'BEGIN VB.UserControl ' for main block."));

            (s, _) = VB6.TakeSpecificString(s, "BEGIN", caseSensitive: true);
            s = HeaderParser.SkipSpace1(s);

            // The FQN parser should handle "VB.UserControl ControlName"
            var fqnResultTuple = VB6FormFile.ParseFullyQualifiedNameInternal(s);
            if (fqnResultTuple == null ||
                !fqnResultTuple.Value.fqn.Namespace.Equals("VB", StringComparison.OrdinalIgnoreCase) ||
                !fqnResultTuple.Value.fqn.Kind.Equals("UserControl", StringComparison.OrdinalIgnoreCase))
            {
                throw s.CreateExceptionFromCheckpoint(initialCheckpoint, VB6ErrorKind.NoForm, // Re-use NoForm, or create NoUserControl
                    innerException: new Exception("Could not parse UserControl's fully qualified name (expected VB.UserControl Name)."));
            }

            stream = fqnResultTuple.Value.stream;
            return fqnResultTuple.Value.fqn;
        }
    }
}