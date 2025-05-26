using System.ComponentModel;
using System.Reflection;

namespace VBSix.Errors
{
    public enum PropertyErrorType
    {
        [Description("Appearance can only be a 0 (Flat) or a 1 (ThreeD)")]
        AppearanceInvalid,

        [Description("BorderStyle can only be a 0 (None) or 1 (FixedSingle)")]
        BorderStyleInvalid,

        [Description("ClipControls can only be a 0 (false) or a 1 (true)")]
        ClipControlsInvalid,

        [Description("DragMode can only be 0 (Manual) or 1 (Automatic)")]
        DragModeInvalid,

        [Description("Enabled can only be 0 (false) or a 1 (true)")]
        EnabledInvalid,

        [Description("MousePointer can only be 0 (Default), 1 (Arrow), 2 (Cross), 3 (IBeam), 6 (SizeNESW), 7 (SizeNS), 8 (SizeNWSE), 9 (SizeWE), 10 (UpArrow), 11 (Hourglass), 12 (NoDrop), 13 (ArrowHourglass), 14 (ArrowQuestion), 15 (SizeAll), or 99 (Custom)")]
        MousePointerInvalid,

        [Description("OLEDropMode can only be 0 (None), or 1 (Manual)")]
        OLEDropModeInvalid,

        [Description("RightToLeft can only be 0 (false) or a 1 (true)")]
        RightToLeftInvalid,

        [Description("Visible can only be 0 (false) or a 1 (true)")]
        VisibleInvalid,

        [Description("Unknown property in header file")]
        UnknownProperty,

        [Description("Invalid property value. Only 0 or -1 are valid for this property")]
        InvalidPropertyValueZeroNegOne,

        [Description("Unable to parse the property name")]
        NameUnparsable,

        [Description("Unable to parse the resource file name")]
        ResourceFileNameUnparsable,

        [Description("Unable to parse the offset into the resource file for property")]
        OffsetUnparsable,

        [Description("Invalid property value. Only True or False are valid for this property")]
        InvalidPropertyValueTrueFalse,
    }

    public enum VB6ErrorKind
    {
        [Description("Property parsing error")]
        Property,

        [Description("Resource file parsing error")]
        ResourceFile,

        [Description("Error reading the source file")]
        SourceFileError,

        [Description("The file contains more than a significant number of non-ASCII characters. This file was likely saved in a non-English character set. The vb6parse crate currently does not support non-english vb6 files.")]
        LikelyNonEnglishCharacterSet,

        [Description("The reference line has too many elements")]
        ReferenceExtraSections,

        [Description("The reference line has too few elements")]
        ReferenceMissingSections,

        [Description("The first line of a VB6 project file must be a project 'Type' entry.")]
        FirstLineNotProject,

        [Description("Line type is unknown.")]
        LineTypeUnknown,

        [Description("Project type is not Exe, OleDll, Control, or OleExe")]
        ProjectTypeUnknown,

        [Description("Project lacks a version number.")]
        NoVersion,

        [Description("Project parse error while processing an Object line.")]
        NoObjects, // Could also be UnrecognizedObjectFormat or more specific depending on context

        [Description("Form parse error. No Form found in form file.")]
        NoForm,

        [Description("Parse error while processing Form attributes.")]
        AttributeParseError,

        [Description("Parse error while attempting to parse Form tokens.")]
        TokenParseError,
        
        [Description("Project parse error, failure to find BEGIN element.")]
        NoBegin,

        [Description("Project line entry is not ended with a recognized line ending.")]
        NoLineEnding,

        [Description("Unable to parse the Uuid")]
        UnableToParseUuid,

        [Description("Unable to find a semicolon ';' in this line.")]
        NoSemicolonSplit,

        [Description("Unable to find an equal '=' in this line.")]
        NoEqualSplit,

        [Description("While trying to parse the offset into the resource file, no colon ':' was found.")]
        NoColonForOffsetSplit,

        [Description("No key value divider found in the line.")]
        NoKeyValueDividerFound,

        [Description("Unknown parser error")]
        Unparseable,

        [Description("Major version is not a number.")]
        MajorVersionUnparseable,

        [Description("Unable to parse hex address from DllBaseAddress key")]
        DllBaseAddressUnparseable,

        [Description("The Startup object is not a valid parameter. Must be a quoted startup method/object, \"(None)\", !(None)!, \"\", or \"!!\"")]
        StartupUnparseable,

        [Description("The Name parameter is invalid. Must be a quoted name, \"(None)\", !(None)!, \"\", or \"!!\"")]
        NameUnparsable,

        [Description("The CommandLine parameter is invalid. Must be a quoted command line, \"(None)\", !(None)!, \"\", or \"!!\"")]
        CommandLineUnparseable,

        [Description("The HelpContextId parameter is not a valid parameter line. Must be a quoted help context id, \"(None)\", !(None)!, \"\", or \"!!\"")]
        HelpContextIdUnparseable,

        [Description("Minor version is not a number.")]
        MinorVersionUnparseable,

        [Description("Revision version is not a number.")]
        RevisionVersionUnparseable,

        [Description("Unable to parse the value after ThreadingModel key")]
        ThreadingModelUnparseable,

        [Description("ThreadingModel can only be 0 (Apartment Threaded), or 1 (Single Threaded)")]
        ThreadingModelInvalid,

        [Description("No property name found after BeginProperty keyword.")]
        NoPropertyName,

        [Description("Unable to parse the RelatedDoc property line.")]
        RelatedDocLineUnparseable,

        [Description("AutoIncrement can only be a 0 (false) or a -1 (true)")]
        AutoIncrementUnparseable,

        [Description("CompatibilityMode can only be a 0 (CompatibilityMode::NoCompatibility), 1 (CompatibilityMode::Project), or 2 (CompatibilityMode::CompatibleExe)")]
        CompatibilityModeUnparseable,

        [Description("NoControlUpgrade can only be a 0 (UpgradeControls::Upgrade) or a 1 (UpgradeControls::NoUpgrade)")]
        NoControlUpgradeUnparsable,

        [Description("ServerSupportFiles can only be a 0 (false) or a -1 (true)")]
        ServerSupportFilesUnparseable,

        [Description("Comment line was unparsable")]
        CommentUnparseable,

        [Description("PropertyPage line was unparsable")]
        PropertyPageUnparseable,

        [Description("CompilationType can only be a 0 (false) or a -1 (true)")]
        CompilationTypeUnparseable,

        [Description("OptimizationType can only be a 0 (FastCode) or 1 (SmallCode), or 2 (NoOptimization)")]
        OptimizationTypeUnparseable,

        [Description("FavorPentiumPro(tm) can only be a 0 (false) or a -1 (true)")]
        FavorPentiumProUnparseable,

        [Description("Designer line is unparsable")]
        DesignerLineUnparseable,

        [Description("Form line is unparsable")]
        FormLineUnparseable,

        [Description("UserControl line is unparsable")]
        UserControlLineUnparseable,

        [Description("UserDocument line is unparsable")]
        UserDocumentLineUnparseable,

        [Description("Period expected in version number")]
        PeriodExpectedInVersionNumber,

        [Description("CodeViewDebugInfo can only be a 0 (false) or a -1 (true)")]
        CodeViewDebugInfoUnparseable,

        [Description("NoAliasing can only be a 0 (false) or a -1 (true)")]
        NoAliasingUnparseable,

        [Description("RemoveUnusedControlInfo can only be 0 (UnusedControlInfo::Retain) or -1 (UnusedControlInfo::Remove)")]
        UnusedControlInfoUnparseable,

        [Description("BoundsCheck can only be a 0 (false) or a -1 (true)")]
        BoundsCheckUnparseable,

        [Description("OverflowCheck can only be a 0 (false) or a -1 (true)")]
        OverflowCheckUnparseable,

        [Description("FlPointCheck can only be a 0 (false) or a -1 (true)")]
        FlPointCheckUnparseable,

        [Description("FDIVCheck can only be a 0 (PentiumFDivBugCheck::CheckPentiumFDivBug) or a -1 (PentiumFDivBugCheck::NoPentiumFDivBugCheck)")]
        FDIVCheckUnparseable,

        [Description("UnroundedFP can only be a 0 (UnroundedFloatingPoint::DoNotAllow) or a -1 (UnroundedFloatingPoint::Allow)")]
        UnroundedFPUnparseable,

        [Description("StartMode can only be a 0 (StartMode::StandAlone) or a 1 (StartMode::Automation)")]
        StartModeUnparseable,

        [Description("Unattended can only be a 0 (Unattended::False) or a -1 (Unattended::True)")]
        UnattendedUnparseable,

        [Description("Retained can only be a 0 (Retained::UnloadOnExit) or a 1 (Retained::RetainedInMemory)")]
        RetainedUnparseable,

        [Description("Unable to parse the ShortCut property.")]
        ShortCutUnparseable,

        [Description("DebugStartup can only be a 0 (false) or a -1 (true)")]
        DebugStartupOptionUnparseable,

        [Description("UseExistingBrowser can only be a 0 (UseExistingBrowser::DoNotUse) or a -1 (UseExistingBrowser::Use)")]
        UseExistingBrowserUnparseable,

        [Description("AutoRefresh can only be a 0 (false) or a -1 (true)")]
        AutoRefreshUnparseable,

        [Description("Data control Connection type is not valid.")]
        ConnectionTypeUnparseable,

        [Description("Thread Per Object is not a number.")]
        ThreadPerObjectUnparseable,

        [Description("Unknown attribute in class header file. Must be one of: VB_Name, VB_GlobalNameSpace, VB_Creatable, VB_PredeclaredId, VB_Exposed, VB_Description, VB_Ext_KEY")]
        UnknownAttribute,

        [Description("Error parsing header")]
        Header,

        [Description("No name in the attribute section of the VB6 file")]
        MissingNameAttribute,

        [Description("Keyword not found")]
        KeywordNotFound,

        [Description("Error parsing true/false from header. Must be a 0 (false), -1 (true), or 1 (true)")]
        TrueFalseOneZeroNegOneUnparseable,

        [Description("Error parsing the VB6 file contents")]
        FileContent,

        [Description("Max Threads is not a number.")]
        MaxThreadsUnparseable,

        [Description("No EndProperty found after BeginProperty")]
        NoEndProperty,

        [Description("No line ending after EndProperty")]
        NoLineEndingAfterEndProperty,

        [Description("Expected namespace after Begin keyword")]
        NoNamespaceAfterBegin,

        [Description("No dot found after namespace")]
        NoDotAfterNamespace,

        [Description("No User Control name found after namespace and '.'")]
        NoUserControlNameAfterDot,

        [Description("No space after Control kind")]
        NoSpaceAfterControlKind,

        [Description("No control name found after Control kind")]
        NoControlNameAfterControlKind,

        [Description("No line ending after Control name")]
        NoLineEndingAfterControlName,

        [Description("Unknown token")]
        UnknownToken,

        [Description("Title text was unparsable")]
        TitleUnparseable,

        [Description("Unable to parse hex color value")]
        HexColorParseError,

        [Description("Unknown control in control list")]
        UnknownControlKind,

        [Description("Property name is not a valid ASCII string")]
        PropertyNameAsciiConversionError,

        [Description("String is unterminated")]
        UnterminatedString,

        [Description("Unable to parse VB6 string.")]
        StringParseError,

        [Description("Property value is not a valid ASCII string")]
        PropertyValueAsciiConversionError,

        [Description("Key value pair format is incorrect")]
        KeyValueParseError,

        [Description("Namespace is not a valid ASCII string")]
        NamespaceAsciiConversionError,

        [Description("Control kind is not a valid ASCII string")]
        ControlKindAsciiConversionError,

        [Description("Qualified control name is not a valid ASCII string")]
        QualifiedControlNameAsciiConversionError,

        [Description("Variable names must be less than 255 characters in VB6.")]
        VariableNameTooLong,

        [Description("Internal Parser Error - please report this issue to the developers.")]
        InternalParseError,

        // Additional C# specific errors
        [Description("Unrecognized format for Object line")]
        UnrecognizedObjectFormat,
        [Description("Expected opening brace '{' for GUID in Object line")]
        ExpectedOpeningBraceForGuid,
        [Description("Expected closing quote '\"' for Object line")]
        ExpectedClosingQuote,
        [Description("Expected asterisk '*' for path-based Object line")]
        ExpectedAsterisk,
        [Description("Expected path separator '\\' after asterisk in Object line")]
        ExpectedPathSeparator,
        [Description("No END keyword found to close a BEGIN block.")]
        NoEndKeyword,
        [Description("No equals sign found after property key.")]
        NoEqualsAfterPropertyKey,
        [Description("A BeginProperty block was started but no matching EndProperty was found.")]
        UnterminatedPropertyGroup,
    }

    public class VB6ParseException : Exception
    {
        public string FileName { get; }
        public string SourceCode { get; } // Storing entire source code can be memory intensive for large files. Consider snippet.
        public int SourceOffset { get; }
        public int Column { get; }
        public int LineNumber { get; }
        public VB6ErrorKind Kind { get; }
        public PropertyErrorType? PropertyErrorDetail { get; } // For VB6ErrorKind.Property

        private static string GetEnumDescription(Enum value)
        {
            FieldInfo? fieldInfo = value.GetType().GetField(value.ToString());
            if (fieldInfo == null) return value.ToString(); // Should not happen for valid enum values
            DescriptionAttribute? attribute = fieldInfo.GetCustomAttribute<DescriptionAttribute>(false);
            return attribute != null ? attribute.Description : value.ToString();
        }

        public override string Message
        {
            get
            {
                string baseMessage = GetEnumDescription(Kind);
                if (Kind == VB6ErrorKind.Property && PropertyErrorDetail.HasValue)
                {
                    return $"{baseMessage}: {GetEnumDescription(PropertyErrorDetail.Value)}";
                }
                return baseMessage;
            }
        }

        // Main constructor
        public VB6ParseException(
            VB6ErrorKind kind,
            string fileName,
            string sourceCode,
            int sourceOffset,
            int column,
            int lineNumber,
            PropertyErrorType? propertyErrorDetail = null,
            Exception? innerException = null)
            : base(null, innerException) // Message is overridden
        {
            Kind = kind;
            FileName = fileName;
            SourceCode = sourceCode; 
            SourceOffset = sourceOffset;
            Column = column;
            LineNumber = lineNumber;
            PropertyErrorDetail = propertyErrorDetail;

            if (kind == VB6ErrorKind.Property && !propertyErrorDetail.HasValue)
            {
                // This implies a programming error if this constructor is called directly for Property.
                // Factory methods FromPropertyError and FromPropertyErrorWithoutStream should be used.
            }
        }

        // Constructor for Property errors (simplifies creation)
        public static VB6ParseException FromPropertyError(
            PropertyErrorType propertyErrorDetail,
            string fileName,
            string sourceCode,
            int sourceOffset,
            int column,
            int lineNumber,
            Exception? innerException = null)
        {
            return new VB6ParseException(VB6ErrorKind.Property, fileName, sourceCode, sourceOffset, column, lineNumber, propertyErrorDetail, innerException);
        }
        
        // Constructor for IO errors (simplifies creation)
        public static VB6ParseException FromIOException(
            VB6ErrorKind kind, // Should be ResourceFile or SourceFileError
            IOException ioException,
            string fileName,
            string sourceCode = "", // Default if source not read yet
            int sourceOffset = 0,
            int column = 0,
            int lineNumber = 0)
        {
            if (kind != VB6ErrorKind.ResourceFile && kind != VB6ErrorKind.SourceFileError)
            {
                throw new ArgumentException("Kind must be ResourceFile or SourceFileError for this factory method.", nameof(kind));
            }
            return new VB6ParseException(kind, fileName, sourceCode, sourceOffset, column, lineNumber, null, ioException);
        }

        public VB6ParseException(VB6ErrorKind kind, Exception? innerException = null) : this(kind, "unknown", string.Empty, 0, 0, 0, null, innerException)
        {
        }
        
        public static VB6ParseException FromPropertyErrorWithoutStream(PropertyErrorType propertyErrorDetail, Exception? innerException = null)
        {
            return new VB6ParseException(VB6ErrorKind.Property, "unknown", string.Empty, 0, 0, 0, propertyErrorDetail, innerException);
        }
    }
}