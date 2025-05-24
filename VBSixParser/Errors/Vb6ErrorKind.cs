namespace VB6Parse.Errors
{
    /// <summary>
    /// Specifies the general type of error that occurred during VB6 parsing.
    /// </summary>
    public enum Vb6ErrorKind
    {
        /// <summary>Error: "Property parsing error"</summary>
        /// <remarks>In Rust: Property(#[from] PropertyError)</remarks>
        Property,

        /// <summary>Error: "Resource file parsing error"</summary>
        /// <remarks>In Rust: ResourceFile(#[from] std::io::Error)</remarks>
        ResourceFile,

        /// <summary>Error: "Error reading the source file"</summary>
        /// <remarks>In Rust: SourceFileError(std::io::Error)</remarks>
        SourceFileError,

        /// <summary>Error: "The file contains more than a significant number of non-ASCII characters. This file was likely saved in a non-English character set. The vb6parse crate currently does not support non-english vb6 files."</summary>
        LikelyNonEnglishCharacterSet,

        /// <summary>Error: "The reference line has too many elements"</summary>
        ReferenceExtraSections,

        /// <summary>Error: "The reference line has too few elements"</summary>
        ReferenceMissingSections,

        /// <summary>Error: "The first line of a VB6 project file must be a project 'Type' entry."</summary>
        FirstLineNotProject,

        /// <summary>Error: "Line type is unknown."</summary>
        LineTypeUnknown,

        /// <summary>Error: "Project type is not Exe, OleDll, Control, or OleExe"</summary>
        ProjectTypeUnknown,

        /// <summary>Error: "Project lacks a version number."</summary>
        NoVersion,

        /// <summary>Error: "Project parse error while processing an Object line."</summary>
        NoObjects,

        /// <summary>Error: "Form parse error. No Form found in form file."</summary>
        NoForm,

        /// <summary>Error: "Parse error while processing Form attributes."</summary>
        AttributeParseError,

        /// <summary>Error: "Parse error while attempting to parse Form tokens."</summary>
        TokenParseError,

        /// <summary>Error: "Project parse error, failure to find BEGIN element."</summary>
        NoBegin,

        /// <summary>Error: "Project line entry is not ended with a recognized line ending."</summary>
        NoLineEnding,

        /// <summary>Error: "Unable to parse the Uuid"</summary>
        UnableToParseUuid,

        /// <summary>Error: "Unable to find a semicolon ';' in this line."</summary>
        NoSemicolonSplit,

        /// <summary>Error: "Unable to find an equal '=' in this line."</summary>
        NoEqualSplit,

        /// <summary>Error: "While trying to parse the offset into the resource file, no colon ':' was found."</summary>
        NoColonForOffsetSplit,

        /// <summary>Error: "No key value divider found in the line."</summary>
        NoKeyValueDividerFound,

        /// <summary>Error: "Unknown parser error"</summary>
        Unparseable, // Generic catch-all

        /// <summary>Error: "Major version is not a number."</summary>
        MajorVersionUnparseable,

        /// <summary>Error: "Unable to parse hex address from DllBaseAddress key"</summary>
        DllBaseAddressUnparseable,

        /// <summary>Error: "The Startup object is not a valid parameter. Must be a qouted startup method/object, \"(None)\", !(None)!, \"\", or \"!!\""</summary>
        StartupUnparseable,

        /// <summary>Error: "The Name parameter is invalid. Must be a qouted name, \"(None)\", !(None)!, \"\", or \"!!\""</summary>
        NameUnparseable, // Note: Also a PropertyError, consider if this specific context is different

        /// <summary>Error: "The CommandLine parameter is invalid. Must be a qouted command line, \"(None)\", !(None)!, \"\", or \"!!\""</summary>
        CommandLineUnparseable,

        /// <summary>Error: "The HelpContextId parameter is not a valid parameter line. Must be a qouted help context id, \"(None)\", !(None)!, \"\", or \"!!\""</summary>
        HelpContextIdUnparseable,

        /// <summary>Error: "Minor version is not a number."</summary>
        MinorVersionUnparseable,

        /// <summary>Error: "Revision version is not a number."</summary>
        RevisionVersionUnparseable,

        /// <summary>Error: "Unable to parse the value after ThreadingModel key"</summary>
        ThreadingModelUnparseable,

        /// <summary>Error: "ThreadingModel can only be 0 (Apartment Threaded), or 1 (Single Threaded)"</summary>
        ThreadingModelInvalid,

        /// <summary>Error: "No property name found after BeginProperty keyword."</summary>
        NoPropertyName,

        /// <summary>Error: "Unable to parse the RelatedDoc property line."</summary>
        RelatedDocLineUnparseable,

        /// <summary>Error: "AutoIncrement can only be a 0 (false) or a -1 (true)"</summary>
        AutoIncrementUnparseable,

        /// <summary>Error: "CompatibilityMode can only be a 0 (CompatibilityMode::NoCompatibility), 1 (CompatibilityMode::Project), or 2 (CompatibilityMode::CompatibleExe)"</summary>
        CompatibilityModeUnparseable,

        /// <summary>Error: "NoControlUpgrade can only be a 0 (UpgradeControls::Upgrade) or a 1 (UpgradeControls::NoUpgrade)"</summary>
        NoControlUpgradeUnparseable,

        /// <summary>Error: "ServerSupportFiles can only be a 0 (false) or a -1 (true)"</summary>
        ServerSupportFilesUnparseable,

        /// <summary>Error: "Comment line was unparsable"</summary>
        CommentUnparseable,

        /// <summary>Error: "PropertyPage line was unparsable"</summary>
        PropertyPageUnparseable,

        /// <summary>Error: "CompilationType can only be a 0 (false) or a -1 (true)"</summary>
        CompilationTypeUnparseable,

        /// <summary>Error: "OptimizationType can only be a 0 (FastCode) or 1 (SmallCode), or 2 (NoOptimization)"</summary>
        OptimizationTypeUnparseable,

        /// <summary>Error: "FavorPentiumPro(tm) can only be a 0 (false) or a -1 (true)"</summary>
        FavorPentiumProUnparseable,

        /// <summary>Error: "Designer line is unparsable"</summary>
        DesignerLineUnparseable,

        /// <summary>Error: "Form line is unparsable"</summary>
        FormLineUnparseable,

        /// <summary>Error: "UserControl line is unparsable"</summary>
        UserControlLineUnparseable,

        /// <summary>Error: "UserDocument line is unparsable"</summary>
        UserDocumentLineUnparseable,

        /// <summary>Error: "Period expected in version number"</summary>
        PeriodExpectedInVersionNumber,

        /// <summary>Error: "CodeViewDebugInfo can only be a 0 (false) or a -1 (true)"</summary>
        CodeViewDebugInfoUnparseable,

        /// <summary>Error: "NoAliasing can only be a 0 (false) or a -1 (true)"</summary>
        NoAliasingUnparseable,

        /// <summary>Error: "RemoveUnusedControlInfo can only be 0 (UnusedControlInfo::Retain) or -1 (UnusedControlInfo::Remove)"</summary>
        UnusedControlInfoUnparseable,

        /// <summary>Error: "BoundsCheck can only be a 0 (false) or a -1 (true)"</summary>
        BoundsCheckUnparseable,

        /// <summary>Error: "OverflowCheck can only be a 0 (false) or a -1 (true)"</summary>
        OverflowCheckUnparseable,

        /// <summary>Error: "FlPointCheck can only be a 0 (false) or a -1 (true)"</summary>
        FlPointCheckUnparseable,

        /// <summary>Error: "FDIVCheck can only be a 0 (PentiumFDivBugCheck::CheckPentiumFDivBug) or a -1 (PentiumFDivBugCheck::NoPentiumFDivBugCheck)"</summary>
        FDIVCheckUnparseable,

        /// <summary>Error: "UnroundedFP can only be a 0 (UnroundedFloatingPoint::DoNotAllow) or a -1 (UnroundedFloatingPoint::Allow)"</summary>
        UnroundedFPUnparseable,

        /// <summary>Error: "StartMode can only be a 0 (StartMode::StandAlone) or a 1 (StartMode::Automation)"</summary>
        StartModeUnparseable,

        /// <summary>Error: "Unattended can only be a 0 (Unattended::False) or a -1 (Unattended::True)"</summary>
        UnattendedUnparseable,

        /// <summary>Error: "Retained can only be a 0 (Retained::UnloadOnExit) or a 1 (Retained::RetainedInMemory)"</summary>
        RetainedUnparseable,

        /// <summary>Error: "Unable to parse the ShurtCut property."</summary>
        ShortCutUnparseable, // Spelling "ShurtCut" is from original Rust.

        /// <summary>Error: "DebugStartup can only be a 0 (false) or a -1 (true)"</summary>
        DebugStartupOptionUnparseable,

        /// <summary>Error: "UseExistingBrowser can only be a 0 (UseExistingBrowser::DoNotUse) or a -1 (UseExistingBrowser::Use)"</summary>
        UseExistingBrowserUnparseable,

        /// <summary>Error: "AutoRefresh can only be a 0 (false) or a -1 (true)"</summary>
        AutoRefreshUnparseable,

        /// <summary>Error: "Data control Connection type is not valid."</summary>
        ConnectionTypeUnparseable,

        /// <summary>Error: "Thread Per Object is not a number."</summary>
        ThreadPerObjectUnparseable,

        /// <summary>Error: "Unknown attribute in class header file. Must be one of: VB_Name, VB_GlobalNameSpace, VB_Creatable, VB_PredeclaredId, VB_Exposed, VB_Description, VB_Ext_KEY"</summary>
        UnknownAttribute,

        /// <summary>Error: "Error parsing header"</summary>
        Header,

        /// <summary>Error: "No name in the attribute section of the VB6 file"</summary>
        MissingNameAttribute,

        /// <summary>Error: "Keyword not found"</summary>
        KeywordNotFound,

        /// <summary>Error: "Error parsing true/false from header. Must be a 0 (false), -1 (true), or 1 (true)"</summary>
        TrueFalseOneZeroNegOneUnparseable,

        /// <summary>Error: "Error parsing the VB6 file contents"</summary>
        FileContent,

        /// <summary>Error: "Max Threads is not a number."</summary>
        MaxThreadsUnparseable,

        /// <summary>Error: "No EndProperty found after BeginProperty"</summary>
        NoEndProperty,

        /// <summary>Error: "No line ending after EndProperty"</summary>
        NoLineEndingAfterEndProperty,

        /// <summary>Error: "Expected namespace after Begin keyword"</summary>
        NoNamespaceAfterBegin,

        /// <summary>Error: "No dot found after namespace"</summary>
        NoDotAfterNamespace,

        /// <summary>Error: "No User Control name found after namespace and '.'"</summary>
        NoUserControlNameAfterDot,

        /// <summary>Error: "No space after Control kind"</summary>
        NoSpaceAfterControlKind,

        /// <summary>Error: "No control name found after Control kind"</summary>
        NoControlNameAfterControlKind,

        /// <summary>Error: "No line ending after Control name"</summary>
        NoLineEndingAfterControlName,

        /// <summary>Error: "Unknown token"</summary>
        UnknownToken,

        /// <summary>Error: "Title text was unparsable"</summary>
        TitleUnparseable,

        /// <summary>Error: "Unable to parse hex color value"</summary>
        HexColorParseError,

        /// <summary>Error: "Unknown control in control list"</summary>
        UnknownControlKind,

        /// <summary>Error: "Property name is not a valid ASCII string"</summary>
        PropertyNameAsciiConversionError,

        /// <summary>Error: "String is unterminated"</summary>
        UnterminatedString,

        /// <summary>Error: "Unable to parse VB6 string."</summary>
        StringParseError,

        /// <summary>Error: "Property value is not a valid ASCII string"</summary>
        PropertyValueAsciiConversionError,

        /// <summary>Error: "Key value pair format is incorrect"</summary>
        KeyValueParseError,

        /// <summary>Error: "Namespace is not a valid ASCII string"</summary>
        NamespaceAsciiConversionError,

        /// <summary>Error: "Control kind is not a valid ASCII string"</summary>
        ControlKindAsciiConversionError,

        /// <summary>Error: "Qualified control name is not a valid ASCII string"</summary>
        QualifiedControlNameAsciiConversionError,

        /// <summary>Error: "Variable names must be less than 255 characters in VB6."</summary>
        VariableNameTooLong,

        /// <summary>Error: "Internal Parser Error - please report this issue to the developers."</summary>
        InternalParseError
    }
}