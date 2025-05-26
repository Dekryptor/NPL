# TODO List for VBSix

This file tracks pending tasks, improvements, and potential future features for the VBSix C# library.

**Legend:**
*   `[ ]` - To Do
*   `[/]` - In Progress / Partially Done
*   `[x]` - Done (or sufficiently addressed for current scope)

## Core Parser Enhancements

*   [ ] **FRM Parser - `ParseControlBlock` and `BuildControl`**:
    *   [/] **`ParseControlBlock` Logic**: Core recursive structure is in place. Needs extensive testing with complex forms, deeply nested controls, and edge cases for property/group parsing.
    *   [/] **`BuildControl` Logic**: Basic structure exists. Systematically ensure all standard VB6 control kinds are correctly instantiated. Verify population of all `TypedProperties` for each control kind using `PropertyParserHelpers`.
    *   [/] **FRX Resource Resolution in `BuildControl`**:
        *   Ensure `TextRTF` for `RichTextBox` is correctly resolved.
        *   Systematically add FRX resolution for `Picture`, `Icon`, `MouseIcon`, `DragIcon`, `List`, `ItemData`, `ObjectVerbs`, etc., for all applicable control kinds.
        *   Ensure `Resolved...Data` fields in `VB6...Kind` classes are populated.
    *   [ ] Thoroughly test parsing of forms with various ActiveX controls (beyond RichTextBox and Winsock if more are specifically handled).
    *   [ ] Test forms with control arrays extensively.
*   [/] **UserControl Parsing (`.ctl` files)**:
    *   [x] Define `UserControlProperties` and `VB6UserControlKind` (Done based on recent prompts).
    *   [/] Implement `VB6UserControlFile.cs` parser. Basic structure exists, leveraging `VB6FormFile`'s helpers.
    *   [ ] Verify parsing of UserControl-specific attributes (e.g., `VB_ControlTooOle`, `VB_ToolboxBitmap` link to `.ctx`, `VB_Public`, `VB_WindowLess`).
    *   [ ] Ensure `.ctx` file resolution for `ToolboxBitmap` works with `DefaultResourceFileResolver` (may need to add `.ctx` to recognized extensions in regex if not already).
*   [ ] **UserDocument Parsing (`.dob` files for ActiveX Documents)**:
    *   [ ] Implement `VB6UserDocumentFile.cs` parser.
    *   [ ] Define `UserDocumentProperties` and `VB6UserDocumentKind`.
*   [ ] **PropertyPage Parsing (`.pag` files)**:
    *   [ ] Implement `VB6PropertyPageFile.cs` parser.
    *   [ ] Define `PropertyPageProperties` and `VB6PropertyPageKind`.
*   [ ] **Designer Parsing (`.dsr` files - Data Environment, Data Report, WebClass, DHTMLPage)**:
    *   [ ] Initial investigation: Identify common header structure (`VERSION`, `Attribute VB_Designer = True`).
    *   [ ] Determine strategy for parsing `BEGIN...END` blocks (generic storage vs. dedicated parsers per designer type). This is a large and complex task.
*   [ ] **Encoding Handling**:
    *   [/] Currently uses `Encoding.Default` or `Encoding.GetEncoding(1252)`. This is a reasonable default for many Western VB6 files.
    *   [ ] Consider adding an optional parameter to top-level `Parse` methods to allow users to specify the input encoding.
    *   [ ] Review `IsEnglishCode` heuristic for robustness or if it should be optional.
*   [ ] **VB6Stream Enhancements**:
    *   [/] Line/column tracking seems mostly fine based on current implementation. Continuous review during testing is good.
    *   [ ] Current `FindSlice` and `CompareSlice` overloads are basic. Evaluate if more advanced or byte-span based versions are needed for performance or flexibility.
*   [ ] **Error Reporting**:
    *   [/] `VB6ParseException` with contextual info is in place.
    *   [ ] Continuously refine error messages for clarity and specificity during testing and as new parsing logic is added. Ensure `CreateExceptionFromCheckpoint` is used effectively.

## Language Feature Support

*   [ ] **VB6 Tokenizer (`Vb6.cs`)**:
    *   [/] Core keywords, symbols, literals, and identifiers are handled.
    *   [ ] Verify handling of all less common VB6 operators or multi-character symbols if any were missed.
    *   [x] Numeric literals with type suffixes and hex/octal prefixes seem to be covered by `NumberParse`. Test thoroughly.
    *   [ ] Test date literals (`#date#`) parsing specifically (currently might parse as individual tokens: `#`, `date string`, `#`).
    *   [ ] Thoroughly test line continuation (`_`) in various contexts (end of line, before comment, mid-statement).
*   [ ] **Control Property Parsing (`XYZProperties` classes)**:
    *   [/] Constructors taking `IDictionary<string, PropertyValue>` are implemented for many controls.
    *   [ ] Systematically review each `XYZProperties` class against VB6 documentation and `.frm` examples to ensure all common design-time properties are included and parsed correctly.
    *   [ ] Double-check default values in C# property classes against VB6 IDE defaults.
    *   [/] `Font` property groups are parsed by `ParsePropertyGroup` and stored in `VB6...Kind.PropertyGroups`. Consider adding a typed `FontProperties` helper if direct access to font attributes is frequently needed.
*   [ ] **Project File (`.vbp`) Properties**:
    *   [/] `Project.cs` handles many common VBP keys.
    *   [ ] Review against a comprehensive list of VBP file properties for completeness (especially for different project types like ActiveX DLL/EXE/Control).
    *   [x] Numeric vs. string vs. boolean-like (-1/0) parsing is generally handled.
*   [ ] **ActiveX Control References in VBP/FRM**:
    *   [/] `HeaderParser.ParseObjectLine` and `Project.ParseObjectReference` handle `Object=` lines.
    *   [ ] Test with various OCX/DLL references, including those with and without quoted filenames or GUIDs.

## FRX Resource Resolving

*   [x] **`DefaultResourceFileResolver`**:
    *   [/] Core logic, handling multiple FRX record header types.
    *   [ ] Needs extensive testing with diverse `.frx` (and `.ctx`) files:
        *   [ ] Pictures (Bitmaps, JPEGs, GIFs, Icons, Cursors).
        *   [ ] ListBox/ComboBox `List` and `ItemData`.
        *   [ ] RichTextBox `TextRTF`.
        *   [ ] OLE `ObjectData`, `ObjectVerbs`.
        *   [ ] Custom control data.
    *   [ ] Verify the "off-by-one" hack logic with actual FRX files.
    *   [ ] Improve error messages if an FRX record is detected but malformed.
*   [ ] **Image Handling**:
    *   [/] Resolved image data is currently `byte[]`.
    *   [ ] *Optional Feature*: Integrate with `System.Drawing.Common` (Windows-only .NET Core/5+) or a cross-platform library like `SixLabors.ImageSharp` to convert `ResolvedPictureData`, `ResolvedIconData`, etc., into usable image objects. This would likely be a separate utility or an optional part of the `Kind` classes.

## Advanced Features & Future Scope

*   [ ] **Code Analysis**: Basic metrics (line counts, control counts). More advanced: finding declarations, usage, call graphs.
*   [ ] **Abstract Syntax Tree (AST)**: Implement a full VB6 AST parser from the token stream. (Major undertaking)
*   [ ] **Code Generation/Transformation**: Dependent on AST. (Major undertaking)
*   [ ] **Language Server Protocol (LSP)**: Potential future integration for IDE support.
*   [ ] **Performance Optimization**: Profile and optimize if parsing large projects becomes slow.
*   [ ] **User Interface**: A simple diagnostic UI to load and inspect parsed VB6 files.
*   [ ] **VB Group File Parsing (`.vbg`)**: Add parser for `.vbg` files to understand project groups.
*   [ ] **Binary FRM File Parsing**: VB6 could also save forms in a binary format (`.bi?`). This is a completely different parsing challenge. (Low priority unless needed).

## Documentation & Testing

*   [ ] **Unit Tests**:
    *   [ ] **Critical Path**: Add comprehensive unit tests for `VB6Stream`, `HeaderParser`, `Vb6` (tokenizer), `PropertyParserHelpers`.
    *   [ ] Create a suite of diverse test files (`.frm`, `.cls`, `.bas`, `.vbp`, `.frx`, `.ctx`) that cover:
        *   All standard controls and their common properties.
        *   Nested controls.
        *   Control arrays.
        *   Various FRX resource types.
        *   Different project types and VBP settings.
        *   Edge cases (empty files, malformed lines, unusual syntax if any).
        *   Files with and without code sections.
*   [ ] **API Documentation**:
    *   [/] Basic XML documentation comments are present in some classes.
    *   [ ] Systematically add/review XML docs for all public types and members.
    *   [x] Update `README.md` with basic structure and usage examples (Done). Needs expansion as features stabilize.
*   [ ] **Sample Projects**: Create 1-2 small C# console applications demonstrating library usage for common parsing tasks.
