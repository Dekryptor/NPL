# VBSix Project Structure

VBSix/
├── VBSix.sln						  (...)
├── VBSix.csproj                      (Standard C# project file - not generated here)
├── Errors/
│   └── Errors.cs                     (Contains VB6ErrorKind, PropertyErrorType, VB6ParseException)
├── Language/
│   ├── Controls/
│   │   ├── VB6Control.cs             (...)
│   │   ├── VB6CtrlEnums.cs           (All control-related enums like Appearance, BorderStyle, etc.)
│   │   ├── CheckBox.cs               (CheckBoxProperties class)
│   │   ├── ComboBox.cs               (ComboBoxProperties class)
│   │   ├── CommandButton.cs          (CommandButtonProperties class)
│   │   ├── CustomCtrl.cs             (CustomControlProperties class)
│   │   ├── DataCtrl.cs               (DataProperties class)
│   │   ├── DirListBox.cs             (DirListBoxProperties class)
│   │   ├── DriveListBox.cs           (DriveListBoxProperties class)
│   │   ├── FileListBox.cs            (FileListBoxProperties class)
│   │   ├── Form.cs                   (FormProperties class)
│   │   ├── Frame.cs                  (FrameProperties class)
│   │   ├── Image.cs                  (ImageProperties class)
│   │   ├── Label.cs                  (LabelProperties class)
│   │   ├── Line.cs                   (LineProperties class)
│   │   ├── ListBox.cs                (ListBoxProperties class)
│   │   ├── ListView.cs               (ListViewProperties class)
│   │   ├── MDIForm.cs                (MDIFormProperties class)
│   │   ├── Menus.cs                  (MenuProperties class, VB6MenuControl class)
│   │   ├── OLE.cs                    (OLEProperties class, OleVerbInfo class)
│   │   ├── OptionButton.cs           (OptionButtonProperties class)
│   │   ├── PictureBox.cs             (PictureBoxProperties class)
│   │   ├── ScrollBar.cs              (ScrollBarProperties class)
│   │   ├── Shape.cs                  (ShapeProperties class)
│   │   ├── TextBox.cs                (TextBoxProperties class)
│   │   ├── UserControl.cs            (UserControlProperties class)
│   │   ├── RichTextBox.cs            (RichTextBoxProperties class)
│   │   ├── Winsock.cs                (WinsockProperties class)
│   │   └── Timer.cs                  (TimerProperties class)
│   ├── VB6Color.cs                   (VB6Color class and predefined color constants)
│	├── VB6Tokens.cs                  (...)
│   └── Tokens.cs                     (VB6Token class, VB6TokenType enum)
├── Parsers/
│   ├── HeaderData.cs                 (VB6FileFormatVersion, HeaderKind, VB6FileAttributes, VB6Attribute)
│   ├── HeaderParser.cs               (Static class for parsing common header sections)
│   ├── VB6Stream.cs                  (VB6Stream class for managing input byte stream, VB6StreamCheckpoint struct)
│   ├── Properties.cs                 (Properties class, a wrapper around Dictionary for control properties from .frm)
│   ├── VB6.cs                        (Core VB6 tokenization logic, including keyword/symbol parsing helpers, IsEnglishCode)
│   ├── CompilationSettings.cs        (Enums and classes related to VBP compilation settings)
│   ├── VB6ObjectReference.cs         (Class representing `Object=` lines in VBP/FRM)
│   ├── VB6ClassFile.cs               (Parser and data structures for .cls files, including VB6ClassHeader, VB6ClassProperties, and related enums)
│   ├── VB6ModuleFile.cs              (Parser and data structures for .bas files)
│   ├── VB6UserControlFile.cs         (Parser and data structures for .ctl files)
│   ├── VB6FormFile.cs                (Parser and data structures for .frm files. This includes PropertyValue, VB6FullyQualifiedName, VB6PropertyGroup, VB6Control, VB6ControlKind and its variants like VB6FormKind, VB6CheckBoxKind, etc.)
│   └── Project.cs                    (Parser and data structures for .vbp files, including project-specific enums and classes like VB6ProjectReferenceBase and its derivatives)
└── Utilities/
│   ├── PropertyParserHelpers.cs      (Extension methods for IDictionary<string, PropertyValue> to get typed properties)
│   └── TwipsConverter.cs             (Static class for converting between Twips and other units like Pixels)
├── Tests/
│   ├── TestProj.vbp
│   ├── TestClass.cls
│   └── TestModule.bas