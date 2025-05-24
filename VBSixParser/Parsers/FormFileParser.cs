// Namespace: VB6Parse.Parsers
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using VB6Parse.Model;
using VB6Parse.Language.Controls;
using VB6Parse.Utilities; // For PropertyParsingHelpers (indirectly via XProperties)

namespace VB6Parse.Parsers
{
    public class FormFileParser
    {
        private Vb6StreamReader _streamReader;
        private Vb6FormFile _formFile;
        private string _filePath;
        private string _frxFilePath; // Path to the .frx file, if one exists

        private Stack<ControlParseState> _controlStack;

        public delegate byte[] FrxResourceResolverDelegate(string frxFilePath, int offset);
        private readonly FrxResourceResolverDelegate _frxResolver;

        // Regexes (some might need refinement for comments or complex quoted strings)
        private static readonly Regex VersionRegex = new Regex(@"^VERSION\s+(\d+)\.(\d+)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private static readonly Regex ObjectRegex = new Regex(@"^Object\s*=\s*(.+)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private static readonly Regex AttributeRegex = new Regex(@"^\s*Attribute\s+([^=]+)\s*=\s*(.+)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        
        // Captures: 1=TypeFullName (e.g. VB.Form), 2=ControlName
        private static readonly Regex BeginControlRegex = new Regex(@"^\s*Begin\s+([A-Za-z0-9_\.]+)\s+([A-Za-z_][A-Za-z0-9_]*)", RegexOptions.Compiled | RegexOptions.CultureInvariant); 
        private static readonly Regex EndControlRegex = new Regex(@"^\s*End\s*$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        
        // Captures: 1=GroupName, 2=OptionalGUID
        private static readonly Regex BeginPropertyRegex = new Regex(@"^\s*BeginProperty\s+([A-Za-z_][A-Za-z0-9_]*)(?:\s+\{([0-9A-Fa-f\-]+)\})?", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private static readonly Regex EndPropertyRegex = new Regex(@"^\s*EndProperty\s*$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        
        // Captures: 1=PropertyName, 2=PropertyValue
        private static readonly Regex PropertyLineRegex = new Regex(@"^\s*([A-Za-z_][A-Za-z0-9_#]*)\s*=\s*(.+)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        
        // Captures: 1=PropertyName, 2=FRXFileName, 3=HexOffset
        private static readonly Regex ResourcePropertyLineRegex = new Regex(@"^\s*([A-Za-z_][A-Za-z0-9_#]*)\s*=\s*""([^""]+\.(frx|ctx|pag|dsx|vbx|oca))""\s*:\s*([0-9A-Fa-f]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private class ControlParseState
        {
            public Vb6Control ControlBeingBuilt { get; }
            public string TypeFullName { get; } // e.g., "VB.Form"
            public Dictionary<string, string> TextualProperties { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, byte[]> BinaryProperties { get; } = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            public List<Vb6PropertyGroup> PropertyGroups { get; } = new List<Vb6PropertyGroup>();

            public ControlParseState(string typeFullName, string name)
            {
                TypeFullName = typeFullName;
                ControlBeingBuilt = new Vb6Control { Name = name }; 
                // Index parsing from name (e.g., Option1(0)) would happen when finalizing control.
            }
        }

        public FormFileParser(FrxResourceResolverDelegate frxResolver = null)
        {
            _frxResolver = frxResolver ?? DefaultFrxResolver;
        }
        
        private static byte[] DefaultFrxResolver(string frxFilePath, int offset)
        {
            // This is a placeholder. Real FRX parsing is complex.
            // See VB6Parse_Rust/src/parsers/form.rs resource_file_resolver function.
            if (string.IsNullOrEmpty(frxFilePath) || !File.Exists(frxFilePath)) return null;
            try
            {
                // The Rust implementation has sophisticated logic to read FRX records.
                // This simplified version just reads a chunk, which is incorrect for most FRX data.
                // For a robust solution, the Rust FRX parsing logic needs to be ported.
                byte[] frxBytes = File.ReadAllBytes(frxFilePath);
                if (offset >= frxBytes.Length) return null;

                // Determine record length based on FRX format (highly simplified)
                // This is a placeholder and will not work correctly for many FRX files.
                // The Rust code checks for various header types (0xFF, "lt\0\0", etc.)
                // and reads lengths from those headers.
                int lengthToRead = 0;
                if (frxBytes[offset] == 0xFF && offset + 2 < frxBytes.Length) // Simplified 16-bit record
                {
                    lengthToRead = BitConverter.ToUInt16(frxBytes, offset + 1);
                    offset += 3; // Skip 0xFF and 2-byte length
                }
                else if (offset + 4 <= frxBytes.Length) // Simplified 32-bit length record (very common)
                {
                     // This is often a size *preceded* by a header, not the first bytes at offset.
                     // The offset in FRM points to start of a record structure.
                     // The Rust code has more detail. For now, this is a guess.
                    lengthToRead = (int)BitConverter.ToUInt32(frxBytes, offset); // This is likely wrong.
                    offset +=4; // skip length
                } else {
                    return null; // Cannot determine length
                }


                if (lengthToRead <= 0 || offset + lengthToRead > frxBytes.Length) return new byte[0]; // Empty or invalid

                byte[] data = new byte[lengthToRead];
                Array.Copy(frxBytes, offset, data, 0, lengthToRead);
                return data;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error reading FRX file {frxFilePath} at offset {offset}: {ex.Message}");
                return null;
            }
        }


        public Vb6FormFile Parse(string filePath)
        {
            _filePath = filePath;
            _frxFilePath = Path.ChangeExtension(filePath, ".frx"); // Common convention
            if (!File.Exists(_frxFilePath)) _frxFilePath = null; // Reset if no FRX

            _formFile = new Vb6FormFile(filePath);
            _controlStack = new Stack<ControlParseState>();
            
            _streamReader = new Vb6StreamReader(filePath, Encoding.GetEncoding("Windows-1252")); 

            ParseGlobalHeader();
            ParseControlBlocksAndAttributes(); // Combines control parsing and form-level attributes
            ParseCodeModule(); // Parses what's left as code

            _streamReader.Close();
            return _formFile;
        }

        private void ParseGlobalHeader()
        {
            string line;
            while ((line = _streamReader.ReadLine()) != null)
            {
                Match match;
                string trimmedLine = line.Trim();

                if (string.IsNullOrWhiteSpace(trimmedLine)) continue;

                match = VersionRegex.Match(trimmedLine);
                if (match.Success)
                {
                    _formFile.FormatVersion.Major = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
                    _formFile.FormatVersion.Minor = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
                    continue;
                }

                match = ObjectRegex.Match(trimmedLine);
                if (match.Success)
                {
                    var objRef = new Vb6ObjectReference(match.Groups[1].Value.Trim());
                    // TODO: Parse the raw string in objRef to populate its fields (GUID, Version, File)
                    _formFile.ObjectReferences.Add(objRef);
                    continue;
                }

                if (trimmedLine.StartsWith("Begin", StringComparison.OrdinalIgnoreCase))
                {
                    _streamReader.PushBack(line); 
                    break;
                }
                // Silently ignore other lines in header, or log as unexpected.
            }
        }
        
        private void ParseControlBlocksAndAttributes()
        {
            string line;
            while ((line = _streamReader.ReadLine()) != null)
            {
                string trimmedLine = line.Trim();
                if (string.IsNullOrWhiteSpace(trimmedLine)) continue;

                Match match;

                match = BeginControlRegex.Match(trimmedLine);
                if (match.Success)
                {
                    string typeFullName = match.Groups[1].Value;
                    string name = match.Groups[2].Value;
                    var newState = new ControlParseState(typeFullName, name);
                    _controlStack.Push(newState);
                    continue;
                }

                match = EndControlRegex.Match(trimmedLine);
                if (match.Success)
                {
                    if (_controlStack.Count > 0)
                    {
                        ControlParseState completedState = _controlStack.Pop();
                        Vb6Control builtControl = BuildControlFromState(completedState);

                        if (_controlStack.Count == 0) // This was the top-level control (the Form itself)
                        {
                            _formFile.Form = builtControl;
                            // After the main Form's "End", subsequent lines are Attributes for the Form or code.
                            // We can now proceed to parse Form-level attributes.
                            ParseFormAttributes();
                            return; // Control block parsing is done.
                        }
                        else // This was a nested control
                        {
                            ControlParseState parentState = _controlStack.Peek();
                            if (builtControl.SpecificProperties is MenuProperties || builtControl.TypeKind == "Menu") // Check if it's a menu
                            {
                                // Convert Vb6Control to Vb6MenuControl and add to parent's menus
                                parentState.ControlBeingBuilt.Menus.Add(ConvertToMenuControl(builtControl));
                            }
                            else
                            {
                                parentState.ControlBeingBuilt.Controls.Add(builtControl);
                            }
                        }
                    }
                    continue;
                }
                
                // Must be inside a Begin...End block for these
                if (_controlStack.Count > 0)
                {
                    ControlParseState currentState = _controlStack.Peek();

                    match = BeginPropertyRegex.Match(trimmedLine);
                    if (match.Success)
                    {
                        string groupName = match.Groups[1].Value;
                        Guid? groupGuid = match.Groups[2].Success ? Guid.Parse(match.Groups[2].Value) : (Guid?)null;
                        currentState.PropertyGroups.Add(ParsePropertyGroup(groupName, groupGuid));
                        continue;
                    }

                    match = ResourcePropertyLineRegex.Match(trimmedLine);
                    if (match.Success)
                    {
                        string propName = match.Groups[1].Value;
                        // string frxFile = match.Groups[2].Value; // Could verify against _frxFilePath
                        int offset = int.Parse(match.Groups[4].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                        
                        byte[] data = _frxResolver?.Invoke(_frxFilePath, offset);
                        if (data != null)
                        {
                            currentState.BinaryProperties[propName] = data;
                        }
                        continue;
                    }

                    match = PropertyLineRegex.Match(trimmedLine);
                    if (match.Success)
                    {
                        string propName = match.Groups[1].Value;
                        string propValue = match.Groups[2].Value.Trim();
                        // Values are often quoted, remove leading/trailing quotes.
                        // VB also uses "" inside a string to represent a single quote.
                        if (propValue.StartsWith("\"") && propValue.EndsWith("\""))
                        {
                            propValue = propValue.Substring(1, propValue.Length - 2).Replace("\"\"", "\"");
                        }
                        currentState.TextualProperties[propName] = propValue;
                        continue;
                    }
                     // Line is not a recognized control structure or property inside a Begin block.
                     // Could be a comment (VB uses ' at start of line or after code)
                     // Simple comment stripping:
                    int commentCharPos = trimmedLine.IndexOf('\'');
                    if (commentCharPos == 0) continue; // Whole line is a comment
                    if (commentCharPos > 0) {
                        // Potentially a property line with a trailing comment.
                        // Regexes should ideally handle this. For now, this is a fallback.
                        // This part needs more robust handling of comments within property lines.
                        // The current PropertyLineRegex will grab the comment as part of value.
                        // A better regex or pre-stripping comments is needed.
                    }

                } else {
                     // Lines outside any Begin...End block, and not a Begin itself.
                     // This could be an Attribute line for the Form, if we are past the Form's End block.
                     // Or it's an error / unexpected line.
                     // The logic to switch to ParseFormAttributes handles this.
                }
            }
        }
        
        private Vb6PropertyGroup ParsePropertyGroup(string groupName, Guid? groupGuid)
        {
            var propertyGroup = new Vb6PropertyGroup(groupName) { Guid = groupGuid };
            string line;
            while ((line = _streamReader.ReadLine()) != null)
            {
                string trimmedLine = line.Trim();
                if (string.IsNullOrWhiteSpace(trimmedLine)) continue;

                if (EndPropertyRegex.IsMatch(trimmedLine))
                {
                    return propertyGroup;
                }

                Match match;
                match = BeginPropertyRegex.Match(trimmedLine); // Nested property group
                if (match.Success)
                {
                    string nestedGroupName = match.Groups[1].Value;
                    Guid? nestedGroupGuid = match.Groups[2].Success ? Guid.Parse(match.Groups[2].Value) : (Guid?)null;
                    propertyGroup.Properties[nestedGroupName] = ParsePropertyGroup(nestedGroupName, nestedGroupGuid);
                    continue;
                }
                
                // Resource properties within property groups are less common but possible
                match = ResourcePropertyLineRegex.Match(trimmedLine);
                if (match.Success)
                {
                    string propName = match.Groups[1].Value;
                    int offset = int.Parse(match.Groups[4].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    byte[] data = _frxResolver?.Invoke(_frxFilePath, offset);
                    if (data != null)
                    {
                        // Storing byte[] directly. Vb6PropertyGroup.Properties is Dictionary<string, object>
                        propertyGroup.Properties[propName] = data; 
                    }
                    continue;
                }

                match = PropertyLineRegex.Match(trimmedLine);
                if (match.Success)
                {
                    string propName = match.Groups[1].Value;
                    string propValue = match.Groups[2].Value.Trim();
                    if (propValue.StartsWith("\"") && propValue.EndsWith("\""))
                    {
                        propValue = propValue.Substring(1, propValue.Length - 2).Replace("\"\"", "\"");
                    }
                    propertyGroup.Properties[propName] = propValue;
                    continue;
                }
                // Log or throw error for unexpected line in property group
            }
            throw new FormatException($"Unterminated 'BeginProperty {groupName}' block at line {_streamReader.LineNumber}.");
        }

        private Vb6Control BuildControlFromState(ControlParseState state)
        {
            Vb6Control control = state.ControlBeingBuilt; // Name is already set

            // Parse TypeFullName into Namespace and Kind
            string[] typeParts = state.TypeFullName.Split('.');
            control.TypeNamespace = typeParts.Length > 1 ? typeParts[0] : "VB"; // Default to VB if no namespace
            control.TypeKind = typeParts.Length > 1 ? typeParts[1] : typeParts[0];

            // Handle control arrays, e.g., "Command1(0)" -> Name="Command1", Index=0
            Match nameIndexMatch = Regex.Match(control.Name, @"^([A-Za-z_][A-Za-z0-9_]*)_?\((\d+)\)$");
            if (nameIndexMatch.Success)
            {
                control.Name = nameIndexMatch.Groups[1].Value;
                control.Index = int.Parse(nameIndexMatch.Groups[2].Value, CultureInfo.InvariantCulture);
            }

            // Assign specific properties based on TypeKind
            // This is where a factory or a large switch statement would go.
            // For now, a simplified example:
            switch (control.TypeKind)
            {
                case "Form":
                    var formProps = new FormProperties();
                    formProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = formProps;
                    break;
                case "CommandButton":
                    var cmdProps = new CommandButtonProperties();
                    cmdProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = cmdProps;
                    break;
                case "CheckBox":
                    var chkProps = new CheckBoxProperties();
                    chkProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = chkProps;
                    break;
                case "Label":
                    var lblProps = new LabelProperties();
                    lblProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = lblProps;
                    break;
                case "TextBox":
                    var txtProps = new TextBoxProperties();
                    txtProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = txtProps;
                    break;
                case "Menu":
                    var menuProps = new MenuProperties();
                    menuProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = menuProps;
                    break;
                case "OptionButton":
					var optProps = new MenuProperties();
                    optProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = optProps;
					break;
                case "ListBox":
					var lbProps = new ListBoxProperties();
                    lbProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = lbProps;
					break;
                case "ComboBox":
					var cmbProps = new ComboBoxProperties();
                    cmbProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = cmbProps;
					break;
                case "Frame":
					var frameProps = new FrameProperties();
                    frameProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = frameProps;
					break;
                case "Timer":
					var timerProps = new TimerProperties();
                    timerProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = timerProps;
					break;
                case "PictureBox":
					var picProps = new PictureBoxProperties();
                    picProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = picProps;
					break;
                case "Image":
					var imgProps = new ImageProperties();
                    imgProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = imgProps;
					break;
                case "Line":
					var lineProps = new LineProperties();
                    lineProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = lineProps;
					break;
                case "Shape":
					var shapeProps = new ShapeProperties();
                    shapeProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = shapeProps;
					break;
                case "HScrollBar":
					var hscrollProps = new HScrollBarProperties();
                    hscrollProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = hscrollProps;
					break;
                case "VScrollBar":
					var vscrollProps = new VScrollBarProperties();
                    vscrollProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = vscrollProps;
					break;
                default:
                    // For custom controls (e.g., MSComctlLib.ListView) or unhandled standard controls
                    // The TypeFullName (e.g. "MSComctlLib.ListViewCtrl") is passed to CustomControlProperties.
                    var customProps = new CustomControlProperties(state.TypeFullName);
                    customProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = customProps;
                    break;
            }
            
            // Common properties like Tag (if not handled by SpecificProperties)
            if (state.TextualProperties.TryGetValue("Tag", out var tagValue))
            {
                control.Tag = tagValue;
            }

            return control;
        }
        
        private Vb6MenuControl ConvertToMenuControl(Vb6Control control)
        {
            if (!(control.SpecificProperties is MenuProperties menuProps))
            {
                // This case should ideally not happen if type checking is correct before calling.
                // Create default MenuProperties if somehow it's missing.
                menuProps = new MenuProperties { Caption = control.Name }; // Basic fallback
            }

            var menuControl = new Vb6MenuControl
            {
                Name = control.Name,
                Index = control.Index,
                Tag = control.Tag,
                Properties = menuProps,
                // SubMenus are already Vb6MenuControl if parsed recursively correctly
                SubMenus = control.Menus 
            };
            return menuControl;
        }


        private void ParseFormAttributes()
        {
            // Called after the main Form's "End" is processed.
            // Reads Attribute lines for the Form itself.
            string line;
            while ((line = _streamReader.ReadLine()) != null)
            {
                string trimmedLine = line.Trim();
                if (string.IsNullOrWhiteSpace(trimmedLine)) continue;

                Match attrMatch = AttributeRegex.Match(trimmedLine);
                if (attrMatch.Success)
                {
                    _formFile.FileAttributes.AddAttribute(attrMatch.Groups[1].Value.Trim(), attrMatch.Groups[2].Value.Trim().Trim('"'));
                }
                else
                {
                    // First line that is not an attribute is considered start of code.
                    _streamReader.PushBack(line); // Push back to be read by ParseCodeModule
                    break; 
                }
            }
        }

        private void ParseCodeModule()
        {
            StringBuilder codeModuleBuilder = new StringBuilder();
            string line;
            while ((line = _streamReader.ReadLine()) != null)
            {
                codeModuleBuilder.AppendLine(line);
            }
            _formFile.CodeModuleText = codeModuleBuilder.ToString();
        }
    }


    // Definition for XProperties base or interface, and specific properties classes
    // would require each XProperties class (e.g., FormProperties, CommandButtonProperties)
    // to have a method like:
    // public virtual void PopulateFromDictionaries(
    //    IReadOnlyDictionary<string, string> textualProps,
    //    IReadOnlyDictionary<string, byte[]> binaryProps,
    //    IReadOnlyList<Vb6PropertyGroup> propertyGroups)
    // {
    //    // Default implementation or make it abstract
    //    this.Caption = PropertyParsingHelpers.GetString(textualProps, "Caption", "");
    //    // ... use PropertyParsingHelpers for other common ones ...
    //    // Handle Font property from propertyGroups or flat textualProps
    //    var fontGroup = propertyGroups.FirstOrDefault(pg => pg.Name.Equals("Font", StringComparison.OrdinalIgnoreCase));
    //    if (fontGroup != null) this.Font = PropertyParsingHelpers.ParseFontFromGroup(fontGroup);
    //    else this.Font = PropertyParsingHelpers.ParseFontFromFlatProperties(textualProps, "Font");
    // }
    // This method would be implemented in each specific XProperties class.

    /// <summary>
    /// Helper class to read lines and allow pushing a line back onto the stream.
    /// Also tracks line numbers for error reporting.
    /// </summary>
    public class Vb6StreamReader : IDisposable
    {
        private readonly StreamReader _reader;
        private string _pushedBackLine;
        public int LineNumber { get; private set; }
        public string FilePath { get; }

        public Vb6StreamReader(string filePath, Encoding encoding)
        {
            FilePath = filePath;
            _reader = new StreamReader(filePath, encoding);
            LineNumber = 0;
        }

        public string ReadLine()
        {
            if (_pushedBackLine != null)
            {
                string temp = _pushedBackLine;
                _pushedBackLine = null;
                // LineNumber was already incremented when this line was originally read.
                return temp;
            }
            
            string line = _reader.ReadLine();
            if (line != null)
            {
                LineNumber++;
            }
            return line;
        }

        public void PushBack(string line)
        {
            if (_pushedBackLine != null)
            {
                throw new InvalidOperationException("Cannot push back more than one line.");
            }
            _pushedBackLine = line;
            // Adjust LineNumber if we consider a "pushed back" line as not yet "re-read".
            // However, for simple use, the LineNumber reflects the last physical read.
        }
        
        public bool EndOfStream => _reader.EndOfStream && _pushedBackLine == null;

        public void Close() => _reader.Close();
        public void Dispose() => _reader.Dispose();
    }
}