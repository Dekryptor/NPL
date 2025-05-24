// Namespace: VB6Parser.Parsers
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using VB6Parser.Models;
using VB6Parser.Language.Controls;
using VB6Parser.Utilities; // For PropertyParsingHelpers (indirectly via XProperties)

namespace VB6Parser.Parsers
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
        private static readonly Regex ObjectRegex = new Regex(@"^Object\s*=\s*(.*)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private static readonly Regex AttributeRegex = new Regex(@"^\s*Attribute\s+([^=]+)\s*=\s*(.*)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        
        // Captures: 1=TypeFullName (e.g. VB.Form), 2=ControlName
        private static readonly Regex BeginControlRegex = new Regex(@"^\s*Begin\s+([A-Za-z0-9_\.]+)\s+([A-Za-z_][A-Za-z0-9_]*)", RegexOptions.Compiled | RegexOptions.CultureInvariant); 
        private static readonly Regex EndControlRegex = new Regex(@"^\s*End\s*$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        
        // Captures: 1=GroupName, 2=OptionalGUID
        private static readonly Regex BeginPropertyRegex = new Regex(@"^\s*BeginProperty\s+([A-Za-z_][A-Za-z0-9_]*)(?:\s+\{([0-9A-Fa-f\-]+)\})?", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private static readonly Regex EndPropertyRegex = new Regex(@"^\s*EndProperty\s*$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        
        // Captures: 1=PropertyName, 2=PropertyValue
        private static readonly Regex PropertyLineRegex = new Regex(@"^\s*([A-Za-z_][A-Za-z0-9_#]*)\s*=\s*(.*)", RegexOptions.Compiled | RegexOptions.CultureInvariant);
        
        // Captures: 1=PropertyName, 2=FRXFileName, 3=HexOffsetValue (group 4 is the extension, group 3 is just the hex digits)
        private static readonly Regex ResourcePropertyLineRegex = new Regex(@"^\s*([A-Za-z_][A-Za-z0-9_#]*)\s*=\s*""([^""]+\.(frx|ctx|pag|dsx|vbx|ocx))""\s*:\s*([0-9A-Fa-f]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant);

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
        
		private static byte[] DefaultFrxResolver(string frxFilePath, int frxOffset) // Renamed offset to frxOffset for clarity
		{
			if (string.IsNullOrEmpty(frxFilePath) || !File.Exists(frxFilePath))
			{
				// Console.Error.WriteLine($"FRX file not found or path is null/empty: {frxFilePath}");
				return null;
			}

			try
			{
				byte[] buffer = File.ReadAllBytes(frxFilePath);

				if (frxOffset < 0 || frxOffset >= buffer.Length) // Added check for frxOffset < 0
				{
					// Console.Error.WriteLine($"Offset {frxOffset} is out of bounds for FRX file {frxFilePath} (length {buffer.Length}).");
					return null; 
				}

				// Type 1: 12-byte Header Record (Signature "lt\0\0")
				// Check signature at frxOffset + 4 to frxOffset + 8
				if (frxOffset + 8 <= buffer.Length) // Ensure enough bytes for signature part
				{
					// The Rust code checks `buffer.len() >= 12` globally for this block,
					// and then `binary_blob_signature = buffer[offset + 4..offset + 8]`.
					// We need to ensure `frxOffset + 4` up to `frxOffset + 8` is valid.
					// And for sizes, `frxOffset` to `frxOffset+4` and `frxOffset+8` to `frxOffset+12`.
					// So, effectively, `frxOffset + 12 <= buffer.Length` for the full header.
					if (frxOffset + 12 <= buffer.Length &&
						buffer[frxOffset + 4] == 'l' && buffer[frxOffset + 5] == 't' &&
						buffer[frxOffset + 6] == 0 && buffer[frxOffset + 7] == 0)
					{
						uint bufferSize1 = BitConverter.ToUInt32(buffer, frxOffset);
						uint bufferSize2 = BitConverter.ToUInt32(buffer, frxOffset + 8);

						if (bufferSize1 == 8 && bufferSize2 == 0)
						{
							return Array.Empty<byte>(); // Special empty case
						}

						if (bufferSize2 != bufferSize1 - 8)
						{
							// Console.Error.WriteLine($"FRX Corrupted (Type 1): Size mismatch in {frxFilePath} at offset {frxOffset}. {bufferSize2} != {bufferSize1 - 8}");
							return null; // Corrupted
						}

						int headerSize = 12;
						int recordStart = frxOffset + headerSize;
						int recordLength = (int)bufferSize2;

						if (recordStart + recordLength > buffer.Length)
						{
							// Console.Error.WriteLine($"FRX Corrupted (Type 1): Record end out of bounds in {frxFilePath} at offset {frxOffset}.");
							return null; // Corrupted
						}
						byte[] data = new byte[recordLength];
						Array.Copy(buffer, recordStart, data, 0, recordLength);
						return data;
					}
				}

				// Type 2: 0xFF Prefixed Record (16-bit length)
				if (buffer[frxOffset] == 0xFF)
				{
					if (frxOffset + 3 > buffer.Length) // Need 1 byte for 0xFF and 2 for length
					{
						// Console.Error.WriteLine($"FRX Corrupted (Type 2): Not enough bytes for header in {frxFilePath} at offset {frxOffset}.");
						return null;
					}
						
					int recordSize = BitConverter.ToUInt16(buffer, frxOffset + 1);
					
					// VB6 quirk: off-by-one for recordSize
					// Rust: if header_size_element_offset (frxOffset+1) + record_size > buffer.len()
					if ((frxOffset + 1) + record_size > buffer.Length) // Note: Rust uses record_offset (frxOffset+1+2) for this check in one place, but (frxOffset+1) seems more direct from context
					{
						 if (record_size > 0) record_size--; // Ensure not to make it negative
					}

					int recordStart = frxOffset + 3; // Skip 0xFF and 2-byte length
					if (recordStart + recordSize > buffer.Length)
					{
						// Console.Error.WriteLine($"FRX Corrupted (Type 2): Record end out of bounds after adjustment in {frxFilePath} at offset {frxOffset}. RecordSize: {recordSize}");
						return null; // Still out of bounds
					}
					 if (recordSize < 0) return Array.Empty<byte>(); // Should not happen with uint16 but defensive

					byte[] data = new byte[recordSize];
					if (recordSize > 0) // Array.Copy will throw if recordSize is 0 and recordStart is at end of buffer
					{
						 Array.Copy(buffer, recordStart, data, 0, recordSize);
					}
					return data;
				}

				// Type 3: List Items Record (Signatures [0x03, 0x00] or [0x07, 0x00])
				// Check signature at frxOffset + 2 to frxOffset + 4
				if (frxOffset + 4 <= buffer.Length) // Need at least 4 bytes for this header type
				{
					bool isListType1 = buffer[frxOffset + 2] == 0x03 && buffer[frxOffset + 3] == 0x00;
					bool isListType2 = buffer[frxOffset + 2] == 0x07 && buffer[frxOffset + 3] == 0x00;

					if (isListType1 || isListType2)
					{
						int listItemCount = BitConverter.ToUInt16(buffer, frxOffset);
						int currentOffset = frxOffset + 4; // Start of first list item's length
						int totalBlockEndOffset = currentOffset; // Will track the end of the entire list block relative to frxOffset

						for (int i = 0; i < listItemCount; i++)
						{
							if (currentOffset + 2 > buffer.Length) // Not enough for item length
							{
								// Console.Error.WriteLine($"FRX Corrupted (Type 3): Truncated list item length in {frxFilePath} at item {i}, offset {currentOffset}.");
								return null;
							}
							int listItemSize = BitConverter.ToUInt16(buffer, currentOffset);
							currentOffset += 2; // Move past item length

							if (currentOffset + listItemSize > buffer.Length) // Not enough for item data
							{
								// Console.Error.WriteLine($"FRX Corrupted (Type 3): Truncated list item data in {frxFilePath} at item {i}, offset {currentOffset}.");
								return null;
							}
							currentOffset += listItemSize; // Move past item data
						}
						totalBlockEndOffset = currentOffset;
						
						int recordLength = totalBlockEndOffset - frxOffset;
						byte[] data = new byte[recordLength];
						Array.Copy(buffer, frxOffset, data, 0, recordLength);
						return data; // Returns the whole block including header and all items
					}
				}
				
				// Type 4: 4-byte Header Record (Contains Null Byte in first 4 bytes)
				// Rust condition: buffer.len() >= 12 && buffer[(offset)..(offset + 4)].contains(&0u8)
				// The buffer.len() >= 12 seems like a general safety for some complex types, might not be strictly for this one.
				// Let's check if frxOffset + 4 is valid first.
				if (frxOffset + 4 <= buffer.Length)
				{
					bool hasNullByteInHeader = false;
					for(int k=0; k<4; ++k) {
						if (buffer[frxOffset + k] == 0) {
							hasNullByteInHeader = true;
							break;
						}
					}

					if (hasNullByteInHeader)
					{
						int recordSize = (int)BitConverter.ToUInt32(buffer, frxOffset);
						int headerSize = 4;
						int recordStart = frxOffset + headerSize;
						
						if (recordStart + recordSize > buffer.Length) {
							// Console.Error.WriteLine($"FRX Corrupted (Type 4): Record end out of bounds in {frxFilePath} at offset {frxOffset}.");
							return null; 
						}
						 if (recordSize < 0) return Array.Empty<byte>();

						byte[] data = new byte[recordSize];
						if (recordSize > 0) {
							Array.Copy(buffer, recordStart, data, 0, recordSize);
						}
						return data;
					}
				}

				// Type 5: 1-byte Header Record (Default/Fallback)
				// Ensure frxOffset is valid before trying to access buffer[frxOffset]
				if (frxOffset < buffer.Length) 
				{
					int recordSize = buffer[frxOffset];
					int headerSize = 1;
					int recordStart = frxOffset + headerSize;

					if (recordStart > buffer.Length) // Check if recordStart itself is out of bounds
					{
						// Console.Error.WriteLine($"FRX Corrupted (Type 5): Record start out of bounds in {frxFilePath} at offset {frxOffset}.");
						return Array.Empty<byte>(); // Or null, Rust returns error. Empty might be safer for caller.
					}

					// Off-by-one hack (Rust: _ if record_size >= buffer.len() => record_start + record_size - 1)
					// This means if the declared size would exceed the buffer *when added to record_start*
					// The Rust code's condition `record_size >= buffer.len()` is a bit confusing in its direct form.
					// It should be `record_start + record_size > buffer.len()` implies an off-by-one.
					if (record_start + recordSize > buffer.Length)
					{
						 if (record_size > 0) record_size--;
					}
					
					if (recordStart + recordSize > buffer.Length)
					{
						// Console.Error.WriteLine($"FRX Corrupted (Type 5): Record end out of bounds after adjustment in {frxFilePath} at offset {frxOffset}. RecordSize: {recordSize}");
						return Array.Empty<byte>(); // Or null
					}
					if (recordSize < 0) return Array.Empty<byte>();

					byte[] data = new byte[recordSize];
					if (recordSize > 0)
					{
						Array.Copy(buffer, recordStart, data, 0, recordSize);
					}
					return data;
				}

				// If no type matched or offset was at the very end initially
				// Console.Error.WriteLine($"FRX: Could not determine record type or invalid offset {frxOffset} for file {frxFilePath}.");
				return null; // Should not be reached if offset is valid and one type matches
			}
			catch (Exception ex)
			{
				Console.Error.WriteLine($"Error processing FRX file {frxFilePath} at offset {frxOffset}: {ex.Message}");
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
                        string rawValuePart = match.Groups[2].Value; // This is everything after " = "
                        string propValue;
                        
                        int commentCharPos = -1;
                        bool inQuotes = false;
                        for (int i = 0; i < rawValuePart.Length; i++)
                        {
                            if (rawValuePart[i] == '"')
                            {
                                if (i + 1 < rawValuePart.Length && rawValuePart[i+1] == '"')
                                {
                                    i++; 
                                }
                                else
                                {
                                    inQuotes = !inQuotes;
                                }
                            }
                            else if (rawValuePart[i] == '\'' && !inQuotes)
                            {
                                commentCharPos = i;
                                break;
                            }
                        }

                        if (commentCharPos != -1)
                        {
                            propValue = rawValuePart.Substring(0, commentCharPos).TrimEnd();
                        }
                        else
                        {
                            propValue = rawValuePart.TrimEnd();
                        }
                        propValue = propValue.TrimStart();

                        if (propValue.StartsWith("\"") && propValue.EndsWith("\""))
                        {
                            if (propValue.Length >= 2)
                            {
                                 propValue = propValue.Substring(1, propValue.Length - 2).Replace("\"\"", "\"");
                            }
                            else 
                            {
                                propValue = string.Empty; 
                            }
                        }
                        currentState.TextualProperties[propName] = propValue;
                        continue;
                    }
                    // Line is not a recognized control structure or property inside a Begin block.
                    // Could be a comment (VB uses ' at start of line or after code)
                    // Simple comment stripping:
                    if (trimmedLine.StartsWith("'")) continue; // Whole line is a comment
                    // More complex handling might be needed if comments can be anywhere or if regexes need to ignore them.
                    // For now, unhandled lines within a block are skipped or could be logged.
                }
                else
                {
                    // Line is outside any Begin/End block.
                    // This should ideally not happen if ParseGlobalHeader correctly found the first "Begin".
                    // Or it's after the final "End" of the form, which ParseFormAttributes will handle.
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
                
                match = ResourcePropertyLineRegex.Match(trimmedLine);
                if (match.Success)
                {
                    string propName = match.Groups[1].Value;
                    int offset = int.Parse(match.Groups[4].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    byte[] data = _frxResolver?.Invoke(_frxFilePath, offset);
                    if (data != null)
                    {
                        propertyGroup.Properties[propName] = data; 
                    }
                    continue;
                }

                match = PropertyLineRegex.Match(trimmedLine);
                if (match.Success)
                {
                    string propName = match.Groups[1].Value;
                    string rawValuePart = match.Groups[2].Value;
                    string propValue;

                    int commentCharPos = -1;
                    bool inQuotes = false;
                    for (int i = 0; i < rawValuePart.Length; i++)
                    {
                        if (rawValuePart[i] == '"')
                        {
                            if (i + 1 < rawValuePart.Length && rawValuePart[i+1] == '"') { i++; }
                            else { inQuotes = !inQuotes; }
                        }
                        else if (rawValuePart[i] == '\'' && !inQuotes) { commentCharPos = i; break; }
                    }

                    if (commentCharPos != -1) { propValue = rawValuePart.Substring(0, commentCharPos).TrimEnd(); }
                    else { propValue = rawValuePart.TrimEnd(); }
                    propValue = propValue.TrimStart();

                    if (propValue.StartsWith("\"") && propValue.EndsWith("\""))
                    {
                        if (propValue.Length >=2) { propValue = propValue.Substring(1, propValue.Length - 2).Replace("\"\"", "\""); }
                        else { propValue = string.Empty;}
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

            string[] typeParts = state.TypeFullName.Split('.');
            control.TypeNamespace = typeParts.Length > 1 ? typeParts[0] : "VB"; // Default to VB if no namespace
            control.TypeKind = typeParts.Length > 1 ? typeParts[1] : typeParts[0];

            Match nameIndexMatch = Regex.Match(control.Name, @"([A-Za-z_][A-Za-z0-9_]*)\_?\((d+)\)");
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
                     var optProps = new OptionButtonProperties(); // Assuming this class exists
                     optProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                     control.SpecificProperties = optProps;
                     break;
                case "ListBox":
                    var lbProps = new ListBoxProperties(); // Assuming this class exists
                    lbProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = lbProps;
                    break;
                case "ComboBox":
                    var cmbProps = new ComboBoxProperties(); // Assuming this class exists
                    cmbProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = cmbProps;
                    break;
                case "Frame":
                    var frameProps = new FrameProperties(); // Assuming this class exists
                    frameProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = frameProps;
                    break;
                case "Timer":
                    var timerProps = new TimerProperties(); // Assuming
                    timerProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = timerProps;
                    break;
                case "PictureBox":
                    var picProps = new PictureBoxProperties(); // Assuming
                    picProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = picProps;
                    break;
                case "Image":
                    var imgProps = new ImageProperties(); // Assuming
                    imgProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = imgProps;
                    break;
                case "Line":
                    var lineProps = new LineProperties(); // Assuming
                    lineProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = lineProps;
                    break;
                case "Shape":
                    var shapeProps = new ShapeProperties(); // Assuming
                    shapeProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = shapeProps;
                    break;
                case "HScrollBar":
                    var hscrollProps = new HScrollBarProperties(); // Assuming
                    hscrollProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = hscrollProps;
                    break;
                case "VScrollBar":
                    var vscrollProps = new VScrollBarProperties(); // Assuming
                    vscrollProps.PopulateFromDictionaries(state.TextualProperties, state.BinaryProperties, state.PropertyGroups);
                    control.SpecificProperties = vscrollProps;
                    break;
                default:
                    // For custom controls (e.g., MSComctlLib.ListViewCtrl) or unhandled standard controls
                    // The TypeFullName is passed to CustomControlProperties.
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


    // Helper class to read lines and allow pushing a line back onto the stream.
    // Also tracks line numbers for error reporting.
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