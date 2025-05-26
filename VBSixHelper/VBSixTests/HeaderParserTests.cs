using System.Text;
using VBSix.Errors;
using VBSix.Parsers;

namespace VBSixTests
{
    public static class HeaderParserTests
    {
        // Helper to create a stream from string content using Windows-1252 encoding
        // This can be a local helper or moved to a shared test utilities class.
        private static VB6Stream CreateStreamForTest(string content, string fileName = "test.txt")
        {
            // Encoding registration should happen once at app startup (e.g., TestRunner.Main or VB6Stream static ctor)
            // Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); 
            byte[] bytes = Encoding.GetEncoding(1252).GetBytes(content);
            return new VB6Stream(fileName, bytes);
        }

        // --- Test Suite Entry Point ---
        public static void RunAllHeaderParserTests(Action<string, Action> runTestHelper) // Pass RunTest delegate
        {
            Console.WriteLine("\n===== Running HeaderParser Tests =====");
            TestParseVersionLine(runTestHelper);
            TestParseAttributes(runTestHelper);
            TestParseKeyValueLine(runTestHelper);
            TestParseKeyResourceOffsetLine(runTestHelper);
            TestParseObjectLine(runTestHelper);
            Console.WriteLine("===== Finished HeaderParser Tests =====");
        }

        // --- Individual Test Groups ---

        private static void TestParseVersionLine(Action<string, Action> RunTest)
        {
            RunTest("HeaderParser.ParseVersionLine_Class_Valid", () =>
            {
                VB6Stream stream = CreateStreamForTest("VERSION 1.0 CLASS\r\n", "test.cls");
                (VB6Stream nextStream, VB6FileFormatVersion version, HeaderKind kind) = HeaderParser.ParseVersionLine(stream, HeaderKind.Class);
                if (version.Major != 1) throw new Exception($"Expected Major=1, got {version.Major}");
                if (version.Minor != 0) throw new Exception($"Expected Minor=0, got {version.Minor}");
                if (kind != HeaderKind.Class) throw new Exception($"Expected Kind=Class, got {kind}");
                if (!nextStream.IsEmpty) throw new Exception("Stream was not fully consumed after VERSION line.");
            });

            RunTest("HeaderParser.ParseVersionLine_Form_Valid_NoKeyword", () =>
            {
                VB6Stream stream = CreateStreamForTest("VERSION 5.00  'Comment\r\n", "test.frm");
                (VB6Stream nextStream, VB6FileFormatVersion version, HeaderKind kind) = HeaderParser.ParseVersionLine(stream, HeaderKind.Form);
                if (version.Major != 5) throw new Exception($"Expected Major=5, got {version.Major}");
                if (version.Minor != 0) throw new Exception($"Expected Minor=00, got {version.Minor}");
                if (kind != HeaderKind.Form) throw new Exception($"Expected Kind=Form (context), got {kind}");
                if (!nextStream.IsEmpty) throw new Exception("Stream was not fully consumed after VERSION line.");
            });

            RunTest("HeaderParser.ParseVersionLine_Module_Valid_WithKeyword", () =>
            {
                VB6Stream stream = CreateStreamForTest("VERSION 2.0 MODULE\r\n", "test.bas");
                (VB6Stream nextStream, VB6FileFormatVersion version, HeaderKind kind) = HeaderParser.ParseVersionLine(stream, HeaderKind.Module);
                if (version.Major != 2) throw new Exception($"Expected Major=2, got {version.Major}");
                if (version.Minor != 0) throw new Exception($"Expected Minor=0, got {version.Minor}");
                if (kind != HeaderKind.Module) throw new Exception($"Expected Kind=Module, got {kind}");
                if (!nextStream.IsEmpty) throw new Exception("Stream not fully consumed.");
            });

            RunTest("HeaderParser.ParseVersionLine_MissingVersionKeyword_Throws", () =>
            {
                VB6Stream stream = CreateStreamForTest("1.0 CLASS\r\n", "test.cls");
                bool threw = false;
                try { HeaderParser.ParseVersionLine(stream, HeaderKind.Class); }
                catch (VB6ParseException ex) { if (ex.Kind == VB6ErrorKind.KeywordNotFound && ex.InnerException.Message.Contains("VERSION")) threw = true; else throw; }
                catch { throw; }
                if (!threw) throw new Exception("Expected VB6ParseException for missing VERSION keyword.");
            });
            RunTest("HeaderParser.ParseVersionLine_InvalidMajor_Throws", () =>
            {
                VB6Stream stream = CreateStreamForTest("VERSION X.0 CLASS\r\n");
                bool threw = false;
                try { HeaderParser.ParseVersionLine(stream, HeaderKind.Class); }
                catch (VB6ParseException ex) { if (ex.Kind == VB6ErrorKind.MajorVersionUnparseable) threw = true; else throw; }
                catch { throw; }
                if (!threw) throw new Exception("Expected MajorVersionUnparseable.");
            });
        }

        private static void TestParseAttributes(Action<string, Action> RunTest)
        {
            RunTest("HeaderParser.ParseAttributes_SingleVBName", () =>
            {
                VB6Stream stream = CreateStreamForTest("Attribute VB_Name = \"MyComponent\"\r\n");
                (VB6Stream nextStream, VB6FileAttributes attrs) = HeaderParser.ParseAttributes(stream);
                if (attrs.Name != "MyComponent") throw new Exception($"Expected Name='MyComponent', got '{attrs.Name}'");
                if (!nextStream.IsEmpty) throw new Exception("Stream not fully consumed.");
            });

            RunTest("HeaderParser.ParseAttributes_MultipleAttributes", () =>
            {
                string content =
                    "Attribute VB_Name = \"MyClass\"\r\n" +
                    "Attribute VB_GlobalNameSpace = False\r\n" +
                    "Attribute VB_Creatable = True\r\n" +
                    "Attribute VB_PredeclaredId = False\r\n" +
                    "Attribute VB_Exposed = True 'Comment\r\n";
                VB6Stream stream = CreateStreamForTest(content);
                (VB6Stream nextStream, VB6FileAttributes attrs) = HeaderParser.ParseAttributes(stream);

                if (attrs.Name != "MyClass") throw new Exception("Name mismatch");
                if (attrs.GlobalNameSpace != NameSpaceKind.Local) throw new Exception("GlobalNameSpace mismatch");
                if (attrs.Creatable != CreatableKind.True) throw new Exception("Creatable mismatch");
                if (attrs.PreDeclaredID != PreDeclaredIDKind.False) throw new Exception("PredeclaredId mismatch");
                if (attrs.Exposed != ExposedKind.True) throw new Exception("Exposed mismatch");
                if (!nextStream.IsEmpty) throw new Exception("Stream not fully consumed.");
            });

            RunTest("HeaderParser.ParseAttributes_VBExtKey", () =>
            {
                string content =
                    "Attribute VB_Name = \"Test\"\r\n" +
                    "Attribute VB_Ext_KEY = \"MyCustomKey\", \"MyCustomValue\"\r\n";
                VB6Stream stream = CreateStreamForTest(content);
                (VB6Stream nextStream, VB6FileAttributes attrs) = HeaderParser.ParseAttributes(stream);

                if (attrs.Name != "Test") throw new Exception("Name mismatch");
                if (!attrs.ExtKey.ContainsKey("MyCustomKey")) throw new Exception("ExtKey MyCustomKey not found");
                if (attrs.ExtKey["MyCustomKey"] != "MyCustomValue") throw new Exception("ExtKey MyCustomKey value mismatch");
                if (!nextStream.IsEmpty) throw new Exception("Stream not fully consumed.");
            });

            RunTest("HeaderParser.ParseAttributes_VBExtKey_Multiple", () =>
            {
                string content =
                    "Attribute VB_Name = \"Test\"\r\n" +
                    "Attribute VB_Ext_KEY = \"Key1\", \"Value1\"\r\n" +
                    "Attribute VB_Ext_KEY = \"Key2\", \"Value2\"\r\n";
                VB6Stream stream = CreateStreamForTest(content);
                (VB6Stream nextStream, VB6FileAttributes attrs) = HeaderParser.ParseAttributes(stream);

                if (attrs.Name != "Test") throw new Exception("Name mismatch");
                if (!attrs.ExtKey.TryGetValue("Key1", out var val1) || val1 != "Value1") throw new Exception("ExtKey Key1 mismatch");
                if (!attrs.ExtKey.TryGetValue("Key2", out var val2) || val2 != "Value2") throw new Exception("ExtKey Key2 mismatch");
                if (!nextStream.IsEmpty) throw new Exception("Stream not fully consumed.");
            });

            RunTest("HeaderParser.ParseAttributes_MissingVBName_WhenOtherAttributesExist_Throws", () =>
            {
                string content = "Attribute VB_GlobalNameSpace = False\r\n";
                VB6Stream stream = CreateStreamForTest(content);
                bool threw = false;
                try { HeaderParser.ParseAttributes(stream); }
                catch (VB6ParseException ex) { if (ex.Kind == VB6ErrorKind.MissingNameAttribute) threw = true; else throw; }
                catch { throw; }
                if (!threw) throw new Exception("Expected MissingNameAttribute.");
            });

            RunTest("HeaderParser.ParseAttributes_NoAttributes_ReturnsEmptyAndOriginalStream", () =>
            {
                string content = "SomeOtherLine=123\r\n";
                VB6Stream stream = CreateStreamForTest(content);
                var initialLength = stream.Length;
                (VB6Stream nextStream, VB6FileAttributes attrs) = HeaderParser.ParseAttributes(stream);
                if (attrs.Name != string.Empty) throw new Exception("Default Name mismatch");
                if (attrs.ExtKey.Any()) throw new Exception("ExtKey should be empty");
                if (nextStream.Length != initialLength) throw new Exception("Stream should not have been consumed.");
            });

            RunTest("HeaderParser.ParseAttributes_MalformedAttribute_NoEquals_Throws", () =>
            {
                string content = "Attribute VB_Name \"MyComponent\"\r\n"; // Missing '='
                VB6Stream stream = CreateStreamForTest(content);
                bool threwCorrectly = false;
                try
                {
                    HeaderParser.ParseAttributes(stream);
                }
                catch (VB6ParseException ex)
                {
                    // Check the outer exception kind
                    if (ex.Kind == VB6ErrorKind.AttributeParseError &&
                        ex.InnerException is VB6ParseException innerEx &&
                        innerEx.Kind == VB6ErrorKind.NoEqualSplit) // This condition needs to match what ParseAttributes actually throws
                    {
                        threwCorrectly = true;
                    }
                    else
                    {
                        // This part is executed if the above condition isn't met
                        throw new Exception($"Expected AttributeParseError with Inner NoEqualSplit, but got Kind={ex.Kind} and Inner={ex.InnerException?.GetType().Name} (InnerKind={(ex.InnerException as VB6ParseException)?.Kind})");
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Expected VB6ParseException, got {ex.GetType().Name}");
                }
                if (!threwCorrectly) throw new Exception("Expected VB6ParseException (AttributeParseError -> NoEqualSplit) was not thrown correctly.");
            });
        }

        private static void TestParseKeyValueLine(Action<string, Action> RunTest)
        {
            RunTest("HeaderParser.ParseKeyValueLine_Simple", () =>
            {
                VB6Stream stream = CreateStreamForTest("MyKey = MyValue\r\n");
                (VB6Stream nextStream, string key, string value) = HeaderParser.ParseKeyValueLine(stream);
                if (key != "MyKey") throw new Exception("Key mismatch");
                if (value != "MyValue") throw new Exception("Value mismatch");
                if (!nextStream.IsEmpty) throw new Exception("Stream not consumed.");
            });

            RunTest("HeaderParser.ParseKeyValueLine_QuotedValue", () =>
            {
                VB6Stream stream = CreateStreamForTest("Title = \"My Project Title\"\r\n");
                (VB6Stream _, string key, string value) = HeaderParser.ParseKeyValueLine(stream);
                if (key != "Title") throw new Exception("Key mismatch");
                if (value != "My Project Title") throw new Exception($"Value mismatch, got '{value}'");
            });

            RunTest("HeaderParser.ParseKeyValueLine_QuotedValueWithInternalQuotes", () =>
            {
                VB6Stream stream = CreateStreamForTest("Caption = \"Form with \"\"Quotes\"\"\"\r\n");
                (VB6Stream _, string key, string value) = HeaderParser.ParseKeyValueLine(stream);
                if (key != "Caption") throw new Exception("Key mismatch");
                if (value != "Form with \"Quotes\"") throw new Exception($"Value mismatch, got '{value}'");
            });

            RunTest("HeaderParser.ParseKeyValueLine_WithComment", () =>
            {
                VB6Stream stream = CreateStreamForTest("Key = Value ' This is a comment\r\n");
                (VB6Stream _, string key, string value) = HeaderParser.ParseKeyValueLine(stream);
                if (key != "Key") throw new Exception("Key mismatch");
                if (value != "Value") throw new Exception("Value mismatch");
            });

            RunTest("HeaderParser.ParseKeyValueLine_NoSpaceAroundEquals", () =>
            {
                VB6Stream stream = CreateStreamForTest("Key=Value\r\n");
                (VB6Stream _, string key, string value) = HeaderParser.ParseKeyValueLine(stream);
                if (key != "Key") throw new Exception("Key mismatch");
                if (value != "Value") throw new Exception("Value mismatch");
            });

            RunTest("HeaderParser.ParseKeyValueLine_EmptyValue_ParsesAsEmpty", () =>
            {
                VB6Stream stream = CreateStreamForTest("Key = \r\n");
                (VB6Stream _, string key, string value) = HeaderParser.ParseKeyValueLine(stream);
                if (key != "Key") throw new Exception("Key mismatch");
                if (value != "") throw new Exception($"Expected empty value, got '{value}'");
            });

            RunTest("HeaderParser.ParseKeyValueLine_MissingEquals_Throws", () =>
            {
                VB6Stream stream = CreateStreamForTest("Key Value\r\n");
                bool threw = false;
                try { HeaderParser.ParseKeyValueLine(stream); }
                catch (VB6ParseException ex) { if (ex.Kind == VB6ErrorKind.NoKeyValueDividerFound) threw = true; else throw; }
                catch { throw; }
                if (!threw) throw new Exception("Expected NoKeyValueDividerFound.");
            });
        }

        private static void TestParseKeyResourceOffsetLine(Action<string, Action> RunTest)
        {
            RunTest("HeaderParser.ParseKeyResourceOffsetLine_Valid", () =>
            {
                VB6Stream stream = CreateStreamForTest("Picture = \"MyForm.frx\":00A0\r\n");
                (VB6Stream _, string key, string resFile, uint offset) = HeaderParser.ParseKeyResourceOffsetLine(stream);
                if (key != "Picture") throw new Exception("Key mismatch");
                if (resFile != "MyForm.frx") throw new Exception("Resource file mismatch");
                if (offset != 0x00A0u) throw new Exception("Offset mismatch");
            });

            RunTest("HeaderParser.ParseKeyResourceOffsetLine_Valid_WithSpacesAndComment", () =>
            {
                VB6Stream stream = CreateStreamForTest("TextRTF =   \"MyUserControl.ctx\"  :  01BC  'Comment\r\n");
                (VB6Stream _, string key, string resFile, uint offset) = HeaderParser.ParseKeyResourceOffsetLine(stream);
                if (key != "TextRTF") throw new Exception("Key mismatch");
                if (resFile != "MyUserControl.ctx") throw new Exception($"Resource file mismatch, got '{resFile}'");
                if (offset != 0x01BCu) throw new Exception("Offset mismatch");
            });

            RunTest("HeaderParser.ParseKeyResourceOffsetLine_InvalidOffset_Throws", () =>
            {
                VB6Stream stream = CreateStreamForTest("Icon = \"MyForm.frx\":XYZ\r\n");
                bool threw = false;
                try { HeaderParser.ParseKeyResourceOffsetLine(stream); }
                catch (VB6ParseException ex) { if (ex.Kind == VB6ErrorKind.Property && ex.PropertyErrorDetail == PropertyErrorType.OffsetUnparsable) threw = true; else throw; }
                catch { throw; }
                if (!threw) throw new Exception("Expected OffsetUnparsable.");
            });

            RunTest("HeaderParser.ParseKeyResourceOffsetLine_MissingColon_Throws", () =>
            {
                VB6Stream stream = CreateStreamForTest("Picture = \"MyForm.frx\"00A0\r\n");
                bool threw = false;
                try { HeaderParser.ParseKeyResourceOffsetLine(stream); }
                catch (VB6ParseException ex) { if (ex.Kind == VB6ErrorKind.NoColonForOffsetSplit) threw = true; else throw; }
                catch { throw; }
                if (!threw) throw new Exception("Expected NoColonForOffsetSplit.");
            });

            RunTest("HeaderParser.ParseKeyResourceOffsetLine_UnquotedFileName_Throws", () => {
                VB6Stream stream = CreateStreamForTest("Picture = MyForm.frx:00A0\r\n");
                bool threwCorrectly = false;
                try
                {
                    HeaderParser.ParseKeyResourceOffsetLine(stream);
                }
                catch (VB6ParseException ex)
                {
                    if (ex.Kind == VB6ErrorKind.Property && ex.PropertyErrorDetail == PropertyErrorType.ResourceFileNameUnparsable)
                    {
                        // Optionally check ex.InnerException for more details if Vb6.ParseVB6String provides it
                        threwCorrectly = true;
                    }
                    else
                    {
                        throw new Exception($"Expected Property/ResourceFileNameUnparsable, got Kind={ex.Kind}, Detail={ex.PropertyErrorDetail}");
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Expected VB6ParseException, got {ex.GetType().Name}");
                }
                if (!threwCorrectly) throw new Exception("Expected VB6ParseException for unquoted FRX filename.");
            });
        }

        private static void TestParseObjectLine(Action<string, Action> RunTest)
        {
            RunTest("HeaderParser.ParseObjectLine_CompiledGuid_VBPStyle", () =>
            {
                string lineContent = "{00020430-0000-0000-C000-000000000046}#2.0#0; STDOLE2.TLB";
                VB6Stream stream = CreateStreamForTest(lineContent + "\r\n");
                (VB6Stream _, VB6ObjectReference objRef) = HeaderParser.ParseObjectLine(stream, "Object=" + lineContent);
                if (objRef.IsSubProjectReference) throw new Exception("Should be compiled ref");
                if (!objRef.IsCompiledReference) throw new Exception("Should be compiled ref flag");
                if (objRef.ObjectGuidString?.ToUpperInvariant() != "{00020430-0000-0000-C000-000000000046}") throw new Exception("GUID mismatch");
                if (objRef.Version != "2.0") throw new Exception("Version mismatch");
                if (objRef.LocaleID != "0") throw new Exception("Unknown1 mismatch");
                if (objRef.FileName != "STDOLE2.TLB") throw new Exception("FileName mismatch");
            });

            RunTest("HeaderParser.ParseObjectLine_SubProject_VBPStyle", () =>
            {
                string lineContent = "*\\AMyProject.vbp";
                VB6Stream stream = CreateStreamForTest(lineContent + "\r\n");
                (VB6Stream _, VB6ObjectReference objRef) = HeaderParser.ParseObjectLine(stream, "Object=" + lineContent);
                if (!objRef.IsSubProjectReference) throw new Exception("Should be sub-project ref");
                if (objRef.Path != "*\\AMyProject.vbp") throw new Exception("Path mismatch");
            });

            RunTest("HeaderParser.ParseObjectLine_ActiveX_FRMStyle_QuotedGuidAndFile", () => {
                // This is a common format for ActiveX controls listed in the OBJECT = {} part of a Form file header
                // The GUID block can be quoted, and the filename can be quoted.
                string lineContent = "\"{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0\"; \"mscomctl.ocx\"";
                VB6Stream stream = CreateStreamForTest(lineContent + "\r\n");
                // Assuming ParseObjectLine is called with the content *after* "Object = "
                (VB6Stream _, VB6ObjectReference objRef) = HeaderParser.ParseObjectLine(stream, "Object = " + lineContent);

                if (!objRef.IsCompiledReference) throw new Exception("Should be compiled ref");
                if (objRef.IsSubProjectReference) throw new Exception("Should not be sub-project ref");

                // Assertions should match the input "lineContent"
                if (objRef.ObjectGuidString?.ToUpperInvariant() != "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}")
                    throw new Exception($"GUID mismatch. Expected {{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}}, got {objRef.ObjectGuidString}");
                if (objRef.Version != "2.1")
                    throw new Exception($"Version mismatch. Expected 2.1, got {objRef.Version}");
                if (objRef.LocaleID != "0") // Changed from Unknown1 to LocaleID for clarity
                    throw new Exception($"Unknown1 mismatch. Expected 0, got {objRef.LocaleID}");
                if (objRef.FileName != "mscomctl.ocx") // Vb6.ParseVB6String unquotes this
                    throw new Exception($"FileName mismatch. Expected mscomctl.ocx, got '{objRef.FileName}'");
            });

            RunTest("HeaderParser.ParseObjectLine_ActiveX_FRMStyle_UnquotedGuidQuotedFile", () =>
            {
                string lineContent = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0; MSWINSCK.OCX";
                VB6Stream stream = CreateStreamForTest(lineContent + "\r\n");
                (VB6Stream _, VB6ObjectReference objRef) = HeaderParser.ParseObjectLine(stream, "Object = " + lineContent);
                if (!objRef.IsCompiledReference) throw new Exception("Should be compiled ref");
                if (objRef.ObjectGuidString?.ToUpperInvariant() != "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}") throw new Exception("GUID mismatch");
                if (objRef.Version != "1.0") throw new Exception("Version mismatch");
                if (objRef.LocaleID != "0") throw new Exception("Unknown1 mismatch");
                if (objRef.FileName != "MSWINSCK.OCX") throw new Exception($"FileName mismatch, got '{objRef.FileName}' (expected 'MSWINSCK.OCX')");
            });
        }
    }
}