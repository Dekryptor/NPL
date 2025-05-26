using System.Text;
using VBSix.Errors;
using VBSix.Language;
using VBSix.Parsers;

namespace VBSixTests
{
    internal class Program
    {
        private static int _testsRun = 0;
        private static int _testsFailed = 0;

        public static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Console.WriteLine("Starting VB6Parse_CSharp Tests...\n");

            //TestVB6ColorParsing();
            //TestHeaderParser_VersionLine();
            HeaderParserTests.RunAllHeaderParserTests(RunTest);

            Console.WriteLine("\n-----------------------------------");
            Console.WriteLine($"Tests Run: {_testsRun}");
            Console.WriteLine($"Tests Failed: {_testsFailed}");
            Console.ForegroundColor = _testsFailed == 0 ? ConsoleColor.Green : ConsoleColor.Red;
            Console.WriteLine(_testsFailed == 0 ? "All tests passed!" : "SOME TESTS FAILED!");
            Console.ResetColor();
            Console.WriteLine("-----------------------------------");
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey(); // Keep console open
        }

        // Helper for assertions
        private static void RunTest(string testName, Action testAction)
        {
            _testsRun++;
            Console.WriteLine($"--- Running Test: {testName} ---");
            try
            {
                testAction();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($"[PASS] {testName}");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                _testsFailed++;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"[FAIL] {testName}");
                Console.WriteLine($"     Error: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"     Inner Error: {ex.InnerException.Message}");
                }
                // Console.WriteLine($"     Stack Trace: {ex.StackTrace}"); // Optional: for more detail
                Console.ResetColor();
            }
            finally
            {
                Console.WriteLine("-----------------------------------\n");
            }
        }

        // --- Example Test Methods ---

        public static void TestVB6ColorParsing()
        {
            RunTest("VB6Color.FromHex_ValidRGBColor", () => {
                string hexString = "&H00FF8040&"; // R=40, G=80, B=FF
                VB6Color color = VB6Color.FromHex(hexString);

                if (color.Type != VB6ColorType.RGB) throw new Exception($"Expected RGB type, got {color.Type}");
                if (color.R != 0x40) throw new Exception($"Expected R=0x40, got 0x{color.R:X2}");
                if (color.G != 0x80) throw new Exception($"Expected G=0x80, got 0x{color.G:X2}");
                if (color.B != 0xFF) throw new Exception($"Expected B=0xFF, got 0x{color.B:X2}");
                if (!hexString.Equals(color.ToString(), StringComparison.OrdinalIgnoreCase)) throw new Exception($"Expected ToString()='{hexString}', got '{color}'");
            });

            RunTest("VB6Color.FromHex_ValidSystemColor", () => {
                string hexString = "&H80000005&"; // System color index 5
                VB6Color color = VB6Color.FromHex(hexString);

                if (color.Type != VB6ColorType.System) throw new Exception($"Expected System type, got {color.Type}");
                if (color.SystemIndex != 5) throw new Exception($"Expected SystemIndex=5, got {color.SystemIndex}");
                if (!hexString.Equals(color.ToString(), StringComparison.OrdinalIgnoreCase)) throw new Exception($"Expected ToString()='{hexString}', got '{color}'");
            });

            RunTest("VB6Color.FromHex_InvalidFormat_ThrowsVB6ParseException", () => {
                string hexString = "&H00FF80&"; // Too short
                bool exceptionThrown = false;
                try
                {
                    VB6Color.FromHex(hexString);
                }
                catch (VB6ParseException ex)
                {
                    if (ex.Kind == VB6ErrorKind.HexColorParseError) exceptionThrown = true;
                    else throw new Exception($"Expected VB6ErrorKind.HexColorParseError, but got {ex.Kind}");
                }
                catch (Exception ex)
                {
                    throw new Exception($"Expected VB6ParseException, but got {ex.GetType().Name}");
                }
                if (!exceptionThrown) throw new Exception("Expected VB6ParseException was not thrown for invalid hex color.");
            });
        }

        public static void TestHeaderParser_VersionLine()
        {
            RunTest("HeaderParser.ParseVersionLine_Class_Valid", () => {
                string content = "VERSION 1.0 CLASS\r\n";
                byte[] bytes = Encoding.Default.GetBytes(content);
                VB6Stream stream = new("./Tests/CSC_FormProperties.cls", bytes);

                (VB6Stream nextStream, VB6FileFormatVersion version, HeaderKind kind) = HeaderParser.ParseVersionLine(stream, HeaderKind.Class);

                if (version.Major != 1) throw new Exception($"Expected Major=1, got {version.Major}");
                if (version.Minor != 0) throw new Exception($"Expected Minor=0, got {version.Minor}");
                if (kind != HeaderKind.Class) throw new Exception($"Expected Kind=Class, got {kind}");
                if (!nextStream.IsEmpty) throw new Exception("Stream was not fully consumed after VERSION line.");
            });

            RunTest("HeaderParser.ParseVersionLine_Form_Valid_NoKeyword", () => {
                string content = "VERSION 5.00  'Comment\r\n";
                byte[] bytes = Encoding.Default.GetBytes(content);
                VB6Stream stream = new("./Tests/CSF_Status.frm", bytes);

                (VB6Stream nextStream, VB6FileFormatVersion version, HeaderKind kind) = HeaderParser.ParseVersionLine(stream, HeaderKind.Form); // Pass expected kind

                if (version.Major != 5) throw new Exception($"Expected Major=5, got {version.Major}");
                if (version.Minor != 0) throw new Exception($"Expected Minor=00, got {version.Minor}"); // "00" parses to 0
                if (kind != HeaderKind.Form) throw new Exception($"Expected Kind=Form (context), got {kind}");
                if (!nextStream.IsEmpty) throw new Exception("Stream was not fully consumed after VERSION line.");
            });

            RunTest("HeaderParser.ParseVersionLine_MissingVersionKeyword_Throws", () => {
                string content = "1.0 CLASS\r\n";
                byte[] bytes = Encoding.Default.GetBytes(content);
                VB6Stream stream = new("./Tests/CSC_FormProperties.cls", bytes);
                bool threw = false;
                try { HeaderParser.ParseVersionLine(stream, HeaderKind.Class); }
                catch (VB6ParseException ex) { if (ex.Kind == VB6ErrorKind.KeywordNotFound) threw = true; else throw; }
                catch { throw; } // Rethrow other exceptions
                if (!threw) throw new Exception("Expected VB6ParseException for missing VERSION keyword.");
            });
        }

        // public static void TestVB6Stream_Advancing() { ... }
        // public static void TestVb6_TokenizingSimpleCode() { ... }
        // public static void TestProjectParser_MinimalVbp() { ... }
        // public static void TestFrmParser_SimpleForm() { ... }
        // etc.
    }
}