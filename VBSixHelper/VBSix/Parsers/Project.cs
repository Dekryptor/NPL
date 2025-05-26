using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using VBSix.Errors;

namespace VBSix.Parsers
{
    // --- Enums specific to Project parsing ---
    public enum Retained : short { UnloadOnExit = 0, RetainedInMemory = 1 }
    public enum UseExistingBrowser : short { DoNotUse = 0, Use = -1 }
    public enum StartMode : short { StandAlone = 0, Automation = 1 } // Project StartMode, distinct from control StartMode
    public enum InteractionMode : short { Interactive = 0, Unattended = -1 }
    public enum ServerSupportFiles : short { Local = 0, Remote = 1 } // VB6 has Remote as 1, not -1
    public enum UpgradeControls : short { Upgrade = 0, NoUpgrade = 1 }
    public enum UnusedControlInfo : short { Retain = 0, Remove = -1 } // VB6 has Remove as -1
    public enum CompatibilityMode : short { NoCompatibility = 0, Project = 1, CompatibleExe = 2 }
    public enum DebugStartupOption : short { WaitForComponentCreation = 0, StartComponent = 1, StartProgram = 2, StartBrowser = 3 }
    public enum CompileTargetType { Exe, Control, OleExe, OleDll } // Project's target type
    public enum ThreadingModel : short { SingleThreaded = 0, ApartmentThreaded = 1 }

    public class VersionInformation
    {
        public ushort Major { get; set; }
        public ushort Minor { get; set; }
        public ushort Revision { get; set; }
        public ushort AutoIncrementRevision { get; set; } // This is actually boolean-like (0 or -1/1) in VBP
        public string? CompanyName { get; set; } = string.Empty;
        public string? FileDescription { get; set; } = string.Empty;
        public string? Copyright { get; set; } = string.Empty;
        public string? Trademark { get; set; } = string.Empty;
        public string? ProductName { get; set; } = string.Empty;
        public string? Comments { get; set; } = string.Empty;
    }

    public abstract class VB6ProjectReferenceBase
    {
        public string OriginalValue { get; set; } = string.Empty; // The raw line content after "Reference=" or "Object="
    }

    public class CompiledProjectReference : VB6ProjectReferenceBase
    {
        public Guid ObjectGuid { get; set; }
        public string? Version { get; set; }    // e.g., "2.8"
        public string? Lcid { get; set; }       // e.g., "0" (Locale ID)
        public string? Path { get; set; }       // Path to the TLB/DLL/OCX
        public string? Description { get; set; } // Friendly name
    }

    public class SubProjectReference : VB6ProjectReferenceBase
    {
        public string Path { get; set; } = string.Empty; // Path to the .VBP file, e.g., "*\\AMySubProject.vbp"
    }

    public class VB6ProjectModule
    {
        public string? Name { get; set; }
        public string? Path { get; set; }
    }

    public class VB6ProjectClass
    {
        public string? Name { get; set; }
        public string? Path { get; set; }
    }

    public partial class VB6Project
    {
        public CompileTargetType? ProjectType { get; set; } // Nullable if not parsed yet
        public List<VB6ProjectReferenceBase> References { get; set; } = [];
        public List<VB6ObjectReference> Objects { get; set; } = [];
        public List<VB6ProjectModule> Modules { get; set; } = [];
        public List<VB6ProjectClass> Classes { get; set; } = [];
        public List<string> RelatedDocuments { get; set; } = [];
        public List<string> Designers { get; set; } = [];
        public List<string> Forms { get; set; } = [];
        public List<string> UserControls { get; set; } = [];
        public List<string> UserDocuments { get; set; } = [];
        public List<string> ProjectWindowLayout { get; set; } = []; // For ProjWin= lines
        public Dictionary<string, Dictionary<string, string>> OtherProperties { get; set; } = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        public UnusedControlInfo UnusedControlInfo { get; set; } = UnusedControlInfo.Remove;
        public UpgradeControls UpgradeControls { get; set; } = UpgradeControls.Upgrade;
        public string? ResFile32Path { get; set; } = string.Empty;
        public string? IconForm { get; set; } = string.Empty;
        public string? Startup { get; set; } // Can be "Sub Main", "Form1", or "(None)"
        public string? HelpFilePath { get; set; } = string.Empty;
        public string? Title { get; set; } = string.Empty;
        public string? Exe32FileName { get; set; } = string.Empty;
        public string? Exe32Compatible { get; set; } = string.Empty; // CompatibleEXE32 property
        public uint DllBaseAddress { get; set; } = 0x11000000;
        public string? Path32 { get; set; } = string.Empty;
        public string? CommandLineArguments { get; set; }
        public string? Name { get; set; } // Project Name, can be "(None)"
        public string? Description { get; set; } = string.Empty;
        public string? DebugStartupComponent { get; set; } = string.Empty;
        public string? HelpContextId { get; set; }
        public CompatibilityMode CompatibilityMode { get; set; } = CompatibilityMode.Project;
        public string? Version32Compatibility { get; set; } = string.Empty; // VersionCompatible32 property
        public VersionInformation VersionInfo { get; set; } = new VersionInformation();
        public ServerSupportFiles ServerSupportFiles { get; set; } = ServerSupportFiles.Local;
        public string? ConditionalCompile { get; set; } = string.Empty;
        public CompilationTypeInfoContainer CompilationType { get; set; } = new CompilationTypeInfoContainer(); // Defaults to PCode
        public int? CompileTarget { get; set; } // Stores the raw -1/0 from VBP for CompilationType

        public StartMode StartMode { get; set; } = StartMode.StandAlone;
        public InteractionMode Unattended { get; set; } = InteractionMode.Interactive;
        public Retained Retained { get; set; } = Retained.UnloadOnExit;
        public ushort? ThreadPerObject { get; set; }
        public ThreadingModel ThreadingModel { get; set; } = ThreadingModel.ApartmentThreaded;
        public ushort MaxNumberOfThreads { get; set; } = 1;
        public DebugStartupOption DebugStartupOption { get; set; } = DebugStartupOption.WaitForComponentCreation;
        public UseExistingBrowser UseExistingBrowser { get; set; } = UseExistingBrowser.Use;
        public string? PropertyPage { get; set; } = string.Empty; // For PropertyPage= lines (without GUID)
        public string? ProjectVersion { get; set; } // For ProjVer= in .VBG files
        public bool StartupOption { get; set; } // For StartupOption= in .VBG files

        // Regex for PropertyPage(GUID)=Name pattern, defined for VBP files
        private static readonly Regex PropertyPageGuidRegex = new Regex(@"^PropertyPage\(([^)]+)\)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);


        // --- Parsing Methods ---

        public static VB6Project ParseFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("FilePath cannot be empty.", nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException("Project file not found.", filePath);

            // VBP files are typically ANSI. Encoding.Default is often a good guess on Windows.
            // For robustness, allow specifying encoding or use a common one like Windows-1252 (CodePage 1252).
            Encoding registeredEncoding = Encoding.GetEncoding(1252);
            string[] lines = File.ReadAllLines(filePath, registeredEncoding);
            return ParseLines(lines, filePath);
        }

        public static VB6Project ParseFromBytes(string filePathForContext, byte[] sourceCodeBytes)
        {
            if (string.IsNullOrWhiteSpace(filePathForContext)) throw new ArgumentException("FilePath cannot be empty.", nameof(filePathForContext));
            if (sourceCodeBytes == null) throw new ArgumentNullException(nameof(sourceCodeBytes));

            Encoding registeredEncoding = Encoding.GetEncoding(1252);
            string[] lines = registeredEncoding.GetString(sourceCodeBytes).Split(["\r\n", "\r", "\n"], StringSplitOptions.None);
            return ParseLines(lines, filePathForContext);
        }

        private static VB6Project ParseLines(string[] lines, string filePathForContext)
        {
            var project = new VB6Project { FilePath = filePathForContext };
            string? currentSection = null;
            int lineNumber = 0;

            foreach (var rawLine in lines)
            {
                lineNumber++;
                string line = rawLine.Trim(); // Trim leading/trailing whitespace from the whole line first
                if (string.IsNullOrEmpty(line)) continue; // Skip empty lines

                if (line.StartsWith('[') && line.EndsWith(']'))
                {
                    currentSection = line[1..^1];
                    if (!project.OtherProperties.ContainsKey(currentSection))
                    {
                        project.OtherProperties[currentSection] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    }
                    continue;
                }

                var parts = line.Split(['='], 2);
                string key = parts[0].Trim(); // Key should also be trimmed
                string valueWithComment = parts.Length > 1 ? parts[1] : string.Empty;
                
                string value = Utilities.PropertyParserHelpers.UnquoteAndStripComment(valueWithComment);

                if (currentSection != null)
                {
                    project.OtherProperties[currentSection][key] = value;
                }
                else
                {
                    try
                    {
                        ProcessProjectProperty(project, key, value, rawLine);
                    }
                    catch (VB6ParseException) { throw; } // Rethrow specific parser errors
                    catch (Exception ex)
                    {
                        throw new VB6ParseException(
                            VB6ErrorKind.LineTypeUnknown,
                            filePathForContext,
                            rawLine, 
                            0, 0, lineNumber,
                            null,
                            new Exception($"Error processing project line (Line {lineNumber}): '{key}={value}' from raw line '{rawLine}'", ex)
                        );
                    }
                }
            }

            // Finalize CompilationType based on CompileTarget
            if (project.CompileTarget.HasValue)
            {
                project.CompilationType.Type = project.CompileTarget.Value == 0 ? CompilationOption.NativeCode : CompilationOption.PCode;
            }
            else // If CompileTarget was not set, default to PCode
            {
                project.CompilationType.Type = CompilationOption.PCode;
            }
            
            // Ensure NativeCodeSettings or PCodeSettings exists based on final type
            if (project.CompilationType.Type == CompilationOption.NativeCode && project.CompilationType.NativeCodeSettings == null)
            {
                project.CompilationType.NativeCodeSettings = new NativeCodeCompilation();
            }
            else if (project.CompilationType.Type == CompilationOption.PCode && project.CompilationType.PCodeSettings == null)
            {
                project.CompilationType.PCodeSettings = new PCodeCompilation();
            }


            if (project.ProjectType == null)
            {
                 throw new VB6ParseException(VB6ErrorKind.FirstLineNotProject, filePathForContext, string.Join(Environment.NewLine, lines),0,0,1);
            }

            return project;
        }

        private static void ProcessProjectProperty(VB6Project project, string key, string value, string rawLineForContext)
        {
            // Using StringComparison.OrdinalIgnoreCase for all key comparisons for robustness.
            if (key.Equals("Type", StringComparison.OrdinalIgnoreCase))
            {
                project.ProjectType = value.ToLowerInvariant() switch {
                    "exe" => CompileTargetType.Exe, "oledll" => CompileTargetType.OleDll,
                    "control" => CompileTargetType.Control, "oleexe" => CompileTargetType.OleExe,
                    _ => throw new VB6ParseException(VB6ErrorKind.ProjectTypeUnknown, new Exception($"Unknown project type: {value}"))
                };
            }
            else if (key.StartsWith("Reference", StringComparison.OrdinalIgnoreCase))
            {
                project.References.Add(ParseProjectReference(key, value, rawLineForContext));
            }
            else if (key.StartsWith("Object", StringComparison.OrdinalIgnoreCase))
            {
                project.Objects.Add(ParseObjectReference(key, value, rawLineForContext));
            }
            else if (key.Equals("Module", StringComparison.OrdinalIgnoreCase))
            {
                var parts = value.Split(';'); project.Modules.Add(new VB6ProjectModule { Name = parts[0].Trim(), Path = parts.Length > 1 ? parts[1].Trim() : null });
            }
            else if (key.Equals("Class", StringComparison.OrdinalIgnoreCase))
            {
                var parts = value.Split(';'); project.Classes.Add(new VB6ProjectClass { Name = parts[0].Trim(), Path = parts.Length > 1 ? parts[1].Trim() : null });
            }
            else if (key.Equals("Form", StringComparison.OrdinalIgnoreCase)) project.Forms.Add(value);
            else if (key.Equals("UserControl", StringComparison.OrdinalIgnoreCase)) project.UserControls.Add(value);
            else if (key.Equals("UserDocument", StringComparison.OrdinalIgnoreCase)) project.UserDocuments.Add(value);
            else if (key.Equals("Designer", StringComparison.OrdinalIgnoreCase)) project.Designers.Add(value);
            else if (key.Equals("RelatedDoc", StringComparison.OrdinalIgnoreCase)) project.RelatedDocuments.Add(value);
            else if (key.Equals("ProjWin", StringComparison.OrdinalIgnoreCase)) project.ProjectWindowLayout.Add(value);
            
            else if (key.Equals("Name", StringComparison.OrdinalIgnoreCase)) project.Name = (value == "!(None)!" || value == "") ? null : value;
            else if (key.Equals("Description", StringComparison.OrdinalIgnoreCase)) project.Description = value;
            else if (key.Equals("HelpFile", StringComparison.OrdinalIgnoreCase) || key.Equals("HelpFile32", StringComparison.OrdinalIgnoreCase)) project.HelpFilePath = value;
            else if (key.Equals("HelpContextID", StringComparison.OrdinalIgnoreCase)) project.HelpContextId = (value == "!(None)!" || value == "") ? null : value;
            else if (key.Equals("Title", StringComparison.OrdinalIgnoreCase)) project.Title = value;
            else if (key.Equals("ExeName32", StringComparison.OrdinalIgnoreCase)) project.Exe32FileName = value;
            else if (key.Equals("Path32", StringComparison.OrdinalIgnoreCase)) project.Path32 = value;
            else if (key.Equals("Command32", StringComparison.OrdinalIgnoreCase)) project.CommandLineArguments = (value == "!(None)!" || value == "") ? null : value;
            else if (key.Equals("Startup", StringComparison.OrdinalIgnoreCase)) project.Startup = (value == "!(None)!" || value == "") ? null : value;
            else if (key.Equals("IconForm", StringComparison.OrdinalIgnoreCase)) project.IconForm = value;
            else if (key.Equals("ResFile32", StringComparison.OrdinalIgnoreCase)) project.ResFile32Path = value;
            else if (key.Equals("CondComp", StringComparison.OrdinalIgnoreCase)) project.ConditionalCompile = value;
            else if (key.Equals("CompatibleEXE32", StringComparison.OrdinalIgnoreCase)) project.Exe32Compatible = value;
            else if (key.Equals("VersionCompatible32", StringComparison.OrdinalIgnoreCase)) project.Version32Compatibility = value;
            else if (key.Equals("DebugStartupComponent", StringComparison.OrdinalIgnoreCase)) project.DebugStartupComponent = value;
            else if (key.Equals("PropertyPage", StringComparison.OrdinalIgnoreCase)) project.PropertyPage = value; // Non-GUID version

            else if (PropertyPageGuidRegex.IsMatch(key)) // PropertyPage(GUID)=Name
            {
                Match match = PropertyPageGuidRegex.Match(key); // Re-match to extract GUID
                string guid = match.Groups[1].Value;
                if (!project.OtherProperties.ContainsKey("PropertyPageGuid")) 
                    project.OtherProperties["PropertyPageGuid"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                project.OtherProperties["PropertyPageGuid"][guid] = value;
            }

            else if (key.Equals("CompatibleMode", StringComparison.OrdinalIgnoreCase)) project.CompatibilityMode = short.TryParse(value, out var cm) ? (CompatibilityMode)cm : CompatibilityMode.NoCompatibility;
            else if (key.Equals("StartMode", StringComparison.OrdinalIgnoreCase)) project.StartMode = short.TryParse(value, out var sm) ? (StartMode)sm : StartMode.StandAlone;
            else if (key.Equals("Unattended", StringComparison.OrdinalIgnoreCase)) project.Unattended = short.TryParse(value, out var im) ? (InteractionMode)im : InteractionMode.Interactive;
            else if (key.Equals("Retained", StringComparison.OrdinalIgnoreCase)) project.Retained = short.TryParse(value, out var ret) ? (Retained)ret : Retained.UnloadOnExit;
            else if (key.Equals("ThreadingModel", StringComparison.OrdinalIgnoreCase)) project.ThreadingModel = short.TryParse(value, out var tm) ? (ThreadingModel)tm : ThreadingModel.ApartmentThreaded;
            else if (key.Equals("NoControlUpgrade", StringComparison.OrdinalIgnoreCase)) project.UpgradeControls = short.TryParse(value, out var ucv) && ucv == 1 ? UpgradeControls.NoUpgrade : UpgradeControls.Upgrade;
            else if (key.Equals("ServerSupportFiles", StringComparison.OrdinalIgnoreCase)) project.ServerSupportFiles = short.TryParse(value, out var ssf) && ssf != 0 ? ServerSupportFiles.Remote : ServerSupportFiles.Local; // -1 or 1 for remote in some contexts, 0 for local
            else if (key.Equals("RemoveUnusedControlInfo", StringComparison.OrdinalIgnoreCase)) project.UnusedControlInfo = short.TryParse(value, out var ruc) && ruc == 0 ? UnusedControlInfo.Retain : UnusedControlInfo.Remove;
            else if (key.Equals("UseExistingBrowser", StringComparison.OrdinalIgnoreCase)) project.UseExistingBrowser = short.TryParse(value, out var ueb) && ueb == 0 ? UseExistingBrowser.DoNotUse : UseExistingBrowser.Use;
            else if (key.Equals("DebugStartupOption", StringComparison.OrdinalIgnoreCase)) project.DebugStartupOption = short.TryParse(value, out var dso) ? (DebugStartupOption)dso : DebugStartupOption.WaitForComponentCreation;

            else if (key.Equals("DllBaseAddress", StringComparison.OrdinalIgnoreCase) || key.Equals("BaseAddress", StringComparison.OrdinalIgnoreCase))
                project.DllBaseAddress = uint.TryParse(value.StartsWith("&H", StringComparison.OrdinalIgnoreCase) ? value[2..] : value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var dba) ? dba : 0x11000000;
            
            else if (key.Equals("ThreadPerObject", StringComparison.OrdinalIgnoreCase)) project.ThreadPerObject = (value == "-1" || !ushort.TryParse(value, out var tpo)) ? (ushort?)null : tpo;
            else if (key.Equals("MaxNumberOfThreads", StringComparison.OrdinalIgnoreCase) || key.Equals("MaxThreads", StringComparison.OrdinalIgnoreCase)) project.MaxNumberOfThreads = ushort.TryParse(value, out var mnt) ? mnt : (ushort)1;
            
            // Version Info
            else if (key.Equals("MajorVer", StringComparison.OrdinalIgnoreCase)) project.VersionInfo.Major = ushort.TryParse(value, out var maj) ? maj : (ushort)0;
            else if (key.Equals("MinorVer", StringComparison.OrdinalIgnoreCase)) project.VersionInfo.Minor = ushort.TryParse(value, out var min) ? min : (ushort)0;
            else if (key.Equals("RevisionVer", StringComparison.OrdinalIgnoreCase)) project.VersionInfo.Revision = ushort.TryParse(value, out var rev) ? rev : (ushort)0;
            else if (key.Equals("AutoIncrementVer", StringComparison.OrdinalIgnoreCase)) project.VersionInfo.AutoIncrementRevision = (ushort)((value == "-1" || value == "1") ? 1 : 0); // AutoIncrement is 0 or -1 (True)

            else if (key.Equals("VersionCompanyName", StringComparison.OrdinalIgnoreCase)) project.VersionInfo.CompanyName = value;
            else if (key.Equals("VersionProductName", StringComparison.OrdinalIgnoreCase)) project.VersionInfo.ProductName = value;
            else if (key.Equals("VersionFileDescription", StringComparison.OrdinalIgnoreCase)) project.VersionInfo.FileDescription = value;
            else if (key.Equals("VersionLegalCopyright", StringComparison.OrdinalIgnoreCase)) project.VersionInfo.Copyright = value;
            else if (key.Equals("VersionLegalTrademarks", StringComparison.OrdinalIgnoreCase)) project.VersionInfo.Trademark = value;
            else if (key.Equals("VersionComments", StringComparison.OrdinalIgnoreCase)) project.VersionInfo.Comments = value;
            
            // Compilation Settings
            else if (key.Equals("CompileTarget", StringComparison.OrdinalIgnoreCase)) 
            {
                project.CompileTarget = short.TryParse(value, out var ct) ? ct : (short)-1;
            }
            else if (key.Equals("OptimizationType", StringComparison.OrdinalIgnoreCase) || key.Equals("Optimization", StringComparison.OrdinalIgnoreCase))
            {
                project.CompilationType.NativeCodeSettings ??= new NativeCodeCompilation();
                project.CompilationType.NativeCodeSettings.OptimizationTypeValue = short.TryParse(value, out var ot) ? (OptimizationType)ot : OptimizationType.FavorFastCode;
            }
            else if (key.Equals("FavorPentiumPro(tm)", StringComparison.OrdinalIgnoreCase))
            {
                project.CompilationType.NativeCodeSettings ??= new NativeCodeCompilation();
                project.CompilationType.NativeCodeSettings.FavorPentiumProValue = short.TryParse(value, out var fpp) && fpp == -1 ? FavorPentiumPro.True : FavorPentiumPro.False;
            }
            else if (key.Equals("CodeViewDebugInfo", StringComparison.OrdinalIgnoreCase))
            {
                project.CompilationType.NativeCodeSettings ??= new NativeCodeCompilation();
                project.CompilationType.NativeCodeSettings.CodeViewDebugInfoValue = short.TryParse(value, out var cvdi) && cvdi == -1 ? CodeViewDebugInfo.Created : CodeViewDebugInfo.NotCreated;
            }
            else if (key.Equals("NoAliasing", StringComparison.OrdinalIgnoreCase))
            {
                project.CompilationType.NativeCodeSettings ??= new NativeCodeCompilation();
                project.CompilationType.NativeCodeSettings.AliasingValue = short.TryParse(value, out var na) && na == -1 ? Aliasing.AssumeNoAliasing : Aliasing.AssumeAliasing;
            }
            else if (key.Equals("BoundsCheck", StringComparison.OrdinalIgnoreCase))
            {
                project.CompilationType.NativeCodeSettings ??= new NativeCodeCompilation();
                project.CompilationType.NativeCodeSettings.BoundsCheckValue = short.TryParse(value, out var bc) && bc == -1 ? BoundsCheck.NoCheck : BoundsCheck.CheckBounds;
            }
            else if (key.Equals("OverflowCheck", StringComparison.OrdinalIgnoreCase))
            {
                project.CompilationType.NativeCodeSettings ??= new NativeCodeCompilation();
                project.CompilationType.NativeCodeSettings.OverflowCheckValue = short.TryParse(value, out var oc) && oc == -1 ? OverflowCheck.NoCheck : OverflowCheck.CheckOverflow;
            }
            else if (key.Equals("FlPointCheck", StringComparison.OrdinalIgnoreCase))
            {
                project.CompilationType.NativeCodeSettings ??= new NativeCodeCompilation();
                project.CompilationType.NativeCodeSettings.FlPointCheckValue = short.TryParse(value, out var fpc) && fpc == -1 ? FloatingPointErrorCheck.NoCheck : FloatingPointErrorCheck.CheckFloatingPointError;
            }
            else if (key.Equals("FDIVCheck", StringComparison.OrdinalIgnoreCase))
            {
                project.CompilationType.NativeCodeSettings ??= new NativeCodeCompilation();
                project.CompilationType.NativeCodeSettings.FDIVCheckValue = short.TryParse(value, out var fdc) && fdc == -1 ? PentiumFDivBugCheck.NoPentiumFDivBugCheck : PentiumFDivBugCheck.CheckPentiumFDivBug;
            }
            else if (key.Equals("UnroundedFP", StringComparison.OrdinalIgnoreCase))
            {
                project.CompilationType.NativeCodeSettings ??= new NativeCodeCompilation();
                project.CompilationType.NativeCodeSettings.UnroundedFPValue = short.TryParse(value, out var ufp) && ufp == -1 ? UnroundedFloatingPoint.Allow : UnroundedFloatingPoint.DoNotAllow;
            }
            else // Store unhandled properties
            {
                if (!project.OtherProperties.ContainsKey("General")) project.OtherProperties["General"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                project.OtherProperties["General"][key] = value;
            }
        }
        
        private static VB6ProjectReferenceBase ParseProjectReference(string key, string value, string rawLineForContext)
        {
            if (value.StartsWith("*\\A", StringComparison.OrdinalIgnoreCase) || value.StartsWith("*/A", StringComparison.OrdinalIgnoreCase))
            {
                return new SubProjectReference { Path = value, OriginalValue = rawLineForContext };
            }
            else if (value.StartsWith("*\\G{", StringComparison.OrdinalIgnoreCase) || value.StartsWith("*/G{", StringComparison.OrdinalIgnoreCase))
            {
                // Example: "*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation"
                var parts = value[3..].Split('#'); // Skip "*\G" or "*/G"
                string guidStr = parts[0];
                if (guidStr.EndsWith('}')) guidStr = guidStr[..^1];
                
                if (!Guid.TryParse(guidStr, out var guid))
                {
                    throw new VB6ParseException(VB6ErrorKind.UnableToParseUuid, new Exception($"Invalid GUID in reference: {guidStr} from line: {rawLineForContext}"));
                }

                return new CompiledProjectReference
                {
                    ObjectGuid = guid,
                    Version = parts.Length > 1 ? parts[1].Trim() : null,
                    Lcid = parts.Length > 2 ? parts[2].Trim() : null, 
                    Path = parts.Length > 3 ? parts[3].Trim() : null,    
                    Description = parts.Length > 4 ? parts[4].Trim() : null,
                    OriginalValue = rawLineForContext
                };
            }
            throw new VB6ParseException(VB6ErrorKind.ReferenceMissingSections, new Exception($"Unknown reference format: {value} from line: {rawLineForContext}"));
        }

        private static VB6ObjectReference ParseObjectReference(string key, string value, string rawLineForContext)
        {
            // Value is already unquoted and comment-stripped
            var objRef = new VB6ObjectReference { OriginalValue = rawLineForContext };

            // Check for sub-project reference first, similar to VB6ProjectReference
            // Object=*\\A..\\MyProj.vbp
            if (value.StartsWith("*\\A", StringComparison.OrdinalIgnoreCase) || value.StartsWith("*/A", StringComparison.OrdinalIgnoreCase))
            {
                objRef.Path = value; // The whole value is the path
            }
            // GUID-based compiled object: Object={GUID}#Version#Unknown1;FileName
            else 
            {
                var mainParts = value.Split([';'], 2); // Split by first semicolon
                var refPart = mainParts[0].Trim();
                objRef.FileName = mainParts.Length > 1 ? mainParts[1].Trim() : null;

                var refDetails = refPart.Split('#');
                string guidStr = refDetails[0].Trim();
                if (guidStr.StartsWith('{') && guidStr.EndsWith('}')) guidStr = guidStr[1..^1];
                
                if (!Guid.TryParse(guidStr, out var parsedGuid))
                    throw new VB6ParseException(VB6ErrorKind.UnableToParseUuid, new Exception($"Invalid GUID in object reference: {guidStr} from line {rawLineForContext}"));
                objRef.ObjectGuidString = parsedGuid.ToString("B").ToUpper(); // Store with braces for standard format if desired, or just guidStr
                
                objRef.Version = refDetails.Length > 1 ? refDetails[1].Trim() : null;
                objRef.LocaleID = refDetails.Length > 2 ? refDetails[2].Trim() : null;
            }
            return objRef;
        }

        // Adding GetSubprojectReferences and GetCompiledReferences as instance methods
        public List<SubProjectReference> GetSubprojectReferences()
        {
            return [.. References.OfType<SubProjectReference>()];
        }

        public List<CompiledProjectReference> GetCompiledReferences()
        {
            return [.. References.OfType<CompiledProjectReference>()];
        }
    }
    
    // Ensure FilePath is available on VB6Project
    public partial class VB6Project
    {
        public string? FilePath { get; internal set; }
    }
}