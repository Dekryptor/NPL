namespace VBSix.Parsers
{
    /// <summary>
    /// Specifies the primary compilation output type for a VB6 project.
    /// This corresponds to the `CompilationType` property in a .VBP file (key "CompileTarget" in IDE, value is -1 for P-Code or 0 for Native Code).
    /// </summary>
    public enum CompilationOption : short
    {
        /// <summary>
        /// Compile to P-Code. Value in VBP: -1.
        /// This is the default compilation type if not otherwise specified or if Native Code options are absent.
        /// </summary>
        PCode = -1,

        /// <summary>
        /// Compile to Native Code. Value in VBP: 0.
        /// </summary>
        NativeCode = 0,
    }

    /// <summary>
    /// Specifies the optimization strategy for Native Code compilation.
    /// Corresponds to the `OptimizationType` property in a .VBP file (often named "Optimization" in the VBP file).
    /// VB6 IDE Values: 0 = No Optimization, 1 = Favor Fast Code, 2 = Favor Small Code.
    /// </summary>
    public enum OptimizationType : short
    {
        /// <summary>No optimization. VBP Value: 0.</summary>
        NoOptimization = 0, // Matched to typical VBP value for "No Optimization"
        /// <summary>Optimize for fast code. VBP Value: 1. This is often the VB6 IDE default.</summary>
        FavorFastCode = 1,  // Matched to typical VBP value for "Favor Fast Code"
        /// <summary>Optimize for small code. VBP Value: 2.</summary>
        FavorSmallCode = 2, // Matched to typical VBP value for "Favor Small Code"
    }

    /// <summary>
    /// Specifies whether to favor the Pentium Pro(tm) processor for optimizations.
    /// Corresponds to `FavorPentiumPro(tm)` in a .VBP file.
    /// VBP Values: 0 = False, -1 = True.
    /// </summary>
    public enum FavorPentiumPro : short
    {
        /// <summary>Do not favor Pentium Pro. VBP Value: 0. (Default)</summary>
        False = 0,
        /// <summary>Favor Pentium Pro. VBP Value: -1.</summary>
        True = -1
    }

    /// <summary>
    /// Specifies the level of debug information to include.
    /// Corresponds to `CodeViewDebugInfo` in a .VBP file.
    /// VBP Values: 0 = No Debug Info, -1 = Full Debug Info.
    /// (VB6 IDE also supports "Symbolic Debug Info" which might map to 1).
    /// </summary>
    public enum CodeViewDebugInfo : short
    {
        /// <summary>No debug information. VBP Value: 0. (Default in many contexts)</summary>
        NotCreated = 0,
        /// <summary>Full debug information. VBP Value: -1.</summary>
        Created = -1,
        // Symbolic = 1, // If "Symbolic Debug Info" (value 1 in some IDEs) needs to be distinct
    }

    /// <summary>
    /// Specifies whether the compiler should assume no aliasing for pointers/references.
    /// Corresponds to the `NoAliasing` property in a .VBP file.
    /// VBP Values: 0 = Assume Aliasing (NoAliasing is False), -1 = Assume No Aliasing (NoAliasing is True).
    /// </summary>
    public enum Aliasing : short 
    {
        /// <summary>Compiler assumes aliasing can occur. VBP NoAliasing=0. (Default)</summary>
        AssumeAliasing = 0,
        /// <summary>Compiler assumes no aliasing occurs. VBP NoAliasing=-1.</summary>
        AssumeNoAliasing = -1
    }

    /// <summary>
    /// Specifies whether to perform array bounds checking.
    /// Corresponds to the `BoundsCheck` property in a .VBP file.
    /// VBP Values: 0 = No Bounds Check (i.e., "Remove array bounds checks" is True).
    ///             -1 = Perform Bounds Check (i.e., "Remove array bounds checks" is False - Default for safety).
    /// </summary>
    public enum BoundsCheck : short
    {
        /// <summary>Do not perform array bounds checking. VBP BoundsCheck=0.</summary>
        NoCheck = 0, 
        /// <summary>Perform array bounds checking. VBP BoundsCheck=-1. (Safer default)</summary>
        CheckBounds = -1 
    }

    /// <summary>
    /// Specifies whether to perform integer overflow checking.
    /// Corresponds to the `OverflowCheck` property in a .VBP file.
    /// VBP Values: 0 = No Overflow Check, -1 = Perform Overflow Check (Default).
    /// </summary>
    public enum OverflowCheck : short
    {
        /// <summary>Do not perform integer overflow checking. VBP OverflowCheck=0.</summary>
        NoCheck = 0,
        /// <summary>Perform integer overflow checking. VBP OverflowCheck=-1. (Safer default)</summary>
        CheckOverflow = -1
    }

    /// <summary>
    /// Specifies whether to perform floating-point error checking.
    /// Corresponds to the `FlPointCheck` property in a .VBP file.
    /// VBP Values: 0 = No Floating Point Check, -1 = Perform Floating Point Check (Default).
    /// </summary>
    public enum FloatingPointErrorCheck : short 
    {
        /// <summary>Do not perform floating-point error checking. VBP FlPointCheck=0.</summary>
        NoCheck = 0,
        /// <summary>Perform floating-point error checking. VBP FlPointCheck=-1. (Safer default)</summary>
        CheckFloatingPointError = -1
    }

    /// <summary>
    /// Specifies whether to check for the Pentium FDIV bug.
    /// Corresponds to the `FDIVCheck` property in a .VBP file.
    /// VBP Values: 0 = Check for bug, -1 = Do not check (Default).
    /// </summary>
    public enum PentiumFDivBugCheck : short 
    {
        /// <summary>Check for Pentium FDIV bug. VBP FDIVCheck=0.</summary>
        CheckPentiumFDivBug = 0,
        /// <summary>Do not check for Pentium FDIV bug. VBP FDIVCheck=-1. (Default)</summary>
        NoPentiumFDivBugCheck = -1
    }

    /// <summary>
    /// Specifies whether to allow unrounded floating-point operations.
    /// Corresponds to the `UnroundedFP` property in a .VBP file.
    /// VBP Values: 0 = Do Not Allow (Default), -1 = Allow.
    /// </summary>
    public enum UnroundedFloatingPoint : short 
    {
        /// <summary>Do not allow unrounded floating-point operations. VBP UnroundedFP=0. (Default)</summary>
        DoNotAllow = 0,
        /// <summary>Allow unrounded floating-point operations. VBP UnroundedFP=-1.</summary>
        Allow = -1
    }

    /// <summary>
    /// Container for settings specific to Native Code compilation.
    /// Default values align with typical VB6 IDE settings for "Optimize for Fast Code"
    /// with standard safety checks enabled.
    /// </summary>
    public class NativeCodeCompilation
    {
        public OptimizationType OptimizationTypeValue { get; set; } = OptimizationType.FavorFastCode;
        public FavorPentiumPro FavorPentiumProValue { get; set; } = FavorPentiumPro.False;
        public CodeViewDebugInfo CodeViewDebugInfoValue { get; set; } = CodeViewDebugInfo.NotCreated; 
        public Aliasing AliasingValue { get; set; } = Aliasing.AssumeAliasing; 
        public BoundsCheck BoundsCheckValue { get; set; } = BoundsCheck.CheckBounds; 
        public OverflowCheck OverflowCheckValue { get; set; } = OverflowCheck.CheckOverflow; 
        public FloatingPointErrorCheck FlPointCheckValue { get; set; } = FloatingPointErrorCheck.CheckFloatingPointError; 
        public PentiumFDivBugCheck FDIVCheckValue { get; set; } = PentiumFDivBugCheck.NoPentiumFDivBugCheck; 
        public UnroundedFloatingPoint UnroundedFPValue { get; set; } = UnroundedFloatingPoint.DoNotAllow; 
    }

    /// <summary>
    /// Placeholder class for P-Code compilation settings.
    /// VB6 P-Code compilation typically has fewer fine-grained options exposed in the VBP file
    /// compared to Native Code.
    /// </summary>
    public class PCodeCompilation
    {
        // Currently no specific P-Code sub-settings are parsed from VBP files
    }

    /// <summary>
    /// A container class to hold either Native Code or P-Code compilation settings,
    /// </summary>
    public class CompilationTypeInfoContainer
    {
        private CompilationOption _type = CompilationOption.PCode; // Default to PCode

        /// <summary>
        /// Gets or sets the primary type of compilation (P-Code or Native Code).
        /// When set, it automatically initializes or clears the appropriate settings object
        /// (<see cref="NativeCodeSettings"/> or <see cref="PCodeSettings"/>).
        /// </summary>
        public CompilationOption Type
        {
            get => _type;
            set
            {
                if (_type != value)
                {
                    _type = value;
                    InitializeSettingsBasedOnType();
                }
            }
        }

        /// <summary>
        /// Settings applicable if <see cref="Type"/> is <see cref="CompilationOption.NativeCode"/>.
        /// Will be null if type is P-Code.
        /// </summary>
        public NativeCodeCompilation? NativeCodeSettings { get; set; }

        /// <summary>
        /// Settings applicable if <see cref="Type"/> is <see cref="CompilationOption.PCode"/>.
        /// Will be null if type is Native Code.
        /// </summary>
        public PCodeCompilation? PCodeSettings { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CompilationTypeInfoContainer"/> class.
        /// The default compilation type is P-Code.
        /// </summary>
        public CompilationTypeInfoContainer()
        {
            InitializeSettingsBasedOnType(); // Initialize with default PCode
        }
        
        /// <summary>
        /// Initializes or re-initializes the settings objects based on the current <see cref="Type"/>.
        /// </summary>
        private void InitializeSettingsBasedOnType()
        {
            if (_type == CompilationOption.NativeCode)
            {
                NativeCodeSettings = new NativeCodeCompilation();
                PCodeSettings = null;
            }
            else // PCode
            {
                PCodeSettings = new PCodeCompilation();
                NativeCodeSettings = null;
            }
        }
    }
}