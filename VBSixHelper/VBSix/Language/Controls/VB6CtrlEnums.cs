namespace VBSix.Language.Controls
{
    // --- Common Enums (used by multiple controls) ---

    /// <summary>
    /// Determines if graphics are repainted from a persistent bitmap (Automatic) or if Paint events are raised (Manual).
    /// Affects Form, PictureBox.
    /// </summary>
    public enum AutoRedraw { Manual = 0, Automatic = -1 }

    /// <summary>
    /// Specifies the text display direction, primarily for systems supporting right-to-left languages.
    /// Affects many controls with text. VB6 Property: RightToLeft.
    /// Note: VB6 actually has three values: 0 (vbContextual), 1 (vbLeftToRight), 2 (vbRightToLeft).
    /// </summary>
    public enum TextDirection { LeftToRight = 0, RightToLeft = -1 }

    /// <summary>
    /// Determines if a control automatically resizes to fit its contents.
    /// Affects Label, PictureBox.
    /// </summary>
    public enum AutoSize { Fixed = 0, Resize = -1 }

    /// <summary>
    /// Determines if a control can respond to user-generated events.
    /// Affects most controls. VB6 Property: Enabled.
    /// </summary>
    public enum Activation { Disabled = 0, Enabled = -1 }

    /// <summary>
    /// Determines if a control can receive focus via the Tab key.
    /// Affects most interactive controls. VB6 Property: TabStop.
    /// </summary>
    public enum TabStop { ProgrammaticOnly = 0, Included = -1 }

    /// <summary>
    /// Determines if a control is visible or hidden.
    /// Affects most controls. VB6 Property: Visible.
    /// </summary>
    public enum Visibility { Hidden = 0, Visible = -1 }

    /// <summary>
    /// For Form/PictureBox, indicates if it has its own persistent device context (hDC).
    /// VB6 Property: HasDC.
    /// </summary>
    public enum HasDeviceContext { No = 0, Yes = -1 }

    /// <summary>
    /// Determines if the MaskColor property is used to create transparent areas in an image.
    /// Affects CheckBox, CommandButton, OptionButton. VB6 Property: UseMaskColor.
    /// </summary>
    public enum UseMaskColor { DoNotUseMaskColor = 0, UseMaskColor = -1 }

    /// <summary>
    /// Determines if focus shifting from this control triggers the Validate event.
    /// Affects most input-capable controls. VB6 Property: CausesValidation.
    /// </summary>
    public enum CausesValidation { No = 0, Yes = -1 }

    /// <summary>
    /// Determines if a Form or MDIForm can be moved by the user.
    /// VB6 Property: Moveable.
    /// </summary>
    public enum Movability { Fixed = 0, Moveable = -1 }

    /// <summary>
    /// Determines if text background is opaque or transparent on Form/PictureBox.
    /// VB6 Property: FontTransparent.
    /// </summary>
    public enum FontTransparency { Opaque = 0, Transparent = -1 }

    /// <summary>
    /// Determines if context-sensitive Help uses "What's This?" pop-up or F1 Help window.
    /// Affects Form, MDIForm. VB6 Property: WhatsThisHelp.
    /// </summary>
    public enum WhatsThisHelp { F1Help = 0, WhatsThisHelpMode = -1 }

    /// <summary>
    /// Type of DDE (Dynamic Data Exchange) link.
    /// Affects Form. VB6 Property: LinkMode.
    /// </summary>
    public enum FormLinkMode { None = 0, Source = 1 }

    /// <summary>
    /// The display state of a Form or MDIForm (Normal, Minimized, Maximized).
    /// VB6 Property: WindowState.
    /// </summary>
    public enum WindowState { Normal = 0, Minimized = 1, Maximized = 2 }

    /// <summary>
    /// How a control is aligned within its container if the container supports alignment (e.g., PictureBox on a Form).
    /// Affects Data, PictureBox. VB6 Property: Align.
    /// </summary>
    public enum Align { None = 0, Top = 1, Bottom = 2, Left = 3, Right = 4 }

    /// <summary>
    /// Text alignment relative to the control (CheckBox, OptionButton).
    /// VB6 Property: Alignment.
    /// </summary>
    public enum JustifyAlignment { LeftJustify = 0, RightJustify = 1 }

    /// <summary>
    /// Text alignment within the control (Label, TextBox).
    /// VB6 Property: Alignment.
    /// </summary>
    public enum Alignment { LeftJustify = 0, RightJustify = 1, Center = 2 }

    /// <summary>
    /// Background style (Opaque or Transparent).
    /// Affects Label, Shape, OLE. VB6 Property: BackStyle.
    /// </summary>
    public enum BackStyle { Transparent = 0, Opaque = 1 }

    /// <summary>
    /// Visual appearance (Flat or 3D).
    /// Affects many controls. VB6 Property: Appearance.
    /// </summary>
    public enum Appearance { Flat = 0, ThreeD = 1 }

    /// <summary>
    /// Border style for many controls (None or FixedSingle).
    /// Affects Frame, Image, Label, OLE, PictureBox, TextBox. VB6 Property: BorderStyle.
    /// </summary>
    public enum BorderStyle { None = 0, FixedSingle = 1 }

    /// <summary>
    /// Drag-and-drop initiation mode.
    /// Affects most controls. VB6 Property: DragMode.
    /// </summary>
    public enum DragMode { Manual = 0, Automatic = 1 }

    /// <summary>
    /// How the pen interacts with the background for graphics methods.
    /// Affects Form, Line, PictureBox, Shape. VB6 Property: DrawMode.
    /// </summary>
    public enum DrawMode { Blackness=1, NotMergePen=2, MaskNotPen=3, NotCopyPen=4, MaskPenNot=5, Invert=6, XorPen=7, NotMaskPen=8, MaskPen=9, NotXorPen=10, Nop=11, MergeNotPen=12, CopyPen=13, MergePenNot=14, MergePen=15, Whiteness=16 }

    /// <summary>
    /// Line style for graphics methods or border.
    /// Affects Form, Line, PictureBox, Shape. VB6 Property: DrawStyle (for graphics), BorderStyle (for Line control).
    /// </summary>
    public enum DrawStyle { Solid=0, Dash=1, Dot=2, DashDot=3, DashDotDot=4, Transparent=5, InsideSolid=6 }

    /// <summary>
    /// The appearance of the mouse cursor when over a control.
    /// Affects most controls. VB6 Property: MousePointer.
    /// </summary>
    public enum MousePointer { Default=0, Arrow=1, Cross=2, IBeam=3, Icon=4, Size=5, SizeNESW=6, SizeNS=7, SizeNWSE=8, SizeWE=9, UpArrow=10, Hourglass=11, NoDrop=12, ArrowHourglass=13, ArrowQuestion=14, SizeAll=15, Custom=99 }

    /// <summary>
    /// OLE drag-and-drop initiation mode.
    /// Affects controls that can be OLE drag sources. VB6 Property: OLEDragMode.
    /// </summary>
    public enum OLEDragMode { Manual = 0, Automatic = 1 }

    /// <summary>
    /// Whether/how a control accepts OLE drops.
    /// Affects controls that can be OLE drop targets. VB6 Property: OLEDropMode.
    /// </summary>
    public enum OLEDropMode { None = 0, Manual = 1 } // VB6 also has Automatic=2

    /// <summary>
    /// Visual style (Standard or Graphical with pictures).
    /// Affects CheckBox, CommandButton, OptionButton. VB6 Property: Style.
    /// </summary>
    public enum Style { Standard = 0, Graphical = 1 }

    /// <summary>
    /// How the interior of a shape or control is filled for graphics methods.
    /// Affects Form, PictureBox, Shape. VB6 Property: FillStyle.
    /// </summary>
    public enum FillStyle { Solid=0, Transparent=1, HorizontalLine=2, VerticalLine=3, UpwardDiagonal=4, DownwardDiagonal=5, Cross=6, DiagonalCross=7 }

    /// <summary>
    /// DDE link mode for data exchange.
    /// Affects Label, TextBox, PictureBox, OLE. VB6 Property: LinkMode (distinct from FormLinkMode).
    /// </summary>
    public enum LinkMode { None = 0, Automatic = 1, Manual = 2, Notify = 3 }

    /// <summary>
    /// Item selection mode for ListBox-like controls.
    /// Affects ListBox, FileListBox. VB6 Property: MultiSelect.
    /// </summary>
    public enum MultiSelect { None = 0, Simple = 1, Extended = 2 }

    /// <summary>
    /// Coordinate system used for graphics methods and control positioning.
    /// Affects Form, PictureBox. VB6 Property: ScaleMode.
    /// </summary>
    public enum ScaleMode { User=0, Twip=1, Point=2, Pixel=3, Character=4, Inches=5, Millimeter=6, Centimeter=7 }

    /// <summary>
    /// How an OLE object is sized within its container.
    /// Affects OLE container. VB6 Property: SizeMode.
    /// (PictureBox has a boolean AutoSize and Stretch, not this enum directly for sizing its picture)
    /// </summary>
    public enum SizeMode { Clip=0, Stretch=1, AutoSize=2, Zoom=3 }

    /// <summary>
    /// Whether child controls are clipped to the container's boundaries during painting.
    /// Affects Form, Frame, PictureBox. VB6 Property: ClipControls.
    /// </summary>
    public enum ClipControls { Unbounded = 0, Clipped = 1 }


    // --- Control-Specific Enums ---

    // From: CheckBox
    public enum CheckBoxValue { Unchecked = 0, Checked = 1, Grayed = 2 }

    // From: ComboBox
    public enum ComboBoxStyle { DropDownCombo = 0, SimpleCombo = 1, DropDownList = 2 }

    // From: Data control
    public enum BOFAction { MoveFirst = 0, Bof = 1 }
    public enum Connection { Access, DBaseIII, DBaseIV, DBase5_0, Excel3_0, Excel4_0, Excel5_0, Excel8_0, FoxPro2_0, FoxPro2_5, FoxPro2_6, FoxPro3_0, LotusWk1, LotusWk3, LotusWk4, Paradox3X, Paradox4X, Paradox5X, Text }
    public enum DefaultCursorType { Default = 0, Odbc = 1, ServerSide = 2 }
    public enum DefaultType { UseODBC = 1, UseJet = 2 } // For Data control's DefaultType property
    public enum EOFAction { MoveLast = 0, Eof = 1, AddNew = 2 }
    public enum RecordSetType { Table = 0, Dynaset = 1, Snapshot = 2 }

    // From: FileListBox
    public enum ArchiveAttribute { Exclude = 0, Include = -1 }
    public enum HiddenAttribute { Exclude = 0, Include = -1 }
    public enum NormalAttribute { Exclude = 0, Include = -1 }
    public enum ReadOnlyAttribute { Exclude = 0, Include = -1 } // Named to avoid conflict with System.IO.FileAttributes.ReadOnly
    public enum SystemAttribute { Exclude = 0, Include = -1 } // Named to avoid conflict with System.IO.FileAttributes.System

    // From: Form
    public enum PaletteMode { HalfTone = 0, UseZOrder = 1, Custom = 2 }
    public enum FormBorderStyle { None = 0, FixedSingle = 1, Sizable = 2, FixedDialog = 3, FixedToolWindow = 4, SizableToolWindow = 5 } // Distinct from common BorderStyle
    public enum ControlBox { Excluded = 0, Included = -1 }
    public enum MaxButton { Excluded = 0, Included = -1 }
    public enum MinButton { Excluded = 0, Included = -1 }
    public enum WhatsThisButton { Excluded = 0, Included = -1 }
    public enum ShowInTaskbar { Hide = 0, Show = -1 }
    public enum StartUpPositionType { Manual = 0, CenterOwner = 1, CenterScreen = 2, WindowsDefault = 3 } // Used with FormProperties

    // From: Label
    public enum WordWrap { NonWrapping = 0, Wrapping = -1 }

    // From: ListBox
    public enum ListBoxStyle { Standard = 0, Checkbox = 1 }

    // From: Menus
    public enum NegotiatePosition { None = 0, Left = 1, Middle = 2, Right = 3 }
    public enum ShortCut 
    { 
        None,
        CtrlA, CtrlB, CtrlC, CtrlD, CtrlE, CtrlF, CtrlG, CtrlH, CtrlI, CtrlJ, CtrlK, CtrlL, CtrlM,
        CtrlN, CtrlO, CtrlP, CtrlQ, CtrlR, CtrlS, CtrlT, CtrlU, CtrlV, CtrlW, CtrlX, CtrlY, CtrlZ,
        F1, F2, F3, F4, F5, F6, F7, F8, F9, /* F10 is invalid */ F11, F12,
        CtrlF1, CtrlF2, CtrlF3, CtrlF4, CtrlF5, CtrlF6, CtrlF7, CtrlF8, CtrlF9, /* CtrlF10 */ CtrlF11, CtrlF12,
        ShiftF1, ShiftF2, ShiftF3, ShiftF4, ShiftF5, ShiftF6, ShiftF7, ShiftF8, ShiftF9, /* ShiftF10 */ ShiftF11, ShiftF12,
        ShiftCtrlF1, ShiftCtrlF2, ShiftCtrlF3, ShiftCtrlF4, ShiftCtrlF5, ShiftCtrlF6, ShiftCtrlF7, ShiftCtrlF8, ShiftCtrlF9, /* ShiftCtrlF10 */ ShiftCtrlF11, ShiftCtrlF12,
        CtrlIns, ShiftIns, Del, ShiftDel, AltBKsp 
    }

    // From: OLE
    public enum OLETypeAllowed { Link = 0, Embedded = 1, Either = 2 }
    public enum UpdateOptions { Automatic = 0, Frozen = 1, Manual = 2 }
    public enum AutoActivate { Manual = 0, GetFocus = 1, DoubleClick = 2, Automatic = 3 }
    public enum DisplayType { Content = 0, Icon = 1 }

    // OptionButton
    public enum OptionButtonValue { UnSelected = 0, Selected = 1 }

    // From: Shape
    public enum ShapeType { Rectangle = 0, Square = 1, Oval = 2, Circle = 3, RoundedRectangle = 4, RoundSquare = 5 }

    // From: TextBox
    public enum ScrollBars { None = 0, Horizontal = 1, Vertical = 2, Both = 3 } // For TextBox scrollbars
    public enum MultiLine { SingleLine = 0, MultiLine = -1 } // For TextBox
}