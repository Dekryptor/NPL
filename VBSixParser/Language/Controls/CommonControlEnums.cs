// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Determines if the control is redrawn automatically or manually.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245029(v=vs.60)
    /// </summary>
    public enum AutoRedraw
    {
        /// <summary>Disables automatic repainting. Paint event invoked when necessary. (Default)</summary>
        Manual = 0,
        /// <summary>Enables automatic repainting. Graphics written to screen and memory image.</summary>
        Automatic = -1 // VB uses -1 for True often
    }

    /// <summary>
    /// Determines the direction in which text is displayed.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa442921(v=vs.60)
    /// </summary>
    public enum TextDirection
    {
        /// <summary>Text is ordered from left to right. (Default)</summary>
        LeftToRight = 0,
        /// <summary>Text is ordered from right to left.</summary>
        RightToLeft = -1
    }

    /// <summary>
    /// Determines if the control is automatically resized to fit its contents (Label, PictureBox).
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245034(v=vs.60)
    /// </summary>
    public enum AutoSize
    {
        /// <summary>Keeps the size of the control constant. (Default)</summary>
        Fixed = 0,
        /// <summary>Automatically resizes the control to display its entire contents.</summary>
        Resize = -1
    }

    /// <summary>
    /// Determines if a control or form can respond to user-generated events.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267301(v=vs.60)
    /// </summary>
    public enum Activation // Corresponds to Enabled property in VB
    {
        /// <summary>The control is disabled.</summary>
        Disabled = 0, // VB False
        /// <summary>The control is enabled. (Default)</summary>
        Enabled = -1  // VB True
    }

    /// <summary>
    /// Determines if the control is included in the tab order.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445721(v=vs.60)
    /// </summary>
    public enum TabStop
    {
        /// <summary>Bypasses the object when tabbing.</summary>
        ProgrammaticOnly = 0, // VB False
        /// <summary>Designates the object as a tab stop. (Default)</summary>
        Included = -1         // VB True
    }

    /// <summary>
    /// Determines if the control is visible or hidden.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445768(v=vs.60)
    /// </summary>
    public enum Visibility // Corresponds to Visible property
    {
        /// <summary>The control is not visible.</summary>
        Hidden = 0,  // VB False
        /// <summary>The control is visible. (Default)</summary>
        Visible = -1 // VB True
    }

    /// <summary>
    /// Determines if the control has a device context.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245860(v=vs.60)
    /// </summary>
    public enum HasDeviceContext // Corresponds to HasDC property
    {
        /// <summary>The control does not have a device context.</summary>
        No = 0,
        /// <summary>The control has a device context. (Default)</summary>
        Yes = -1
    }

    /// <summary>
    /// Determines whether the color assigned in the MaskColor property is used as a mask.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445753(v=vs.60)
    /// </summary>
    public enum UseMaskColor
    {
        /// <summary>The control does not use the mask color. (Default)</summary>
        DoNotUseMaskColor = 0,
        /// <summary>The MaskColor property is used as a mask.</summary>
        UseMaskColor = -1
    }

    /// <summary>
    /// Determines if the control causes validation to occur on the control losing focus.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245065(v=vs.60)
    /// </summary>
    public enum CausesValidation
    {
        /// <summary>The control does not cause validation.</summary>
        No = 0,
        /// <summary>The control causes validation. (Default)</summary>
        Yes = -1
    }

    /// <summary>
    /// Determines whether the form can be moved by the user.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa235194(v=vs.60)
    /// </summary>
    public enum Movability // Corresponds to Moveable property of a Form
    {
        /// <summary>The form is not moveable.</summary>
        Fixed = 0,
        /// <summary>The form is moveable. (Default)</summary>
        Moveable = -1
    }

    /// <summary>
    /// Determines whether background text and graphics are displayed in spaces around characters.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267490(v=vs.60)
    /// </summary>
    public enum FontTransparency // Corresponds to FontTransparent property
    {
        /// <summary>Masks existing background graphics and text around characters.</summary>
        Opaque = 0,
        /// <summary>Permits background graphics and text to show. (Default)</summary>
        Transparent = -1
    }

    /// <summary>
    /// Determines whether context-sensitive Help uses What's This pop-up or main Help window.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445772(v=vs.60)
    /// </summary>
    public enum WhatsThisHelp
    {
        /// <summary>Application uses F1 key for Windows Help. (Default)</summary>
        F1Help = 0,
        /// <summary>Application uses "What's This?" access techniques.</summary>
        WhatsThisHelpEnabled = -1 // Renamed from "WhatsThisHelp" to avoid conflict with enum name
    }

    /// <summary>
    /// Type of link for DDE conversation.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa235154(v=vs.60)
    /// </summary>
    public enum FormLinkMode // Corresponds to LinkMode property of a Form
    {
        /// <summary>No DDE interaction. (Default)</summary>
        None = 0,
        /// <summary>Allows controls to supply data to DDE destination. Notifies on change.</summary>
        Source = 1
    }

    /// <summary>
    /// Controls the display state of a form (Normal, Minimized, Maximized).
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445778(v=vs.60)
    /// </summary>
    public enum WindowState
    {
        /// <summary>Normal state. (Default)</summary>
        Normal = 0,
        /// <summary>Minimized to an icon.</summary>
        Minimized = 1,
        /// <summary>Maximized to full screen.</summary>
        Maximized = 2
    }

    // Note: StartUpPosition in Rust is an enum with associated data for 'Manual'.
    // In C#, this is better represented by a class hierarchy or separate properties.
    // For now, a simple enum for the mode, and separate properties on the FormProperties class for manual coordinates.
    /// <summary>
    /// Initial position of the form when it first appears.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445708(v=vs.60)
    /// </summary>
    public enum StartUpPositionMode
    {
        /// <summary>Positioned based on ClientLeft, ClientTop, ClientWidth, ClientHeight. (VB Value: 0)</summary>
        Manual = 0,
        /// <summary>Centered in the parent Form or MDIForm. (VB Value: 1)</summary>
        CenterOwner = 1,
        /// <summary>Centered on the screen. (VB Value: 2)</summary>
        CenterScreen = 2,
        /// <summary>Position in upper-left corner of screen. (Default for VB, VB Value: 3)</summary>
        WindowsDefault = 3
    }

    /// <summary>
    /// Determines how a control is aligned on a form or container.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267259(v=vs.60)
    /// </summary>
    public enum Align
    {
        /// <summary>Not docked. (Default on non-MDI form)</summary>
        None = 0,
        /// <summary>Aligns to top. (Default on MDI form)</summary>
        Top = 1,
        /// <summary>Aligns to bottom.</summary>
        Bottom = 2,
        /// <summary>Aligns to left.</summary>
        Left = 3,
        /// <summary>Aligns to right.</summary>
        Right = 4
    }

    /// <summary>
    /// Alignment for CheckBox and OptionButton controls (text vs. control graphic).
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267261(v=vs.60)
    /// </summary>
    public enum JustifyAlignment // For CheckBox, OptionButton.Alignment
    {
        /// <summary>Text is left-aligned, control is right-aligned. (Default)</summary>
        LeftJustify = 0,
        /// <summary>Text is right-aligned, control is left-aligned.</summary>
        RightJustify = 1
    }

    /// <summary>
    /// Text alignment within Label and TextBox controls.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267261(v=vs.60)
    /// </summary>
    public enum Alignment // For Label, TextBox.Alignment
    {
        /// <summary>Text is left-aligned. (Default)</summary>
        LeftJustify = 0,
        /// <summary>Text is right-aligned.</summary>
        RightJustify = 1,
        /// <summary>Text is centered.</summary>
        Center = 2
    }

    /// <summary>
    /// Background style (transparent or opaque) for Label or Shape.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245038(v=vs.60)
    /// </summary>
    public enum BackStyle
    {
        /// <summary>Background is transparent.</summary>
        Transparent = 0,
        /// <summary>BackColor fills the control. (Default)</summary>
        Opaque = 1
    }

    /// <summary>
    /// Appearance of a control (flat or 3D).
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa244932(v=vs.60)
    /// </summary>
    public enum Appearance
    {
        /// <summary>Control is painted with a flat style.</summary>
        Flat = 0,
        /// <summary>Control is painted with a 3D style. (Default)</summary>
        ThreeD = 1
    }

    /// <summary>
    /// Border style of a control.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245047(v=vs.60)
    /// </summary>
    public enum BorderStyle // General BorderStyle, not FormBorderStyle
    {
        /// <summary>No border. (Default for Image, Label)</summary>
        None = 0,
        /// <summary>Single-line border. (Default for PictureBox, TextBox, OLE)</summary>
        FixedSingle = 1
    }

    /// <summary>
    /// Drag and drop mode (manual or automatic).
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267292(v=vs.60) (DragMode)
    /// </summary>
    public enum DragMode
    {
        /// <summary>Manual drag initiation. (Default)</summary>
        Manual = 0,
        /// <summary>Automatic drag initiation on mouse press.</summary>
        Automatic = 1
    }

    /// <summary>
    /// How the pen interacts with the background during drawing.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267294(v=vs.60) (DrawMode)
    /// </summary>
    public enum DrawMode
    {
        Blackness = 1,
        NotMergePen = 2,
        MaskNotPen = 3,
        NotCopyPen = 4,
        MaskPenNot = 5,
        Invert = 6,
        XorPen = 7,
        NotMaskPen = 8,
        MaskPen = 9,
        NotXorPen = 10,
        Nop = 11,
        MergeNotPen = 12,
        /// <summary>ForeColor applied over background. (Default)</summary>
        CopyPen = 13,
        MergePenNot = 14,
        MergePen = 15,
        Whiteness = 16
    }

    /// <summary>
    /// Line style for graphic methods.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267296(v=vs.60) (DrawStyle)
    /// </summary>
    public enum DrawStyle
    {
        /// <summary>Solid line. (Default)</summary>
        Solid = 0,
        Dash = 1,
        Dot = 2,
        DashDot = 3,
        DashDotDot = 4,
        /// <summary>Transparent line/border.</summary>
        Transparent = 5, // VB: vbInvisible, but value is 5 for DrawStyle
        InsideSolid = 6
    }

    /// <summary>
    /// Appearance of the mouse pointer over the control.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267508(v=vs.60) (MousePointer)
    /// </summary>
    public enum MousePointer
    {
        /// <summary>Standard pointer. (Default)</summary>
        Default = 0,
        Arrow = 1,
        Cross = 2,
        IBeam = 3,
        Icon = 4,       // Uses MouseIcon property
        Size = 5,       // VB: vbSizeAll
        SizeNESW = 6,
        SizeNS = 7,
        SizeNWSE = 8,
        SizeWE = 9,
        UpArrow = 10,
        Hourglass = 11,
        NoDrop = 12,
        ArrowHourglass = 13,
        ArrowQuestion = 14,
        SizeAll = 15,   // Same as vbSizeAll (5)
        Custom = 99     // Uses MouseIcon property (same as Icon (4))
    }

    /// <summary>
    /// OLE drag and drop mode.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242173(v=vs.60) (OLEDragMode)
    /// </summary>
    public enum OLEDragMode
    {
        /// <summary>Manual OLE drag/drop handling. (Default)</summary>
        Manual = 0,
        /// <summary>Automatic OLE drag/drop handling.</summary>
        Automatic = 1
    }

    /// <summary>
    /// OLE drop mode.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242176(v=vs.60) (OLEDropMode)
    /// </summary>
    public enum OLEDropMode
    {
        /// <summary>Does not accept OLE drops. (Default)</summary>
        None = 0,
        /// <summary>Manual OLE drop handling.</summary>
        Manual = 1
        // VB also has a 'vbDropEffectCopyOrMove = 2' implicitly handled by events
    }

    /// <summary>
    /// Determines if controls are clipped to the bounds of the parent.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245086(v=vs.60) (ClipControls)
    /// </summary>
    public enum ClipControls
    {
        /// <summary>Controls are not clipped.</summary>
        Unbounded = 0, // VB False
        /// <summary>Controls are clipped. (Default)</summary>
        Clipped = 1    // VB True
    }

    /// <summary>
    /// Control style (standard or graphical from picture properties).
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266603(v=vs.60) (Style for CommandButton, CheckBox, OptionButton)
    /// </summary>
    public enum Style // For CommandButton, CheckBox, OptionButton
    {
        /// <summary>Standard styling. (Default)</summary>
        Standard = 0,
        /// <summary>Graphical styling from picture properties.</summary>
        Graphical = 1
    }

    /// <summary>
    /// Fill style for drawing (Form, PictureBox).
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267487(v=vs.60) (FillStyle)
    /// </summary>
    public enum FillStyle
    {
        Solid = 0,
        /// <summary>Transparent background. (Default)</summary>
        Transparent = 1,
        HorizontalLine = 2,
        VerticalLine = 3,
        UpwardDiagonal = 4,
        DownwardDiagonal = 5,
        Cross = 6,
        DiagonalCross = 7
    }

    /// <summary>
    /// Type of link for DDE or OLE. (General LinkMode, not specific FormLinkMode)
    /// Reference: (e.g. Label.LinkMode) https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa235154(v=vs.60)
    /// </summary>
    public enum LinkMode
    {
        /// <summary>No DDE interaction. (Default)</summary>
        None = 0,
        /// <summary>Automatic update for DDE link.</summary>
        Automatic = 1, // DDE specific
        /// <summary>Manual update for DDE link.</summary>
        ManualLink = 2, // DDE specific, renamed from Manual to avoid clash
        /// <summary>Notify for DDE link.</summary>
        Notify = 3      // DDE specific
    }

    /// <summary>
    /// Multi-selection behavior for controls like ListBox.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267511(v=vs.60) (MultiSelect)
    /// </summary>
    public enum MultiSelect
    {
        /// <summary>No multiselection. (Default)</summary>
        None = 0,
        /// <summary>Simple multiselection (click toggles, no Shift/Ctrl).</summary>
        Simple = 1,
        /// <summary>Extended multiselection (Shift/Ctrl keys).</summary>
        Extended = 2
    }

    /// <summary>
    /// Unit of measure for coordinates and sizes.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267542(v=vs.60) (ScaleMode)
    /// </summary>
    public enum ScaleMode
    {
        /// <summary>User-defined. ScaleHeight, ScaleWidth, ScaleLeft, ScaleTop set explicitly.</summary>
        User = 0,
        /// <summary>Twip (1/1440 inch). (Default)</summary>
        Twip = 1,
        /// <summary>Point (1/72 inch).</summary>
        Point = 2,
        /// <summary>Pixel.</summary>
        Pixel = 3,
        /// <summary>Character (horizontal = 120 twips/char; vertical = 240 twips/char).</summary>
        Character = 4,
        /// <summary>Inch.</summary>
        Inch = 5, // Corrected from Inches
        /// <summary>Millimeter.</summary>
        Millimeter = 6,
        /// <summary>Centimeter.</summary>
        Centimeter = 7
    }

    /// <summary>
    /// How an image is displayed in a control (e.g., PictureBox, Image).
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa442618(v=vs.60) (PictureBox.SizeMode, not Image.Stretch)
    /// </summary>
    public enum SizeMode // For PictureBox.AutoSize = False
    {
        /// <summary>Clip picture if larger than control. (Default)</summary>
        Clip = 0, // vbPictureBoxSizeModeClip
        /// <summary>Stretch picture to fit control.</summary>
        Stretch = 1, // vbPictureBoxSizeModeStretchImage
        /// <summary>Resize control to fit picture (actually handled by AutoSize=True for PictureBox).</summary>
        AutoSize = 2, // vbPictureBoxSizeModeAutoSize - Note: This means the *control* resizes. If AutoSize is False, this value might not be directly used.
        /// <summary>Zoom picture to fit control, maintain aspect ratio.</summary>
        Zoom = 3 // vbPictureBoxSizeModeZoom
    }
	
	/// <summary>
    /// Specifies the alignment of a CheckBox or OptionButton caption relative to its graphic.
    /// </summary>
    public enum JustifyAlignment
    {
        /// <summary>The graphic is to the left of the caption. (Default for CheckBox)</summary>
        LeftJustify = 0,
        /// <summary>The graphic is to the right of the caption.</summary>
        RightJustify = 1
    }
	
	/// <summary>
    /// Specifies the style of a line or border for graphical controls or methods.
    /// </summary>
    public enum DrawStyle
    {
        /// <summary>Solid line. (Default for Line/Shape BorderStyle, DrawStyle method)</summary>
        Solid = 0,
        /// <summary>Dashed line.</summary>
        Dash = 1,
        /// <summary>Dotted line.</summary>
        Dot = 2,
        /// <summary>Dash-dot sequence.</summary>
        DashDot = 3,
        /// <summary>Dash-dot-dot sequence.</summary>
        DashDotDot = 4,
        /// <summary>Transparent (no line visible).</summary>
        Transparent = 5, // vbTransparent for Shape.FillStyle/BorderStyle; Line control's BorderStyle can be this.
        /// <summary>Line is drawn inside the shape's border.</summary>
        InsideSolid = 6 // vbInsideSolid for Shape.BorderStyle, DrawStyle method
    }

    /// <summary>
    /// Specifies how a line or shape is drawn in relation to existing background colors.
    /// </summary>
    public enum DrawMode
    {
        /// <summary>Blackness (all 0s).</summary>
        Blackness = 1, // vbBlackness
        /// <summary>Not Merge Pen (inverse of MergePen).</summary>
        NotMergePen = 2, // vbNotMergePen
        /// <summary>Mask Not Pen (combination of display and pen colors, inverse of MaskPenNot).</summary>
        MaskNotPen = 3, // vbMaskNotPen
        /// <summary>Not Copy Pen (inverse of CopyPen).</summary>
        NotCopyPen = 4, // vbNotCopyPen
        /// <summary>Mask Pen Not (combination of pen and display colors, inverse of MaskNotPen).</summary>
        MaskPenNot = 5, // vbMaskPenNot
        /// <summary>Invert (inverse of display color).</summary>
        Invert = 6, // vbInvert
        /// <summary>XOR Pen (XOR of pen and display colors).</summary>
        XorPen = 7, // vbXorPen
        /// <summary>Not Mask Pen (inverse of MaskPen).</summary>
        NotMaskPen = 8, // vbNotMaskPen
        /// <summary>Mask Pen (AND of pen and display colors).</summary>
        MaskPen = 9, // vbMaskPen
        /// <summary>Not Xor Pen (inverse of XorPen).</summary>
        NotXorPen = 10, // vbNotXorPen
        /// <summary>No Operation (display remains unchanged).</summary>
        Nop = 11, // vbNop
        /// <summary>Merge Not Pen (combination of display color and inverse of pen color).</summary>
        MergeNotPen = 12, // vbMergeNotPen
        /// <summary>Copy Pen (color specified by ForeColor). (Default)</summary>
        CopyPen = 13, // vbCopyPen
        /// <summary>Merge Pen Not (combination of pen color and inverse of display color).</summary>
        MergePenNot = 14, // vbMergePenNot
        /// <summary>Merge Pen (OR of pen and display colors).</summary>
        MergePen = 15, // vbMergePen
        /// <summary>Whiteness (all 1s).</summary>
        Whiteness = 16 // vbWhiteness
    }
	
	/// <summary>
    /// Specifies whether the background of a control is transparent or opaque.
    /// </summary>
    public enum BackStyle
    {
        /// <summary>Opaque: Background is filled with BackColor. (Default for some controls)</summary>
        Opaque = 1, // vbOpaque
        /// <summary>Transparent: Background of container shows through. (Default for Label, Shape)</summary>
        Transparent = 0 // vbTransparent
    }

    /// <summary>
    /// Specifies the pattern used to fill a Shape control or graphical output.
    /// </summary>
    public enum FillStyleType // Named to distinguish from BorderStyle/DrawStyle
    {
        /// <summary>Solid fill with FillColor. (Default for Shape.FillStyle is Transparent)</summary>
        Solid = 0, // vbFSSolid
        /// <summary>Transparent fill (BackColor of container shows through). (Default for Shape)</summary>
        Transparent = 1, // vbFSTransparent
        /// <summary>Horizontal lines.</summary>
        HorizontalLine = 2, // vbFSHorizontalLine
        /// <summary>Vertical lines.</summary>
        VerticalLine = 3, // vbFSVerticalLine
        /// <summary>Upward diagonal lines (bottom-left to top-right).</summary>
        UpwardDiagonal = 4, // vbFSUpwardDiagonal
        /// <summary>Downward diagonal lines (top-left to bottom-right).</summary>
        DownwardDiagonal = 5, // vbFSDownwardDiagonal
        /// <summary>Crosshatch (horizontal and vertical lines).</summary>
        Cross = 6, // vbFSCross
        /// <summary>Diagonal crosshatch.</summary>
        DiagonalCross = 7 // vbFSDiagonalCross
    }
	
	/// <summary>
    /// Determines whether graphics are clipped to the boundaries of a container control.
    /// True (Clipped) is generally more efficient.
    /// </summary>
    public enum ClipControls
    {
        /// <summary>Graphics are clipped to the container. (Corresponds to VB6 True - Default)</summary>
        Clipped = 0, // True in VB6
        /// <summary>Graphics are not clipped. (Corresponds to VB6 False)</summary>
        Unclipped = 1  // False in VB6
    }
	
	/// <summary>
    /// Specifies whether graphics output to a control is persistent.
    /// </summary>
    public enum AutoRedraw
    {
        /// <summary>Graphics are not persistent; repainting is manual or event-driven. (Corresponds to VB6 False/0)</summary>
        Manual = 0, // Default for PictureBox in Rust is Manual (False)
        /// <summary>Graphics are persistent; the control automatically repaints itself. (Corresponds to VB6 True/-1)</summary>
        Automatic = 1
    }

    /// <summary>
    /// Specifies whether the background of text printed on a control is transparent or opaque.
    /// </summary>
    public enum FontTransparency // VB6 FontTransparent property
    {
        /// <summary>Text background is opaque, filled with BackColor. (Corresponds to VB6 False/0)</summary>
        Opaque = 0,
        /// <summary>Text background is transparent. (Corresponds to VB6 True/-1). Default in Rust.</summary>
        Transparent = 1
    }

    /// <summary>
    /// Specifies if a control has a persistent Device Context (DC).
    /// For PictureBox, this is always Yes (True).
    /// </summary>
    public enum HasDeviceContext
    {
        /// <summary>Control does not have its own persistent DC. (Corresponds to VB6 False/0)</summary>
        No = 0,
        /// <summary>Control has its own persistent DC. (Corresponds to VB6 True/-1). Default in Rust.</summary>
        Yes = 1
    }

    /// <summary>
    /// Specifies the type of DDE (Dynamic Data Exchange) link for a control.
    /// </summary>
    public enum LinkMode
    {
        /// <summary>No DDE interaction. (Default)</summary>
        None = 0, // vbLinkNone
        /// <summary>Automatic DDE link: control is updated when linked data changes.</summary>
        Automatic = 1, // vbLinkAutomatic (for destination controls)
        /// <summary>Manual DDE link: control is updated only when LinkRequest method is called.</summary>
        Manual = 2, // vbLinkManual (for destination controls)
        /// <summary>Notify DDE link: control notifies application when linked data changes, but doesn't update automatically.</summary>
        Notify = 3 // vbLinkNotify (for destination controls)
        // Note: Source controls use LinkMode = 1 (vbLinkSource), but this enum is often used for destination link types.
        // The Rust LinkMode enum seems to align with destination modes.
    }
	
	/// <summary>
    /// Specifies the DDE (Dynamic Data Exchange) link mode for a Form or MDIForm,
    /// particularly when it acts as a DDE source.
    /// </summary>
    public enum FormLinkMode
    {
        /// <summary>No DDE interaction. (Default for MDIForm LinkMode)</summary>
        None = 0, // vbLinkNone
        /// <summary>The form can be a source for a DDE link.</summary>
        Source = 1 // vbLinkSource
    }

    /// <summary>
    /// Specifies whether a Form or MDIForm can be moved by the user at runtime.
    /// </summary>
    public enum Movability // VB6 "Moveable" property
    {
        /// <summary>The form cannot be moved by the user. (Corresponds to VB6 False/0)</summary>
        NotMoveable = 0,
        /// <summary>The form can be moved by the user. (Corresponds to VB6 True/-1). Default for MDIForm.</summary>
        Moveable = 1
    }

    /// <summary>
    /// Determines whether context-sensitive Help uses What's This pop-up or the main Help window.
    /// Corresponds to VB6 WhatsThisHelp property on a Form.
    /// </summary>
    public enum WhatsThisHelpMode
    {
        /// <summary>(Default) Application uses F1 key for standard Windows Help.</summary>
        F1Help = 0,
        /// <summary>Application uses "What's This?" pop-up help.</summary>
        WhatsThis = -1 // VB6 True is -1
    }
	
	/// <summary>
    /// Determines whether background text and graphics on a Form or PictureBox
    /// are displayed in the spaces around characters.
    /// Corresponds to VB6 FontTransparent property.
    /// </summary>
    public enum FontTransparencyMode
    {
        /// <summary>Masks existing background graphics and text around the characters (Default for some contexts).</summary>
        Opaque = 0, // vbFontOpaque
        /// <summary>Permits background graphics and text to show through spaces around characters (Default for some contexts).</summary>
        Transparent = -1 // vbFontTransparent (VB6 True is -1)
        // Note: Default depends on control. Rust default is Transparent (-1).
    }

    /// <summary>
    /// Determines the type of DDE link used for a Form and activates the connection.
    /// Corresponds to VB6 LinkMode property on a Form.
    /// </summary>
    public enum FormLinkModeType
    {
        /// <summary>(Default) No DDE interaction.</summary>
        None = 0,   // vbLinkNone
        /// <summary>Allows controls on the form to supply data via DDE (Form is DDE source).</summary>
        Source = 1  // vbLinkSource
        // Note: Controls have a more extensive LinkMode (Automatic, Manual, Notify for DDE client links)
    }

    /// <summary>
    /// Determines whether a Form can be moved by the user at runtime.
    /// Corresponds to VB6 Movable property.
    /// </summary>
    public enum MovabilityType
    {
        /// <summary>The form cannot be moved by the user.</summary>
        Fixed = 0,
        /// <summary>(Default) The form can be moved by the user.</summary>
        Movable = -1 // VB6 True (Movable) is -1
    }
}