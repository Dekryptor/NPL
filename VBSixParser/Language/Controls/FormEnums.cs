// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// The palette drawing mode of a form (for 256-color displays).
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa733659(v=vs.60)
    /// </summary>
    public enum PaletteMode
    {
        /// <summary>System halftone palette. (Default)</summary>
        HalfTone = 0,
        /// <summary>Palette of the topmost control has precedence.</summary>
        UseZOrder = 1,
        /// <summary>Uses a custom palette from an image assigned to the Palette property.</summary>
        Custom = 2
    }

    /// <summary>
    /// The appearance of a form's border.
    /// </summary>
    public enum FormBorderStyle // Distinct from the general BorderStyle for controls
    {
        /// <summary>No border or title bar.</summary>
        None = 0,
        /// <summary>Fixed, single-line border.</summary>
        FixedSingle = 1,
        /// <summary>Sizable border. (Default)</summary>
        Sizable = 2,
        /// <summary>Fixed dialog box border.</summary>
        FixedDialog = 3,
        /// <summary>Fixed tool window border.</summary>
        FixedToolWindow = 4,
        /// <summary>Sizable tool window border.</summary>
        SizableToolWindow = 5
    }

    /// <summary>
    /// Determines if the control box is displayed in the form's title bar.
    /// </summary>
    public enum ControlBox
    {
        /// <summary>Control box is not displayed.</summary>
        Excluded = 0, // VB False
        /// <summary>Control box is displayed. (Default)</summary>
        Included = -1 // VB True
    }

    /// <summary>
    /// Determines if the Maximize button is displayed in the form's title bar.
    /// </summary>
    public enum MaxButton
    {
        /// <summary>Maximize button is not displayed.</summary>
        Excluded = 0, // VB False
        /// <summary>Maximize button is displayed. (Default)</summary>
        Included = -1 // VB True
    }

    /// <summary>
    /// Determines if the Minimize button is displayed in the form's title bar.
    /// </summary>
    public enum MinButton
    {
        /// <summary>Minimize button is not displayed.</summary>
        Excluded = 0, // VB False
        /// <summary>Minimize button is displayed. (Default)</summary>
        Included = -1 // VB True
    }

    /// <summary>
    /// Determines if the 'What's This?' button is displayed in the form's title bar.
    /// </summary>
    public enum WhatsThisButton
    {
        /// <summary>'What's This?' button is not displayed. (Default)</summary>
        Excluded = 0, // VB False
        /// <summary>'What's This?' button is displayed.</summary>
        Included = -1 // VB True
    }

    /// <summary>
    /// Determines if the form is shown in the Windows taskbar.
    /// </summary>
    public enum ShowInTaskbar
    {
        /// <summary>Form is not shown in the taskbar.</summary>
        Hide = 0, // VB False
        /// <summary>Form is shown in the taskbar. (Default)</summary>
        Show = -1  // VB True
    }

    /// <summary>
    /// Specifies the DDE link mode for a Form.
    /// (This was mentioned in the imports of form.rs as `FormLinkMode`)
    /// </summary>
    public enum FormLinkMode // From VB6.hlp: LinkMode Property (Form)
    {
        /// <summary>(Default) No DDE interaction.</summary>
        None = 0,
        /// <summary>Automatic DDE link (destination updated whenever linked data changes).</summary>
        Automatic = 1, // Also DDE LinkAutomatic
        /// <summary>Manual DDE link (destination requests updates).</summary>
        Manual = 2,    // Also DDE LinkManual
        /// <summary>DDE link for which destination notifies source when data changes (rarely used for Forms).</summary>
        Notify = 3     // Also DDE LinkNotify
    }
}