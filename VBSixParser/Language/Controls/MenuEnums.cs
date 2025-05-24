// Namespace: VB6Parse.Language.Controls
using System.ComponentModel; // For DefaultValue attribute if needed, or just comments

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Determines whether top-level Menu controls are displayed on the menu bar
    /// while a linked or embedded object is active.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa278135(v=vs.60)
    /// </summary>
    public enum NegotiatePosition
    {
        /// <summary>The menu is not displayed on the menu bar. (Default)</summary>
        None = 0,
        /// <summary>The menu is displayed at the left end of the menu bar.</summary>
        Left = 1,
        /// <summary>The menu is displayed in the middle of the menu bar.</summary>
        Middle = 2,
        /// <summary>The menu is displayed at the right end of the menu bar.</summary>
        Right = 3
    }

    /// <summary>
    /// Represents a keyboard shortcut for a menu item.
    /// Note: F10, Ctrl+F10, Shift+F10, Ctrl+Shift+F10 are not valid.
    /// Access keys (& in Caption) are separate.
    /// </summary>
    public enum ShortCut
    {
        // Values are typically 0 for None/Default, then sequential.
        // VB6 stores these as specific integer values in .frm files (0 for none, then 1-N for specific keys)
        // The Rust enum maps to string representations like "^A", "{F1}".
        // For C#, a simple enum is fine. The mapping from integer/string in the .frm file
        // would happen during parsing. For now, we list them.
        // We can add attributes or a helper class later if direct mapping to VB6 int values is needed.

        None = 0, // Assuming 0 is no shortcut, which is typical.

        CtrlA, CtrlB, CtrlC, CtrlD, CtrlE, CtrlF, CtrlG, CtrlH, CtrlI, CtrlJ,
        CtrlK, CtrlL, CtrlM, CtrlN, CtrlO, CtrlP, CtrlQ, CtrlR, CtrlS, CtrlT,
        CtrlU, CtrlV, CtrlW, CtrlX, CtrlY, CtrlZ,

        F1, F2, F3, F4, F5, F6, F7, F8, F9, /* F10 is invalid */ F11, F12,

        CtrlF1, CtrlF2, CtrlF3, CtrlF4, CtrlF5, CtrlF6, CtrlF7, CtrlF8, CtrlF9, /* CtrlF10 invalid */ CtrlF11, CtrlF12,

        ShiftF1, ShiftF2, ShiftF3, ShiftF4, ShiftF5, ShiftF6, ShiftF7, ShiftF8, ShiftF9, /* ShiftF10 invalid */ ShiftF11, ShiftF12,

        ShiftCtrlF1, ShiftCtrlF2, ShiftCtrlF3, ShiftCtrlF4, ShiftCtrlF5, ShiftCtrlF6,
        ShiftCtrlF7, ShiftCtrlF8, ShiftCtrlF9, /* ShiftCtrlF10 invalid */ ShiftCtrlF11, ShiftCtrlF12,

        CtrlIns,
        ShiftIns,
        Del, // VB: vbKeyDelete
        ShiftDel,
        AltBksp // VB: vbKeyAltBack
    }
}