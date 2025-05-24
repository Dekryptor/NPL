namespace VB6Parse.Errors
{
    /// <summary>
    /// Specifies the type of error that occurred during VB6 property parsing.
    /// </summary>
    public enum PropertyErrorKind
    {
        /// <summary>Error: "Appearance can only be a 0 (Flat) or a 1 (ThreeD)"</summary>
        AppearanceInvalid,

        /// <summary>Error: "BorderStyle can only be a 0 (None) or 1 (FixedSingle)"</summary>
        BorderStyleInvalid,

        /// <summary>Error: "ClipControls can only be a 0 (false) or a 1 (true)"</summary>
        ClipControlsInvalid,

        /// <summary>Error: "DragMode can only be 0 (Manual) or 1 (Automatic)"</summary>
        DragModeInvalid,

        /// <summary>Error: "Enabled can only be 0 (false) or a 1 (true)"</summary>
        EnabledInvalid,

        /// <summary>Error: "MousePointer can only be 0 (Default), 1 (Arrow), 2 (Cross), 3 (IBeam), 6 (SizeNESW), 7 (SizeNS), 8 (SizeNWSE), 9 (SizeWE), 10 (UpArrow), 11 (Hourglass), 12 (NoDrop), 13 (ArrowHourglass), 14 (ArrowQuestion), 15 (SizeAll), or 99 (Custom)"</summary>
        MousePointerInvalid,

        /// <summary>Error: "OLEDropMode can only be 0 (None), or 1 (Manual)"</summary>
        OLEDropModeInvalid,

        /// <summary>Error: "RightToLeft can only be 0 (false) or a 1 (true)"</summary>
        RightToLeftInvalid,

        /// <summary>Error: "Visible can only be 0 (false) or a 1 (true)"</summary>
        VisibleInvalid,

        /// <summary>Error: "Unknown property in header file"</summary>
        UnknownProperty,

        /// <summary>Error: "Invalid property value. Only 0 or -1 are valid for this property"</summary>
        InvalidPropertyValueZeroNegOne,

        /// <summary>Error: "Unable to parse the property name"</summary>
        NameUnparsable,

        /// <summary>Error: "Unable to parse the resource file name"</summary>
        ResourceFileNameUnparseable,

        /// <summary>Error: "Unable to parse the offset into the resource file for property"</summary>
        OffsetUnparseable,

        /// <summary>Error: "Invalid property value. Only True or False are valid for this property"</summary>
        InvalidPropertyValueTrueFalse
    }
}