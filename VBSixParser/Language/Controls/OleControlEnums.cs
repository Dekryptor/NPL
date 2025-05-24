// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Specifies the type of OLE object the OLE container control can contain.
    /// Corresponds to VB6 OLETypeAllowed property.
    /// </summary>
    public enum OLETypeAllowedEnum
    {
        /// <summary>Can contain only a linked object.</summary>
        Link = 0,     // vbOLELinked
        /// <summary>Can contain only an embedded object.</summary>
        Embedded = 1, // vbOLEEmbedded
        /// <summary>(Default) Can contain either a linked or an embedded object.</summary>
        Either = 2    // vbOLEEither
    }

    /// <summary>
    /// Specifies how a linked OLE object is updated when its source data changes.
    /// Corresponds to VB6 UpdateOptions property.
    /// </summary>
    public enum OLEUpdateOptionsEnum
    {
        /// <summary>(Default) Object is updated each time linked data changes.</summary>
        Automatic = 0, // vbOLEUpdateAutomatic
        /// <summary>Object is updated when the user saves the linked data from its source application.</summary>
        Frozen = 1,    // vbOLEUpdateFrozen (VB6 docs say this is for when source app saves)
                       // Rust says "whenever the user saves the linked data from within the application in which it was created"
        /// <summary>Object is updated only by using the Update method.</summary>
        Manual = 2     // vbOLEUpdateManual
    }

    /// <summary>
    /// Specifies how an OLE object is activated.
    /// Corresponds to VB6 AutoActivate property.
    /// </summary>
    public enum OLEAutoActivateEnum
    {
        /// <summary>Object is not automatically activated. Use DoVerb method.</summary>
        Manual = 0,        // vbOLEActivateManual
        /// <summary>Object is activated when the OLE control receives focus (if object supports single-click activation).</summary>
        Focus = 1,         // vbOLEActivateGetFocus
        /// <summary>(Default) Object is activated on double-click or when Enter is pressed on the control.</summary>
        DoubleClick = 2,   // vbOLEActivateDoubleClick
        /// <summary>Object is activated based on its normal activation method (focus or double-click).</summary>
        Automatic = 3      // vbOLEActivateAuto (Not a standard VB6 constant, but implies object's default)
                           // VB6 documentation lists Manual, GetFocus, DoubleClick.
                           // Rust's Automatic(3) might be a custom value or from a specific OLE interface.
                           // MSDN for AutoActivate only shows 0, 1, 2. Let's stick to those for C# primary mapping.
                           // For now, including all Rust values.
    }

    /// <summary>
    /// Specifies whether an OLE object displays its content or an icon.
    /// Corresponds to VB6 DisplayType property.
    /// </summary>
    public enum OLEDisplayTypeEnum
    {
        /// <summary>(Default) Object's data is displayed.</summary>
        Content = 0, // vbOLEDisplayContent
        /// <summary>Object's icon is displayed.</summary>
        Icon = 1     // vbOLEDisplayIcon
    }
}