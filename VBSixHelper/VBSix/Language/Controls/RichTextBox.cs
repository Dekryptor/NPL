using System.Text;
using VBSix.Parsers;
using VBSix.Utilities;

namespace VBSix.Language.Controls
{
    // Assuming ScrollBars, Appearance, BorderStyle, etc. enums are defined in VB6ControlEnums.cs

    /// <summary>
    /// Represents the properties specific to a VB6 RichTextBox control.
    /// (Typically from RICHTX32.OCX)
    /// </summary>
    public class RichTextBoxProperties
    {
        public int Height { get; set; } = 2295; // Default Height in Twips for a new RichTextBox
        public int Width { get; set; } = 4335;  // Default Width in Twips
        public int Left { get; set; } = 240;
        public int Top { get; set; } = 240;
        public Appearance Appearance { get; set; } = Controls.Appearance.Flat; // RichTextBox defaults to Flat
        public BorderStyle BorderStyle { get; set; } = Controls.BorderStyle.FixedSingle;
        public ScrollBars ScrollBars { get; set; } = Controls.ScrollBars.Both; // Common default

        /// <summary>
        /// Gets or sets whether selected text remains highlighted when the control loses focus.
        /// VB6 Property: HideSelection. Default is True.
        /// </summary>
        public bool HideSelection { get; set; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether the user can change the control's contents.
        /// VB6 Property: Locked. Default is False.
        /// </summary>
        public bool Locked { get; set; } = false;

        /// <summary>
        /// Gets or sets the maximum number of characters the control can hold.
        /// 0 means limited by system resources (approx 64KB for older RTB, more for newer).
        /// VB6 Property: MaxLength.
        /// </summary>
        public int MaxLength { get; set; } = 0;

        /// <summary>
        /// Gets or sets a value indicating whether the control can display multiple lines of text.
        /// For RichTextBox, this is effectively always True, but the property exists.
        /// VB6 Property: MultiLine.
        /// </summary>
        public MultiLine MultiLine { get; set; } = Controls.MultiLine.MultiLine; // RichTextBox is inherently multiline (True/-1)

        /// <summary>
        /// Gets or sets the plain text content of the control.
        /// If design-time content is set via RTF, this property might reflect the plain text equivalent.
        /// VB6 Property: Text.
        /// </summary>
        public string Text { get; set; } = "";

        /// <summary>
        /// Stores the raw string value of the TextRTF property as found in the .frm file.
        /// This might be an FRX link (e.g., "MyForm.frx:00A0" or "$MyForm.frx:00A0")
        /// or, less commonly for design-time, inline RTF.
        /// </summary>
        public string TextRTF_SourceString { get; set; } = "";

        // --- Standard Control Properties ---
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        public TabStop TabStop { get; set; } = Controls.TabStop.Included;
        public int TabIndex { get; set; } = 0;
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public string ToolTipText { get; set; } = "";
        public int HelpContextID { get; set; } = 0;
        public int WhatsThisHelpID { get; set; } = 0; // RichTextBox doesn't directly support WhatsThisHelp button
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.IBeam; // Default for text entry areas

        public VB6Color BackColor { get; set; } = VB6Color.VbWindowBackground; // Default: &H80000005&
        public VB6Color ForeColor { get; set; } = VB6Color.VbWindowText;     // Default: &H80000008&

        // --- OLE Drag/Drop Properties ---
        public bool AllowDrop { get; set; } = false; // VB6 Property: OLEDropAllowed (for general files)
        public OLEDragMode OLEDragMode { get; set; } = Controls.OLEDragMode.Manual;
        public OLEDropMode OLEDropMode { get; set; } = Controls.OLEDropMode.None;

        // --- RichTextBox Specific Properties (Examples from snippet, others may exist) ---
        public int BulletIndent { get; set; } = 0;
        public int SelAlignment { get; set; } = 0; // Could be an enum: 0=Left, 1=Right, 2=Center (vbLeftJustify, vbRightJustify, vbCenterJustify)
                                                   // Let's use a simple int for now, mapping can be done by consumer.
        public bool AutoVerbMenu { get; set; } = true; // Common for OLE-capable controls

        // --- ActiveX Internal Properties ---
        public int _ExtentX { get; set; } = 7000; // Approx default in Twips, from snippet
        public int _ExtentY { get; set; } = 2000; // Approx default in Twips, from snippet
        public int _Version { get; set; } = 393217; // Version of the control/persistence format

        public RichTextBoxProperties() { }

        public RichTextBoxProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Height = rawProps.GetInt32("Height", Height);
            Width = rawProps.GetInt32("Width", Width);
            Left = rawProps.GetInt32("Left", Left);
            Top = rawProps.GetInt32("Top", Top);
            Appearance = rawProps.GetEnum("Appearance", Appearance);
            BorderStyle = rawProps.GetEnum("BorderStyle", BorderStyle);
            ScrollBars = rawProps.GetEnum("ScrollBars", ScrollBars);
            HideSelection = rawProps.GetBoolean("HideSelection", HideSelection);
            Locked = rawProps.GetBoolean("Locked", Locked);
            MaxLength = rawProps.GetInt32("MaxLength", MaxLength); // For plain text limit
            MultiLine = rawProps.GetEnum("MultiLine", MultiLine);

            // Text property (plain text)
            if (rawProps.TryGetValue("Text", out var textProp) && !textProp.IsResource)
            {
                Text = textProp.AsString() ?? "";
            }

            // TextRTF property (usually an FRX link or inline RTF string)
            if (rawProps.TryGetValue("TextRTF", out var rtfProp) && !rtfProp.IsResource)
            {
                TextRTF_SourceString = rtfProp.AsString() ?? "";
            }
            else if (rtfProp != null && rtfProp.IsResource)
            {
                // This case implies TextRTF pointed to an FRX resource directly in the main properties,
                // which is unusual. Normally it's a string link.
                // The VB6RichTextBoxKind constructor handles resolving the FRX link string.
                // If TextRTF was directly binary here, it's an edge case.
                // For now, we assume TextRTF in rawProps is a string link.
                // Console.Error.WriteLine("Warning: TextRTF property found as direct resource in rawProps, expected string link.");
            }


            Enabled = rawProps.GetEnum("Enabled", Enabled);
            TabStop = rawProps.GetEnum("TabStop", TabStop);
            TabIndex = rawProps.GetInt32("TabIndex", TabIndex);
            Visible = rawProps.GetEnum("Visible", Visible);
            ToolTipText = rawProps.GetString("ToolTipText", ToolTipText);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID); // Though not visually active
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            BackColor = rawProps.GetVB6Color("BackColor", BackColor);
            ForeColor = rawProps.GetVB6Color("ForeColor", ForeColor);

            AllowDrop = rawProps.GetBoolean("OLEDropAllowed", AllowDrop); // Map OLEDropAllowed from VBP/ActiveX to AllowDrop
            OLEDragMode = rawProps.GetEnum("OLEDragMode", OLEDragMode);
            OLEDropMode = rawProps.GetEnum("OLEDropMode", OLEDropMode);

            BulletIndent = rawProps.GetInt32("BulletIndent", BulletIndent);
            SelAlignment = rawProps.GetInt32("SelAlignment", SelAlignment);
            AutoVerbMenu = rawProps.GetBoolean("AutoVerbMenu", AutoVerbMenu);

            _ExtentX = rawProps.GetInt32("_ExtentX", _ExtentX);
            _ExtentY = rawProps.GetInt32("_ExtentY", _ExtentY);
            _Version = rawProps.GetInt32("_Version", _Version);

            // Font is handled as a PropertyGroup by the main parser
        }
    }
}
