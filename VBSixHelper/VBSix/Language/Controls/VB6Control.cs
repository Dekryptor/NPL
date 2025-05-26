using System.Text;
using VBSix.Parsers;

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Abstract base class for all specific VB6 control kind definitions.
    /// This allows VB6Control.Kind to hold any type of control, providing a common type.
    /// </summary>
    public abstract class VB6ControlKind { }

    // --- Standard VB Control Kind Variants ---

    public class VB6FormKind : VB6ControlKind
    {
        public FormProperties TypedProperties { get; set; } = new FormProperties();
        public List<VB6Control> Controls { get; set; } = [];
        public List<VB6MenuControl> Menus { get; set; } = [];
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];

        // FRX Resolved Data
        public byte[]? ResolvedIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }
        public byte[]? ResolvedPictureData { get; set; }
        public byte[]? ResolvedPaletteData { get; set; }

        public VB6FormKind() { }

        // Constructor for use in BuildControl
        public VB6FormKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups, List<VB6Control> childControls, List<VB6MenuControl> childMenus)
        {
            TypedProperties = new FormProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            Controls = childControls ?? [];
            Menus = childMenus ?? [];

            // Resolve FRX directly from rawProps within this constructor
            if (rawProps.TryGetValue("Icon", out var iconProp) && iconProp.IsResource) ResolvedIconData = iconProp.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mIconProp) && mIconProp.IsResource) ResolvedMouseIconData = mIconProp.AsResource();
            if (rawProps.TryGetValue("Picture", out var picProp) && picProp.IsResource) ResolvedPictureData = picProp.AsResource();
            if (rawProps.TryGetValue("Palette", out var palProp) && palProp.IsResource) ResolvedPaletteData = palProp.AsResource();
        }
    }

    public class VB6MDIFormKind : VB6ControlKind
    {
        public MDIFormProperties TypedProperties { get; set; } = new MDIFormProperties();
        public List<VB6Control> Controls { get; set; } = []; // MDI Forms don't directly host controls in same way as Forms
        public List<VB6MenuControl> Menus { get; set; } = [];
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];

        public byte[]? ResolvedIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }
        public byte[]? ResolvedPictureData { get; set; }

        public VB6MDIFormKind() { }

        public VB6MDIFormKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups, List<VB6Control> childControls, List<VB6MenuControl> childMenus)
        {
            TypedProperties = new MDIFormProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            Controls = childControls ?? []; // MDI Children are Forms, not arbitrary controls
            Menus = childMenus ?? [];

            if (rawProps.TryGetValue("Icon", out var iProp) && iProp.IsResource) ResolvedIconData = iProp.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var miProp) && miProp.IsResource) ResolvedMouseIconData = miProp.AsResource();
            if (rawProps.TryGetValue("Picture", out var pProp) && pProp.IsResource) ResolvedPictureData = pProp.AsResource();
        }
    }

    public class VB6MenuKind : VB6ControlKind
    {
        public MenuProperties TypedProperties { get; set; } = new MenuProperties();
        public List<VB6MenuControl> SubMenus { get; set; } = [];
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = []; // For Font

        public VB6MenuKind() { }

        public VB6MenuKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups, List<VB6MenuControl> subMenus)
        {
            TypedProperties = new MenuProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            SubMenus = subMenus ?? [];
        }
    }

    public class VB6FrameKind : VB6ControlKind
    {
        public FrameProperties TypedProperties { get; set; } = new FrameProperties();
        public List<VB6Control> Controls { get; set; } = [];
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }

        public VB6FrameKind() { }

        public VB6FrameKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups, List<VB6Control> childControls)
        {
            TypedProperties = new FrameProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            Controls = childControls ?? [];
            if (rawProps.TryGetValue("DragIcon", out var diProp) && diProp.IsResource) ResolvedDragIconData = diProp.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var miProp) && miProp.IsResource) ResolvedMouseIconData = miProp.AsResource();
        }
    }

    public class VB6CheckBoxKind : VB6ControlKind
    {
        public CheckBoxProperties TypedProperties { get; set; } = new CheckBoxProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDisabledPictureData { get; set; }
        public byte[]? ResolvedDownPictureData { get; set; }
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }
        public byte[]? ResolvedPictureData { get; set; }

        public VB6CheckBoxKind() { }

        public VB6CheckBoxKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new CheckBoxProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DisabledPicture", out var dp) && dp.IsResource) ResolvedDisabledPictureData = dp.AsResource();
            if (rawProps.TryGetValue("DownPicture", out var downp) && downp.IsResource) ResolvedDownPictureData = downp.AsResource();
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
            if (rawProps.TryGetValue("Picture", out var p) && p.IsResource) ResolvedPictureData = p.AsResource();
        }
    }

    public class VB6ComboBoxKind : VB6ControlKind
    {
        public ComboBoxProperties TypedProperties { get; set; } = new ComboBoxProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public List<string> ResolvedListItems { get; set; } = [];
        public List<int> ResolvedItemData { get; set; } = [];
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }

        public VB6ComboBoxKind() { }

        public VB6ComboBoxKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new ComboBoxProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("List", out var l) && l.IsResource) { byte[]? d = l.AsResource(); if (d != null) ResolvedListItems = VB6FormFile.ListResolver(d); }
            if (rawProps.TryGetValue("ItemData", out var id) && id.IsResource) { byte[]? d = id.AsResource(); if (d != null) ResolvedItemData = VB6FormFile.ItemDataResolver(d); }
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
        }
    }

    public class VB6CommandButtonKind : VB6ControlKind
    {
        public CommandButtonProperties TypedProperties { get; set; } = new CommandButtonProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDisabledPictureData { get; set; }
        public byte[]? ResolvedDownPictureData { get; set; }
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }
        public byte[]? ResolvedPictureData { get; set; }

        public VB6CommandButtonKind() { }

        public VB6CommandButtonKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new CommandButtonProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DisabledPicture", out var dp) && dp.IsResource) ResolvedDisabledPictureData = dp.AsResource();
            if (rawProps.TryGetValue("DownPicture", out var downp) && downp.IsResource) ResolvedDownPictureData = downp.AsResource();
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
            if (rawProps.TryGetValue("Picture", out var p) && p.IsResource) ResolvedPictureData = p.AsResource();
        }
    }

    public class VB6DataKind : VB6ControlKind
    {
        public DataProperties TypedProperties { get; set; } = new DataProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }

        public VB6DataKind() { }

        public VB6DataKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new DataProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
        }
    }

    public class VB6DirListBoxKind : VB6ControlKind
    {
        public DirListBoxProperties TypedProperties { get; set; } = new DirListBoxProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }

        public VB6DirListBoxKind() { }

        public VB6DirListBoxKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new DirListBoxProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
        }
    }

    public class VB6DriveListBoxKind : VB6ControlKind
    {
        public DriveListBoxProperties TypedProperties { get; set; } = new DriveListBoxProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }

        public VB6DriveListBoxKind() { }

        public VB6DriveListBoxKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new DriveListBoxProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
        }
    }

    public class VB6FileListBoxKind : VB6ControlKind
    {
        public FileListBoxProperties TypedProperties { get; set; } = new FileListBoxProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }

        public VB6FileListBoxKind() { }

        public VB6FileListBoxKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new FileListBoxProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
        }
    }

    public class VB6ImageKind : VB6ControlKind
    {
        public ImageProperties TypedProperties { get; set; } = new ImageProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = []; // Image doesn't have Font group usually
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }
        public byte[]? ResolvedPictureData { get; set; }

        public VB6ImageKind() { }

        public VB6ImageKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new ImageProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
            if (rawProps.TryGetValue("Picture", out var p) && p.IsResource) ResolvedPictureData = p.AsResource();
        }
    }

    public class VB6LabelKind : VB6ControlKind
    {
        public LabelProperties TypedProperties { get; set; } = new LabelProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }

        public VB6LabelKind() { }

        public VB6LabelKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new LabelProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
        }
    }

    public class VB6LineKind : VB6ControlKind
    {
        public LineProperties TypedProperties { get; set; } = new LineProperties();
        // Line typically has no FRX resources or complex groups

        public VB6LineKind() { }

        public VB6LineKind(IDictionary<string, PropertyValue> rawProps)
        {
            TypedProperties = new LineProperties(rawProps);
        }
    }

    public class VB6ListBoxKind : VB6ControlKind
    {
        public ListBoxProperties TypedProperties { get; set; } = new ListBoxProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public List<string> ResolvedListItems { get; set; } = [];
        public List<int> ResolvedItemData { get; set; } = [];
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }

        public VB6ListBoxKind() { }

        public VB6ListBoxKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new ListBoxProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("List", out var l) && l.IsResource) { byte[]? d = l.AsResource(); if (d != null) ResolvedListItems = VB6FormFile.ListResolver(d); }
            if (rawProps.TryGetValue("ItemData", out var id) && id.IsResource) { byte[]? d = id.AsResource(); if (d != null) ResolvedItemData = VB6FormFile.ItemDataResolver(d); }
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
        }
    }

    public class VB6OleKind : VB6ControlKind
    {
        public OLEProperties TypedProperties { get; set; } = new OLEProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }
        public byte[]? ResolvedSourceDocData { get; set; } // If SourceDoc is an FRX link
        public byte[]? ResolvedObjectData { get; set; }    // Raw embedded object data from FRX
        public List<OleVerbInfo> ResolvedObjectVerbs { get; set; } = [];
        public byte[]? ResolvedObjectVerbFlagsData { get; set; }

        public VB6OleKind() { }

        public VB6OleKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new OLEProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
            if (rawProps.TryGetValue("SourceDoc", out var sd) && sd.IsResource) ResolvedSourceDocData = sd.AsResource();
            if (rawProps.TryGetValue("ObjectVerbs", out var ov) && ov.IsResource) { byte[]? d = ov.AsResource(); if (d != null) ResolvedObjectVerbs = VB6FormFile.OleVerbResolver(d); }
            if (rawProps.TryGetValue("ObjectVerbFlags", out var ovf) && ovf.IsResource) ResolvedObjectVerbFlagsData = ovf.AsResource();
            if (rawProps.TryGetValue("Object", out var o) && o.IsResource) ResolvedObjectData = o.AsResource(); // Or "ObjectData"
        }
    }

    public class VB6OptionButtonKind : VB6ControlKind
    {
        public OptionButtonProperties TypedProperties { get; set; } = new OptionButtonProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDisabledPictureData { get; set; }
        public byte[]? ResolvedDownPictureData { get; set; }
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }
        public byte[]? ResolvedPictureData { get; set; }

        public VB6OptionButtonKind() { }

        public VB6OptionButtonKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new OptionButtonProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DisabledPicture", out var dp) && dp.IsResource) ResolvedDisabledPictureData = dp.AsResource();
            if (rawProps.TryGetValue("DownPicture", out var downp) && downp.IsResource) ResolvedDownPictureData = downp.AsResource();
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
            if (rawProps.TryGetValue("Picture", out var p) && p.IsResource) ResolvedPictureData = p.AsResource();
        }
    }

    public class VB6PictureBoxKind : VB6ControlKind
    {
        public PictureBoxProperties TypedProperties { get; set; } = new PictureBoxProperties();
        public List<VB6Control> Controls { get; set; } = [];
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }
        public byte[]? ResolvedPictureData { get; set; }

        public VB6PictureBoxKind() { }

        public VB6PictureBoxKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups, List<VB6Control> childControls)
        {
            TypedProperties = new PictureBoxProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            Controls = childControls ?? [];
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
            if (rawProps.TryGetValue("Picture", out var p) && p.IsResource) ResolvedPictureData = p.AsResource();
        }
    }

    public class VB6HScrollBarKind : VB6ControlKind
    {
        public ScrollBarProperties TypedProperties { get; set; } = new ScrollBarProperties();
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];

        public VB6HScrollBarKind() { }

        public VB6HScrollBarKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new ScrollBarProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
        }
    }

    public class VB6VScrollBarKind : VB6ControlKind
    {
        public ScrollBarProperties TypedProperties { get; set; } = new ScrollBarProperties();
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];

        public VB6VScrollBarKind() { }

        public VB6VScrollBarKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new ScrollBarProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
        }
    }

    public class VB6ShapeKind : VB6ControlKind
    {
        public ShapeProperties TypedProperties { get; set; } = new ShapeProperties();
        // Shape typically no FRX or complex groups

        public VB6ShapeKind() { }

        public VB6ShapeKind(IDictionary<string, PropertyValue> rawProps)
        {
            TypedProperties = new ShapeProperties(rawProps);
        }
    }

    public class VB6TextBoxKind : VB6ControlKind
    {
        public TextBoxProperties TypedProperties { get; set; } = new TextBoxProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }

        public VB6TextBoxKind() { }

        public VB6TextBoxKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new TextBoxProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
        }
    }

    public class VB6TimerKind : VB6ControlKind
    {
        public TimerProperties TypedProperties { get; set; } = new TimerProperties();
        // Timer is non-visual, no FRX or complex groups

        public VB6TimerKind() { }

        public VB6TimerKind(IDictionary<string, PropertyValue> rawProps)
        {
            TypedProperties = new TimerProperties(rawProps);
        }
    }

    public class VB6UserControlKind : VB6ControlKind // Defined earlier, ensuring it's here
    {
        public UserControlProperties TypedProperties { get; set; } = new UserControlProperties();
        public List<VB6Control> ConstituentControls { get; set; } = [];
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedToolboxBitmapBytes { get; set; }
        public byte[]? ResolvedPictureData { get; set; }
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }

        public VB6UserControlKind() { }

        public VB6UserControlKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups, List<VB6Control> constituentControls, string ctlFilePath, ResourceFileResolver resourceResolver)
        {
            TypedProperties = new UserControlProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            ConstituentControls = constituentControls ?? [];
            string parentDirectory = Path.GetDirectoryName(ctlFilePath) ?? string.Empty;

            if (!string.IsNullOrEmpty(TypedProperties.ToolboxBitmap_SourceString))
            {
                if (VB6FormFile.TryParseFrxLinkInternal(TypedProperties.ToolboxBitmap_SourceString, out string? ctxFileName, out int ctxOffset) && ctxFileName != null)
                {
                    string ctxFullPath = Path.Combine(parentDirectory, ctxFileName);
                    ResolvedToolboxBitmapBytes = resourceResolver(ctxFullPath, ctxOffset);
                }
            }
            if (rawProps.TryGetValue("Picture", out var p) && p.IsResource) ResolvedPictureData = p.AsResource();
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
        }
    }

    public class VB6RichTextBoxKind : VB6ControlKind
    {
        public RichTextBoxProperties TypedProperties { get; set; } = new RichTextBoxProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedTextRTFBytes { get; set; }
        public string? ResolvedTextRTF { get; set; }
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }

        public VB6RichTextBoxKind() { }

        public VB6RichTextBoxKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups, string formFilePath, ResourceFileResolver resourceResolver)
        {
            TypedProperties = new RichTextBoxProperties(rawProps);
            PropertyGroups = propGroups ?? [];
            string parentDirectory = Path.GetDirectoryName(formFilePath) ?? string.Empty;

            if (!string.IsNullOrEmpty(TypedProperties.TextRTF_SourceString))
            {
                if (VB6FormFile.TryParseFrxLinkInternal(TypedProperties.TextRTF_SourceString, out string? frxFileName, out int frxOffset) && frxFileName != null)
                {
                    string frxFullPath = Path.Combine(parentDirectory, frxFileName);
                    ResolvedTextRTFBytes = resourceResolver(frxFullPath, frxOffset);
                    if (ResolvedTextRTFBytes != null && ResolvedTextRTFBytes.Length > 0)
                    {
                        Encoding registeredEncoding = Encoding.GetEncoding(1252);
                        try { ResolvedTextRTF = registeredEncoding.GetString(ResolvedTextRTFBytes); }
                        catch (Exception ex) { ResolvedTextRTF = $"[Error decoding RTF: {ex.Message}]"; }
                    }
                    else if (ResolvedTextRTFBytes != null) ResolvedTextRTF = string.Empty;
                }
            }
            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
        }
    }

    public class VB6WinsockKind : VB6ControlKind
    {
        public WinsockProperties TypedProperties { get; set; } = new WinsockProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = []; // Unlikely to have Font etc.
        // Winsock typically doesn't have FRX resources for its core function.
        // DragIcon/MouseIcon might be possible if added to all ActiveX bases.

        public VB6WinsockKind() { }

        public VB6WinsockKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new WinsockProperties(rawProps);
            PropertyGroups = propGroups ?? [];
        }
    }

    public class VB6ListViewKind : VB6ControlKind
    {
        public ListViewProperties TypedProperties { get; set; } = new ListViewProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        public byte[]? ResolvedDragIconData { get; set; }
        public byte[]? ResolvedMouseIconData { get; set; }
        // Icons and SmallIcons are string names of ImageList controls, not direct FRX for ListView itself.
        // ColumnHeaders and ListItems are often handled via PropertyGroups if defined at design-time.

        public VB6ListViewKind() { }

        public VB6ListViewKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups, string formFilePath, ResourceFileResolver resourceResolver)
        {
            TypedProperties = new ListViewProperties(rawProps);
            PropertyGroups = propGroups ?? [];

            if (rawProps.TryGetValue("DragIcon", out var di) && di.IsResource) ResolvedDragIconData = di.AsResource();
            if (rawProps.TryGetValue("MouseIcon", out var mi) && mi.IsResource) ResolvedMouseIconData = mi.AsResource();
            // Note: Icons and SmallIcons properties in ListViewProperties store the *name* of an ImageList control.
            // Actual ImageList data isn't directly part of the ListView's FRX.
            // Parsing ColumnHeaders from PropertyGroups would be a more advanced step if needed.
        }
    }

    /// <summary>
    /// Represents a custom (non-standard VB or third-party) control.
    /// </summary>
    public class VB6CustomKind : VB6ControlKind
    {
        public CustomControlProperties TypedProperties { get; set; } = new CustomControlProperties();
        public List<VB6PropertyGroup> PropertyGroups { get; set; } = [];
        // FRX for custom controls would be handled by accessing rawProps within TypedProperties
        // and checking specific property names known for that custom control.

        public VB6CustomKind() { }

        public VB6CustomKind(IDictionary<string, PropertyValue> rawProps, List<VB6PropertyGroup> propGroups)
        {
            TypedProperties = new CustomControlProperties(rawProps);
            PropertyGroups = propGroups ?? [];
        }
    }


    /// <summary>
    /// Represents a generic VB6 control instance in a form's control hierarchy
    /// or as the main object in a UserControl file.
    /// </summary>
    public class VB6Control
    {
        /// <summary>
        /// The programmatic name of the control (e.g., "Command1", "txtName").
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// User-defined data associated with the control.
        /// </summary>
        public string Tag { get; set; } = string.Empty;

        /// <summary>
        /// If the control is part of a control array, this is its 0-based index.
        /// If not part of an array, this value might be 0 or a sentinel like -1
        /// depending on how VB6 stores it (often absent for non-array, implying 0 or not applicable).
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// The specific kind of this control (e.g., TextBox, CommandButton, Form, Custom ActiveX).
        /// This will hold an instance of a class derived from <see cref="VB6ControlKind"/>.
        /// </summary>
        public VB6ControlKind? Kind { get; set; }

        public VB6Control() { }

        /// <summary>
        /// Convenience constructor.
        /// </summary>
        public VB6Control(string name, int index, string tag, VB6ControlKind kind)
        {
            Name = name;
            Index = index;
            Tag = tag;
            Kind = kind;
        }

        public override string ToString()
        {
            string kindName = Kind?.GetType().Name.Replace("VB6", "").Replace("Kind", "") ?? "Unknown";
            return $"Control: {Name}{(Index != 0 ? $"({Index})" : "")} (Kind: {kindName})";
        }
    }
}