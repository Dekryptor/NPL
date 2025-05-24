// Namespace: VB6Parse.Language.Controls
namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Defines the specific kind of a VB6 control.
    /// This enum corresponds to the variants of the Rust VB6ControlKind.
    /// </summary>
    public enum Vb6ControlKindEnum
    {
        Unknown, // Default/Error case

        // Standard Controls
        CheckBox,
        ComboBox,
        CommandButton,
        Custom, // For ActiveX/OCX controls
        Data,
        DirListBox,
        DriveListBox,
        FileListBox,
        Form,
        Frame,
        HScrollBar, // Horizontal ScrollBar
        Image,
        Label,
        Line,
        ListBox,
        MDIForm,
        Menu, // Represents a menu item or menu header
        Ole,  // OLE Container
        OptionButton,
        PictureBox,
        Shape,
        TextBox,
        Timer,
        VScrollBar  // Vertical ScrollBar
    }
}