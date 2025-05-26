using System.Text;
using VBSix.Parsers;
using VBSix.Utilities;

namespace VBSix.Language.Controls
{
    // No specific enums typically needed for Winsock design-time properties beyond standard ones.
    // Protocol (TCP/UDP) is often set at runtime or has a default.

    /// <summary>
    /// Represents the properties specific to a VB6 Winsock control (MSWINSCK.OCX).
    /// The Winsock control provides TCP and UDP network communication capabilities.
    /// </summary>
    public class WinsockProperties
    {
        // Design-time position of the icon
        public int Left { get; set; } = 120;   // Typical default Left in Twips
        public int Top { get; set; } = 120;    // Typical default Top in Twips

        // ActiveX Internal Properties
        public int _ExtentX { get; set; } = 741; // Default from MSWINSCK.OCX often
        public int _ExtentY { get; set; } = 741;
        public int _Version { get; set; } = 393216; // Version of the control/persistence format

        // Winsock Specific Properties settable at design-time
        // Note: Many Winsock properties are more relevant or only settable at runtime.
        // We capture what's typically found in a .frm file.

        /// <summary>
        /// The local port number to bind to for listening (TCP server) or sending/receiving (UDP).
        /// VB6 Property: LocalPort. Default is 0 (system assigns).
        /// </summary>
        public int LocalPort { get; set; } = 0;

        /// <summary>
        /// The hostname or IP address of the remote computer for client connections or sending UDP data.
        /// VB6 Property: RemoteHost. Default is an empty string.
        /// </summary>
        public string RemoteHost { get; set; } = "";

        /// <summary>
        /// The port number on the remote computer for client connections or sending UDP data.
        /// VB6 Property: RemotePort. Default is 0.
        /// </summary>
        public int RemotePort { get; set; } = 0;

        // Index and Name are on VB6Control itself. Tag is also common.
        // Other common properties for invisible controls (like Enabled) are often not explicitly in .frm
        // unless changed from their defaults, but can be added if observed.
        // public Activation Enabled { get; set; } = Controls.Activation.Enabled; // Usually enabled by default

        /// <summary>
        /// Default constructor initializing with common defaults for a Winsock control.
        /// </summary>
        public WinsockProperties() { }

        /// <summary>
        /// Initializes WinsockProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public WinsockProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Left = rawProps.GetInt32("Left", Left);
            Top = rawProps.GetInt32("Top", Top);
            _ExtentX = rawProps.GetInt32("_ExtentX", _ExtentX);
            _ExtentY = rawProps.GetInt32("_ExtentY", _ExtentY);
            _Version = rawProps.GetInt32("_Version", _Version);

            LocalPort = rawProps.GetInt32("LocalPort", LocalPort);
            RemoteHost = rawProps.GetString("RemoteHost", RemoteHost); // Stored as quoted string in .frm
            RemotePort = rawProps.GetInt32("RemotePort", RemotePort);

            // Example for other common properties if they appear for Winsock in .frm
            // Enabled = rawProps.GetEnum("Enabled", Enabled);
        }
    }
}
