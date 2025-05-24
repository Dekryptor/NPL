// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Determines the action taken by the Data control when the BOF (Beginning Of File) property is True.
    /// Corresponds to VB6 BOFAction property.
    /// </summary>
    public enum BOFActionType
    {
        /// <summary>(Default) Stays on the first record. MovePrevious button is disabled.</summary>
        MoveFirst = 0, // vbBOFActionMoveFirst
        /// <summary>Triggers Validate and Reposition events. MovePrevious button is disabled.</summary>
        BOF = 1        // vbBOFActionBOF
    }

    /// <summary>
    /// Specifies the type of database or data source for the Data control's Connect property.
    /// </summary>
    public enum ConnectType
    {
        Access,     // Default in Rust
        DBaseIII,
        DBaseIV,
        DBase5_0,   // "dBase 5.0"
        Excel3_0,   // "Excel 3.0"
        Excel4_0,   // "Excel 4.0"
        Excel5_0,   // "Excel 5.0" / "Excel 95"
        Excel8_0,   // "Excel 8.0" / "Excel 97"
        FoxPro2_0,  // "FoxPro 2.0"
        FoxPro2_5,  // "FoxPro 2.5"
        FoxPro2_6,  // "FoxPro 2.6"
        FoxPro3_0,  // "FoxPro 3.0"
        LotusWk1,   // "Lotus WK1"
        LotusWk3,   // "Lotus WK3"
        LotusWk4,   // "Lotus WK4"
        Paradox3X,  // "Paradox 3.x"
        Paradox4X,  // "Paradox 4.x"
        Paradox5X,  // "Paradox 5.x"
        Text        // "Text"
        // Note: VB6 also supports "ODBC;" for general ODBC connection strings.
        // And "" (empty string) which defaults to Jet for Access.
        // The Rust enum doesn't have an empty/ODBC generic option, defaults to Access.
    }

    /// <summary>
    /// Specifies the cursor library used for an ODBCDirect connection.
    /// Corresponds to VB6 DefaultCursorType property.
    /// </summary>
    public enum CursorTypeDefault // DefaultCursorType in VB6
    {
        /// <summary>(Default) Driver default cursor.</summary>
        DefaultCursor = 0, // vbUseDefaultCursor
        /// <summary>Use ODBC cursor library.</summary>
        ODBCCursor = 1,    // vbUseODBCCursor
        /// <summary>Use server-side cursors.</summary>
        ServerSideCursor = 2 // vbUseServerSideCursor
        // Note: VB6 also has vbUseClientBatchCursor = 3 for ADO Data Control, not standard Data Control.
    }

    /// <summary>
    /// Specifies the data access method (Jet or ODBCDirect) used by the Data control.
    /// Corresponds to VB6 DefaultType property.
    /// </summary>
    public enum DataAccessType // DefaultType in VB6
    {
        // Values seem swapped in Rust compared to VB6 constants.
        // VB6: vbUseJet = 2 (Default), vbUseODBC = 1
        // Rust: UseODBC = 1, UseJet = 2 (Default) -- Rust default matches VB6 value.
        /// <summary>Use ODBCDirect.</summary>
        UseODBC = 1, // vbUseODBC
        /// <summary>(Default) Use Microsoft Jet database engine.</summary>
        UseJet = 2   // vbUseJet
    }

    /// <summary>
    /// Determines the action taken by the Data control when the EOF (End Of File) property is True.
    /// Corresponds to VB6 EOFAction property.
    /// </summary>
    public enum EOFActionType
    {
        /// <summary>(Default) Stays on the last record. MoveNext button is disabled.</summary>
        MoveLast = 0, // vbEOFActionMoveLast
        /// <summary>Triggers Validate and Reposition events. MoveNext button is disabled.</summary>
        EOF = 1,      // vbEOFActionEOF
        /// <summary>Triggers Validate, then an AddNew operation, then Reposition. </summary>
        AddNew = 2    // vbEOFActionAddNew
    }

    /// <summary>
    /// Specifies the type of Recordset object created by the Data control.
    /// Corresponds to VB6 RecordsetType property.
    /// </summary>
    public enum RecordsetTypeEnum // RecordsetType in VB6, values from RecordsetTypeConstants
    {
        /// <summary>Table-type Recordset (Jet direct table access).</summary>
        Table = 0,    // vbRSTypeTable
        /// <summary>(Default) Dynaset-type Recordset.</summary>
        Dynaset = 1,  // vbRSTypeDynaset
        /// <summary>Snapshot-type Recordset.</summary>
        Snapshot = 2  // vbRSTypeSnapshot
        // Note: DAO also supports ForwardOnly (3) and Dynamic (4) for OpenRecordset,
        // but Data control UI only offers Table, Dynaset, Snapshot.
    }
}