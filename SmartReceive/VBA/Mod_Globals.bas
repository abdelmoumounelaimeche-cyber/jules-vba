Attribute VB_Name = "Mod_Globals"
Option Explicit

' Sheet Names
Public Const SH_DASHBOARD As String = "Dashboard"
Public Const SH_ORIGINAL_PL As String = "Original PL"
Public Const SH_FIRST_SCAN As String = "First scan"
Public Const SH_DATA_WHOLE As String = "data Whole"
Public Const SH_SCAN_WHOLE As String = "SCAN WHOLE"
Public Const SH_PARTIAL As String = "PARTIAL"
Public Const SH_SEGREGATION As String = "segregation"
Public Const SH_SCAN_PARTIAL As String = "scan partial"
Public Const SH_DATA_RECEPTION As String = "data after reception"
Public Const SH_OPEN_BOXES As String = "OPEN Boxes"
Public Const SH_MISSING As String = "Missing"
Public Const SH_SETTINGS As String = "Settings"

' Table Names
Public Const TBL_SETTINGS As String = "tblSettings"
Public Const TBL_ORIGINAL_PL As String = "tblOriginalPL"
Public Const TBL_STORES As String = "tblStores"
Public Const TBL_PROVIDERS As String = "tblProviders"

' Classification Constants
Public Const CLASS_WHOLE As String = "WHOLE"
Public Const CLASS_PARTIAL As String = "PARTIAL"

' Status Constants
Public Const STATUS_PENDING As String = "PENDING"
Public Const STATUS_IN_PROGRESS As String = "IN PROGRESS"
Public Const STATUS_COMPLETED As String = "COMPLETED"
Public Const STATUS_MISSING As String = "MISSING/DAMAGED"

' Global State (Stored in Settings sheet for persistence)
Public Property Get CurrentCartonID() As String
    CurrentCartonID = GetSetting("CurrentCartonID")
End Property

Public Property Let CurrentCartonID(ByVal value As String)
    SetSetting "CurrentCartonID", value
End Property
