Attribute VB_Name = "M_Globals"
Option Explicit

' --- Sheet Names ---
Public Const SH_DASHBOARD As String = "Dashboard"
Public Const SH_ORIGINAL_PL As String = "Original PL"
Public Const SH_FIRST_SCAN As String = "First scan"
Public Const SH_DATA_WHOLE As String = "data Whole"
Public Const SH_SCAN_WHOLE As String = "SCAN WHOLE"
Public Const SH_SEGREGATION As String = "segregation"
Public Const SH_SCAN_PARTIAL As String = "scan partial"
Public Const SH_DATA_AFTER As String = "data after reception"
Public Const SH_OPEN_BOXES As String = "OPEN Boxes"
Public Const SH_MISSING As String = "Missing"
Public Const SH_SETTINGS As String = "Settings"

' --- Table Names ---
Public Const TBL_PL As String = "tblOriginalPL"
Public Const TBL_FIRST_SCAN As String = "tblFirstScan"
Public Const TBL_DATA_WHOLE As String = "tblDataWhole"
Public Const TBL_SEGREGATION As String = "tblSegregation"
Public Const TBL_SCAN_PARTIAL As String = "tblScanPartial"
Public Const TBL_DATA_AFTER As String = "tblDataAfterReception"
Public Const TBL_OPEN_BOXES As String = "tblOpenBoxes"
Public Const TBL_MISSING As String = "tblMissing"
Public Const TBL_SETTINGS As String = "tblSettings"

' --- Enums ---
Public Enum ScannerMode
    smFirstScan = 1
    smWholeScan = 2
    smSegregationCarton = 3
    smSegregationItem = 4
    smPartialScan = 5
End Enum

' --- Global Variables ---
Public CurrentOperator As String
Public ActiveSourceCarton As String ' Used during segregation
