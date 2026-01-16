Attribute VB_Name = "AV_Constants"
Option Explicit

' ======================================================
' AV_Constants
' All constants and magic numbers in one place
' NO MORE HARDCODED VALUES IN OTHER MODULES
' ======================================================

' ======================================================
' TABLE NAMES - Configuration Tables
' ======================================================

' Primary Configuration Tables
Public Const TBL_VALIDATION_TARGETS As String = "ValidationTargets"
Public Const TBL_AUTO_VAL_MAPPING As String = "AutoValidationCommentPrefixMappingTable"
Public Const TBL_AUTO_FORMAT As String = "AutoFormatOnFullValidation"
Public Const TBL_AUTO_CHECK_VALIDATION As String = "AutoCheckDataValidationTable"

' Menu Field Tables (EN/FR)
Public Const TBL_EN_MENU_FIELDS As String = "ENMenuSelectionMenuFields"
Public Const TBL_FR_MENU_FIELDS As String = "FRMenuSelectionMenuFields"
Public Const TBL_ENFR_HEADER_MAPPING As String = "ENFRHeaderNamesTable"

' Validation Rule Tables
Public Const TBL_GIW_VALIDATION As String = "GIWValidationTable"
Public Const TBL_ELECTRICITY_PAIRS As String = "ElectricityPairValidation"
Public Const TBL_PLUMBING_PAIRS As String = "PlumbingPairValidation"
Public Const TBL_HEAT_SOURCE_PAIRS As String = "HeatSourcePairValidation"
Public Const TBL_HEAT_ANY_REF As String = "HeatSourceANYRefTable"

' Supporting Tables
Public Const TBL_GLOBAL_DEBUG As String = "GlobalDebugOptions"
Public Const TBL_DEBUG_CONTROLS As String = "DebugControls"
Public Const TBL_FORCE_VALIDATION As String = "ForceValidationTable"
Public Const TBL_REVIEW_REF_COLUMNS As String = "ReviewRefColumnTable"
Public Const TBL_REVIEW_STATUS As String = "ReviewStatusTable"
Public Const TBL_DDM_FIELDS_INFO As String = "DDMFieldsInfo"

' Common Error Fix Tables
Public Const TBL_BUILDING_TYPE_ERRORS As String = "BuildingTypeCommonErrorTbl"
Public Const TBL_MAIN_USAGE_TYPE_ERRORS As String = "MainUsageTypeCommonErrorTable"

' ======================================================
' VALIDATION SETTINGS
' ======================================================

' Validation Timeouts & Intervals
Public Const VALIDATION_TIMEOUT_SECONDS As Long = 10000
Public Const VALIDATION_PROGRESS_UPDATE_INTERVAL As Long = 10     ' Update UI every N rows
Public Const VALIDATION_DETAILED_LOG_INTERVAL As Long = 25        ' Detailed log every N rows

' Validation Modes (from ValidationTargets)
Public Const VALIDATION_MODE_BOTH As String = "Both"
Public Const VALIDATION_MODE_TRIGGER As String = "Trigger"
Public Const VALIDATION_MODE_BULK As String = "Bulk"

' ======================================================
' VALIDATION LIMITS
' ======================================================

Public Const MAX_GIW_VALUE As Long = 1000
Public Const MIN_CONSTRUCTION_YEAR As Long = 1800
Public Const MAX_CONSTRUCTION_YEAR As Long = 2100

' ======================================================
' LANGUAGE SETTINGS
' ======================================================

Public Const LANGUAGE_ENGLISH As String = "English"
Public Const LANGUAGE_FRENCH As String = "Français"

' ======================================================
' SYSTEM TAGS & MARKERS
' ======================================================

Public Const SYSTEM_TAG_START As String = "[[SYS_TAG"
Public Const SYSTEM_TAG_END As String = "]]"
Public Const SYSTEM_COMMENT_TAG As String = "[[SYS_COMMENT]]"

' ======================================================
' FORMAT TYPES
' ======================================================

Public Const FORMAT_DEFAULT As String = "Default"
Public Const FORMAT_ERROR As String = "Error"
Public Const FORMAT_AUTOCORRECT As String = "Autocorrect"

' ======================================================
' REVIEW STATUS VALUES
' ======================================================

Public Const STATUS_AUTO_CORRECTED As String = "Auto Corrected"
Public Const STATUS_ERROR As String = "Error"
Public Const STATUS_NO_ERRORS As String = "No Errors Found"

' ======================================================
' CONFIGURATION SHEET
' ======================================================

Public Const CONFIG_SHEET_NAME As String = "Config"

' ======================================================
' LEGACY CELL REFERENCES (DEPRECATED - DO NOT USE)
' These are maintained for documentation only
' Use LoadValidationConfig() instead
' ======================================================

' DEPRECATED: Target sheet name - now from ValidationTargets table
' Public Const CONFIG_TARGET_SHEET_CELL As String = "B3"

' DEPRECATED: Start row - now from target table structure
' Public Const CONFIG_START_ROW_CELL As String = "B4"

' DEPRECATED: Row count - now from target table structure
' Public Const CONFIG_ROW_COUNT_CELL As String = "D4"

' DEPRECATED: Key column - now from ValidationTargets.KeyColumn
' Public Const CONFIG_KEY_COLUMN_CELL As String = "B5"

' DEPRECATED: Language - now from ValidationSettings table
' Public Const CONFIG_LANGUAGE_CELL As String = "M1"

' ======================================================
' COLUMN NAMES - ValidationTargets Table
' ======================================================

Public Const COL_VT_TABLE_NAME As String = "TableName"
Public Const COL_VT_ENABLED As String = "Enabled"
Public Const COL_VT_MODE As String = "Mode"
Public Const COL_VT_KEY_COLUMN As String = "Key Column (Header Name)"

' ======================================================
' COLUMN NAMES - AutoValidationCommentPrefixMappingTable
' ======================================================

Public Const COL_AVCPM_DEV_FUNC_NAMES As String = "Dev Function Names"
Public Const COL_AVCPM_DROP_COLUMN As String = "Drop in Column"
Public Const COL_AVCPM_PREFIX_EN As String = "Prefix to message"
Public Const COL_AVCPM_PREFIX_FR As String = "(FR) Prefix to message"
Public Const COL_AVCPM_RULE_TABLE As String = "RuleTableName"
Public Const COL_AVCPM_AUTO_VALIDATE As String = "AutoValidate"
Public Const COL_AVCPM_REVIEW_COLUMN_HEADER As String = "ReviewSheet Column Header"

' ======================================================
' COLUMN NAMES - AutoFormatOnFullValidation Table
' ======================================================

Public Const COL_AF_FORMAT_KEY As String = "Formatting Key"
Public Const COL_AF_AUTO_FORMATTING As String = "Autoformatting"
Public Const COL_AF_PRIORITY As String = "KeyFlagPriority"

' ======================================================
' COLUMN NAMES - ENFRHeaderNamesTable
' ======================================================

Public Const COL_ENFR_EN_HEADER As String = "EN - ENMenuSelectionMenuFields Table Header"
Public Const COL_ENFR_FR_HEADER As String = "FR - ENMenuSelectionMenuFields Table Header"

' ======================================================
' COLUMN NAMES - AutoCheckDataValidationTable
' ======================================================

Public Const COL_ACDV_COLUMN_NAME As String = "Column Name"
Public Const COL_ACDV_COLUMN_NAME_FR As String = "Column Name (FR)"
Public Const COL_ACDV_REVIEW_COLUMN_NAME As String = "ReviewSheet Column Name"
Public Const COL_ACDV_MENU_FIELD_EN As String = "MenuField Column (EN)"
Public Const COL_ACDV_MENU_FIELD_FR As String = "MenuField Column (FR)"
Public Const COL_ACDV_AUTO_CHECK As String = "AutoCheck"
Public Const COL_ACDV_AUTO_COMMENT As String = "AutoComment Column"

' ======================================================
' VALIDATION MESSAGES - Common Errors
' ======================================================

' Error: Missing Configuration
Public Const ERR_CONFIG_TABLE_MISSING As String = "Critical configuration table missing: {0}" & vbCrLf & _
                                                   "Please check Config sheet has all required tables."

Public Const ERR_TARGET_SHEET_MISSING As String = "Target sheet not found: {0}" & vbCrLf & _
                                                   "Check ValidationTargets table."

Public Const ERR_TARGET_TABLE_MISSING As String = "Target table not found: {0}" & vbCrLf & _
                                                   "Sheet {1} must contain a ListObject table."

Public Const ERR_COLUMN_NOT_FOUND As String = "Column not found: {0}" & vbCrLf & _
                                               "Table: {1}"

' Error: Invalid Configuration
Public Const ERR_INVALID_LANGUAGE As String = "Invalid language setting: {0}" & vbCrLf & _
                                               "Must be 'English' or 'Français'."

Public Const ERR_NO_VALIDATION_TARGETS As String = "No validation targets enabled." & vbCrLf & _
                                                    "Check ValidationTargets table has at least one Enabled=TRUE row."

' ======================================================
' HELPER FUNCTION - String Formatting
' Simple string replacement for error messages
' ======================================================

Public Function FormatString(template As String, ParamArray values() As Variant) As String
    Dim result As String
    Dim i As Long
    
    result = template
    For i = LBound(values) To UBound(values)
        result = Replace(result, "{" & i & "}", CStr(values(i)))
    Next i
    
    FormatString = result
End Function
