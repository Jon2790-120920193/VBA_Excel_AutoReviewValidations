Attribute VB_Name = "MetaTableConfigTool"
Option Explicit

' ======================================================
' MetaTableConfigTool.bas
' Version: 1.2
' Purpose: Extract table metadata to structured XML for
'          AI assistant parsing and project documentation
'
' STANDALONE MODULE - No dependencies on project code
'
' CHANGELOG:
' v1.2 - Added comprehensive <AIContext> section at top of XML
'      - Provides full context for Claude to understand file purpose
'      - Includes project context, structure guide, common mistakes
' v1.1 - Fixed whitespace/hidden char trimming on table names
'      - Added flexible column name matching for Description
'      - Enhanced error messages with more diagnostic detail
'      - Added MsgBox error summary with specific issues
' ======================================================

Private Const MODULE_NAME As String = "MetaTableConfigTool"
Private Const META_TABLE_NAME As String = "MetaVBAMappingTable"
Private Const XML_BASE_FILENAME As String = "TableMetaExport"

' Column names in MetaVBAMappingTable
Private Const COL_TABLE_NAMES As String = "TableNames"
Private Const COL_TABLE_INFO As String = "TableInformation/Description"
Private Const COL_PULL_HEADER_ONLY As String = "PullHeaderOnly"
Private Const COL_GET_FORMAT As String = "GetFormatFromColumn"
Private Const COL_FORMAT_HEADER As String = "FormatColumnHeaderName"

' Error collection for end-of-process reporting
Private mErrors As Collection

' ======================================================
' MAIN ENTRY POINT
' ======================================================

Public Sub ExportTableMetaToXML()
    
    Dim xmlContent As String
    Dim metaTable As ListObject
    Dim wsConfig As Worksheet
    Dim outputPath As String
    Dim version As Long
    Dim r As ListRow
    
    ' Initialize error collection
    Set mErrors = New Collection
    
    On Error GoTo ErrHandler
    
    ' Find MetaVBAMappingTable
    Set metaTable = FindMetaTable()
    If metaTable Is Nothing Then
        Debug.Print "FATAL: " & META_TABLE_NAME & " not found in workbook."
        MsgBox "Cannot find table '" & META_TABLE_NAME & "' in this workbook.", vbCritical, MODULE_NAME
        Exit Sub
    End If
    
    ' Validate MetaVBAMappingTable columns exist
    ValidateMetaTableColumns metaTable
    
    ' Determine output path and version
    version = GetNextVersion(ThisWorkbook.Path)
    outputPath = BuildOutputPath(ThisWorkbook.Path, version)
    
    ' Build XML content
    xmlContent = BuildXMLContent(metaTable, version)
    
    ' Write to file
    WriteXMLFile outputPath, xmlContent
    
    ' Report results
    ReportResults outputPath, version
    
    Exit Sub
    
ErrHandler:
    Debug.Print "FATAL ERROR in ExportTableMetaToXML: " & Err.Number & " - " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, MODULE_NAME
End Sub

' ======================================================
' VALIDATE META TABLE COLUMNS
' ======================================================

Private Sub ValidateMetaTableColumns(metaTable As ListObject)
    Dim colIndex As Long
    Dim colName As Variant
    Dim requiredCols As Variant
    Dim foundCols As String
    Dim missingCols As String
    
    requiredCols = Array(COL_TABLE_NAMES, COL_TABLE_INFO, COL_PULL_HEADER_ONLY, COL_GET_FORMAT, COL_FORMAT_HEADER)
    
    Debug.Print "=== MetaTable Column Validation ==="
    Debug.Print "Looking for columns in: " & metaTable.Name
    Debug.Print ""
    
    ' List all actual columns in the table
    Debug.Print "Actual columns found:"
    Dim col As ListColumn
    For Each col In metaTable.ListColumns
        Debug.Print "  [" & col.Index & "] """ & col.Name & """ (Len=" & Len(col.Name) & ")"
    Next col
    Debug.Print ""
    
    ' Check each required column
    For Each colName In requiredCols
        colIndex = GetColumnIndexFlexible(metaTable, CStr(colName))
        If colIndex > 0 Then
            foundCols = foundCols & "  OK: " & colName & " -> Index " & colIndex & vbCrLf
        Else
            missingCols = missingCols & "  MISSING: " & colName & vbCrLf
            LogError "MetaTable column not found: '" & colName & "' - check for typos or special characters"
        End If
    Next colName
    
    Debug.Print "Column mapping results:"
    If Len(foundCols) > 0 Then Debug.Print foundCols
    If Len(missingCols) > 0 Then Debug.Print missingCols
    Debug.Print "==================================="
End Sub

' ======================================================
' FIND META TABLE
' ======================================================

Private Function FindMetaTable() As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, META_TABLE_NAME, vbTextCompare) = 0 Then
                Set FindMetaTable = lo
                Exit Function
            End If
        Next lo
    Next ws
    
    Set FindMetaTable = Nothing
End Function

' ======================================================
' VERSION MANAGEMENT
' ======================================================

Private Function GetNextVersion(folderPath As String) As Long
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim maxVersion As Long
    Dim currentVersion As Long
    Dim fileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        GetNextVersion = 1
        Exit Function
    End If
    
    Set folder = fso.GetFolder(folderPath)
    maxVersion = 0
    
    For Each file In folder.Files
        fileName = file.Name
        If Left(fileName, Len(XML_BASE_FILENAME)) = XML_BASE_FILENAME And Right(fileName, 4) = ".xml" Then
            currentVersion = ExtractVersionFromFilename(fileName)
            If currentVersion > maxVersion Then
                maxVersion = currentVersion
            End If
        End If
    Next file
    
    GetNextVersion = maxVersion + 1
End Function

Private Function ExtractVersionFromFilename(fileName As String) As Long
    Dim vPos As Long
    Dim dotPos As Long
    Dim versionStr As String
    
    vPos = InStrRev(fileName, "_v")
    If vPos = 0 Then
        ExtractVersionFromFilename = 0
        Exit Function
    End If
    
    dotPos = InStrRev(fileName, ".")
    If dotPos <= vPos Then
        ExtractVersionFromFilename = 0
        Exit Function
    End If
    
    versionStr = Mid(fileName, vPos + 2, dotPos - vPos - 2)
    
    If IsNumeric(versionStr) Then
        ExtractVersionFromFilename = CLng(versionStr)
    Else
        ExtractVersionFromFilename = 0
    End If
End Function

Private Function BuildOutputPath(folderPath As String, version As Long) As String
    BuildOutputPath = folderPath & "\" & XML_BASE_FILENAME & "_v" & version & ".xml"
End Function

' ======================================================
' XML CONTENT BUILDER
' ======================================================

Private Function BuildXMLContent(metaTable As ListObject, version As Long) As String
    Dim xml As String
    Dim r As ListRow
    Dim tableName As String
    Dim tableDesc As String
    Dim pullHeaderOnly As Boolean
    Dim getFormat As Boolean
    Dim formatColName As String
    Dim targetTable As ListObject
    Dim tableCount As Long
    Dim rowNum As Long
    
    ' XML Declaration and Root
    xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    xml = xml & "<TableMetaExport>" & vbCrLf
    xml = xml & vbCrLf
    
    ' AI Context Section - Comprehensive description for Claude
    xml = xml & "  <!-- ================================================================== -->" & vbCrLf
    xml = xml & "  <!-- AI ASSISTANT: READ THIS SECTION FIRST                             -->" & vbCrLf
    xml = xml & "  <!-- ================================================================== -->" & vbCrLf
    xml = xml & "  <AIContext>" & vbCrLf
    xml = xml & "    <Purpose><![CDATA[" & vbCrLf
    xml = xml & "This file is a structured export of Excel ListObject tables (Excel Tables) from a VBA project." & vbCrLf
    xml = xml & "It provides you (Claude) with accurate, machine-readable visibility into the exact table structures," & vbCrLf
    xml = xml & "column names, data types, and actual data values used in the Auto-Validation system." & vbCrLf
    xml = xml & vbCrLf
    xml = xml & "USE THIS FILE TO:" & vbCrLf
    xml = xml & "- Understand which tables exist and their exact names (case-sensitive)" & vbCrLf
    xml = xml & "- See exact column headers (critical for VBA code that references them)" & vbCrLf
    xml = xml & "- Understand the data structure and valid values in configuration tables" & vbCrLf
    xml = xml & "- Verify your assumptions about table relationships before writing code" & vbCrLf
    xml = xml & "- Debug issues where code references incorrect table/column names" & vbCrLf
    xml = xml & "    ]]></Purpose>" & vbCrLf
    xml = xml & vbCrLf
    xml = xml & "    <ProjectContext><![CDATA[" & vbCrLf
    xml = xml & "PROJECT: Auto-Validation System for Excel" & vbCrLf
    xml = xml & "DESCRIPTION: A VBA-based validation framework that validates building data in Excel." & vbCrLf
    xml = xml & "The system uses configuration tables to define:" & vbCrLf
    xml = xml & "  - Which validation functions to run (AutoValidationCommentPrefixMappingTable)" & vbCrLf
    xml = xml & "  - Valid value pairs for field validation (e.g., ElectricityPairValidation)" & vbCrLf
    xml = xml & "  - Formatting styles for error/success states (AutoFormatOnFullValidation)" & vbCrLf
    xml = xml & "  - Dropdown menu valid values (via DDM tables)" & vbCrLf
    xml = xml & vbCrLf
    xml = xml & "ARCHITECTURE: Table-driven validation where business rules are stored in Excel Tables," & vbCrLf
    xml = xml & "not hardcoded in VBA. This allows configurators to modify validation rules without code changes." & vbCrLf
    xml = xml & "    ]]></ProjectContext>" & vbCrLf
    xml = xml & vbCrLf
    xml = xml & "    <HowToRead><![CDATA[" & vbCrLf
    xml = xml & "STRUCTURE GUIDE:" & vbCrLf
    xml = xml & vbCrLf
    xml = xml & "<Table name=""X"">           - Excel ListObject name (use exactly as shown in VBA)" & vbCrLf
    xml = xml & "  <Location>               - Worksheet containing this table" & vbCrLf
    xml = xml & "  <Description>            - Human/AI-readable explanation of table purpose" & vbCrLf
    xml = xml & "  <RowCount>/<ColumnCount> - Table dimensions" & vbCrLf
    xml = xml & "  <HeaderOnly>             - TRUE = large data table, only headers exported" & vbCrLf
    xml = xml & "                             FALSE = full data exported (config tables)" & vbCrLf
    xml = xml & "  <FormatSource>           - If present, this table contains cells whose FORMATTING" & vbCrLf
    xml = xml & "                             (colors, borders) should be read for validation styling" & vbCrLf
    xml = xml & "  <Columns>                - All column headers with inferred data types" & vbCrLf
    xml = xml & "  <Data>                   - Actual row data (if HeaderOnly=FALSE)" & vbCrLf
    xml = xml & vbCrLf
    xml = xml & "IMPORTANT NOTES:" & vbCrLf
    xml = xml & "- Column names are CASE-SENSITIVE in VBA ListObject references" & vbCrLf
    xml = xml & "- Empty cells appear as empty <Cell> tags" & vbCrLf
    xml = xml & "- Boolean values appear as 'True'/'False' strings" & vbCrLf
    xml = xml & "- The <Description> field contains context I (the user) provided about each table" & vbCrLf
    xml = xml & "- Check <Errors> section at bottom if any tables failed to export" & vbCrLf
    xml = xml & "    ]]></HowToRead>" & vbCrLf
    xml = xml & vbCrLf
    xml = xml & "    <CommonMistakesToAvoid><![CDATA[" & vbCrLf
    xml = xml & "WHEN WRITING VBA CODE, AVOID THESE ERRORS:" & vbCrLf
    xml = xml & "1. Using wrong column names - ALWAYS verify against <Columns> section" & vbCrLf
    xml = xml & "2. Assuming column order - Use column NAMES not index numbers" & vbCrLf
    xml = xml & "3. Case mismatches - 'AutoValidate' is not 'autovalidate'" & vbCrLf
    xml = xml & "4. Missing tables - Check table actually exists before referencing" & vbCrLf
    xml = xml & "5. Wrong data types - Check <DataType> before parsing values" & vbCrLf
    xml = xml & "6. Hardcoding values that exist in config tables - Use table lookups instead" & vbCrLf
    xml = xml & "    ]]></CommonMistakesToAvoid>" & vbCrLf
    xml = xml & "  </AIContext>" & vbCrLf
    xml = xml & vbCrLf
    
    ' Export Metadata
    xml = xml & "  <ExportMetadata>" & vbCrLf
    xml = xml & "    <Version>" & version & "</Version>" & vbCrLf
    xml = xml & "    <ExportDate>" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "</ExportDate>" & vbCrLf
    xml = xml & "    <SourceWorkbook>" & XMLEncode(ThisWorkbook.Name) & "</SourceWorkbook>" & vbCrLf
    xml = xml & "    <GeneratedBy>" & MODULE_NAME & "</GeneratedBy>" & vbCrLf
    xml = xml & "  </ExportMetadata>" & vbCrLf
    xml = xml & vbCrLf
    xml = xml & "  <Tables>" & vbCrLf
    
    tableCount = 0
    rowNum = 0
    
    ' Process each row in MetaVBAMappingTable
    For Each r In metaTable.ListRows
        rowNum = rowNum + 1
        
        ' Get and CLEAN table name (remove all hidden chars, whitespace, line breaks)
        tableName = CleanString(CStr(GetCellValueFlexible(r, metaTable, COL_TABLE_NAMES)))
        
        If Len(tableName) = 0 Then GoTo NextRow
        
        ' Get description with flexible column matching
        tableDesc = Trim(CStr(GetCellValueFlexible(r, metaTable, COL_TABLE_INFO)))
        pullHeaderOnly = (UCase(CleanString(CStr(GetCellValueFlexible(r, metaTable, COL_PULL_HEADER_ONLY)))) = "TRUE")
        getFormat = (UCase(CleanString(CStr(GetCellValueFlexible(r, metaTable, COL_GET_FORMAT)))) = "TRUE")
        formatColName = CleanString(CStr(GetCellValueFlexible(r, metaTable, COL_FORMAT_HEADER)))
        
        ' Debug output for troubleshooting
        Debug.Print "[Row " & rowNum & "] Processing: '" & tableName & "' (Len=" & Len(tableName) & ")"
        
        ' Find the target table
        Set targetTable = FindTableByName(tableName)
        
        If targetTable Is Nothing Then
            ' Table not found - log detailed error
            Dim errDetail As String
            errDetail = "Table '" & tableName & "' not found. " & _
                       "Row=" & rowNum & " in " & META_TABLE_NAME & ". " & _
                       "Check for: hidden characters, extra spaces, typos, or table doesn't exist."
            LogError errDetail
            xml = xml & BuildErrorTableXML(tableName, tableDesc, "TABLE_NOT_FOUND", errDetail, rowNum)
        Else
            ' Build table XML
            xml = xml & BuildTableXML(targetTable, tableDesc, pullHeaderOnly, getFormat, formatColName)
            tableCount = tableCount + 1
        End If
        
NextRow:
    Next r
    
    xml = xml & "  </Tables>" & vbCrLf
    xml = xml & vbCrLf
    
    ' Add error summary if any errors occurred
    If mErrors.Count > 0 Then
        xml = xml & "  <Errors>" & vbCrLf
        xml = xml & "    <!-- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! -->" & vbCrLf
        xml = xml & "    <!-- !!!ATTENTION: ERRORS DETECTED DURING EXPORT!!! -->" & vbCrLf
        xml = xml & "    <!-- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! -->" & vbCrLf
        Dim errMsg As Variant
        Dim errIdx As Long
        errIdx = 0
        For Each errMsg In mErrors
            errIdx = errIdx + 1
            xml = xml & "    <Error index=""" & errIdx & """><![CDATA[" & CStr(errMsg) & "]]></Error>" & vbCrLf
        Next errMsg
        xml = xml & "  </Errors>" & vbCrLf
    End If
    
    xml = xml & "  <Summary>" & vbCrLf
    xml = xml & "    <TablesProcessed>" & tableCount & "</TablesProcessed>" & vbCrLf
    xml = xml & "    <TablesRequested>" & metaTable.ListRows.Count & "</TablesRequested>" & vbCrLf
    xml = xml & "    <ErrorCount>" & mErrors.Count & "</ErrorCount>" & vbCrLf
    xml = xml & "  </Summary>" & vbCrLf
    xml = xml & "</TableMetaExport>"
    
    BuildXMLContent = xml
End Function

Private Function BuildTableXML(tbl As ListObject, description As String, _
                               headerOnly As Boolean, getFormat As Boolean, _
                               formatColName As String) As String
    Dim xml As String
    Dim col As ListColumn
    Dim r As ListRow
    Dim cellVal As Variant
    Dim rowIndex As Long
    
    xml = "    <Table name=""" & XMLEncode(tbl.Name) & """>" & vbCrLf
    xml = xml & "      <Location>" & XMLEncode(tbl.Parent.Name) & "</Location>" & vbCrLf
    xml = xml & "      <Description><![CDATA[" & description & "]]></Description>" & vbCrLf
    xml = xml & "      <RowCount>" & tbl.ListRows.Count & "</RowCount>" & vbCrLf
    xml = xml & "      <ColumnCount>" & tbl.ListColumns.Count & "</ColumnCount>" & vbCrLf
    xml = xml & "      <HeaderOnly>" & IIf(headerOnly, "TRUE", "FALSE") & "</HeaderOnly>" & vbCrLf
    
    ' Format column info
    If getFormat And Len(formatColName) > 0 Then
        xml = xml & "      <FormatSource>" & vbCrLf
        xml = xml & "        <Enabled>TRUE</Enabled>" & vbCrLf
        xml = xml & "        <ColumnName>" & XMLEncode(formatColName) & "</ColumnName>" & vbCrLf
        xml = xml & "        <Note>Cell formatting should be read from this column for validation styling</Note>" & vbCrLf
        xml = xml & "      </FormatSource>" & vbCrLf
    End If
    
    ' Columns section
    xml = xml & "      <Columns>" & vbCrLf
    For Each col In tbl.ListColumns
        xml = xml & "        <Column index=""" & col.Index & """>" & vbCrLf
        xml = xml & "          <Name>" & XMLEncode(col.Name) & "</Name>" & vbCrLf
        xml = xml & "          <DataType>" & InferDataType(col) & "</DataType>" & vbCrLf
        xml = xml & "        </Column>" & vbCrLf
    Next col
    xml = xml & "      </Columns>" & vbCrLf
    
    ' Data section (if not header only)
    If Not headerOnly Then
        xml = xml & "      <Data>" & vbCrLf
        
        If tbl.ListRows.Count = 0 Then
            xml = xml & "        <!-- No data rows -->" & vbCrLf
        Else
            rowIndex = 0
            For Each r In tbl.ListRows
                rowIndex = rowIndex + 1
                xml = xml & "        <Row index=""" & rowIndex & """>" & vbCrLf
                
                For Each col In tbl.ListColumns
                    cellVal = r.Range.Cells(1, col.Index).Value
                    xml = xml & "          <Cell column=""" & XMLEncode(col.Name) & """>"
                    xml = xml & XMLEncode(CellValueToString(cellVal))
                    xml = xml & "</Cell>" & vbCrLf
                Next col
                
                xml = xml & "        </Row>" & vbCrLf
            Next r
        End If
        
        xml = xml & "      </Data>" & vbCrLf
    Else
        xml = xml & "      <Data><!-- HeaderOnly=TRUE: Data rows not exported --></Data>" & vbCrLf
    End If
    
    xml = xml & "    </Table>" & vbCrLf
    
    BuildTableXML = xml
End Function

Private Function BuildErrorTableXML(tableName As String, description As String, _
                                    errorType As String, errorDetail As String, _
                                    sourceRow As Long) As String
    Dim xml As String
    
    xml = "    <Table name=""" & XMLEncode(tableName) & """>" & vbCrLf
    xml = xml & "      <!-- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! -->" & vbCrLf
    xml = xml & "      <!-- !!!ERROR: " & errorType & "!!! -->" & vbCrLf
    xml = xml & "      <!-- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! -->" & vbCrLf
    xml = xml & "      <Error>" & vbCrLf
    xml = xml & "        <Type>" & errorType & "</Type>" & vbCrLf
    xml = xml & "        <SourceRow>" & sourceRow & "</SourceRow>" & vbCrLf
    xml = xml & "        <Detail><![CDATA[" & errorDetail & "]]></Detail>" & vbCrLf
    xml = xml & "      </Error>" & vbCrLf
    xml = xml & "      <Description><![CDATA[" & description & "]]></Description>" & vbCrLf
    xml = xml & "    </Table>" & vbCrLf
    
    BuildErrorTableXML = xml
End Function

' ======================================================
' FLEXIBLE COLUMN MATCHING
' Handles variations in column names (with/without special chars)
' ======================================================

Private Function GetColumnIndexFlexible(tbl As ListObject, colName As String) As Long
    Dim col As ListColumn
    Dim cleanedTarget As String
    Dim cleanedCol As String
    
    ' First try exact match
    On Error Resume Next
    GetColumnIndexFlexible = tbl.ListColumns(colName).Index
    On Error GoTo 0
    
    If GetColumnIndexFlexible > 0 Then Exit Function
    
    ' Try case-insensitive match
    cleanedTarget = LCase(Trim(colName))
    
    For Each col In tbl.ListColumns
        cleanedCol = LCase(Trim(col.Name))
        
        ' Exact match (case-insensitive)
        If cleanedCol = cleanedTarget Then
            GetColumnIndexFlexible = col.Index
            Exit Function
        End If
        
        ' Match ignoring special characters (/ becomes space or removed)
        If Replace(cleanedCol, "/", "") = Replace(cleanedTarget, "/", "") Then
            GetColumnIndexFlexible = col.Index
            Exit Function
        End If
        
        If Replace(cleanedCol, "/", " ") = Replace(cleanedTarget, "/", " ") Then
            GetColumnIndexFlexible = col.Index
            Exit Function
        End If
        
        ' Partial match for Description column specifically
        If InStr(1, cleanedTarget, "description", vbTextCompare) > 0 Then
            If InStr(1, cleanedCol, "description", vbTextCompare) > 0 Then
                GetColumnIndexFlexible = col.Index
                Exit Function
            End If
            If InStr(1, cleanedCol, "info", vbTextCompare) > 0 Then
                GetColumnIndexFlexible = col.Index
                Exit Function
            End If
        End If
    Next col
    
    GetColumnIndexFlexible = 0
End Function

Private Function GetCellValueFlexible(r As ListRow, tbl As ListObject, colName As String) As Variant
    Dim colIndex As Long
    
    colIndex = GetColumnIndexFlexible(tbl, colName)
    
    If colIndex = 0 Then
        GetCellValueFlexible = ""
        Exit Function
    End If
    
    GetCellValueFlexible = r.Range.Cells(1, colIndex).Value
End Function

' ======================================================
' HELPER FUNCTIONS
' ======================================================

Private Function FindTableByName(tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim cleanedSearch As String
    Dim cleanedTable As String
    
    cleanedSearch = CleanString(tableName)
    
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            cleanedTable = CleanString(lo.Name)
            
            ' Try exact match first
            If StrComp(cleanedTable, cleanedSearch, vbTextCompare) = 0 Then
                Set FindTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
    
    ' Debug: List all available tables if not found
    Debug.Print "  Table '" & tableName & "' not found. Available tables:"
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            Debug.Print "    - " & ws.Name & "!" & lo.Name & " (Len=" & Len(lo.Name) & ")"
        Next lo
    Next ws
    
    Set FindTableByName = Nothing
End Function

Private Function CleanString(text As String) As String
    Dim result As String
    Dim i As Long
    Dim c As String
    Dim ascCode As Long
    
    result = ""
    
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        ascCode = AscW(c)
        
        ' Keep only printable ASCII characters (32-126) and common extended chars
        If ascCode >= 32 And ascCode <= 126 Then
            result = result & c
        ElseIf ascCode >= 192 And ascCode <= 255 Then
            ' Keep accented characters
            result = result & c
        End If
        ' Skip: line breaks, tabs, null chars, other control chars
    Next i
    
    ' Also trim leading/trailing whitespace
    CleanString = Trim(result)
End Function

Private Function InferDataType(col As ListColumn) As String
    Dim sampleCell As Range
    Dim sampleValue As Variant
    Dim i As Long
    Dim hasData As Boolean
    
    If col.DataBodyRange Is Nothing Then
        InferDataType = "Unknown"
        Exit Function
    End If
    
    For i = 1 To Application.Min(10, col.DataBodyRange.Rows.Count)
        sampleValue = col.DataBodyRange.Cells(i, 1).Value
        
        If Not IsEmpty(sampleValue) And Len(Trim(CStr(sampleValue))) > 0 Then
            hasData = True
            
            If IsDate(sampleValue) Then
                InferDataType = "Date"
                Exit Function
            ElseIf IsNumeric(sampleValue) Then
                If InStr(CStr(sampleValue), ".") > 0 Then
                    InferDataType = "Decimal"
                Else
                    InferDataType = "Integer"
                End If
                Exit Function
            ElseIf UCase(Trim(CStr(sampleValue))) = "TRUE" Or UCase(Trim(CStr(sampleValue))) = "FALSE" Then
                InferDataType = "Boolean"
                Exit Function
            Else
                InferDataType = "String"
                Exit Function
            End If
        End If
    Next i
    
    If hasData Then
        InferDataType = "String"
    Else
        InferDataType = "Empty"
    End If
End Function

Private Function CellValueToString(val As Variant) As String
    If IsNull(val) Or IsEmpty(val) Then
        CellValueToString = ""
    ElseIf IsError(val) Then
        CellValueToString = "#ERROR#"
    ElseIf IsDate(val) Then
        CellValueToString = Format(val, "yyyy-mm-dd")
    Else
        CellValueToString = CStr(val)
    End If
End Function

Private Function XMLEncode(text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "&", "&amp;")
    result = Replace(result, "<", "&lt;")
    result = Replace(result, ">", "&gt;")
    result = Replace(result, """", "&quot;")
    result = Replace(result, "'", "&apos;")
    XMLEncode = result
End Function

' ======================================================
' ERROR LOGGING
' ======================================================

Private Sub LogError(msg As String)
    mErrors.Add msg
    Debug.Print "[ERROR] " & msg
End Sub

' ======================================================
' FILE OUTPUT
' ======================================================

Private Sub WriteXMLFile(filePath As String, content As String)
    Dim fso As Object
    Dim ts As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, True) ' Overwrite, Unicode
    
    ts.Write content
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing
End Sub

' ======================================================
' RESULTS REPORTING
' ======================================================

Private Sub ReportResults(outputPath As String, version As Long)
    Dim msgText As String
    Dim errSummary As String
    
    Debug.Print "=============================================="
    Debug.Print "MetaTableConfigTool Export Complete"
    Debug.Print "=============================================="
    Debug.Print "Version: " & version
    Debug.Print "Output: " & outputPath
    Debug.Print "Timestamp: " & Now
    Debug.Print ""
    
    If mErrors.Count > 0 Then
        Debug.Print "!!! ERRORS DETECTED !!!"
        Debug.Print "Error Count: " & mErrors.Count
        Debug.Print ""
        Dim errMsg As Variant
        Dim errNum As Long
        errNum = 0
        For Each errMsg In mErrors
            errNum = errNum + 1
            Debug.Print "  [" & errNum & "] " & errMsg
            If errNum <= 5 Then
                errSummary = errSummary & vbCrLf & "  " & errNum & ". " & Left(CStr(errMsg), 80)
                If Len(CStr(errMsg)) > 80 Then errSummary = errSummary & "..."
            End If
        Next errMsg
        If errNum > 5 Then
            errSummary = errSummary & vbCrLf & "  ... and " & (errNum - 5) & " more (see Immediate Window)"
        End If
        Debug.Print ""
    Else
        Debug.Print "No errors detected."
    End If
    
    Debug.Print "=============================================="
    
    ' Build message box text
    msgText = "Export complete!" & vbCrLf & vbCrLf & _
              "File: " & outputPath & vbCrLf & _
              "Version: " & version & vbCrLf & _
              "Errors: " & mErrors.Count
    
    If mErrors.Count > 0 Then
        msgText = msgText & vbCrLf & vbCrLf & "ERRORS:" & errSummary
    End If
    
    MsgBox msgText, _
           IIf(mErrors.Count > 0, vbExclamation, vbInformation), _
           MODULE_NAME & " v1.2"
End Sub
