Attribute VB_Name = "MetaTableConfigTool"
Option Explicit

' ======================================================
' MetaTableConfigTool.bas
' Version: 1.0
' Purpose: Extract table metadata to structured XML for
'          AI assistant parsing and project documentation
'
' STANDALONE MODULE - No dependencies on project code
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
            ' Extract version number from filename like "TableMetaExport_v3.xml"
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
    Dim errorCount As Long
    
    ' XML Declaration and Root
    xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    xml = xml & "<TableMetaExport>" & vbCrLf
    xml = xml & "  <ExportMetadata>" & vbCrLf
    xml = xml & "    <Version>" & version & "</Version>" & vbCrLf
    xml = xml & "    <ExportDate>" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "</ExportDate>" & vbCrLf
    xml = xml & "    <SourceWorkbook>" & XMLEncode(ThisWorkbook.Name) & "</SourceWorkbook>" & vbCrLf
    xml = xml & "    <GeneratedBy>" & MODULE_NAME & "</GeneratedBy>" & vbCrLf
    xml = xml & "  </ExportMetadata>" & vbCrLf
    xml = xml & vbCrLf
    xml = xml & "  <Tables>" & vbCrLf
    
    tableCount = 0
    
    ' Process each row in MetaVBAMappingTable
    For Each r In metaTable.ListRows
        tableName = Trim(CStr(GetCellValue(r, metaTable, COL_TABLE_NAMES)))
        
        If Len(tableName) = 0 Then GoTo NextRow
        
        tableDesc = Trim(CStr(GetCellValue(r, metaTable, COL_TABLE_INFO)))
        pullHeaderOnly = (UCase(Trim(CStr(GetCellValue(r, metaTable, COL_PULL_HEADER_ONLY)))) = "TRUE")
        getFormat = (UCase(Trim(CStr(GetCellValue(r, metaTable, COL_GET_FORMAT)))) = "TRUE")
        formatColName = Trim(CStr(GetCellValue(r, metaTable, COL_FORMAT_HEADER)))
        
        ' Find the target table
        Set targetTable = FindTableByName(tableName)
        
        If targetTable Is Nothing Then
            ' Table not found - log error and add error marker to XML
            LogError "Table not found: " & tableName
            xml = xml & BuildErrorTableXML(tableName, tableDesc, "TABLE_NOT_FOUND")
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
        xml = xml & "    <!-- !!!ATTENTION: ERRORS DETECTED DURING EXPORT!!! -->" & vbCrLf
        Dim errMsg As Variant
        For Each errMsg In mErrors
            xml = xml & "    <Error>" & XMLEncode(CStr(errMsg)) & "</Error>" & vbCrLf
        Next errMsg
        xml = xml & "  </Errors>" & vbCrLf
    End If
    
    xml = xml & "  <Summary>" & vbCrLf
    xml = xml & "    <TablesProcessed>" & tableCount & "</TablesProcessed>" & vbCrLf
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
    Dim colIndex As Long
    Dim formatColIndex As Long
    
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

Private Function BuildErrorTableXML(tableName As String, description As String, errorType As String) As String
    Dim xml As String
    
    xml = "    <Table name=""" & XMLEncode(tableName) & """>" & vbCrLf
    xml = xml & "      <!-- !!!ERROR: " & errorType & "!!! -->" & vbCrLf
    xml = xml & "      <Error>" & errorType & "</Error>" & vbCrLf
    xml = xml & "      <Description><![CDATA[" & description & "]]></Description>" & vbCrLf
    xml = xml & "    </Table>" & vbCrLf
    
    BuildErrorTableXML = xml
End Function

' ======================================================
' HELPER FUNCTIONS
' ======================================================

Private Function FindTableByName(tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
    
    Set FindTableByName = Nothing
End Function

Private Function GetCellValue(r As ListRow, tbl As ListObject, colName As String) As Variant
    Dim colIndex As Long
    
    On Error Resume Next
    colIndex = tbl.ListColumns(colName).Index
    On Error GoTo 0
    
    If colIndex = 0 Then
        GetCellValue = ""
        Exit Function
    End If
    
    GetCellValue = r.Range.Cells(1, colIndex).Value
End Function

Private Function InferDataType(col As ListColumn) As String
    Dim sampleCell As Range
    Dim sampleValue As Variant
    Dim i As Long
    Dim hasData As Boolean
    
    ' Check first few non-empty cells to infer type
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
        For Each errMsg In mErrors
            Debug.Print "  - " & errMsg
        Next errMsg
        Debug.Print ""
    Else
        Debug.Print "No errors detected."
    End If
    
    Debug.Print "=============================================="
    
    MsgBox "Export complete!" & vbCrLf & vbCrLf & _
           "File: " & outputPath & vbCrLf & _
           "Version: " & version & vbCrLf & _
           "Errors: " & mErrors.Count, _
           IIf(mErrors.Count > 0, vbExclamation, vbInformation), _
           MODULE_NAME
End Sub
