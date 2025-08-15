Option Explicit

' Core operations for bulk add/update/delete and history logging.
' Provides public APIs used by the UserForm:
'  - ImportCSVToBulk (path)
'  - ValidateBulkSheet (returns Boolean)
'  - ApplyBulkChanges (action: "addupdate" or "delete")
'  - ExportTemplateCSV (path)

' -----------------------------
' Public API
' -----------------------------
Sub ImportCSVToBulk(csvPath As String)
    On Error GoTo ErrHandler
    If Dir(csvPath) = "" Then
        MsgBox "File CSV không tồn tại: " & csvPath, vbExclamation
        Exit Sub
    End If
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsBulk As Worksheet: Set wsBulk = EnsureSheetExists("Bulk Form")
    wsBulk.Cells.Clear
    
    Dim qt As QueryTable
    Set qt = wsBulk.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=wsBulk.Range("A1"))
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTrimSpace = True
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    MsgBox "Import CSV vào 'Bulk Form' thành công.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Lỗi ImportCSVToBulk: " & Err.Description, vbCritical
End Sub

Function ValidateBulkSheet() As Boolean
    ' Validate required headers and required fields for each row.
    On Error GoTo ErrHandler
    Dim wsBulk As Worksheet: Set wsBulk = EnsureSheetExists("Bulk Form")
    Dim requiredHeaders As Variant
    requiredHeaders = Array("Location", "SEOV Name", "GSCM Name", "Type", "ID Assets", "Model")
    
    Dim missing As Collection: Set missing = New Collection
    Dim h As Variant
    For Each h In requiredHeaders
        If GetColIndex(wsBulk, h) = 0 Then missing.Add h
    Next h
    If missing.Count > 0 Then
        Dim s As String: s = "Thiếu header bắt buộc trong 'Bulk Form': " & vbCrLf
        Dim it As Variant
        For Each it In missing: s = s & "- " & it & vbCrLf: Next it
        MsgBox s, vbExclamation
        ValidateBulkSheet = False
        Exit Function
    End If
    
    ' Check each row has ID Assets
    Dim lastRow As Long: lastRow = LastRow(wsBulk)
    Dim errs As Worksheet: Set errs = EnsureSheetExists("Bulk Errors")
    errs.Cells.Clear
    errs.Range("A1").Value = "Row"
    errs.Range("B1").Value = "Error"
    Dim outRow As Long: outRow = 2
    Dim r As Long
    For r = 2 To lastRow
        Dim idVal As String
        idVal = Trim(CStr(wsBulk.Cells(r, GetColIndex(wsBulk, "ID Assets")).Value))
        If idVal = "" Then
            errs.Cells(outRow, 1).Value = r
            errs.Cells(outRow, 2).Value = "ID Assets trống"
            outRow = outRow + 1
        End If
    Next r
    If outRow > 2 Then
        MsgBox "Phát hiện lỗi dữ liệu. Xem sheet 'Bulk Errors'.", vbExclamation
        ValidateBulkSheet = False
    Else
        MsgBox "Dữ liệu hợp lệ.", vbInformation
        ValidateBulkSheet = True
    End If
    Exit Function
ErrHandler:
    MsgBox "Lỗi ValidateBulkSheet: " & Err.Description, vbCritical
    ValidateBulkSheet = False
End Function

Sub ApplyBulkChanges(actionType As String)
    ' actionType: "addupdate" or "delete"
    On Error GoTo ErrHandler
    Dim wsBulk As Worksheet: Set wsBulk = EnsureSheetExists("Bulk Form")
    Dim wsOffice As Worksheet: Set wsOffice = EnsureSheetExists("Office")
    Dim wsProd As Worksheet: Set wsProd = EnsureSheetExists("Production")
    Dim wsHist As Worksheet: Set wsHist = EnsureSheetExists("Asset Movement History")
    Dim lastRow As Long: lastRow = LastRow(wsBulk)
    If lastRow < 2 Then
        MsgBox "Không có dữ liệu trong 'Bulk Form'.", vbInformation
        Exit Sub
    End If
    
    Dim colMap As Object: Set colMap = GetHeaderMap(wsBulk)
    Dim r As Long
    For r = 2 To lastRow
        Dim idVal As String
        idVal = Trim(CStr(wsBulk.Cells(r, colMap("ID Assets")).Value))
        If idVal = "" Then GoTo ContinueLoop
        If LCase(actionType) = "delete" Then
            ' Delete from Office and Production
            Dim found As Range
            Set found = FindRowByHeaderValue(wsOffice, "ID Assets", idVal)
            If Not found Is Nothing Then
                AppendHistory wsHist, idVal, CStr(wsOffice.Cells(found.Row, GetColIndex(wsOffice, "SEOV Name")).Value), CStr(wsOffice.Cells(found.Row, GetColIndex(wsOffice, "GSCM Name")).Value), CStr(wsOffice.Cells(found.Row, GetColIndex(wsOffice, "Model")).Value), CStr(wsOffice.Cells(found.Row, GetColIndex(wsOffice, "Location")).Value), "", "Bulk delete", Application.UserName, "", 1, "Deleted by bulk"
                found.EntireRow.Delete
            End If
            Set found = FindRowByHeaderValue(wsProd, "ID Assets", idVal)
            If Not found Is Nothing Then
                AppendHistory wsHist, idVal, CStr(wsProd.Cells(found.Row, GetColIndex(wsProd, "SEOV Name")).Value), CStr(wsProd.Cells(found.Row, GetColIndex(wsProd, "GSCM Name")).Value), CStr(wsProd.Cells(found.Row, GetColIndex(wsProd, "Model")).Value), CStr(wsProd.Cells(found.Row, GetColIndex(wsProd, "Location")).Value), "", "Bulk delete", Application.UserName, "", 1, "Deleted by bulk"
                found.EntireRow.Delete
            End If
        Else
            ' Add/Update
            Dim values As Object: Set values = CreateObject("Scripting.Dictionary")
            Dim headers As Variant: headers = Array("Location", "SEOV Name", "GSCM Name", "Type", "ID Assets", "Model")
            Dim i As Long, hh As String
            For i = LBound(headers) To UBound(headers)
                hh = headers(i)
                If colMap.Exists(hh) Then values(hh) = Trim(CStr(wsBulk.Cells(r, colMap(hh)).Value)) Else values(hh) = ""
            Next i
            Dim foundRow As Range
            Set foundRow = FindRowByHeaderValue(wsOffice, "ID Assets", idVal)
            If foundRow Is Nothing Then Set foundRow = FindRowByHeaderValue(wsProd, "ID Assets", idVal)
            If Not foundRow Is Nothing Then
                ' Update
                Dim oldLocation As String
                oldLocation = CStr(foundRow.Parent.Cells(foundRow.Row, GetColIndex(foundRow.Parent, "Location")).Value)
                SetCellByHeader foundRow.Parent, foundRow.Row, "Location", values("Location")
                SetCellByHeader foundRow.Parent, foundRow.Row, "SEOV Name", values("SEOV Name")
                SetCellByHeader foundRow.Parent, foundRow.Row, "GSCM Name", values("GSCM Name")
                SetCellByHeader foundRow.Parent, foundRow.Row, "Type", values("Type")
                SetCellByHeader foundRow.Parent, foundRow.Row, "Model", values("Model")
                If Trim(oldLocation) <> Trim(values("Location")) Then
                    AppendHistory wsHist, idVal, values("SEOV Name"), values("GSCM Name"), values("Model"), oldLocation, values("Location"), "Bulk update", Application.UserName, "", 1, ""
                End If
            Else
                ' Add to Office
                Dim newRow As Long
                newRow = wsOffice.Cells(wsOffice.Rows.Count, 1).End(xlUp).Row + 1
                If newRow < 2 Then newRow = 2
                Dim colNo As Long: colNo = GetColIndex(wsOffice, "No.")
                If colNo > 0 Then wsOffice.Cells(newRow, colNo).Value = newRow - 1
                SetCellByHeader wsOffice, newRow, "ID Assets", idVal
                SetCellByHeader wsOffice, newRow, "Location", values("Location")
                SetCellByHeader wsOffice, newRow, "SEOV Name", values("SEOV Name")
                SetCellByHeader wsOffice, newRow, "GSCM Name", values("GSCM Name")
                SetCellByHeader wsOffice, newRow, "Type", values("Type")
                SetCellByHeader wsOffice, newRow, "Model", values("Model")
                If GetColIndex(wsOffice, "Status") > 0 Then wsOffice.Cells(newRow, GetColIndex(wsOffice, "Status")).Value = "Active"
                AppendHistory wsHist, idVal, values("SEOV Name"), values("GSCM Name"), values("Model"), "", values("Location"), "Bulk add", Application.UserName, "", 1, "Added by bulk via UserForm"
            End If
        End If
ContinueLoop:
    Next r
    MsgBox "Thao tác bulk đã hoàn tất (" & actionType & ").", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Lỗi ApplyBulkChanges: " & Err.Description, vbCritical
End Sub

Sub ExportTemplateCSV(csvPath As String)
    On Error GoTo ErrHandler
    Dim fnum As Integer: fnum = FreeFile
    Dim header As String
    header = "Location,SEOV Name,GSCM Name,Type,ID Assets,Model"
    Open csvPath For Output As #fnum
    Print #fnum, header
    Print #fnum, "Hanoi,Nguyen Van A,Tran B,Desktop,ASSET-0001,Dell OptiPlex 7070"
    Close #fnum
    MsgBox "Template CSV đã xuất: " & csvPath, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Lỗi ExportTemplateCSV: " & Err.Description, vbCritical
End Sub

' -----------------------------
' History & helpers (re-used)
' -----------------------------
Sub AppendHistory(wsHist As Worksheet, idAsset As String, seov As String, gscm As String, model As String, fromLoc As String, toLoc As String, reason As String, reqBy As String, dept As String, numMoves As Long, note As String)
    Dim nextRow As Long
    nextRow = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    Dim mapping As Variant
    mapping = Array("Date", "ID Assets", "SEOV Name", "GSCM Name", "Model", "From", "To", "Reason", "Request By", "Dept", "Number of moves", "Note")
    Dim i As Long, colIndex As Long
    For i = LBound(mapping) To UBound(mapping)
        colIndex = GetColIndex(wsHist, mapping(i))
        If colIndex = 0 Then
            colIndex = wsHist.Cells(1, wsHist.Columns.Count).End(xlToLeft).Column + 1
            wsHist.Cells(1, colIndex).Value = mapping(i)
        End If
    Next i
    wsHist.Cells(nextRow, GetColIndex(wsHist, "Date")).Value = Now
    wsHist.Cells(nextRow, GetColIndex(wsHist, "ID Assets")).Value = idAsset
    wsHist.Cells(nextRow, GetColIndex(wsHist, "SEOV Name")).Value = seov
    wsHist.Cells(nextRow, GetColIndex(wsHist, "GSCM Name")).Value = gscm
    wsHist.Cells(nextRow, GetColIndex(wsHist, "Model")).Value = model
    wsHist.Cells(nextRow, GetColIndex(wsHist, "From")).Value = fromLoc
    wsHist.Cells(nextRow, GetColIndex(wsHist, "To")).Value = toLoc
    wsHist.Cells(nextRow, GetColIndex(wsHist, "Reason")).Value = reason
    wsHist.Cells(nextRow, GetColIndex(wsHist, "Request By")).Value = reqBy
    wsHist.Cells(nextRow, GetColIndex(wsHist, "Dept")).Value = dept
    wsHist.Cells(nextRow, GetColIndex(wsHist, "Number of moves")).Value = numMoves
    wsHist.Cells(nextRow, GetColIndex(wsHist, "Note")).Value = note
End Sub

' -----------------------------
' Generic helpers (re-use previous)
' -----------------------------
Function EnsureSheetExists(sName As String) As Worksheet
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sName
    End If
    Set EnsureSheetExists = ws
End Function

Function LastRow(ws As Worksheet) As Long
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        LastRow = 0
    Else
        LastRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
End Function

Function GetColIndex(ws As Worksheet, headerName As String) As Long
    Dim fc As Range
    Set fc = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not fc Is Nothing Then
        GetColIndex = fc.Column
    Else
        GetColIndex = 0
    End If
End Function

Function GetHeaderMap(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        Dim h As String
        h = Trim(CStr(ws.Cells(1, c).Value))
        If h <> "" Then
            If Not dict.Exists(h) Then dict.Add h, c
        End If
    Next c
    Set GetHeaderMap = dict
End Function

Sub SetCellByHeader(ws As Worksheet, rowNum As Long, headerName As String, value As Variant)
    Dim col As Long
    col = GetColIndex(ws, headerName)
    If col > 0 Then ws.Cells(rowNum, col).Value = value
End Sub
