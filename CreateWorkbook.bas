Option Explicit

' Create a new IT-Inventory.xlsm workbook with required sheets and headers.
Sub CreateITInventoryWorkbook(Optional savePath As String)
    On Error Go To ErrHandler
    Dim wb As Workbook
    Set wb = Workbooks.Add
    ' Remove extras to start clean
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        sh.Delete
    Next sh
    
    Dim headers As Variant
    headers = Array("No.", "ID Code", "User", "Dept", "Location", "SEOV Name", "GSCM Name", "Type", "ID Assets", "Model", "Hostname", "Mac LAN", "Mac Wifi", "Serial number", "Recived Date", "Supplier", "FA", "FA Code", "PO Number", "Kian No", "Status", "Fist checkout date", "Reason checkout", "Checkin date", "Reason Checkin", "Note")
    
    Dim wsOffice As Worksheet: Set wsOffice = wb.Worksheets.Add
    wsOffice.Name = "Office"
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        wsOffice.Cells(1, i + 1).Value = headers(i)
    Next i
    
    Dim wsProd As Worksheet: Set wsProd = wb.Worksheets.Add
    wsProd.Name = "Production"
    For i = LBound(headers) To UBound(headers)
        wsProd.Cells(1, i + 1).Value = headers(i)
    Next i
    
    Dim wsHist As Worksheet: Set wsHist = wb.Worksheets.Add
    wsHist.Name = "Asset Movement History"
    Dim hhist As Variant
    hhist = Array("Date", "ID Assets", "SEOV Name", "GSCM Name", "Model", "From", "To", "Reason", "Request By", "Dept", "Number of moves", "Note")
    For i = LBound(hhist) To UBound(hhist)
        wsHist.Cells(1, i + 1).Value = hhist(i)
    Next i
    
    Dim wsBulk As Worksheet: Set wsBulk = wb.Worksheets.Add
    wsBulk.Name = "Bulk Form"
    Dim hbulk As Variant
    hbulk = Array("Location", "SEOV Name", "GSCM Name", "Type", "ID Assets", "Model")
    For i = LBound(hbulk) To UBound(hbulk)
        wsBulk.Cells(1, i + 1).Value = hbulk(i)
    Next i
    
    Dim wsMenu As Worksheet: Set wsMenu = wb.Worksheets.Add
    wsMenu.Name = "Menu"
    wsMenu.Cells(1, 1).Value = "Menu - Bulk Operations"
    
    ' Save workbook
    Dim finalPath As String
    If savePath = "" Then
        finalPath = Application.DefaultFilePath & Application.PathSeparator & "IT-Inventory.xlsm"
    Else
        finalPath = savePath
    End If
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=finalPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    MsgBox "Workbook mẫu đã được tạo: " & finalPath, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Lỗi CreateITInventoryWorkbook: " & Err.Description, vbCritical
End Sub
