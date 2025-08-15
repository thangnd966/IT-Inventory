' Add a UserForm to the workbook named: ufmBulkImport
' Controls (names recommended):
'  - cmdLoadCSV     (CommandButton)  : "Load CSV"
'  - cmdValidate    (CommandButton)  : "Validate"
'  - cmdApplyAddUpd (CommandButton)  : "Apply Add/Update"
'  - cmdApplyDelete (CommandButton)  : "Apply Delete"
'  - cmdExportTemplate (CommandButton): "Export Template"
'  - lblStatus      (Label)
' Paste this code in the UserForm code window.

Option Explicit

Private Sub cmdLoadCSV_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "CSV Files", "*.csv"
    If fd.Show = -1 Then
        Dim f As String: f = fd.SelectedItems(1)
        Call ImportCSVToBulk(f)
        lblStatus.Caption = "CSV loaded: " & f
    End If
End Sub

Private Sub cmdValidate_Click()
    Dim ok As Boolean
    ok = ValidateBulkSheet()
    If ok Then
        lblStatus.Caption = "Validation OK"
    Else
        lblStatus.Caption = "Validation failed: xem sheet 'Bulk Errors'"
    End If
End Sub

Private Sub cmdApplyAddUpd_Click()
    If ValidateBulkSheet() Then
        If MsgBox("Áp dụng Add/Update cho dữ liệu trong 'Bulk Form'?, vbYesNo + vbQuestion) = vbYes Then
            ApplyBulkChanges "addupdate"
            lblStatus.Caption = "Applied Add/Update"
        End If
    Else
        lblStatus.Caption = "Validation failed - không thực hiện"
    End If
End Sub

Private Sub cmdApplyDelete_Click()
    If ValidateBulkSheet() Then
        If MsgBox("Xóa các ID liệt kê trong 'Bulk Form'?, vbYesNo + vbQuestion) = vbYes Then
            ApplyBulkChanges "delete"
            lblStatus.Caption = "Applied Delete"
        End If
    Else
        lblStatus.Caption = "Validation failed - không thực hiện"
    End Sub

Private Sub cmdExportTemplate_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show = -1 Then
        Dim folder As String: folder = fd.SelectedItems(1)
        Dim path As String: path = folder & Application.PathSeparator & "bulk_template.csv"
        ExportTemplateCSV path
        lblStatus.Caption = "Template xuất: " & path
    End If
End Sub

Private Sub UserForm_Initialize()
    lblStatus.Caption = "Ready"
End Sub
