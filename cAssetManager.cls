' Class module named cAssetManager
' Responsibilities: encapsulate find/add/update/delete on sheets

Option Explicit

Private pOffice As Worksheet
Private pProduction As Worksheet

Public Sub Init()
    Set pOffice = EnsureSheetExists("Office")
    Set pProduction = EnsureSheetExists("Production")
End Sub

Public Function FindAsset(idAsset As String) As Range
    Set FindAsset = Nothing
    If pOffice Is Nothing Then Init
    Dim r As Range
    Set r = pOffice.Columns(GetColIndex(pOffice, "ID Assets")).Find(What:=idAsset, LookIn:=xlValues, LookAt:=xlWhole)
    If Not r Is Nothing Then
        Set FindAsset = r
        Exit Function
    End If
    Set r = pProduction.Columns(GetColIndex(pProduction, "ID Assets")).Find(What:=idAsset, LookIn:=xlValues, LookAt:=xlWhole)
    If Not r Is Nothing Then Set FindAsset = r
End Function

Public Sub DeleteAssetByRow(rng As Range)
    If Not rng Is Nothing Then rng.EntireRow.Delete
End Sub

' You can add more methods here for add/update to keep logic centralized.
