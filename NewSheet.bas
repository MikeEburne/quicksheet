Attribute VB_Name = "NewSheet"
' Quick Sheet - VBA version (for use inside your own workbooks).
' This file CANNOT be uploaded to the Microsoft store - the store only accepts the
' Office.js add-in in this same folder (manifest.xml + the .html/.js files).
'
' To use: in Excel press Alt+F11 > File > Import File... > pick NewSheet.bas,
' then save the workbook as .xlsm and run the macro CreateNewSheet (Alt+F8).
Option Explicit

Public Sub CreateNewSheet()
    Dim nm As String, ws As Worksheet

    nm = InputBox("Name for the new sheet:", "Quick Sheet", "New Sheet")
    If StrPtr(nm) = 0 Then Exit Sub            ' Cancel pressed
    nm = UniqueSheetName(CleanSheetName(nm))

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = nm
    With ws.Range("A1")
        .Value = "Created by Quick Sheet on " & Format(Now, "dd mmm yyyy hh:nn")
        .Font.Bold = True
    End With
    ws.Range("A2").Select
End Sub

' Excel sheet names: max 31 characters, none of  \ / ? * [ ] :  and not blank.
Private Function CleanSheetName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", "?", "*", "[", "]", ":")
        s = Replace(s, ch, " ")
    Next ch
    s = Trim$(s)
    If Len(s) = 0 Then s = "New Sheet"
    CleanSheetName = Trim$(Left$(s, 31))
End Function

' "New Sheet" -> "New Sheet (2)" -> "New Sheet (3)" ... until the name is free.
Private Function UniqueSheetName(ByVal base As String) As String
    Dim nm As String, i As Long, suffix As String
    nm = base
    i = 2
    Do While SheetExists(nm)
        suffix = " (" & i & ")"
        nm = Trim$(Left$(base, 31 - Len(suffix))) & suffix
        i = i + 1
    Loop
    UniqueSheetName = nm
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = ThisWorkbook.Sheets(nm)
    SheetExists = Not sh Is Nothing
    On Error GoTo 0
End Function
