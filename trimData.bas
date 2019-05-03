Attribute VB_Name = "Module4"
Sub TrimData()
'*****Capitalize States
Dim i As Integer
Dim lRow As Integer

'get last row of text
lRow = Cells(Rows.Count, 2).End(xlUp).Row

Set stateCol = Sheets(1).Rows(1).Find(what:="Mail_State")

For i = 2 To lRow
    'Make uppercase
   Cells(i, stateCol.Column).Value = UCase(Cells(i, stateCol.Column).Value)
   'Delete Row if cell empty
   If Cells(i, stateCol.Column).Value = "" Then
        Rows(i).EntireRow.Delete
    End If
Next i

'*****Delete empty mail zip
Set zipCol = Sheets(1).Rows(1).Find(what:="Mail_ZipZip4")
For i = 2 To lRow
    If Cells(i, zipCol.Column).Value = "" Then
        Rows(i).EntireRow.Delete
    ElseIf Len(Cells(i, zipCol.Column).Value) < 5 Then
        Rows(i).EntireRow.Delete
    End If


End Sub
