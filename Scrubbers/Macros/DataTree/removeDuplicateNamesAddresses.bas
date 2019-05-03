Attribute VB_Name = "Module2"
'This sub removes duplicate mail addresses from the "Mail Address Full" column

'Version 1
'5/2/2019
'Provided by Travis Keller
'www.FoxholeAcres.com
'For inquiries or special requests, please email Travis Keller at Travis@FoxholeAcres.com

Sub RemoveDuplicateNamesAddresses()
    Dim aCell As Range
    Dim i As Integer
    Dim lRow As Integer
    
    lRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    Set aCell = Sheet1.Rows(1).Find(what:="OWNERS (ALL)") ', LookIn:=xlValues, _
    'LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
   ' MatchCase:=False, SearchFormat:=False)
    'Remove the duplicate owners
    Cells.RemoveDuplicates Columns:=Array(aCell.Column)
    
    'Get ready to concatenate address info
    Set street = Sheet1.Rows(1).Find(what:="Mail_Street")
    Set city = Sheet1.Rows(1).Find(what:="Mail_City")
    Set sutate = Sheet1.Rows(1).Find(what:="Mail_State")
    Set zip = Sheet1.Rows(1).Find(what:="Mail_ZipZip4")
    'Make a temp column at column z
    Columns("Z").Insert
    Cells(1, "Z").Value = "temp"
    'Concat the address info in each cell of the temp column
    For i = 2 To lRow
        Cells(i, "Z").Value = Cells(i, street.Column).Value _
        & Cells(i, city.Column).Value & Cells(i, sutate.Column) _
        & Cells(i, zip.Column).Value
        If Cells(i, "Z").Value = "" Then
            Rows(i).EntireRow.Delete
        End If
    Next i
    'Remove duplicates from the temp column
    Cells.RemoveDuplicates Columns:=Array(26)
    
    'Now delete the temp column
    Columns("Z").EntireColumn.Delete
End Sub



