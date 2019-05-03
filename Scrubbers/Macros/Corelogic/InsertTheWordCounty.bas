Attribute VB_Name = "Module7"
'This module inserts the word "County" after the county name and before the comma in the "County1" Column

'Version 1
'4/30/2019
'Provided by Travis Keller
'www.FoxholeAcres.com
'For inquiries or special requests, please email Travis Keller at Travis@FoxholeAcres.com

Sub InsertTheWordCounty()

Dim Result() As String
Dim i As Integer
Dim lRow As Integer

'find the last row with text in it
lRow = Cells(Rows.Count, 1).End(xlUp).Row

'find the column with the county1 header
    Set aCell = Sheet1.Rows(1).Find(what:="County1")
'Insert the word "County" after the county name and before the comma
   For i = 2 To lRow
        Cells(i, aCell.Column).Value = Replace(Cells(i, aCell.Column).Value, ",", " County,")
    Next i
End Sub
