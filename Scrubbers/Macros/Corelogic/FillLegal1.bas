Attribute VB_Name = "Module8"
'This module will fill in the legal description by concatenating the lot size and county name
'Note: Currently only works for acreage

'Version 1
'4/30/2019
'Provided by Travis Keller
'www.FoxholeAcres.com
'For inquiries or special requests, please email Travis Keller at Travis@FoxholeAcres.com

Sub FillLegal1()
Dim i As Integer
Dim lRow As Integer


'find the last row with text in it
lRow = Cells(Rows.Count, 1).End(xlUp).Row

'get the columns for acreage, county1, and legal1 based on their headers
Set acreCol = Sheet1.Rows(1).Find(what:="Lot_Acreage")
Set countyCol = Sheet1.Rows(1).Find(what:="County1")
Set legalCol = Sheet1.Rows(1).Find(what:="Legal1")

'fill Legal1 with the information from these other two colum cells
    For i = 2 To lRow
        Cells(i, legalCol.Column).Value = Cells(i, acreCol.Column).Value & " acre(s) in " & Cells(i, countyCol.Column).Value
    Next i
End Sub
