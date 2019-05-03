Attribute VB_Name = "Module10"
'This sub will put the name of the county in the 'County' column without the word 'County' as in the County1 Column

'Version 1
'4/30/2019
'Provided by Travis Keller
'www.FoxholeAcres.com
'For inquiries or special requests, please email Travis Keller at Travis@FoxholeAcres.com

Sub FillCountyName()
    Dim i As Integer
    Dim lRow As Integer
    Dim pos As Integer
    Dim res As String
    
    'find the last row with text in it
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'get the column numbers of the two county columns
    Set county1Col = Sheet1.Rows(1).Find(what:="County1", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=True, SearchFormat:=False)
    Set countyCol = Sheet1.Rows(1).Find(what:="County", LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=True, SearchFormat:=False)
    
    'fill in the county name on the county column
    For i = 2 To lRow
        pos = InStr(1, Cells(i, county1Col.Column).Value, "County,")
        Cells(i, countyCol.Column).Value = Left(Cells(i, county1Col.Column).Value, pos - 2)
    Next i
End Sub
