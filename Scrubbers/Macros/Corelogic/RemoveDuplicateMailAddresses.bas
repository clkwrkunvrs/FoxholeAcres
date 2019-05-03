Attribute VB_Name = "Module1"
'This sub removes duplicate mail addresses from the "Mail Address Full" column

'Version 1
'4/30/2019
'Provided by Travis Keller
'www.FoxholeAcres.com
'For inquiries or special requests, please email Travis Keller at Travis@FoxholeAcres.com

Sub RemoveDuplicateMailAddresses()
    Dim aCell As Range
    
    Set aCell = Sheet1.Rows(1).Find(what:="Mail Address Full") ', LookIn:=xlValues, _
    'LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
   ' MatchCase:=False, SearchFormat:=False)

    Cells.RemoveDuplicates Columns:=Array(aCell.Column)
    
    
    
End Sub


