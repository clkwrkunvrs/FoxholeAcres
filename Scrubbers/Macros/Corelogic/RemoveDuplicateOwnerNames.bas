Attribute VB_Name = "Module2"
'This sub removes duplicates in the "Owner Name (First Name First)" Column

'Version 1
'4/30/2019
'Provided by Travis Keller
'www.FoxholeAcres.com
'For inquiries or special requests, please email Travis Keller at Travis@FoxholeAcres.com

Sub RemoveDuplicateOwnerNamesFNF()
    Dim aCell As Range
    
    'Get the appropriate Column number
    Set aCell = Sheet1.Rows(1).Find(what:="Owner Name (First Name First)")
    If aCell Is Nothing Then
        aCell = Sheet1.Rows(1).Find(what:="Owner_Name")
    End If
    ', LookIn:=xlValues, _
    'LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
   ' MatchCase:=False, SearchFormat:=False)

    Cells.RemoveDuplicates Columns:=Array(aCell.Column)
        
End Sub
