Attribute VB_Name = "Module9"
'This sub will prompt the user to decide what their control number will start at and then fils in the control number incremeting by one for each cell

'Version 1
'4/30/2019
'Provided by Travis Keller
'www.FoxholeAcres.com
'For inquiries or special requests, please email Travis Keller at Travis@FoxholeAcres.com

Sub FillControlNums()
    Dim i As Integer
    Dim start As Variant
    Dim lRow As Integer
    

'find the last row with text in it
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Set controlCol = Sheet1.Rows(1).Find(what:="Control")
    start = InputBox("What number should control numbers start at?") - 1
    
    For i = 2 To lRow
        start = start + 1
        Cells(i, controlCol.Column).Value = start
    Next i
    

End Sub
