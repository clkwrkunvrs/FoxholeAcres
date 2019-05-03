Attribute VB_Name = "Module3"
'This sub will remove any entries that are not in the U.S. or Canada based on what's in the "Mail State" column

'Version 1
'4/30/2019
'Provided by Travis Keller
'www.FoxholeAcres.com
'For inquiries or special requests, please email Travis Keller at Travis@FoxholeAcres.com

Sub RemoveInternationalAddresses()
    Dim aCell As Range
    Dim states As Variant
    Dim i As Integer
    Dim j As Integer
    Dim found As Boolean
    Dim lRow As Integer
    
    'figure out what the last row with text in it is
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    found = False
    states = Array("Al", "Ak", "Az", "Ar", "Ca", "Co", "Ct", "De", "Fl", "Ga", "Hi", "Id", "Il", "In", "Ia", _
                "Ks", "Ky", "La", "Me", "Md", "Ma", "Mi", "Mn", "Ms", "Mo", "Mt", "Ne", "Nv", "Nh", "Nj", "Nm", "Ny", _
                "Nc", "Nd", "Oh", "Ok", "Or", "Pa", "Ri", "Sc", "Sd", "Tn", "Tx", "Ut", "Vt", "Va", "Wa", "Wv", "Wi", "Wy", "Canada")
    
    'Find State Column
    Set aCell = Sheet1.Rows(1).Find(what:=("Mail State"))
    If aCell Is Nothing Then
      Set aCell = Sheet1.Rows(1).Find(what:=("Mail_State"))
    End If
    On Error GoTo Err
    
      For i = 2 To lRow
        found = False 'reset found variable
        For j = 0 To (UBound(states) - LBound(states)) 'iterate to end of array
            If Cells(i, aCell.Column).Value = states(j) Then 'compare cell with states
                found = True 'if you find a match, exit the loop
                Exit For
            End If
        Next j 'End state check loop
        If found = False Then Rows(i).Delete 'if it was not found, delete that row
    Next i 'end cell iteration loop

Done:
    Exit Sub
    
Err:
    msgBox "The header name ""Mail State"" was not found"
     
End Sub

