Attribute VB_Name = "Module6"
'This sub will first pull all of the formatted columns to the left side of the spreadsheet
'It will then insert new desired fields to the right of the already present ones

'Version 1
'4/30/2019
'Provided by Travis Keller
'www.FoxholeAcres.com
'For inquiries or special requests, please email Travis Keller at Travis@FoxholeAcres.com

Sub InsertAndShiftColumns()

Const ZIP As String = "Mail_ZIP_ZIP_4"
Const STATE As String = "Mail_State"
Const CITY As String = "Mail_City"
Const ACREAGE As String = "Lot_Acreage"
Const ADDRESS As String = "Mail_Address"
Const NAME As String = "Owner_Name"
Const COUNTY As String = "County1"
Const APN As String = "APN"

Dim headerName As Variant
Dim i As Integer

'headerName = Array(NAME, APN, ADDRESS, CITY, STATE, ZIP, ACREAGE, COUNTY)
headerName = Array(COUNTY, ACREAGE, ZIP, STATE, CITY, ADDRESS, APN, NAME)

'Insert Columns to the left of Column A
    For i = 0 To 7
        Set aCell = Sheet1.Rows(1).Find(what:=headerName(i), LookIn:=xlValues, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=True, SearchFormat:=True)
        Columns(aCell.Column).Cut
        Columns("A").Insert
    Next i

'Insert new columns
    Columns("I:M").Insert Shift:=xlToRight, _
      CopyOrigin:=xlFormaFromRightOrBelow 'or xlFormatFromRightOrBelow
      
'Color new columns yellow
    For i = 1 To 13
      Cells(1, i).Interior.ColorIndex = 27
    Next i

'Name the new columns
    Cells(1, "I").Value = "Offer_Price"
    Cells(1, "J").Value = "Mailing_Status"
    Cells(1, "K").Value = "Control"
    Cells(1, "L").Value = "County"
    Cells(1, "M").Value = "Legal1"
    
'Swap Old APN column with APN Column
    'Columns("N").Cut
    'Columns("B").Insert
    
    'Columns("C").Cut
    'Columns("O").Insert
    
    'Cells(1, "B").Interior.ColorIndex = 27
    'Cells(1, "N").Interior.ColorIndex = 0
    
    
End Sub

