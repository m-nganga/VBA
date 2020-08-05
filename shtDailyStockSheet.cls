VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtDailyStockSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit ' All variables must be declared


Sub LastRow()
    Dim LastRow As Long
    Dim rng As Range

    ' Use all cells on the sheet
    Set rng = Sheets("Sheet1").Cells

    'Use a range on the sheet
    'Set rng = Sheets("Sheet1").Range("A1:D30")

    ' Find the last row
    LastRow = Last(1, rng)

    ' After the last row with data change the value of the cell in Column A
    rng.Parent.Cells(LastRow + 1, 1).Value = "Hi there"

End Sub


Sub LastColumn()
    Dim LastCol As Long
    Dim rng As Range

    ' Use all cells on the sheet
    Set rng = Sheets("Sheet1").Cells

    'Or use a range on the sheet
    'Set rng = Sheets("Sheet1").Range("A1:D30")

    ' Find the last column
    LastCol = Last(2, rng)

    ' After the last column with data change the value of the cell in row 1
    rng.Parent.Cells(1, LastCol + 1).Value = "Hi there"

End Sub


Sub LastCell()
    Dim LastCell As String
    Dim rng As Range

    ' Use all cells on the sheet
    Set rng = Sheets("Sheet1").Cells

    'Or use a range on the sheet
    'Set rng = Sheets("Sheet1").Range("A1:D30")

    ' Find the last cell
    LastCell = Last(3, rng)

    ' Select from A1 till the last cell in Rng
    With rng.Parent
        .Select
        .Range("A1", LastCell).Select
    End With
End Sub

Function Last(choice As Long, rng As Range)
'Ron de Bruin, 5 May 2008
' 1 = last row
' 2 = last column
' 3 = last cell
    Dim lrw As Long
    Dim lcol As Long

    Select Case choice

    Case 1:
        On Error Resume Next
        Last = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
        On Error GoTo 0

    Case 2:
        On Error Resume Next
        Last = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

    Case 3:
        On Error Resume Next
        lrw = rng.Find(What:="*", _
                       After:=rng.Cells(1), _
                       Lookat:=xlPart, _
                       LookIn:=xlFormulas, _
                       SearchOrder:=xlByRows, _
                       SearchDirection:=xlPrevious, _
                       MatchCase:=False).Row
        On Error GoTo 0

        On Error Resume Next
        lcol = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

        On Error Resume Next
        Last = rng.Parent.Cells(lrw, lcol).Address(False, False)
        If Err.Number > 0 Then
            Last = rng.Cells(1).Address(False, False)
            Err.Clear
        End If
        On Error GoTo 0

    End Select
End Function

Function LastRowColumn(sht As Worksheet, RowColumn As String) As Long
'PURPOSE: Function To Return the Last Row Or Column Number In the Active Spreadsheet
'INPUT: "R" or "C" to determine which direction to search

'This function would be called as : x = LastRowColumn(ActiveSheet, "Row")

Dim rc As Long

Select Case LCase(Left(RowColumn, 1)) 'If they put in 'row' or column instead of 'r' or 'c'.
  Case "c"
    LastRowColumn = sht.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, _
    SearchDirection:=xlPrevious).Column
  Case "r"
    LastRowColumn = sht.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, _
    SearchDirection:=xlPrevious).Row
  Case Else
    LastRowColumn = 1
End Select

End Function


Private Sub cmdEndofDay_Click()

Dim blnEOD As Boolean
Dim businessDate As String
Dim CurrentCol As Long
Dim i As Byte, j As Byte    'variables to control loops

Dim sht As Worksheet 'To be used for selecting the active worksheet
Dim LastColumn As Long  'Keep track of the last column of the worksheet
Dim LastRow As Integer, LastSummaryRow As Integer   'Help to find the last row in the active worksheet

Dim intOpeningStock() As Integer    'track the stock levels
Dim intStockLevel() As Integer   'the dynamic array to keep count of the products
Dim intQtyArray() As Integer     'dynamic 2D array to carry the day's purchased & sold quantities respectively
Dim curStockValue() As Currency  'dynamic 2D array to carry the day's purchased & sold values respectively
Dim DaySummaryValues(1 To 3) As Currency  'Array to keep daily purchases and sales values.

'Housekeeping... Turnoff the screen flashing when youare switching between active windows.
Application.ScreenUpdating = False


'Record todays Date
businessDate = Range("B3").Value
MsgBox "For debuging only..." & vbCrLf & "Today's business date is.. : " & businessDate ' To be deleted

'Check that all values entered for stock are in numbers and that they are not greater than available stock
'Ideally, this can also be done at dataentry using the cell change event.

'Code goes here


'Give a warning here to have updated all the deliveries, sales and expenses for the day in a message box

If MsgBox("Please ensure that you have updated all the sales, purchases and expenses for the day before proceeding" & vbCrLf & "Do You Want To Proceed", vbYesNo, "End of Day Process") = vbNo Then

    End

End If

'Check if EOD has been executed already
'.
'.
'.



'The next step is to get a summary of the Days business sales and deliveries ready for writing to the Stock Report


'PURPOSE: Update the stock levels after the days sales and deliveries, where
'stock level = Opening Stock Add Purchases Less Sales

    'Determine the last row of column A of this worksheet so as to know the size of array to create.
    
    Set sht = ActiveSheet
        
    LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
    ReDim intStockLevel(5 To LastRow)
    ReDim intQtyArray(5 To LastRow, 1 To 2)
    ReDim curStockValue(5 To LastRow, 1 To 2)
    ReDim intOpeningStock(5 To LastRow)
    
    'Loop through the rows to compute the stock levels while also reading the day's deliveries and sales data
    For i = 5 To LastRow
        intStockLevel(i) = sht.Cells(i, 5).Value + sht.Cells(i, 6).Value - sht.Cells(i, 9).Value
        intOpeningStock(i) = sht.Cells(i, 5).Value
        intQtyArray(i, 1) = sht.Cells(i, 6).Value
        intQtyArray(i, 2) = sht.Cells(i, 9).Value
        curStockValue(i, 1) = sht.Cells(i, 7).Value
        curStockValue(i, 2) = sht.Cells(i, 10).Value
    Next i
    
    'Read the sales and purchases numbers data for the day
    DaySummaryValues(1) = sht.Cells(i, 8).Value 'Purchases
    DaySummaryValues(2) = sht.Cells(i, 11).Value 'Sales
    DaySummaryValues(3) = sht.Cells(i, 16).Value 'Expenses
    
    
    'Activate the Inventory Worksheet and update the stock levels with the just computed values
    'these values become the current stock levels as well as the opening stock for the following business day.
    Set sht = ThisWorkbook.Worksheets("Inventory")
    sht.Activate
    
    For i = 5 To LastRow
            sht.Cells(i, 7).Value = intStockLevel(i)
    Next i
    
   
'PURPOSE : Update the Worksheet "Stock Report" with the day's stock status and BP/SP data.

    Set sht = ThisWorkbook.Worksheets("Stock Report")
    
    sht.Activate
    sht.Range("A3").Select
        
    'Determine the last column of row 3 with previous sales / purchases data
    'so as to determine the columns to record days data
    
    LastColumn = sht.Cells(3, sht.Columns.Count).End(xlToLeft).Column       'Ctrl + Shift + End
        
                'Using UsedRange
                  'sht.UsedRange 'Refresh UsedRange
                  'LastColumn = sht.UsedRange.Columns(sht.UsedRange.Columns.Count).Column
                
                'Using Table Range
                  'LastColumn = sht.ListObjects("Table1").Range.Columns.Count
                
                'Using Named Range
                  'LastColumn = sht.Range("MyNamedRange").Columns.Count
                
                'Ctrl + Shift + Right (Range should be first cell in data set)
                  'LastColumn = sht.Range("A1").CurrentRegion.Columns.Count
        
    'Set the date and the headers for the day's sales / purchases data
    
    For j = LastColumn + 1 To LastColumn + 5
        sht.Cells(3, j).Value = businessDate
         If j = LastColumn + 1 Then
             sht.Cells(4, j).Value = "Opening Stock"
             sht.Cells(4, j + 1).Value = "Purchases"
             sht.Cells(4, j + 2).Value = "Buying Price"
             sht.Cells(4, j + 3).Value = "Sales"
             sht.Cells(4, j + 4).Value = "Selling Price"
        End If
    Next j
        
    
    'Write the day's sales / purchases data by looping through the multidimensional arrays for quantity and values respecctively
    For i = 5 To LastRow
        j = LastColumn
        sht.Cells(i, j + 1).Value = intOpeningStock(i)
        sht.Cells(i, j + 2).Value = intQtyArray(i, 1)
        sht.Cells(i, j + 3).Value = curStockValue(i, 1)
        sht.Cells(i, j + 4).Value = intQtyArray(i, 2)
        sht.Cells(i, j + 5).Value = curStockValue(i, 2)
    Next i

    'Write the day's Summary data of Sales, Purchases and Gross Profit/Loss
    
    'Determine the last row in Col 5 so that we can start / continue writing data from the next row.
    LastSummaryRow = sht.Cells(sht.Rows.Count, "E").End(xlUp).Row
    
    sht.Cells(LastSummaryRow + 1, 5).Value = businessDate
    sht.Cells(LastSummaryRow + 1, 6).Value = DaySummaryValues(2)
    sht.Cells(LastSummaryRow + 1, 7).Value = DaySummaryValues(1)
    sht.Cells(LastSummaryRow + 1, 8).Value = DaySummaryValues(2) - DaySummaryValues(1) 'Gross Profit / Loss
    sht.Cells(LastSummaryRow + 1, 9).Value = DaySummaryValues(3)
    sht.Cells(LastSummaryRow + 1, 10).Value = sht.Cells(LastSummaryRow + 1, 8) - DaySummaryValues(3) ' Net Profit / Loss
    
    GeneratePivotTable
    
'PURPOSE : Reset the control variables, advance business date by 1d and write the opening stock of the NBD

    'Set the business date to NBD
    businessDate = DateAdd("d", 1, businessDate)
    MsgBox "For debuging only..." & vbCrLf & "Business date advanced to.. : " & businessDate ' To be deleted
    
    'Activate Worksheet "Parameters" for writing the next Business Day
    Set sht = ThisWorkbook.Worksheets("Parameters")
    
    sht.Activate
    sht.Cells(2, 7).Value = businessDate
    
    'take control back to Daily Stock Worksheet
    Set sht = ThisWorkbook.Worksheets("Daily Stock Sheet")
    sht.Activate
    
    'write the NBD opening stock which is the stock level of the previous day.
    'Also clear the quantities delivered and that sold for the new day.
    For i = 5 To LastRow
        sht.Cells(i, 5).Value = intStockLevel(i)
        sht.Cells(i, 6).Value = 0
        sht.Cells(i, 9).Value = 0
        sht.Cells(i, 14).Value = 0
    Next i
    
'Finally, save the worksheet and we are all good.
Workbooks("Bikisi Wines & Spirits - Stocks and Financial Records ver II.xlsm").Save

'Housekeeping... Turn back on screen flashing, just good practice but technically you do not have to.
Application.ScreenUpdating = True


End Sub

Sub GeneratePivotTable()
Dim PSheet As Worksheet 'Create a sheet for a new pivot table
Dim DSheet As Worksheet  'To use as a data sheet.
Dim PCache As PivotCache 'To use as a name for pivot table cache.
Dim PTable As PivotTable 'Name for the pivot table
Dim PRange As Range ' Define source data range
Dim LastRow As Long, LastCol As Long 'last row and column for the data range



Application.DisplayAlerts = False
On Error Resume Next
'If there exists a Worksheet by the name Summary Report, delete it
Worksheets("Summary Report").Delete
Sheets.Add After:=ActiveSheet
ActiveSheet.Name = "Summary Report"
Application.DisplayAlerts = True

Set PSheet = Worksheets("Summary Report")
Set DSheet = Worksheets("Stock Report")


'if the PTable is being inserted in the same worksheet then you need to check and delete any existing pivot table withthe same name.
    'Set PSheet = Worksheets("PivotTable")
    'Set DSheet = Worksheets(“Data”)
    '
    'Worksheets("PivotTable").Activate
    'On Error Resume Next
    'ActiveSheet.PivotTables("SalesPivotTable").TableRange2.Clear


'Define Data Range
LastRow = DSheet.Cells(Rows.Count, 5).End(xlUp).Row
'LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
LastCol = 10

'Set PRange = DSheet.Cells(149, 5).Resize(LastRow, LastCol)
Set PRange = DSheet.Range("E149", "J" & LastRow) 'Define the range of data source


'Define Pivot Cache
'Normally when you create a pivot table, Excel automatically creates a pivot cache
'without asking you, but when you need to use VBA, you need to write a code for this.

Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
TableName:="SalesPivotTable")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="SalesPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Date")
.Orientation = xlRowField
.Position = 1
End With


'Insert Column Fields
'
'With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Sales")
'.Orientation = xlColumnField
'.Position = 1
'End With



'Insert Data Field
With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Sales")
    .Orientation = xlDataField
    .Position = 1
    .Function = xlSum
    .NumberFormat = "#,##0.00"
    .Name = "Sales"
End With

With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Purchases")
    .Orientation = xlDataField
    .Position = 2
    .Function = xlSum
    .NumberFormat = "#,##0.00"
    .Name = "Purchases "
End With

With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Gross Profit / Loss")
    .Orientation = xlDataField
    .Position = 3
    .Function = xlSum
    .NumberFormat = "#,##0.00"
    .Name = "Gross Profit / Loss "
End With

With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Expenses")
    .Orientation = xlDataField
    .Position = 4
    .Function = xlSum
    .NumberFormat = "#,##0.00"
    .Name = "Expenses "
End With

With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Net Profit")
    .Orientation = xlDataField
    .Position = 5
    .Function = xlSum
    .NumberFormat = "#,##0.00"
    .Name = "Net Profit / Loss "
End With


'Format Pivot Table
ActiveSheet.PivotTables("SalesPivotTable").ShowTableStyleRowStripes = True
ActiveSheet.PivotTables("SalesPivotTable").TableStyle2 = "PivotStyleMedium9"

End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
MsgBox "worksheet change"

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
