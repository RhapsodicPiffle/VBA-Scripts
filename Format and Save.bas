Attribute VB_Name = "OpenPOFormatWithASN"
Sub OpenPOReportFormat()
'
' OpenPOReport Macro
'
' Keyboard Shortcut: Ctrl+Shift+R
'

'***Turn off updating to speed up the macro
Call TurnOffStuff

'''Set Bottom row of Report
Dim LastRow As Integer
LastRow = Cells(3, 2).End(xlDown).row

'***Get current workbook
Dim file As Workbook
Set file = ActiveWorkbook

'''CLEAN DATA
    ''CHANGE NAMES
        'Change Header Names
    Range("A1").Value = "NEED"
    Range("D1").Value = "Ln"
    Range("R1").Value = "CREATE"
    Range("S1").Value = "ORDER"
    Range("T1").Value = "IN TRANSIT"
    Range("U1").Value = "RECEIVED"
    Range("V1").Value = "BALANCE"
    Range("W1").Value = "PRICE"
    Range("X1").Value = "POSITION"

'''FORMAT DATA
    ''MOVE COLUMNS
        'Move PO and Ln column to the left
    Columns("C:D").Cut
    Columns("A:A").Insert Shift:=xlToRight
    
        'Move ASN and LPN columns to the right
    Columns("K:M").Cut
    Columns("Y:Y").Insert Shift:=xlToRight
    
        'Move Category column
    Columns("H:H").Cut
    Columns("C:C").Insert Shift:=xlToRight
    
        'Move Item and Description columns to the left after PO column
    Columns("I:J").Cut
    Columns("D:D").Insert Shift:=xlToRight
    
        'Move Vendor column after Description column
    Columns("I:I").Cut
    Columns("F:F").Insert Shift:=xlToRight
    
        'Move Program column right of Vendor
    Columns("M:M").Cut
    Columns("G:G").Insert Shift:=xlToRight
    
        'Move Create Date
    Columns("O:O").Cut
    Columns("H:H").Insert Shift:=xlToRight
        
        'Move Required Date
'    Columns("J:J").Cut
'    Columns("I:I").Insert Shift:=xlToRight
        
        'Move Need Date
    ' Columns("J:J").Cut
    ' Columns("J:J").Insert Shift:=xlToRight
        
        'Move Buyer column
    Columns("K:K").Cut
    Columns("H:H").Insert Shift:=xlToRight
    
        'Move Order and Price columns
    Columns("P:U").Cut
    Columns("K:K").Insert Shift:=xlToRight
    
        'Move Item Status columns
    Columns("R:R").Cut
    Columns("V:V").Insert Shift:=xlToRight
    
        'Move ORG columns
    Columns("Q:Q").Cut
    Columns("V:V").Insert Shift:=xlToRight
    
        'Move Container date columns
'    Columns("S:S").Cut
'    Columns("V:V").Insert Shift:=xlToRight
    
        'Move ASN and LPN columns to the left
    Columns("V:X").Cut
    Columns("Q:Q").Insert Shift:=xlToRight
    
    ''TABLE HEADERS
    With Rows("1:1")
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    ''COLUMN WIDTH
        'AutoFit width of columns
    Cells.Columns.AutoFit

    ''HIDE COLUMNS
        'Category
    Columns("C:C").EntireColumn.Hidden = True
        'Buyer and Dates
    Columns("H:I").EntireColumn.Hidden = True
        'Price and Position
    Columns("O:P").EntireColumn.Hidden = True
    
        'Format column PO
    With Columns("A:A")
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
        'Format column Item
    With Columns("D:D")
        .NumberFormat = "00000000000"
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Columns("E:E").ColumnWidth = 41 'Description
    Columns("F:F").ColumnWidth = 34 'Vendor
    Columns("G:G").ColumnWidth = 46 'Program
    Columns("T:T").ColumnWidth = 59 'Tracking
    
    'Freeze Panes
    Range("F2").Select
    ActiveWindow.FreezePanes = True
       
    'Zoom to 85%
    ActiveWindow.Zoom = 85
    
    'Add AutoFilter
    Range("A1").AutoFilter
    
    'Change Tab name
    Sheets(1).Name = "Data"
    
    '***Turn on updating
    Call TurnOnStuff
    
    Call Save_File(file)
    
End Sub
Private Sub TurnOffStuff()
'
'   This function turns updating off to increase macro efficiency
'
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
End Sub
Private Sub TurnOnStuff()
'
'   This function turns updating on after macros end
'
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

Private Sub Save_File(file As Workbook)
'
'   This function saves the Open PO File to the month folder
'
    '***Create Directory
    Dim currentFolder As String, saveFolder As String
    currentFolder = "X:\Pugs Legacy X\UTAH\PROCUREMENT\PURCHASING\Open PO Report"
    saveFolder = CreateDirectory(currentFolder)

    Dim Title As String, count As Integer, check As Boolean
    count = 0
    check = False

    Do Until check = True
        Title = saveFolder & "\Open PO Report " & Format(Date, "mm-dd-yy") & " (" & count & ")"
        If Len(Dir(Title & ".xlsx", vbDirectory)) = 0 Then
            file.SaveAs fileName:=Title, FileFormat:=51
            'file.Close SaveChanges:=False
            check = True
        End If
        count = count + 1
        Debug.Print "count: " & count
    Loop

End Sub

Private Function CreateDirectory(currentFolder As String)
'
' This function creates a directory path for the Open PO Report
'
    Dim year, month
    Dim yearFolder As String, monthFolder As String
    year = Format(Date, "yyyy")
    month = Format(Date, "mmmm")
    
    'Checks if month folder exists
    monthFolder = currentFolder & "\" & year & " " & month
    If Len(Dir(currentFolder & "\" & year & " " & month, vbDirectory)) = 0 Then
        MkDir monthFolder 'Make the folder if it doesn't exist
    End If
    
    CreateDirectory = monthFolder
    'Debug.Print monthFolder
    
End Function

