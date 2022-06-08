Attribute VB_Name = "ASCP_VAL_BUTTONS"
Option Explicit

Dim ASCPExportWB As Workbook, ExportWS As Worksheet, OpenPOWS As Worksheet, RawPlannedPOWS As Worksheet, ActualPlannedPOWS As Worksheet
Dim PODataArray() As Variant, RawDataArray() As Variant, ExportDataArray() As Variant, Costfile As String

Sub Load_Data()

    '***Turn off updating to speed up the macro
    Call TurnOffStuff

    '***Start Timer
    Dim StartTime As Double
    StartTime = Timer
    
    Debug.Print "Establishing Workbook and Worksheets"
    '***Create ASCP Export Workbook
    Set ASCPExportWB = Workbooks.Add
    Set ExportWS = ASCPExportWB.Worksheets(1)
    Set OpenPOWS = ASCPExportWB.Worksheets.Add(After:=ExportWS)
    Set RawPlannedPOWS = ASCPExportWB.Worksheets.Add(After:=OpenPOWS)
    Set ActualPlannedPOWS = ASCPExportWB.Worksheets.Add(After:=RawPlannedPOWS)
    
    'Update sheet names
    ExportWS.Name = "ASCP DATA"
    OpenPOWS.Name = "OPEN_PO"
    RawPlannedPOWS.Name = "RAW_PLANNED_PO"
    ActualPlannedPOWS.Name = "ACTUAL PLANNED PO"

    '*** Load Data into Arrays for quicker computing
    Debug.Print "Loading Data into Arrays"
    
    '*****OPEN PO*****
    Dim POfile As String
    MsgBox "Please Select Open PO Report to Open", vbInformation, "OPEN PO File"
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = -1 Then 'if OK is pressed
            POfile = .SelectedItems(1)
        Else
            Exit Sub 'Exit if canceled
        End If
    End With
    
    Call Copy_Data(POfile, OpenPOWS)
    OpenPOWS.Range("S1").Value = "ADJ_WEEK"
    OpenPOWS.Range("T1").Value = "DAYS_CHANGED"
    OpenPOWS.Range("U1").Value = "FLOW"
    OpenPOWS.Range("V1").Value = "CASH_FORECAST"
    Call Format_Data(OpenPOWS)
    
    OpenPOWS.Range("T2").Formula = "=IF([@[ADJ_WEEK]]=" & Chr(34) & Chr(34) & "," & Chr(34) & "CANCEL" & Chr(34) & ",IF(ABS([@[ADJ_WEEK]]-[@[CURRENT_WEEK]])<=14,0,[@[ADJ_WEEK]]-[@[CURRENT_WEEK]]))"
    OpenPOWS.Range("U2").Formula = "=IF(NOT(ISBLANK([@[TRACKING_INFORMATION]]))," & Chr(34) & "TRACKING" & Chr(34) & ",IF(AND(NOT(ISBLANK([@[PROGRAM_NAME]])),NOT(ISNUMBER(SEARCH(" & Chr(34) & "AES" & Chr(34) & ",[@[PROGRAM_NAME]])))), " & Chr(34) & "PROGRAMS" & Chr(34) & ",IF([@ETA]<TODAY()+30," & Chr(34) & "IN TRANSIT" & Chr(34) & ",IF(ISNUMBER(SEARCH(" & Chr(34) & "TERM" & Chr(34) & ",[@NOTES]))," & Chr(34) & "TERM" & Chr(34) & ", IF([@[ADJ_WEEK]]=" & Chr(34) & Chr(34) & "," & Chr(34) & "CANCEL" & Chr(34) & "," & Chr(34) & "ADJUSTABLE PO" & Chr(34) & ")))))"
    OpenPOWS.Range("V2").Formula = "=IF(LEFT([@UPC],6)=" & Chr(34) & "719410" & Chr(34) & "," & Chr(34) & "5HR OPEN PO" & Chr(34) & "," & Chr(34) & "OPEN PO" & Chr(34) & ")"
    
    'Sort
    OpenPOWS.ListObjects("OPEN_PO").Sort.SortFields.Clear
    OpenPOWS.ListObjects("OPEN_PO").Sort.SortFields.Add2 Key:=Range("OPEN_PO[UPC]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    OpenPOWS.ListObjects("OPEN_PO").Sort.SortFields.Add2 Key:=Range("OPEN_PO[PO_RANK]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With OpenPOWS.ListObjects("OPEN_PO").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    OpenPOWS.Range("OPEN_PO[UPC]").NumberFormat = "00000000000"
    OpenPOWS.Range("OPEN_PO[ETA]").EntireColumn.AutoFit

    
    '*****PLANNED PO*****
    Dim Rawfile As String
    MsgBox "Please Select Raw Planned PO Report to Open", vbInformation, "DW PLANNED PO File"
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = -1 Then 'if OK is pressed
            Rawfile = .SelectedItems(1)
        Else
            Exit Sub 'Exit if canceled
        End If
    End With

    Call Copy_Data(Rawfile, RawPlannedPOWS)
    RawPlannedPOWS.Range("O1").Value = "ADJ_WEEK"
    RawPlannedPOWS.Range("P1").Value = "PLACEMENT_DATE"
    RawPlannedPOWS.Range("Q1").Value = "LATE_PLACEMENT"
    RawPlannedPOWS.Range("R1").Value = "CASH_FORECAST"
    
    Call Format_Data(RawPlannedPOWS)
    
    RawPlannedPOWS.Range("P2").Formula = "=IF([@[ADJ_WEEK]]=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",([@[ADJ_WEEK]]-ROUNDUP([@[LT]],0))-WEEKDAY(([@[ADJ_WEEK]]-ROUNDUP([@[LT]],0)),2)+1)"
    RawPlannedPOWS.Range("Q2").Formula = "=AND([@[ADJ_WEEK]]<>" & Chr(34) & Chr(34) & ",[@[PLACEMENT_DATE]]<TODAY())"
    RawPlannedPOWS.Range("R2").Formula = "=IF(LEFT([@[ITEM NAME]],6)=" & Chr(34) & "719410" & Chr(34) & "," & Chr(34) & "5HR ASCP PLAN" & Chr(34) & "," & Chr(34) & "ASCP PLAN" & Chr(34) & ")"

    RawPlannedPOWS.ListObjects("RAW_PLANNED_PO").Sort.SortFields.Clear
    RawPlannedPOWS.ListObjects("RAW_PLANNED_PO").Sort.SortFields.Add2 Key:=Range("RAW_PLANNED_PO[ITEM NAME]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    RawPlannedPOWS.ListObjects("RAW_PLANNED_PO").Sort.SortFields.Add2 Key:=Range("RAW_PLANNED_PO[WEEK]"), SortOn:=xlSortOnValues _
        , Order:=xlAscending, DataOption:=xlSortNormal
    RawPlannedPOWS.ListObjects("RAW_PLANNED_PO").Sort.SortFields.Add2 Key:=Range("RAW_PLANNED_PO[PLAN RANK]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With RawPlannedPOWS.ListObjects("RAW_PLANNED_PO").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    RawPlannedPOWS.Range("RAW_PLANNED_PO[ITEM NAME]").NumberFormat = "00000000000"
    RawPlannedPOWS.Range("RAW_PLANNED_PO[WEEK]").EntireColumn.AutoFit
    RawPlannedPOWS.Range("RAW_PLANNED_PO[PLACEMENT_DATE]").NumberFormat = "m/d/yyyy"
    
    '*****ASCP DATA*****
    Dim Expfile As String
    MsgBox "Please Select ASCP DATA to Open", vbInformation, "ASCP DATA File"
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = -1 Then 'if OK is pressed
            Expfile = .SelectedItems(1)
        Else
            Exit Sub 'Exit if canceled
        End If
    End With

    Call Copy_Data(Expfile, ExportWS)
    ExportWS.Range("P1").Value = "Running Inventory QTY"
    ExportWS.Range("Q1").Value = "PO Rank"
    ExportWS.Range("R1").Value = "PO Quantity"
    ExportWS.Range("S1").Value = "Plan Rank"
    ExportWS.Range("T1").Value = "Plan Quantity"
    ExportWS.Range("U1").Value = "Note"
    ExportWS.Range("V1").Value = "ACTUAL QTY"
    ExportWS.Range("W1").Value = "RUNNING BALANCE QTY"
    ExportWS.Range("X1").Value = "RUNING VALUE"
    ExportWS.Range("Y1").Value = "SS CHECK"
    
    Dim ASCPBotRow As Long
    ASCPBotRow = ExportWS.Cells(1, 1).End(xlDown).row
    ExportWS.Range("A1").CurrentRegion.AutoFilter
    ExportWS.AutoFilter.Sort.SortFields.Clear
    ExportWS.AutoFilter.Sort.SortFields.Add2 Key:=ExportWS.Range("F1:F" & ASCPBotRow), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ASCP DATA").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ExportWS.Range("B:B").NumberFormat = "00000000000"
    ExportWS.Range("A:A").EntireColumn.AutoFit
    
    '***Turn on updating
    Call TurnOnStuff

    'Determine how many seconds code took to run
    Dim MinutesElapsed As String
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    'Notify user in seconds
    MsgBox "This code ran successfully in " & MinutesElapsed & " minutes", vbInformation

End Sub

Private Sub Copy_Data(file As String, OutputWS As Worksheet) 'COMPLETE
'
'   This function will return the array the holds all PO data, both static and calculated
'
    '***Open selected file
    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Open(file)
    Set ws = wb.Worksheets(1)

    'Delete Index column
    ws.Range("A1").EntireColumn.Delete

    'Create Array and add Data
    Dim DataArray As Variant
    DataArray = ws.Range("A1").CurrentRegion

    OutputWS.Range("A1:" & Split(Cells(1, UBound(DataArray, 2)).Address, "$")(1) & UBound(DataArray, 1)).Value = DataArray

    'Close Workbook
    wb.Close SaveChanges:=False
End Sub

Private Sub Format_Data(OutputWS As Worksheet)
'
'   This Sub will create tables for easy formula manipulation
'
    Dim BotRow As Integer
    BotRow = OutputWS.Cells(1, 1).End(xlDown).row
    Sheets(OutputWS.Name).ListObjects.Add(xlSrcRange, OutputWS.Range("$A$1").CurrentRegion, , xlYes).Name = OutputWS.Name

End Sub

Sub Rebalance()

    '***Turn off updating to speed up the macro
    Call TurnOffStuff

    '***Start Timer
    Dim StartTime As Double
    StartTime = Timer
    
    'Calculate DEMAND_RANK, PO_COUNT, PLANNED_ORDER_COUNT on Export Worksheet
    Debug.Print "Calculating Fields. This is the heavy stuff"
    Call ExportCounts(ExportWS, RawPlannedPOWS, OpenPOWS)
    
    '***Turn on updating
    Call TurnOnStuff

    'Determine how many seconds code took to run
    Dim MinutesElapsed As String
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    'Notify user in seconds
    MsgBox "This code ran successfully in " & MinutesElapsed & " minutes", vbInformation

End Sub

Private Sub ExportCounts(ExportWS As Worksheet, RawPlannedPOWS As Worksheet, OpenPOWS As Worksheet)
'
'   This macro will rebalance the POs using data from ASCP and EBS
'
    Dim br As Long 'bottom row
    Dim r As Long 'current row
    Dim upc As Variant
    Dim po As Integer 'current po rank
    Dim pl As Integer 'current planned order rank
    Dim mPo As Integer 'open po count
    Dim mPl As Integer ' planned order count
    Dim riv As Variant 'Running Inventory Value
    Dim poBr As Integer
    Dim plBr As Long
    Dim poR As Integer
    Dim plR As Long
    Dim ch As Integer 'check variable for do loop
    Dim ct As Integer
    Dim uSs As Long, oSs As Long, vSs As Long
    
    'Turn Off Stuff
    Call TurnOffStuff
    
    br = ExportWS.Cells(ExportWS.Rows.Count, 1).End(xlUp).row
    poBr = OpenPOWS.Cells(OpenPOWS.Rows.Count, 1).End(xlUp).row
    plBr = RawPlannedPOWS.Cells(RawPlannedPOWS.Rows.Count, 1).End(xlUp).row

    For r = 2 To br
        If ExportWS.Cells(r, 2).Value <> ExportWS.Cells(r - 1, 2).Value Then 'If new UPC in ASCP DATA
            upc = ExportWS.Cells(r, 2).Value 'Get UPC
            mPo = ExportWS.Cells(r, 14).Value 'Get PO total count for this UPC
            mPl = ExportWS.Cells(r, 15).Value 'Get Planned Order total count for this UPC
            If mPo > 0 Then po = 1 Else po = 0
            If mPl > 0 Then pl = 1 Else pl = 0
    
            riv = ExportWS.Cells(r, 7).Value - ExportWS.Cells(r, 9).Value + ExportWS.Cells(r, 11).Value
            ch = 0
            ct = 0
    'If upc = "4168914508" Then 'Use this line of code to review a specific UPC in the macro
        'Debug.Print r
    'End If
            Do Until ch = 1
                If riv < ExportWS.Cells(r, 12).Value Then 'If running inventory drops below SS value
                    If po <> 0 And mPo >= po Then
                        For poR = 2 To poBr
                            If OpenPOWS.Cells(poR, 3).Value = upc And OpenPOWS.Cells(poR, 15).Value = po Then
                                ExportWS.Cells(r, 17).Value = po
                                ExportWS.Cells(r, 18).Value = OpenPOWS.Cells(poR, 17).Value + ExportWS.Cells(r, 18).Value
                                OpenPOWS.Cells(poR, 19).Value = ExportWS.Cells(r, 1).Value
                                riv = riv + OpenPOWS.Cells(poR, 17).Value
                                po = po + 1
                                poR = poBr
                            End If
                        Next
                    ElseIf pl <> 0 And mPl >= pl Then
                        For plR = 2 To plBr
                            If RawPlannedPOWS.Cells(plR, 2).Value = upc And RawPlannedPOWS.Cells(plR, 10).Value = pl Then
                                ExportWS.Cells(r, 19).Value = pl
                                ExportWS.Cells(r, 20).Value = RawPlannedPOWS.Cells(plR, 11).Value + ExportWS.Cells(r, 20).Value
                                RawPlannedPOWS.Cells(plR, 15).Value = ExportWS.Cells(r, 1).Value
                                riv = riv + RawPlannedPOWS.Cells(plR, 11).Value
                                pl = pl + 1
                                plR = plBr
                            End If
                        Next
                    Else
                        ExportWS.Cells(r, 21).Value = "NEW ORDER NEEDED"
                        ch = 1
                    End If
                Else
                    ch = 1
                End If
              ct = ct + 1
              If ct > 100 Then
                Debug.Print ct
              End If
            Loop
            ExportWS.Cells(r, 16).Value = riv
        Else 'If it is not a new UPC
            'WS Updated the value here
            riv = riv + ExportWS.Cells(r, 7).Value - ExportWS.Cells(r, 9).Value + ExportWS.Cells(r, 11).Value
            ch = 0
            ct = 0
            Do Until ch = 1
                If riv < ExportWS.Cells(r, 12).Value Then 'If running inventory drops below SS value
                    If po <> 0 And mPo >= po Then
                        For poR = 2 To poBr
                            If OpenPOWS.Cells(poR, 3).Value = upc And OpenPOWS.Cells(poR, 15).Value = po Then
                                ExportWS.Cells(r, 17).Value = po
                                ExportWS.Cells(r, 18).Value = OpenPOWS.Cells(poR, 17).Value + ExportWS.Cells(r, 18).Value
                                OpenPOWS.Cells(poR, 19).Value = ExportWS.Cells(r, 1).Value
                                riv = riv + OpenPOWS.Cells(poR, 17).Value
                                po = po + 1
                                poR = poBr
                            End If
                        Next
                    ElseIf pl <> 0 And mPl >= pl Then
                        For plR = 2 To plBr
                            If RawPlannedPOWS.Cells(plR, 2).Value = upc And RawPlannedPOWS.Cells(plR, 10).Value = pl Then
                                ExportWS.Cells(r, 19).Value = pl
                                ExportWS.Cells(r, 20).Value = RawPlannedPOWS.Cells(plR, 11).Value + ExportWS.Cells(r, 20).Value
                                RawPlannedPOWS.Cells(plR, 15).Value = ExportWS.Cells(r, 1).Value
                                riv = riv + RawPlannedPOWS.Cells(plR, 11).Value
                                pl = pl + 1
                                plR = plBr
                            End If
                        Next
                    Else
                        ExportWS.Cells(r, 21).Value = "NEW ORDER NEEDED"
                        ch = 1
                    End If
                Else
                    ch = 1
                End If
                ct = ct + 1
              If ct > 100 Then
                Debug.Print "Issue on " & upc
                MsgBox "ERROR, Please check UPC " & upc
                Exit Sub
              End If
            Loop
            ExportWS.Cells(r, 16).Value = riv
        End If
        
    If Int(r / 10000) = r / 10000 Then
        Debug.Print r
    End If
    
    
    Next
        
    'Turn On Stuff
    Call TurnOnStuff

End Sub
Sub Running_qty()
'
'   This sub will go through each line of the ASCP DATA after procurement has added their notes
'     and remove or change quantities based on additions and deletions from procurement
'
    Dim wb As Workbook, ExWS As Worksheet, PlannedPOWS As Worksheet
    If ExportWS Is Nothing Then
        ' need to initialize obj: '
        Dim file As String
        MsgBox "Please Select file to Open", vbInformation, "DATA File"
        With Application.FileDialog(msoFileDialogFilePicker)
            If .Show = -1 Then 'if OK is pressed
                file = .SelectedItems(1)
            Else
                Exit Sub 'Exit if canceled
            End If
        End With
        
        Set wb = Workbooks.Open(file)
        Set ExWS = wb.Worksheets(1)
        Set PlannedPOWS = wb.Worksheets(4)
    Else
        Set ExWS = ExportWS
        Set PlannedPOWS = ActualPlannedPOWS
    End If
    
    '***Turn off updating to speed up the macro
    Call TurnOffStuff

    '***Start Timer
    Dim StartTime As Double
    StartTime = Timer
    
    'Recalculate based on ACTUAL PLANNED PO worksheet data
    Debug.Print "Recalculating Fields..."
    
    Dim BotRow As Long, row As Long
    BotRow = ExWS.Cells(1, 1).End(xlDown).row
        ExWS.Range("V2").Formula = "=SUMIFS(Table3[ASCP_PLANNED_ORDER],Table3[ITEM NAME],'ASCP DATA'!B2,Table3[APPROVED_DATE],'ASCP DATA'!A2,Table3[NOTE]," & Chr(34) & "LEAVE" & Chr(34) & ")"
        ExWS.Range("V2").AutoFill Destination:=ExWS.Range("V2:V" & BotRow)
        ExWS.Range("W2").Formula = "=IF(B2<>B1,G2+J2+K2+V2-I2,W1+G2+J2+K2+V2-I2)"
        ExWS.Range("W2").AutoFill Destination:=ExWS.Range("W2:W" & BotRow)
        ExWS.Range("X2").Formula = "=W2*M2"
        ExWS.Range("X2").AutoFill Destination:=ExWS.Range("X2:X" & BotRow)
        ExWS.Range("Y2").Formula = "=W2<L2"
        ExWS.Range("Y2").AutoFill Destination:=ExWS.Range("Y2:Y" & BotRow)
        
    '***Turn on updating
    Call TurnOnStuff

    'Determine how many seconds code took to run
    Dim MinutesElapsed As String
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    'Notify user in seconds
    MsgBox "This code ran successfully in " & MinutesElapsed & " minutes", vbInformation

End Sub


Private Sub TurnOffStuff() 'COMPLETE
'
'   This function turns updating off to increase macro efficiency
'
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
End Sub
Private Sub TurnOnStuff() 'COMPLETE
'
'   This function turns updating on after macros end
'
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
