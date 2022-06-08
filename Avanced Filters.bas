Attribute VB_Name = "AdvFilterVendors"
Option Explicit
Sub AdvFilterTest()
'
'   This function is the master outline for the vendor PO Report
'
    'Set up objects and variables from the Vendor PO Template worksheet
    Dim CriteriaWB As Workbook, CriteriaWS As Worksheet
    Set CriteriaWB = ActiveWorkbook
    Set CriteriaWS = ActiveWorkbook.Worksheets("Template")
    Dim early As Integer, late As Integer
    early = CriteriaWS.Range("M7").CurrentRegion
    late = CriteriaWS.Range("O7").CurrentRegion
    
    'Select the Open PO Report of your choice
    MsgBox "Please Select Open PO to Open", vbInformation, "OPEN PO FILE"
    Dim file As String
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = -1 Then 'if OK is pressed
            file = .SelectedItems(1)
        Else
            Exit Sub 'Exit if canceled
        End If
    End With
    
    '***Turn off updating to speed up the macro
    Call TurnOffStuff

    '***Start Timer
    Dim StartTime As Double
    StartTime = Timer

    '***Open selected file
    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Open(file)
    Set ws = wb.Worksheets(1)
    
    '***Create Directory
    Dim currentFolder As String, timeFolder As String
    currentFolder = CriteriaWB.Path
    timeFolder = CreateDirectory(currentFolder)
    
    '***Get list of unique buyers
    Dim BuyersList As Dictionary
    Set BuyersList = ReadBuyerData(ws)
    
    '***Get list of unique vendors
    Dim VendorsList As Dictionary
    Set VendorsList = ReadVendorData(ws)
    
    '***Get vendor data
    Dim vendorName As String, vendor As Variant, Buyer As Variant, col As Integer
    Dim vendorSize As Integer, buyerSize As Integer
    vendorSize = VendorsList.count
    buyerSize = BuyersList.count + 8
    Dim SummaryArray() As Variant
    ReDim SummaryArray(vendorSize, buyerSize)
    col = 1

    'Set up SummaryArray headers
    SummaryArray(0, 0) = "Vendor Name"
    For Each Buyer In BuyersList
        SummaryArray(0, col) = Buyer
        col = col + 1
    Next Buyer
    SummaryArray(0, col) = "Total Lines"
    col = col + 1
    SummaryArray(0, col) = "Total Value"
    col = col + 1
    SummaryArray(0, col) = "Past Due Lines"
    col = col + 1
    SummaryArray(0, col) = "Unconfirmed Lines"
    col = col + 1
    SummaryArray(0, col) = "Early Lines"
    col = col + 1
    SummaryArray(0, col) = "Late Lines"
    col = col + 1
    SummaryArray(0, col) = "Assigned Buyer"

    Dim row As Integer
    row = 1
    For Each vendor In VendorsList.Keys
        
        'Create new worksheet for each vendor
        vendorName = vendor
        Dim NewWB As Workbook, NewWS As Worksheet
        Set NewWB = Workbooks.Add
        Set NewWS = NewWB.Worksheets(1)

        Call AdvancedFilterCopy(CriteriaWS, ws, NewWS, vendorName)
        
        Dim SummaryDictionary As New Dictionary
        Set SummaryDictionary = SummaryProcess(NewWS, BuyersList, vendorName, early, late)

        'Populate the SummaryArray
        Dim var As Variant, count As Integer
        count = 0
        For Each var In SummaryDictionary
            'Go through the dictionary and add that data to the array
            SummaryArray(row, count) = SummaryDictionary(var)
            'If count = 9 Then Debug.Print "totalValue", SummaryDictionary(var)
            count = count + 1
        Next var
        row = row + 1

        'Save Vendor File
        Call SaveFile(NewWB, vendorName, timeFolder, True)
    Next vendor

    '***Create Summary File
    Dim SummaryWB As Workbook, SummaryWS As Worksheet, SummaryRange As Range
    Set SummaryWB = Workbooks.Add
    Set SummaryWS = SummaryWB.Worksheets(1)
    Set SummaryRange = SummaryWS.Range("A1" & ":" & Split(Cells(1, UBound(SummaryArray, 2)).Address, "$")(1) & UBound(SummaryArray, 1))
    SummaryRange.Value = SummaryArray
    
    '***Save Summary File
    SummaryWS.Range("B1:" & Split(Cells(1, col).Address, "$")(1) & "1").EntireColumn.NumberFormat = "#,###"
    SummaryWS.Range("J1").EntireColumn.NumberFormat = "$#,##0.00"
    SummaryWS.Range("A1").CurrentRegion.HorizontalAlignment = xlCenter
    SummaryWS.Range("A1").CurrentRegion.Columns.AutoFit
    SummaryWS.Range("A1").CurrentRegion.Sort Key1:=Range("I1"), Order1:=xlDescending, Header:=xlYes
    Call SaveFile(SummaryWB, "Summary", timeFolder, False)
    wb.Close SaveChanges:=False
    
    '***Turn on updating
    Call TurnOnStuff
    
    'Determine how many seconds code took to run
    Dim MinutesElapsed As String
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    'Notify user in seconds
    MsgBox "This code ran successfully in " & MinutesElapsed & " minutes", vbInformation

    Call EmailBuyers(timeFolder, SummaryWB)

End Sub
Sub AdvancedFilterCopy(CriteriaWS As Worksheet, Data As Worksheet, NewWS As Worksheet, vendorName As String)
'
'   This function copies vendor data to a new worksheet using advanced filter
'
    Dim headers As Range
    CriteriaWS.Range("E2").Value = vendorName 'Set filter criteria
    Set headers = CriteriaWS.Range("A1").CurrentRegion
    
    'Copy headers over for filter
    headers.Copy
    NewWS.Range("A1").PasteSpecial xlPasteAll
    
    'Set ranges
    Dim rgData As Range, rgCriteria As Range, rgOutput As Range
    Set rgData = Data.Range("A3").CurrentRegion
    Set rgCriteria = CriteriaWS.Range("A1").CurrentRegion
    Set rgOutput = NewWS.Range("A1:S1")
    
    'Do the filter
    rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgOutput
    
    'Format the file
    NewWS.Range("A1").Value = "PO #"
    NewWS.Range("B1").Value = "LINE #"
    NewWS.Range("C1").Value = "UPC"
    NewWS.Range("D1").Value = "DESCRIPTION"
    NewWS.Range("E1").Value = "VENDOR"
    NewWS.Range("F1").Value = "CREATE DATE"
    NewWS.Range("G1").Value = "INITIAL CONFIRMATION"
    NewWS.Range("H1").Value = "CURRENT ETA"
    NewWS.Range("I1").Value = "ORDER QUANTITY"
    NewWS.Range("J1").Value = "RECEIVED QUANTITY"
    NewWS.Range("K1").Value = "REMAINING QUANTITY"
    NewWS.Range("L1").Value = "UNIT PRICE"
    NewWS.Range("L1").EntireColumn.NumberFormat = "_($* #,##0.00_)"
    NewWS.Range("M1").Value = "EXTENDED PRICE"
    NewWS.Range("M1").EntireColumn.NumberFormat = "_($* #,##0.00_)"
    NewWS.Range("N1").Value = "CARRIER"
    NewWS.Range("O1").Value = "CONTAINER"
    NewWS.Range("P1").Value = "TRACKING"
    NewWS.Range("Q1").Value = "ORG"
    NewWS.Range("R1").Value = "BUYER"
    NewWS.Range("S1").Value = "SRP COMMENT"
    With NewWS.Range("T1")
        .Value = "VENDOR COMMENT"
        .Font.Bold = True
        .Font.Underline = True
    End With
    'NewWS.Range("R1").EntireColumn.Delete 'Delete Buyer column
    NewWS.Range("A1").CurrentRegion.Columns.AutoFit
    NewWS.Range("A1").CurrentRegion.Replace What:="CAVS", Replacement:="VADC"
End Sub
Private Function SummaryProcess(NewWS As Worksheet, buyerDict As Dictionary, vendorName As String, early As Integer, late As Integer) As Dictionary
'
' This function summarizes all of the vendor information and puts it into a dictionary
'
    'Place vendor data into an array
    Dim vendorArray As Variant
    vendorArray = NewWS.Range("A1").CurrentRegion
    
    'Establishes summary variables to return
    Dim vendorLines As Long, totalValue As Double
    Dim i As Integer
    Dim newBuyer As Buyer, buyerID As String
    Dim requiredDate As Date, ETA As Date
    Dim PastDue As Integer, earlyLines As Integer, lateLines As Integer, unconfirmed As Integer
    PastDue = 0
    earlyLines = 0
    lateLines = 0
    unconfirmed = 0
    totalValue = 0#
    'Creates a new buyer dictionary to count buyer lines per vendor
    Dim dict As New Dictionary, key As Variant
    For Each key In buyerDict.Keys
        dict.Add key, 0
    Next key

    '***Loop through vendor data to summarize
    For i = LBound(vendorArray, 1) + 1 To UBound(vendorArray, 1) 'Loop through vendor lines
        
        'Count the number of lines per buyer
        buyerID = vendorArray(i, 18)
        If dict.Exists(buyerID) = True Then
            dict(buyerID) = dict(buyerID) + 1
        Else 'If the buyer does not exist in the dictionary, add it
            Set newBuyer = New Buyer
            dict.Add buyerID, newBuyer
        End If
        
        totalValue = totalValue + vendorArray(i, 13)
        'Debug.Print "totalValue", totalValue
        
        'Calculate total Past Due lines
        requiredDate = vendorArray(i, 8)
        ETA = vendorArray(i, 9)
        If ETA < Date Then
            PastDue = PastDue + 1
        End If
        
        'Calculate Unconfirmed lines
        If Not IsDate(requiredDate) Or requiredDate = 0 Then
            unconfirmed = unconfirmed + 1
        End If
        
        'Calculate Early Lines
        If ETA < requiredDate - early Then
            earlyLines = earlyLines + 1
        End If
        
        'Calculate Late Lines
        If Not IsDate(requiredDate) Or requiredDate = 0 Then
            lateLines = lateLines
        ElseIf ETA > requiredDate + late Then
            lateLines = lateLines + 1
        End If
    Next i

    'Calculate totalLines
    vendorLines = UBound(vendorArray, 1) - LBound(vendorArray, 1)

    'Assign buyer based on max buyer line count
    Dim max As Integer, assignedBuyer As String
    max = Application.max(dict.Items)
    For Each key In dict
        If dict(key) = max Then
            assignedBuyer = key
            'Debug.Print "max= " & dict(key) & vbTab & "key= " & key
        End If
    Next key
    
    'Populate dictionary for return
    Dim DataDict As New Dictionary
    DataDict.Add "vendor", vendorName
    For Each key In dict
        DataDict.Add key, dict(key)
    Next key
    DataDict.Add "totalLines", vendorLines
    DataDict.Add "totalValue", totalValue
    DataDict.Add "pastDue", PastDue
    DataDict.Add "unconfirmed", unconfirmed
    DataDict.Add "earlyLines", earlyLines
    DataDict.Add "lateLines", lateLines
    DataDict.Add "assignedBuyer", assignedBuyer
    
    Set SummaryProcess = DataDict

End Function
Private Function ReadVendorData(ws As Worksheet) As Dictionary
'
'   This function will create a dictionary of vendors where key is
'   vendor name and value is line count
'
    Dim dict As New Dictionary
    Dim BotRow As Integer

    BotRow = ws.Cells(3, 2).End(xlDown).row
    
    Dim arr As Variant
    Set arr = ws.Range("G4:G" & BotRow) 'This is the column of Vendors
    
    Dim i As Variant, vendorID As String, count As Long
    Dim newVendor As Buyer
    
    For i = 1 To BotRow - 3 'Loop through the rows of vendor lines
        vendorID = arr(i)
        
        'If the buyer exists, grab that vendor object
        If dict.Exists(vendorID) = True Then
            Set newVendor = dict(vendorID)
        Else 'If the buyer does not exist in the dictionary, add it
            Set newVendor = New Buyer
            dict.Add vendorID, newVendor
        End If
        
        newVendor.lineCount = newVendor.lineCount + 1
        
    Next i
    
    'Return the dictionary of vendors
    Set ReadVendorData = dict

End Function
Private Function ReadBuyerData(ws As Worksheet) As Dictionary
'
'   This function will create a dictionary of buyers where key is
'   buyer name and value is line count
'
    Dim dict As New Dictionary
    Dim BotRow As Integer
    
    BotRow = ws.Cells(3, 2).End(xlDown).row
    
    Dim arr As Variant
    Set arr = ws.Range("I4:I" & BotRow) 'This is the column of Buyers
    
    Dim i As Variant, buyerID As String, count As Long
    Dim newBuyer As Buyer
    
    For i = 1 To BotRow - 3 'Loop through the rows of buyer lines
        buyerID = arr(i)
        
        'If the buyer exists, grab that buyer object
        If dict.Exists(buyerID) = True Then
            Set newBuyer = dict(buyerID)
        Else 'If the buyer does not exist in the dictionary, add it
            Set newBuyer = New Buyer
            dict.Add buyerID, newBuyer
        End If
        
        newBuyer.lineCount = newBuyer.lineCount + 1
        
    Next i
    
    'Return the dictionary of buyers
    Set ReadBuyerData = dict

End Function
Private Function CreateDirectory(currentFolder As String)
'
' This function creates a directory path for the Vendor PO Project
'
    Dim year, month
    Dim yearFolder As String, monthFolder As String, timeFolder As String
    year = Format(Date, "yyyy")
    month = Format(Date, "mmmm")
    
    'Checks if year folder exists
    yearFolder = currentFolder & "\" & year
    If Len(Dir(currentFolder & "\" & year, vbDirectory)) = 0 Then
        MkDir yearFolder 'Make the folder if it doesn't exist
    End If
    
    'Checks if month folder exists
    monthFolder = currentFolder & "\" & year & "\" & month
    If Len(Dir(currentFolder & "\" & year & "\" & month, vbDirectory)) = 0 Then
        MkDir monthFolder 'Make the folder if it doesn't exist
    End If
    
    'Checks if time folder exists
    timeFolder = currentFolder & "\" & year & "\" & month & "\" & Format(Now, "mm-dd-yy-hh-mm-ss")
    If Len(Dir(timeFolder, vbDirectory)) = 0 Then
        MkDir timeFolder 'Make the folder if it doesn't exist
    End If
    
    CreateDirectory = timeFolder
    'Debug.Print timeFolder
    
End Function
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
Private Sub SaveFile(file As Workbook, vendorName As String, timeFolder As String, decision As Boolean)
'
'   This function will save the file
'
    Dim Title As String
        vendorName = Replace(vendorName, "/", "-")
        vendorName = Replace(vendorName, ",", "")
        vendorName = Replace(vendorName, ".", "")
        Title = timeFolder & "\" & vendorName & " Open PO " & Format(Date, "mm-dd-yy")
    file.SaveAs Filename:=Title, FileFormat:=51
    file.Close SaveChanges:=decision
End Sub
Private Sub EmailBuyers(timeFolder As String, SummaryWB As Workbook)
'
'   This function will email the buyers and attach the Summary Report
'
    Dim EmailApp As Outlook.Application
    Dim Source As String
    Set EmailApp = New Outlook.Application 'To launch outlook application
    
    Dim EmailItem As Outlook.MailItem 'To refer new outlook email
    Set EmailItem = EmailApp.CreateItem(olMailItem) 'To launch new outlook email
    
    'Set addressees
    EmailItem.To = ""
    EmailItem.CC = ""
    EmailItem.BCC = ""
    
    'Set email subject
    EmailItem.Subject = "Summary Open PO Report "
    
    'Set Body of email
    EmailItem.Body = "Hello all," & vbNewLine & _
                        "Attached is the Summary Open PO Report by vendor" & _
                        "This file can be found at " & timeFolder & _
                        vbNewLine & vbNewLine & _
                        "Regards," & vbNewLine & _
                        Application.UserName
    
    Source = ThisWorkbook.FullName
    EmailItem.Attachments.Add Source
    EmailItem.Display True
    'EmailItem.Send 'Uncomment this line if you want Outlook to automatically send the email
End Sub
