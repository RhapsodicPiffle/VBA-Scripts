Attribute VB_Name = "DomesticEmail"
Option Explicit


Private Sub BtnEmail_Quote()
    Send_Quote
End Sub


Private Sub Send_Quote()
    '-------------< Send_Email() >-------------
    Dim sSubject As String
    Dim sEmail_To As String
    Dim sEmail_CC As String
    Dim wb As Workbook
    
    '----< Send with Outlook >----
    Dim app_Outlook As Outlook.Application
    Set app_Outlook = New Outlook.Application
   
    '--< Edit Email >--
    Dim objEmail As Outlook.MailItem
    Dim countEmails As Integer
    Dim iRow As Integer
    Dim LastRow As Integer
    LastRow = ActiveSheet.UsedRange.Rows.count
    For iRow = 2 To LastRow                     'Loop through the POs
        If Cells(iRow, 1) <> "" Then            'If there is a PO#
            If Cells(iRow, 5) = "" Then         'And they have not been quoted already
                Dim nRow As Integer
                For nRow = 2 To 5               'Loop through Brokers
                    Sheets("LTL Brokers").Select    'Start grabbing Broker information
                    sEmail_To = Cells(nRow, 2)  'Grab Email To
                    sEmail_CC = Cells(nRow, 3)  'Grab Email CC
                    Sheets("POs").Select        'Go back to PO tab
                    Dim Location As String
                    If Cells(iRow, 3) = "OH" Then
                        Location = " OHIO"
                    End If
                    If Cells(iRow, 3) = "UT" Then
                        Location = " UTAH"
                    End If
                    If Cells(iRow, 3) = "OK" Then
                        Location = " OKLAHOMA"
                    End If
                    sSubject = "PO#" & Range("A" & iRow).Value & Location
                    
                    '--< Body Text>--
                    Dim strBody As String
                    strBody = "Hello, " & vbNewLine & _
                    "LTL quote needed.  See attached for information.  Pickup is " & _
                    Range("B" & iRow).Value & ". We need this to " & _
                    Range("C" & iRow).Value & " by Friday " & _
                    Range("D" & iRow).Value & _
                    ". Please double check freight class, as vendor is not as thorough. If you need anything else, please let me know. Thanks!" & vbNewLine & vbNewLine & _
                    "Mike Brown" & vbNewLine
                    '--</ Body Text >--
                    
                    '--< Attachments >--
                    Dim strLocation As String
                    strLocation = "C:\Users\Michael.Brown\Documents\POs\" & "PO_" & _
                        Range("A" & iRow).Value & Location & ".xls"
                    If Len(Dir(strLocation)) = 0 Then
                        strLocation = "C:\Users\Michael.Brown\Documents\POs\" & "PO_" & _
                        Range("A" & iRow).Value & Location & ".xlsx"
                        If Len(Dir(strLocation)) = 0 Then
                            MsgBox "File does not exist"
                            Exit Sub
                        End If
                    End If
                    '--</ Attachments >--
                    
                    '--< Send Email >--
                    Set objEmail = app_Outlook.CreateItem(olMailItem)
                    objEmail.To = sEmail_To
                    objEmail.CC = sEmail_CC
                    objEmail.Subject = sSubject
                    objEmail.Body = strBody
                    objEmail.Attachments.Add (strLocation)
                    objEmail.Display True
                    'objEmail.Send
                    '--</ Send Email >--
                    
                    countEmails = countEmails + 1
                    
                Next
            End If
        End If
    Next
    
    '< End >
    Set objEmail = Nothing
    Set app_Outlook = Nothing
    '</ End>

    
    MsgBox countEmails & " emails sent", vbInformation, "Finished"
    '----</ Send with Outlook >----
    '-------------</ Send_Email() >-------------
End Sub

Private Sub BtnTest()
    TestFunctions
End Sub


Private Sub TestFunctions()
    '-------------< TestFunctions() >-------------
    Dim DataColumns As Variant
    DataColumns = Get_Columns()
    Dim Col_PO_Number As Integer, Col_QR_Number As Integer
    Col_PO_Number = DataColumns(1)(0)
    Col_QR_Number = DataColumns(1)(4)
    
    Dim Email_Subject As String, Location As String, BodyText As String, Attachments As String
    Dim Email_Addresses As Variant, Email_To As String, Email_CC As String
    
    '----< Send with Outlook >----
    Dim app_Outlook As Outlook.Application
    Set app_Outlook = New Outlook.Application
   
    '--< Edit Email >--
    Dim objEmail As Outlook.MailItem
    
    Dim countEmails As Integer
    
    Dim iRow As Integer, LastRow As Integer, LastBroker As Integer
    LastRow = ActiveSheet.UsedRange.Rows.count
    LastBroker = numBroker                                                  'Grab all of the Brokers, regardless of number
    For iRow = 2 To LastRow                                                 'Loop through the POs
        If Cells(iRow, Col_PO_Number) <> "" Then                            'If there is a PO#
            If Cells(iRow, Col_QR_Number) = "" Then                         'And we have not already asked for quotes
                Dim nRow As Integer
                For nRow = 2 To LastBroker                                  'Loop through Brokers To and CC
                    Email_Addresses = Get_Email_Addresses(nRow)             'Get the Email Addresses
                    Email_To = Email_Addresses(0)                           'Get the Email_To
                    Email_CC = Email_Addresses(1)                           'Get the Email_CC
                    Location = Get_Location(iRow)                           'Get the Location
                    Email_Subject = Set_Subject(iRow, Location)             'Get the Email Subject
                    BodyText = Get_Body_Text(iRow)                          'Get the Body Text
                    Attachments = Get_Attachments(iRow, Location, "PO")     'Get the Attachment location
                
                    '--< Send Email>--
                    Set objEmail = app_Outlook.CreateItem(olMailItem)
                    objEmail.To = Email_To
                    objEmail.CC = Email_CC
                    objEmail.Subject = Email_Subject
                    objEmail.HTMLBody = Get_Body_Text(iRow)
                    objEmail.Attachments.Add (Attachments)
                    objEmail.Display True
                    'objEmail.Send
                    '--</ Send Email>--
                    countEmails = countEmails + 1
                Next
            End If
        End If
    Next
    
    '< End >
    Set objEmail = Nothing
    Set app_Outlook = Nothing
    '</ End>

    MsgBox countEmails & " emails sent", vbInformation, "Finished"
    '----</ Send with Outlook >----
    '-------------</ TestFunctions() >-------------
End Sub

Function Get_Email_Addresses(nRow As Integer) As Variant
    
    '--< Grab Email Info >--
    Sheets("Email Test").Select
    Dim Email_To As String, Email_CC As String
    Email_To = Cells(nRow, 2)                           'Grab Email To
    Email_CC = Cells(nRow, 3)                           'Grab Email CC
    Sheets("POs Test Sheet").Select
    Get_Email_Addresses = Array(Email_To, Email_CC)     'Return Email Addresses
    '--</ Grab Email Info >--
    
End Function

Function Set_Subject(iRow As Integer, Location As String) As String

    Dim DataColumns As Variant
    DataColumns = Get_Columns()
    Dim Col_PO_Letter As String
    Col_PO_Letter = DataColumns(0)(0)
    
    '--< Set Subject >--
    Sheets("POs Test Sheet").Select            'Use the PO tab
    Set_Subject = "PO#" & Range(Col_PO_Letter & iRow).Value & Location
    '--</ Set Subject>--
    
End Function

Function Get_Body_Text(iRow As Integer) As String

    Dim DataColumns As Variant
    DataColumns = Get_Columns()
    Dim Col_PUD_Letter As String, Col_LOC_Letter As String, Col_NBD_Letter As String
    Col_PUD_Letter = DataColumns(0)(1)
    Col_LOC_Letter = DataColumns(0)(2)
    Col_NBD_Letter = DataColumns(0)(3)

    '--< Body Text>--
    Sheets("POs Test Sheet").Select            'Use the PO tab
    
    Get_Body_Text = "<p>Hello,<br>LTL quote needed.  See attached for information. " & _
        "Pickup is <span style='background-color: #ffff00'>" & Range(Col_PUD_Letter & iRow).Value & "</span>. " & _
        "We need this to " & Range(Col_LOC_Letter & iRow).Value & " by <span style='background-color: #ffff00'>Friday " & _
        Range(Col_NBD_Letter & iRow).Value & "</span>. If you need anything else, please let me know.<br>" & _
        "<br>Mike Brown <br><br></p>"
    '--</ Body Text >--

End Function

Function Get_Attachments(iRow As Integer, Location As String, Docs As String) As String
    
    Dim DataColumns As Variant
    DataColumns = Get_Columns()
    Dim Col_PO_Letter As String
    Col_PO_Letter = DataColumns(0)(0)
    
    '--< Attachments >--
    Get_Attachments = "C:\Users\Michael.Brown\Documents\" & Docs & "s\" & Docs & "_" & _
        Range(Col_PO_Letter & iRow).Value & Location & ".pdf"
    If Not Dir(Get_Attachments, vbDirectory) = vbNullString Then
        MsgBox "exists"
    ElseIf Dir(Get_Attachments, vbDirectory) = vbNullString Then
        MsgBox "does not exist, looking up .xls files"
        Get_Attachments = "C:\Users\Michael.Brown\Documents\" & Docs & "s\" & Docs & "_" & _
        Range(Col_PO_Letter & iRow).Value & Location & ".xls"
    End If
    MsgBox "Get_Attachments: " & Get_Attachments
    '--</ Attachments >--
    
End Function

Function Get_Location(iRow) As String

    Dim DataColumns As Variant
    DataColumns = Get_Columns()
    Dim Col_LOC_Number As Integer
    Col_LOC_Number = DataColumns(1)(2)

    '--< Determine Location >--
    Sheets("POs Test Sheet").Select            'Use the PO tab
    If Cells(iRow, Col_LOC_Number) = "OK" Then
        Get_Location = " OKLAHOMA"
    ElseIf Cells(iRow, Col_LOC_Number) = "UT" Then
        Get_Location = " UTAH"
    ElseIf Cells(iRow, Col_LOC_Number) = "OH" Then
        Get_Location = " OHIO"
    Else
        Get_Location = ""
    End If
    '--</ Determine Location >--

End Function

Function numBroker() As Integer

    '--< Get Number of Brokers >--
    Sheets("Email Test").Select
    Dim LastBroker As Integer
    numBroker = ActiveSheet.UsedRange.Rows.count       'Grab all of the Brokers, regardless of number
    Sheets("POs Test Sheet").Select
    '--</ Get Number of Brokers>--

End Function

Private Sub BtnReceiving()
    Send_To_Receiving
End Sub


Private Sub Send_To_Receiving()
    '-------------< Send_To_Receiving() >-------------
    Dim DataColumns As Variant
    DataColumns = Get_Columns()
    Dim Col_REC_Number As Integer, Col_PO_Number As Integer, Col_NBD_Letter As String
    Col_REC_Number = DataColumns(1)(9)
    Col_PO_Number = DataColumns(1)(0)
    Col_NBD_Letter = DataColumns(0)(3)
    
    Dim Email_Subject As String, Location As String, BodyText As String, Attachment1 As String, Attachment2 As String
    Dim Email_Addresses As Variant, Email_To As String, Email_CC As String, Name As String
    
    '----< Send with Outlook >----
    Dim app_Outlook As Outlook.Application
    Set app_Outlook = New Outlook.Application
   
    '--< Edit Email >--
    Dim objEmail As Outlook.MailItem
    
    Dim iRow As Integer, LastRow As Integer
    LastRow = ActiveSheet.UsedRange.Rows.count
    For iRow = 2 To LastRow                                         'Loop through the POs
        If Cells(iRow, Col_PO_Number) <> "" Then                    'If there is a PO#
            If Cells(iRow, Col_REC_Number) = "" Then                'And we have not already sent it to receiving
                Location = Get_Location(iRow)                       'Get the Location
                If Location = " OHIO" Then
                    Email_To = "OHIO EMAIL"
                    Name = ""
                ElseIf Location = " UTAH" Then
                    Email_To = "UTAH EMAIL" 
                    Name = ""
                Else
                    Email_To = ""
                    Name = "Unknown"
                End If
                Email_CC = "" 
                Email_Subject = Set_Subject(iRow, Location)     'Get the Email Subject
                BodyText = "Hey " & Name & ", " & vbNewLine & _
                "This will be coming your way the week of " & Range(Col_NBD_Letter & iRow).Value & _
                vbNewLine & vbNewLine & "Mike Brown"
            
                '--< Send Email>--
                Set objEmail = app_Outlook.CreateItem(olMailItem)
                objEmail.To = Email_To
                objEmail.CC = Email_CC
                objEmail.Subject = Email_Subject
                objEmail.Body = BodyText
                Attachment1 = Get_Attachments(iRow, Location, "BOL")    'Get the BOL Attachment
                objEmail.Attachments.Add (Attachment1)
                Attachment2 = Get_Attachments(iRow, Location, "Packing List")     'Get the Packing list Attachment
                objEmail.Attachments.Add (Attachment2)
                objEmail.Display True
                'objEmail.Send
                '--</ Send Email>--
            End If
        End If
    Next
    
    '< End >
    Set objEmail = Nothing
    Set app_Outlook = Nothing
    '</ End>

    MsgBox "Emails Sent", vbInformation, "Finished"
    '----</ Send with Outlook >----
    '-------------</ Send_To_Receiving() >-------------
End Sub

Function Get_Columns() As Variant

'--< Get_Columns >--
    ' Anytime that I am looking for a column, use a variable instead
    ' That way, when I need to make a change (like I added a column,
    '   then I can make the change in one location

    Dim Get_Columns_Letters As Variant, Get_Columns_Numbers As Variant
    Dim Col_PO_Letter As String, Col_PUD_Letter As String, Col_LOC_Letter As String, Col_NBD_Letter As String
    Dim Col_QR_Letter As String, Col_LOQ_Letter As Variant, Col_QA_Letter As String, Col_BOLR_Letter As String
    Dim Col_BOLV_Letter  As String, Col_REC_Letter As String, Col_done_Letter  As String
    
    Dim Col_PO_Number As Integer, Col_PUD_Number As Integer, Col_LOC_Number As Integer, Col_NBD_Number As Integer
    Dim Col_QR_Number As Integer, Col_LOQ_Number As Variant, Col_QA_Number As Integer, Col_BOLR_Number As Integer
    Dim Col_BOLV_Number As Integer, Col_REC_Number As Integer, Col_done_Number As Integer

'' __List of colums__
    Col_PO_Letter = "A"
    Col_PO_Number = 1        'Purchase order
    Col_PUD_Letter = "B"
    Col_PUD_Number = 2       'Pickup date
    Col_LOC_Letter = "C"
    Col_LOC_Number = 3       'Location
    Col_NBD_Letter = "D"
    Col_NBD_Number = 4       'Need-by date
    Col_QR_Letter = "E"
    Col_QR_Number = 5        'Quote requested
    Col_LOQ_Letter = Array("F", "G", "H", "I")
    Col_LOQ_Number = Array(6, 7, 8, 9) 'List of Quotes
    Col_QA_Letter = "J"
    Col_QA_Number = 10       'Quote accepted
    Col_BOLR_Letter = "K"
    Col_BOLR_Number = 11     'BOL Received
    Col_BOLV_Letter = "L"
    Col_BOLV_Number = 12     'BOL to Vendor
    Col_REC_Letter = "M"
    Col_REC_Number = 13      'Sent to Receiving
    Col_done_Letter = "N"
    Col_done_Number = 14     'DONE
    
    Get_Columns_Letters = Array(Col_PO_Letter, Col_PUD_Letter, Col_LOC_Letter, Col_NBD_Letter, Col_QR_Letter, Col_LOQ_Letter, Col_QA_Letter, Col_BOLR_Letter, Col_BOLV_Letter, Col_REC_Letter, Col_done_Letter)
    Get_Columns_Numbers = Array(Col_PO_Number, Col_PUD_Number, Col_LOC_Number, Col_NBD_Number, Col_QR_Number, Col_LOQ_Number, Col_QA_Number, Col_BOLR_Number, Col_BOLV_Number, Col_REC_Number, Col_done_Number)
    Get_Columns = Array(Get_Columns_Letters, Get_Columns_Numbers)

'--</ Get_Columns >--

End Function
