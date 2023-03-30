'To Setup Rules See https://windowsreport.com/outlook-rule-download-attachments/
Sub SaveWorkOrderToDisk()
    'This is for Nicole and Sarah to Automatically Save Work Order Atachments to the T: Drive
    Dim MItem As Outlook.MailItem
    Dim oAttachment As Outlook.Attachment
    Dim sSaveFolder As String
    Dim Msg, Style, Title, Help As String
    Dim Ctxt As Integer
    
    
    For Each MItem In Application.ActiveExplorer.Selection
        For Each oAttachment In MItem.Attachments
            If InStr(oAttachment.DisplayName, "WorkOrder") <> 0 Then
                sSaveFolder = FindWorkOrderDirectory(oAttachment.DisplayName) & "\Invoices & Workorders\"
                
                Msg = "Save to: " & vbCrLf & sSaveFolder & oAttachment.DisplayName  ' Define message.
                Style = vbYesNo Or vbQuestion Or vbDefaultButton1    ' Define buttons.
                Title = "Work Order Attachments to T: Drive (Contracts)"    ' Define title.
                Help = "DEMO.HLP"    ' Define Help file.
                Ctxt = 1000    ' Define topic context.
    
                Response = MsgBox(Msg, Style, Title, Help, Ctxt)
                If Response = vbYes Then    ' User chose Yes.
                    Debug.Print "sSaveFolder: " & sSaveFolder
                    oAttachment.SaveAsFile sSaveFolder & oAttachment.DisplayName
                Else    ' User chose No.
                    MyString = "No"    ' Perform some action.
                End If
            End If
        Next
    Next
End Sub

Function FindWorkOrderDirectory(AttachmentName As String)
    'This is for Nicole and Sarah to Automatically Save Work Order Atachments to the T: Drive
    Dim fldr As String
    Dim DRIVE_PATH As String
    DRIVE_PATH = "T:\" & Left$(AttachmentName, 4)
    fldr = Dir(DRIVE_PATH & "\" & Left$(AttachmentName, 9) & "*", vbDirectory)
    Debug.Print Len(fldr) & " " & fldr
    If Len(fldr) > 0 Then
        FindWorkOrderDirectory = "T:\" & Left$(AttachmentName, 4) & "\" & fldr
    End If
End Function
