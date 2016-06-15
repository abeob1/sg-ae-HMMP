Imports System.Net.Mail
Module modEmail

    Public Function SendEmail(ByVal sBody As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "SendEmail"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Sarting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting SMTP properties", sFuncName)

            Dim smtpClient As SmtpClient = New SmtpClient(p_oCompDef.sSMTPServer, p_oCompDef.sSMTPPort)

            smtpClient.UseDefaultCredentials = False
            smtpClient.Credentials = New System.Net.NetworkCredential(p_oCompDef.sSMTPUser, p_oCompDef.sSMTPPassword)

            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network
            smtpClient.EnableSsl = True

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Function CreateDefaultMailMessage()", sFuncName)

            ' Split Email string based on ";"

            'MsgBox(sEmailTo)
            'sEmailTo = "boorlas@yahoo.com;boorlas@gmail.com"

            Dim sSendTo As String() = p_oCompDef.sEmailTo.Split(";")
            Dim message As MailMessage = CreateDefaultMailMessage(p_oCompDef.sEmailFrom, sSendTo, p_oCompDef.sEmailSubject, sBody, sErrDesc)
            Dim userState As Object = message

            AddHandler smtpClient.SendCompleted, New SendCompletedEventHandler(AddressOf smtp_SendCompleted)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Sending Email Message", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Sending Email Messages to : " & p_oCompDef.sEmailTo, sFuncName)
            smtpClient.Send(message)

            SendEmail = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with Success", sFuncName)

        Catch ex As Exception
            SendEmail = RTN_ERROR
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with Error", sFuncName)
            Call WriteToLogFile(sErrDesc, sFuncName)
            Call WriteToLogFile("Failed sending email to : " & " " & p_oCompDef.sEmailTo, sFuncName)
        Finally

        End Try

    End Function

    Private Sub smtp_SendCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        If (Not (e.Error Is Nothing)) Then
            Dim smessage As String = e.Error.Message
        End If
    End Sub

    Private Function CreateDefaultMailMessage(ByVal MailFrom As String, _
                                              ByVal MailTo As String(), _
                                              ByVal subject As String, _
                                              ByVal body As String, _
                                              ByRef sErrDesc As String) As MailMessage

        Dim message As MailMessage = New MailMessage()
        Dim sUploadFile As String = String.Empty
        Dim sFuncName As String = "CreateDefaultMailMessage"
        Dim sEmailAddress As String = String.Empty
        Dim AttachmentsFileInfo As System.IO.FileInfo() = Nothing
        Dim sAttachements As String = String.Empty
        Dim sFileName As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Sarting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning Email Properties..", sFuncName)

            message.From = New MailAddress(MailFrom)

            For Each sEmailAddress In MailTo
                message.To.Add(New MailAddress(sEmailAddress))
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Email Address" & ":  " & sEmailAddress, sFuncName)
            Next


            message.SubjectEncoding = System.Text.Encoding.UTF8
            message.Subject = subject
            message.BodyEncoding = System.Text.Encoding.UTF8
            message.Body = body
            message.IsBodyHtml = True

            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Attachments..", sFuncName)

            'If Not String.IsNullOrEmpty(sAttachmentPath) Then
            '    If System.IO.Directory.Exists(sAttachmentPath) Then
            '        AttachmentsDirInfo = New System.IO.DirectoryInfo(sAttachmentPath)
            '        AttachmentsFileInfo = AttachmentsDirInfo.GetFiles("*.*")
            '    End If
            'End If

            'For Each File In AttachmentsFileInfo
            '    sFileName = Microsoft.VisualBasic.Left(File.Name, InStr(File.Name, "-") - 2)
            '    If sFileName = sInvoiceNo Then
            '        message.Attachments.Add(New Attachment(sAttachmentPath & "\" & File.Name))
            '    End If
            'Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Successfully..", sFuncName)


            Return message

        Catch ex As Exception
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with Error", sFuncName)
            Call WriteToLogFile(sErrDesc, sFuncName)
            CreateDefaultMailMessage = Nothing
        Finally

        End Try

    End Function

    Public Sub EmailTemplate_Error()

        Dim sFuncName As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sErrDesc As String = String.Empty
        Dim sBody As String = String.Empty

        Try
            sFuncName = "EmailTemplate_Error()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Building email message body steps .... ", sFuncName)

            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
            sBody = sBody & " Dear Valued Customer,<br /><br /> HMMP Interface program had encountered the problem based on the following information. <br /><br />"
            sBody = sBody & " Please contact your system administrator or Technical consultant to assist you for further information.<br />"
            sBody = sBody & " For more details please refer to Error Log file.<br /><br />"

            sBody = sBody & p_SyncDateTime & " <br /><br />"

            ' Table Details
            sBody = sBody & "<table align='left' border=1 cellspacing=0 cellpadding=0 style='font-size:10.0pt;font-family:Arial'>"
            sBody = sBody & "<tr><td><strong style='color: blue; background-color: transparent;'>&nbsp;No&nbsp;</strong></td> "
            sBody = sBody & "<td><strong style='color: blue; background-color: transparent;'>&nbsp;File Name&nbsp;</strong></td> "
            sBody = sBody & "<td><strong style='color: blue'>&nbsp;Status&nbsp;</strong></td>"
            sBody = sBody & "<td><strong style='color: blue'>&nbsp;Error Description&nbsp;</strong></td></tr>"

            For i As Integer = 1 To p_oDtError.Rows.Count
                sBody = sBody & "<tr>"
                ' Sr. No
                sBody = sBody & "<td>&nbsp;" & i & "&nbsp;</td>"
                ' FileName
                sBody = sBody & "<td>&nbsp;" & p_oDtError.Rows(i - 1).Item(0).ToString & "&nbsp;</td>"
                'Status
                sBody = sBody & "<td>&nbsp;" & p_oDtError.Rows(i - 1).Item(1).ToString & "&nbsp;</td>"
                'Error Desc
                sBody = sBody & "<td>&nbsp;" & p_oDtError.Rows(i - 1).Item(2).ToString & "&nbsp;</td>"
                sBody = sBody & "</tr>"
            Next

            sBody = sBody & "</table><br /><br /><br />"
            For j As Integer = 1 To p_oDtError.Rows.Count
                sBody = sBody & "<br />"
            Next
            sBody = sBody & "<br/> Note: This email message is computer generated and it will be used internal purpose usage only.<div/>"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendMail", sFuncName)
            If SendEmail(sBody, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Sub

    Public Sub EmailTemplate_Success()

        Dim sFuncName As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sErrDesc As String = String.Empty
        Dim sBody As String = String.Empty

        Try
            sFuncName = "EmailTemplate_Success()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Building email message body steps .... ", sFuncName)

            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
            sBody = sBody & " Dear Valued Customer,<br /><br /> HMMP Interface program had Successfully uploaded the files into SAP. <br /><br />"
            sBody = sBody & " Please find the below list of files. <br /><br />"
            sBody = sBody & p_SyncDateTime & " <br /><br />"

            ' Table Details
            sBody = sBody & "<table align='left' border=1 cellspacing=0 cellpadding=0 style='font-size:10.0pt;font-family:Arial'>"
            sBody = sBody & "<tr><td><strong style='color: blue; background-color: transparent;'>&nbsp;No&nbsp;</strong></td> "
            sBody = sBody & "<td><strong style='color: blue; background-color: transparent;'>&nbsp;File Name&nbsp;</strong></td> "
            sBody = sBody & "<td><strong style='color: blue'>&nbsp;Status&nbsp;</strong></td></tr>"

            For i As Integer = 1 To p_oDtSuccess.Rows.Count
                sBody = sBody & "<tr>"
                ' Sr. No
                sBody = sBody & "<td>&nbsp;" & i & "&nbsp;</td>"
                ' FileName
                sBody = sBody & "<td>&nbsp;" & p_oDtSuccess.Rows(i - 1).Item(0).ToString & "&nbsp;</td>"
                'Status
                sBody = sBody & "<td>&nbsp;" & p_oDtSuccess.Rows(i - 1).Item(1).ToString & "&nbsp;</td>"
                sBody = sBody & "</tr>"
            Next

            sBody = sBody & "</table><br /><br /><br />"
            For j As Integer = 1 To p_oDtSuccess.Rows.Count
                sBody = sBody & "<br />"
            Next
            sBody = sBody & "<br/> Note: This email message is computer generated and it will be used internal purpose usage only.<div/>"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendMail", sFuncName)
            If SendEmail(sBody, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Sub

    Public Sub EmailTemplate_GeneralError(ByVal sErrDesc As String)
        Dim sFuncName As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sBody As String = String.Empty

        Try
            sFuncName = "EmailTemplate_Error()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Building email message body steps .... ", sFuncName)

            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
            sBody = sBody & " Dear Valued Customer,<br /><br /> HMMP Interface program had encountered the problem based on the following information. <br /><br />"
            sBody = sBody & " Please contact your system administrator or Technical consultant to assist you for further information.<br /><br />"
            sBody = sBody & p_SyncDateTime & " <br />"

            sBody = sBody & sErrDesc

            sBody = sBody & "<br/> <br/> Note: This email message is computer generated and it will be used internal purpose usage only.<div/>"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendMail", sFuncName)
            If SendEmail(sBody, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Sub

End Module
