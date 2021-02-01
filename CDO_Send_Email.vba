'''<summary>
''' Method SendCDOEmail is a function that sends email using CDOSYS (CDOsys.dll) with SMTP protocal.
''' <b>Note</b>: This only works on VBA with Office on Windows.  It does not work with Office for Mac.
'''</summary>
'''<param name="emailTo"><c>emailTo</c> is a string of the email address or addresses to which the mail will be sent. If multiple, separate with semicolon.</param>
'''<param name="emailFrom"><c>emailFrom</c> is the string From email address.  Only a single email address is expected.</param>
'''<param name="SMTPServer"><c>SMTPServer</c> is the string domain name of the SMTP server.</param>
'''<param name="emailSubject"><b>Optional</b> <c>emailSubject</c> is the string subect text of the email.</param>
'''<param name="textBody"><b>Optional</b> <c>textBody</c> text only version of the email body content.</param>
'''<param name="htmlBody"><b>Optional</b> <c>textBody</c> html version of the email body content. One, or the other (text only), or both can be set.</param>
'''<param name="emailCC"><b>Optional</b> <c>emailCC</c> is a String representing the email address or addresses of the Carbon Copy recipients.  If multiple, separate with semicolon.</param>
'''<param name="emailBCC"><b>Optional</b> <c>emailBCC</c> is a String representing the email address or addresses of the Blind Carbon Copy recipients. If multiple, separate with semicolon.</param>
'''<param name="attachmentPath"><b>Optional</b> <c>attachmentPath</c> is the string or array of the file path(s) of the attachment file(s) to be attached to the email</param>
'''<param name="SMTPPort"><b>Optional</b> <c>SMTPPort</c> is the optional integer port number to be used with the SMTP protocal. The default value is 25.</param>
'''<author>David Sullivan</author>
'''<remarks>
''' This method uses elements from the example from https://www.rondebruin.nl/win/s1/cdo.htm.  More examples for CDO at https://www.w3schools.com/asp/asp_send_email.asp
'''</remarks>
Public Sub SendEmail(emailTo As String, _
 ByVal emailFrom As String, _
 ByVal SMTPServer As String, _
 Optional emailSubject As Variant, _
 Optional textBody As Variant = "", _
 Optional htmlBody As Variant = "", _
 Optional emailCC As Variant = "", _
 Optional emailBCC As Variant = "", _
 Optional attachmentPath As Variant = "", _
 Optional ByVal SMTPPort As Integer = 25)
 
    ' Declare CDOsys objects
    Dim iMsg As Object
    Dim iConf As Object
    Dim Flds As Variant
    Dim strTo As String
    
    ' initialize CDO message and configuration objects.
    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    
    iConf.Load -1  ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort
        .Update
    End With
    
    ' Set CDO Message properties
    With iMsg
        Set .Configuration = iConf
        .To = emailTo
        .From = emailFrom
        .Subject = emailSubject
        .textBody = textBody
    End With
    
    ' Check if there is an HTML version of the body.  If it is blank, the email client will probably still override the text body if added.
    If htmlBody <> "" Then
        iMsg.htmlBody = htmlBody
    End If
    
    ' Check if there are CC recipients
    If emailCC <> "" Then
        iMsg.Cc = emailCC
    End If
    
    ' Check if there are BCC recipients
    If emailBCC <> "" Then
        iMsg.Bcc = emailBCC
    End If
    
    ' Check if there are attachments
    If attachmentPath <> "" Then
        If IsArray(attachmentPath) Then
            For Each attachment In attachmentPath
                ' Check if attachment exists (path to file is correct), then add the attachment.
                If Len(Dir(attachment)) Then
                    iMsg.AddAttachment attachment
                Else
                    ' Error 53 "File Not Found"
                    Err.Raise 53
                End If
            Next attachment
        ElseIf VarType(attachmentPath) = 8 Then
            ' emailAttachment is just a string,  add it if the file exists.
            If Len(Dir(attachmentPath)) Then
                iMsg.AddAttachment attachmentPath
            Else
                ' Error 53 "File Not Found"
                Err.Raise 53
            End If
        End If
    End If
    
    iMsg.Send
    
    Set iMsg = Nothing
    Set iConf = Nothing
    Set Flds = Nothing

End Sub
