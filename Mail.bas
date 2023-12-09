Attribute VB_Name = "Module2"
Dim rec1 As DAO.Recordset
Public Function SendEmail(ByVal strSender As String, _
                          ByVal strRecipient As String, _
                          ByVal strSubject As String, _
                          ByVal strBody As String, _
                          Optional ByVal strCc As String, _
                          Optional ByVal strBcc As String, _
                          Optional ByVal colAttachments As Collection _
                        ) As Boolean
    Dim cdoMsg As New CDO.Message
    Dim cdoConf As New CDO.Configuration
    Dim schema As String
    Dim Flds
    Dim attachment
    Dim strHTML

    On Error GoTo errtrap
    Const cdoSendUsingPort = 2

    'Set cdoMsg =  CreateObject("CDO.Message")
    'Set cdoConf = CreateObject("CDO.Configuration")

    Set Flds = cdoConf.Fields
    
    Set rec1 = db.OpenRecordset("select * from mailsetup")
    If Not IsNull(rec1("sendermailid")) And Not IsNull(rec1("smtpserver")) And Not IsNull(rec1("smtpport")) And Not IsNull(rec1("password")) Then
        mailid = rec1("sendermailid")
        serveradd = rec1("smtpserver")
        smtpport = rec1("smtpport")
        pwd = rec1("password")
    Else
        MsgBox "Invalid Mail Setup", vbCritical
        Exit Function
    End If
    schema = "http://schemas.microsoft.com/cdo/configuration/"
    
    With Flds
        .Item(schema & "sendusing") = 2
        .Item(schema & "smtpserver") = serveradd
        .Item(schema & "smtpserverport") = smtpport
        .Item(schema & "smtpauthenticate") = 1
        .Item(schema & "sendusername") = mailid
        .Item(schema & "sendpassword") = pwd
        .Item(schema & "smtpusessl") = 1
        .Update
    End With
    '    With Flds
    '        .Item(schema & "sendusing") = 2
    '        .Item(schema & "smtpserver") = "smtp.rediffmail.com"
    '        .Item(schema & "smtpserverport") = 25
    '        .Item(schema & "smtpauthenticate") = 1
    '        .Item(schema & "sendusername") = "abinashacharya@rediffmail.com"
    '        .Item(schema & "sendpassword") = "pitu1234"
    '        .Item(schema & "smtpusessl") = 1
    '        .Update
    '    End With
    '    With Flds
    '        .Item(schema & "sendusing") = 2
    '        .Item(schema & "smtpserver") = "smtp.mail.yahoo.com"
    '        .Item(schema & "smtpserverport") = 465
    '        .Item(schema & "smtpauthenticate") = 1
    '        .Item(schema & "sendusername") = "yantrainfotech@yahoo.com"
    '        .Item(schema & "sendpassword") = "pass098&6"
    '        .Item(schema & "smtpusessl") = 1
    '        .Update
    '    End With

    ' Apply the settings to the message.
    With cdoMsg
        Set .Configuration = cdoConf
        .To = strRecipient
        .From = strSender
        .Subject = strSubject
        .TextBody = strBody
        If Not colAttachments Is Nothing Then
            For Each attachment In colAttachments
                .AddAttachment attachment
            Next
        End If
        If strCc <> "" Then .CC = strCc
        If strBcc <> "" Then .BCC = strBcc
        .Send
    End With

    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set Flds = Nothing

    SendEmail = True
    Exit Function
errtrap:
    Err.Raise Err.Number, "", "Error from Functions.SendEmail" & Err.Description
    SendEmail = False
End Function




