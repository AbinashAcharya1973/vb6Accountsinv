Attribute VB_Name = "Module1"
Global db As DAO.Database
Global Db1 As DAO.Database
Global formid
Global usertype
Global dbname
Global AccountingPeriod
Global usname
Global SEARCHWORD
Global InvoicePush As Boolean
Global PurchasePush As Boolean
Global VoucherPush As Boolean
Global EnliteBasicPlus As Boolean
Global CompleteSyn As Boolean
Global SoftwareVersion
Global RH
Global SH
Global EC
Global ACYEAR
Global FORMNAME

Public Function ValidateNumeric(strText As String) _
    As Boolean
ValidateNumeric = CBool(strText = "-" _
    Or strText = "-." _
    Or strText = "." _
    Or IsNumeric(strText))
End Function
Public Function NumberToWord(num As String)
    Dim w1(100) As Integer
    Dim w2(100) As String
    Dim x1, x2, rs, X3 As String
    Dim i As Integer
    w1(0) = 0
    w1(1) = 1
    w1(2) = 2
    w1(3) = 3
    w1(4) = 4
    w1(5) = 5
    w1(6) = 6
    w1(7) = 7
    w1(8) = 8
    w1(9) = 9
    w1(10) = 10
    w1(11) = 11
    w1(12) = 12
    w1(13) = 13
    w1(14) = 14
    w1(15) = 15
    w1(16) = 16
    w1(17) = 17
    w1(18) = 18
    w1(19) = 19
    w1(20) = 20
    w1(21) = 21
    w1(22) = 22
    w1(23) = 23
    w1(24) = 24
    w1(25) = 25
    w1(26) = 26
    w1(27) = 27
    w1(28) = 28
    w1(29) = 29
    w1(30) = 30
    w1(31) = 31
    w1(32) = 32
    w1(33) = 33
    w1(34) = 34
    w1(35) = 35
    w1(36) = 36
    w1(37) = 37
    w1(38) = 38
    w1(39) = 39
    w1(40) = 40
    w1(41) = 41
    w1(42) = 42
    w1(43) = 43
    w1(44) = 44
    w1(45) = 45
    w1(46) = 46
    w1(47) = 47
    w1(48) = 48
    w1(49) = 49
    w1(50) = 50
    w1(51) = 51
    w1(52) = 52
    w1(53) = 53
    w1(54) = 54
    w1(55) = 55
    w1(56) = 56
    w1(57) = 57
    w1(58) = 58
    w1(59) = 59
    w1(60) = 60
    w1(61) = 61
    w1(62) = 62
    w1(63) = 63
    w1(64) = 64
    w1(65) = 65
    w1(66) = 66
    w1(67) = 67
    w1(68) = 68
    w1(69) = 69
    w1(70) = 70
    w1(71) = 71
    w1(72) = 72
    w1(73) = 73
    w1(74) = 74
    w1(75) = 75
    w1(76) = 76
    w1(77) = 77
    w1(78) = 78
    w1(79) = 79
    w1(80) = 80
    w1(81) = 81
    w1(82) = 82
    w1(83) = 83
    w1(84) = 84
    w1(85) = 85
    w1(86) = 86
    w1(87) = 87
    w1(88) = 88
    w1(89) = 89
    w1(90) = 90
    w1(91) = 91
    w1(92) = 92
    w1(93) = 93
    w1(94) = 94
    w1(95) = 95
    w1(96) = 96
    w1(97) = 97
    w1(98) = 98
    w1(99) = 99



    w2(0) = "Zero "
    w2(1) = "One "
    w2(2) = "Two "
    w2(3) = "Three "
    w2(4) = "Four "
    w2(5) = "Five "
    w2(6) = "Six "
    w2(7) = "Seven "
    w2(8) = "Eight "
    w2(9) = "Nine "
    w2(10) = "Ten "
    w2(11) = "Eleven "
    w2(12) = "Twelve "
    w2(13) = "Thirteen "
    w2(14) = "Fourteen "
    w2(15) = "Fifteen "
    w2(16) = "Sixteen "
    w2(17) = "Seventeen "
    w2(18) = "Eighteen "
    w2(19) = "Ninteen "
    w2(20) = "Twenty "
    w2(21) = "Twenty One "
    w2(22) = "Twenty Two "
    w2(23) = "Twenty Theee "
    w2(24) = "Twenty Four "
    w2(25) = "Twenty Five "
    w2(26) = "Twenty Six "
    w2(27) = "Twenty Seven "
    w2(28) = "Twenty Eight "
    w2(29) = "Twenty Nine "
    w2(30) = "Thirty "
    w2(31) = "Thirty One "
    w2(32) = "Thirty Two "
    w2(33) = "Thirty Three "
    w2(34) = "Thirty Four "
    w2(35) = "Thirty Five "
    w2(36) = "Thirty Six "
    w2(37) = "Thirty Seven "
    w2(38) = "Thirty Eight "
    w2(39) = "Thirty Nine "
    w2(40) = "Fourty "
    w2(41) = "Fourty One "
    w2(42) = "Fourty Two "
    w2(43) = "Fourty Three "
    w2(44) = "Fourty Four "
    w2(45) = "Fourty Five "
    w2(46) = "Fourty Six "
    w2(47) = "Fourty Seven "
    w2(48) = "Fourty Eight "
    w2(49) = "Fourty Nine "
    w2(50) = "Fifty "
    w2(51) = "Fifty One "
    w2(52) = "Fifty Two "
    w2(53) = "Fifty Three "
    w2(54) = "Fifty Four "
    w2(55) = "Fifty Five "
    w2(56) = "Fifty Six "
    w2(57) = "Fifty Seven "
    w2(58) = "Fifty Eight "
    w2(59) = "Fifty Nine "
    w2(60) = "Sixty "
    w2(61) = "Sixty One "
    w2(62) = "Sixty Two "
    w2(63) = "Sixty Three "
    w2(64) = "Sixty Four "
    w2(65) = "Sixty Five "
    w2(66) = "Sixty Six "
    w2(67) = "Sixty Seven "
    w2(68) = "Sixty Eight "
    w2(69) = "Sixty Nine "
    w2(70) = "Seventy "
    w2(71) = "Seventy One "
    w2(72) = "Seventy Two "
    w2(73) = "Seventy Three "
    w2(74) = "Seventy Four "
    w2(75) = "Seventy Five "
    w2(76) = "Seventy Six "
    w2(77) = "Seventy Seven "
    w2(78) = "Seventy Eight "
    w2(79) = "Seventy Nine "
    w2(80) = "Eighty "
    w2(81) = "Eighty One "
    w2(82) = "Eighty Two "
    w2(83) = "Eighty Three "
    w2(84) = "Eighty Four "
    w2(85) = "Eighty Five "
    w2(86) = "Eighty Six "
    w2(87) = "Eighty Seven "
    w2(88) = "Eighty Eight "
    w2(89) = "Eighty Nine "
    w2(90) = "Ninety  "
    w2(91) = "Ninety One "
    w2(92) = "Ninety Two "
    w2(93) = "Ninety Three "
    w2(94) = "Ninety Four "
    w2(95) = "Ninety Five "
    w2(96) = "Ninety Six "
    w2(97) = "Ninety Seven "
    w2(98) = "Ninety Eight"
    w2(99) = "Ninety Nine"

    X3 = Val(num) - Int(Val(num))
    x1 = Int(Val(num))
    x2 = Right(x1, 2)
    rs = ""
    'Label2.Caption = x3
    For i = 1 To 99
        If w1(i) = Val(x2) Then
            rs = w2(i)
            'k = 1
        End If
    Next i
    If Len(x1) >= 2 Then
        x1 = Left(x1, Len(x1) - 2)
    ElseIf Len(x1) >= 1 Then
        x1 = Left(x1, Len(x1) - 1)
    End If
    x2 = Right(x1, 1)
    For i = 1 To 9
        If w1(i) = Val(x2) Then
            rs = w2(i) + " Hundred " + rs
        End If
    Next i
    If Len(x1) > 0 Then
        x1 = Left(x1, Len(x1) - 1)
    End If
    x2 = Right(x1, 2)
    For i = 1 To 99
        If w1(i) = Val(x2) Then
            rs = w2(i) + " Thousand " + rs
        End If
    Next i
    If Len(x1) >= 2 Then
        x1 = Left(x1, Len(x1) - 2)
    ElseIf Len(x1) >= 1 Then
        x1 = Left(x1, Len(x1) - 1)
    End If
    x2 = Right(x1, 2)
    For i = 1 To 99
        If w1(i) = Val(x2) Then
            rs = w2(i) + " Lakh " + rs
        End If
    Next i
    If Len(x1) >= 2 Then
        x1 = Left(x1, Len(x1) - 2)
    ElseIf Len(x1) >= 1 Then
        x1 = Left(x1, Len(x1) - 1)
    End If
    x2 = Right(x1, 2)
    For i = 1 To 99
        If w1(i) = Val(x2) Then
            rs = w2(i) + " Crore " + rs
        End If
    Next i
    NumberToWord = "Rupees " + rs + " Only"


End Function

Public Function NetConnectStatus() As Boolean
    On Error GoTo err_DoWebRequest
    Dim strurl As String
    strurl = "http://www.google.co.in/"

    Dim objXML As Object
    Set objXML = CreateObject("Microsoft.XMLHTTP")
    objXML.Open "GET", strurl, False
    objXML.Send
    If (objXML.Status = 404) Then
        DoWebRequest = "404 Error"
        DoWebRequest = objXML.ResponseText
    Else
        NetConnectStatus = True
    End If
    Set objXML = Nothing
    Exit Function
err_DoWebRequest:
    NetConnectStatus = "False"
End Function
Public Function getFinancialYear(StrDate As String)
    Dim tMonth
    Dim tdate
    Dim tYear
    tdate = CDate(StrDate)
    tMonth = Month(tdate)
    tYear = Year(tdate)

    'Financial Year=CurrentYear to PreviousYear
    If tMonth >= 1 And tMonth <= 3 Then
        getFinancialYear = Trim(str(tYear - 1)) & "-" & Trim(str(tYear))
    End If
    'Financial Year=Current Year to Next Year
    If tMonth >= 4 Then
        getFinancialYear = Trim(str(tYear)) & "-" & Trim(str(tYear + 1))
    End If

End Function
