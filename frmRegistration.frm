VERSION 5.00
Begin VB.Form frmRegistration 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registration"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtjurisdiction 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   29
      Top             =   6180
      Width           =   3555
   End
   Begin VB.TextBox txtifsc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Top             =   5100
      Width           =   3555
   End
   Begin VB.TextBox txtbankac 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   4680
      Width           =   3555
   End
   Begin VB.TextBox txtbankname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   4260
      Width           =   3555
   End
   Begin VB.TextBox txtaddress1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1740
      Width           =   3555
   End
   Begin VB.TextBox txtdealin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   11
      Top             =   5520
      Width           =   3555
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox txtgstno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   3840
      Width           =   3555
   End
   Begin VB.CommandButton cmdonline 
      Caption         =   "Register Online"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1740
      TabIndex        =   20
      Top             =   7500
      Width           =   1875
   End
   Begin VB.TextBox txtmobile 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   3420
      Width           =   2055
   End
   Begin VB.TextBox txtemail 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   3000
      Width           =   3555
   End
   Begin VB.TextBox txtstate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   2580
      Width           =   2055
   End
   Begin VB.TextBox txtpin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   1380
      Width           =   3555
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   3555
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Jurisdiction"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   6180
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "IFSC"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   5100
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Ac No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4260
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal In"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "GST No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblprid 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   22
      Top             =   420
      Width           =   5295
   End
   Begin VB.Label lblmessage 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   0
      TabIndex        =   21
      Top             =   7020
      Width           =   5295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2580
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Online Registration Form"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WinHttpReq As WinHttp.WinHttpRequest, rec1 As DAO.Recordset
      Private Const LOCALE_SSHORTDATE = &H1F
      Private Const WM_SETTINGCHANGE = &H1A
      'same as the old WM_WININICHANGE
      Private Const HWND_BROADCAST = &HFFFF&

      Private Declare Function SetLocaleInfo Lib "kernel32" Alias _
          "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As _
          Long, ByVal lpLCData As String) As Boolean
      Private Declare Function PostMessage Lib "user32" Alias _
          "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
          ByVal wParam As Long, ByVal lParam As Long) As Long
      Private Declare Function GetSystemDefaultLCID Lib "kernel32" _
          () As Long

Private Sub cmdonline_Click()
On Error GoTo errtrap
    If Me.txtname.Text <> "" And Me.txtaddress.Text <> "" And Me.txtpin.Text <> "" And Me.txtemail.Text <> "" And Me.txtmobile.Text <> "" And Me.txtgstno.Text <> "" And Me.txtpassword.Text <> "" And Me.txtaddress1.Text <> "" Then
        Me.lblmessage.Caption = "Please Wait!"
        Dim CPU As String
        Dim Customer As String
        Dim Adress As String
        Dim City As String
        Dim State As String
        Dim Email As String
        Dim Mobile As String
        Dim pwd As String
        Dim BName As String
        Dim BAcNO As String
        Dim IFSC As String
        Dim Jdiction As String
        
        
        'CPU = GetWmiDeviceSingleValue("Win32_Processor", "ProcessorID")
        Customer = Trim(Me.txtname.Text)
        Adress = Trim(Me.txtaddress.Text)
        City = Trim(Me.txtpin.Text)
        Email = Trim(Me.txtemail.Text)
        Mobile = Trim(Me.txtmobile.Text)
        pwd = Trim(Me.txtpassword.Text)
        State = Trim(Me.txtstate.Text)
        BName = Trim(Me.txtbankname.Text)
        BAcNO = Trim(Me.txtbankac.Text)
        IFSC = Trim(Me.txtifsc.Text)
        Jdiction = Trim(Me.txtjurisdiction.Text)
        ' Create an array to hold the response data.
        ' Assemble an HTTP Request.
        'On Error GoTo ErrorHandler
        WinHttpReq.Open "GET", _
                        "http://techspark.xp3.biz/enlite/mr.php?Customer=" & Customer & "&Adress=" & Adress & "&PIN=" & City & "&State=" & State & "&Email=" & Email & "&Mobile=" & Mobile & "&gstno=" & Me.txtgstno.Text & "&pwd=" & pwd & "&Address1=" & Me.txtaddress1.Text & "&bname=" & BName & "&acno=" & BAcNO & "&ifsc=" & IFSC & "&jdiction=" & Jdiction, False
        ' Send the HTTP Request.
        WinHttpReq.Send

        ' Put status and content type into status text box.
        'Text1.text = WinHttpReq.Status & " - " & WinHttpReq.StatusText
        'Me.Text2.text = WinHttpReq.ResponseText
        If WinHttpReq.ResponseText Like "*ok*" Then
            Me.lblmessage.Caption = "Thank You for Registration"
            Me.cmdonline.Enabled = False
           db.Execute "update companymaster set company='" & Me.txtname.Text & "',address='" & Me.txtaddress.Text & "',address1='" & Me.txtaddress1.Text & "',phone='" & Me.txtmobile.Text & _
                   "',email='" & Me.txtemail.Text & "',taxno='" & Me.txtgstno.Text & "',dealin='" & Me.txtdealin.Text & "',pin='" & Me.txtpin.Text & "',state='" & State & "',bankname='" & BName & "',bankacno='" & BAcNO & "',ifsc='" & IFSC & "',jurisdiction='" & Jdiction & "'"
        End If
        If WinHttpReq.ResponseText Like "*Duplicate*" Then
            Me.lblmessage.Caption = ""
            Me.lblmessage.Caption = "You are Allready registered"
            Me.cmdonline.Enabled = False
        
            
        End If
        If WinHttpReq.ResponseText Like "*Not Registered*" Then
        Me.lblmessage.Caption = "Please Contact techSpark email: info@tech-spark.in"
        End If
            
        
    Else
        MsgBox "Please Fill All These Fields!", vbCritical
    End If
    Exit Sub
errtrap:
    MsgBox "Error :" & Err.Description, vbCritical, "Connection Failed..."
End Sub
Private Sub Form_Load()
On Error GoTo errtrap
    Dim Customer As String
    Dim Adress As String
    Dim City As String
    Dim State As String
    Dim Email As String
    Dim Mobile As String
    Dim gstino As String
    Dim Address As String
    Dim dealin As String
    Dim BankName As String
    Dim BankAcno As String
    Dim IFSC As String
    Dim Jurisdiction As String
    
    ' Create an instance of the WinHTTPRequest ActiveX object.
    Set WinHttpReq = New WinHttpRequest
    
    Set rec1 = db.OpenRecordset("Select * from companymaster")
    If Not rec1.EOF Then
        Me.lblmessage.Caption = "Please Wait, Verifing Registration!"
        Customer = IIf(IsNull(rec1("company")), "", Trim(rec1("company")))
        Adress = IIf(IsNull(rec1("address")), "", Trim(rec1("address")))
        City = IIf(IsNull(rec1("pin")), "", Trim(rec1("pin")))
        Email = IIf(IsNull(rec1("email")), "", Trim(rec1("email")))
        Mobile = IIf(IsNull(rec1("phone")), "", Trim(rec1("phone")))
        gstino = IIf(IsNull(rec1("taxno")), "", Trim(rec1("taxno")))
        Address = IIf(IsNull(rec1("address1")), "", Trim(rec1("address1")))
        dealin = IIf(IsNull(rec1("dealin")), "", Trim(rec1("dealin")))
        State = IIf(IsNull(rec1("state")), "", Trim(rec1("state")))
        BankName = IIf(IsNull(rec1("bankname")), "", Trim(rec1("bankname")))
        BankAcno = IIf(IsNull(rec1("bankacno")), "", Trim(rec1("bankacno")))
        IFSC = IIf(IsNull(rec1("ifsc")), "", Trim(rec1("ifsc")))
        Jurisdiction = IIf(IsNull(rec1("jurisdiction")), "", Trim(rec1("jurisdiction")))
        
        WinHttpReq.Open "GET", _
                        "http://techspark.xp3.biz/enlite/checkreg.php?Customer=" & Customer & "&PIN=" & City & "&Email=" & Email & "&Mobile=" & Mobile & "&gstno=" & gstino, False
        
        WinHttpReq.Send
        If WinHttpReq.ResponseText Like "*Registered*" Then
            Me.lblmessage.Caption = "This Copy of Software is Registered"
            Me.txtname.Text = Customer
            Me.txtaddress.Text = Adress
            Me.txtpin.Text = City
            Me.txtmobile.Text = Mobile
            Me.txtgstno.Text = gstino
            Me.txtemail.Text = Email
            Me.txtaddress1 = Address
            Me.txtdealin.Text = dealin
            Me.txtstate.Text = State
            Me.txtbankname.Text = BankName
            Me.txtbankac.Text = BankAcno
            Me.txtifsc.Text = IFSC
            Me.txtjurisdiction.Text = Jurisdiction
        Else
            Me.lblmessage.Caption = WinHttpReq.ResponseText
            cmdonline.Enabled = True
        End If

    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text3_Change()

End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtpin.SetFocus
End If
End Sub
Private Sub txtcity_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtstate.SetFocus
End If
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtmobile.SetFocus
End If
End Sub

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtaddress.SetFocus
End If
End Sub

Private Sub txtstate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtemail.SetFocus
End If
End Sub
