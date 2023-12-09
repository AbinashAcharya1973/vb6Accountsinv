VERSION 5.00
Begin VB.Form frmPartyDr 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Party Dr Master (Sundry Debtor)"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstregisteredids 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   1440
      TabIndex        =   40
      Top             =   600
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8535
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtgid 
         Height          =   315
         Left            =   3060
         TabIndex        =   42
         Text            =   "0"
         Top             =   7500
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cbostate 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   7440
         Width           =   1455
      End
      Begin VB.TextBox txtdiscount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         TabIndex        =   36
         Text            =   "0"
         Top             =   6960
         Width           =   1215
      End
      Begin VB.TextBox TxtMobile 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   4560
         Width           =   2535
      End
      Begin VB.TextBox TxtPin 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox TxtCity 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox TxtState 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   2640
         Width           =   2535
      End
      Begin VB.ComboBox CboZone 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox TxtDueDay 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Text            =   "0"
         Top             =   6480
         Width           =   1215
      End
      Begin VB.TextBox TxtOstCst 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox txtParty 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   120
         Width           =   4695
      End
      Begin VB.TextBox txtaddress1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtAddress2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox txtemail 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   3120
         Width           =   4695
      End
      Begin VB.TextBox txtPhone 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox txtCreditLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   8040
         Width           =   975
      End
      Begin VB.TextBox txtFax 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   5040
         Width           =   2535
      End
      Begin VB.TextBox txttin 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   6000
         Width           =   2535
      End
      Begin VB.ComboBox cbodr_cr 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   6960
         Width           =   975
      End
      Begin VB.TextBox txtOpeningBalance 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   6960
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Press F2 to List Online Registered Parties"
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
         Left            =   1320
         TabIndex        =   39
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label18 
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   7440
         Width           =   2175
      End
      Begin VB.Label Label17 
         Caption         =   "Tr.Dis%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   35
         Top             =   6960
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Pin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Zone"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Due Days"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   29
         Top             =   6480
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "OST / CST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Address1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Address2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Email @"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Credit Limit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "GSTIN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Op. Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   6960
         Width           =   2175
      End
   End
   Begin VB.Label lbldealin 
      Caption         =   "Label20"
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
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   8640
      Width           =   6135
   End
End
Attribute VB_Name = "frmPartyDr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WinHttpReq As WinHttp.WinHttpRequest
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

Dim rs As Recordset, rec1 As Recordset, rec2 As Recordset, JSONRec As Object



Private Sub cbodr_cr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtdiscount.SetFocus
    End If
End Sub

Private Sub cbostate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cmdSave.SetFocus
End If
End Sub

Private Sub CboZone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtPhone.SetFocus
    End If
End Sub

Private Sub CmdSave_Click()
On Error GoTo errtrap
    ans = MsgBox("Save This?", vbYesNo)
    If ans = 6 Then
        Set rec1 = db.OpenRecordset("select * from Groups where GroupName='Sundry Debtor'")
        If Not rec1.EOF Then
            temp_tran_type = rec1("groupnature")
            Set rec2 = db.OpenRecordset("select max(AccID) as max_slno from LedgerMaster")
            If Not IsNull(rec2!max_slno) Then
                temp_accid = rec2!max_slno + 1
            Else
                temp_accid = 1
            End If
            db.Execute ("insert into PartyDr (AccId,Party,Address,Address2,City,Pin,State,EMail,Zone,Phone,Mobile,Fax,OstCst,Tin,CrLimit,DueDays,OPBalance,OPType,ZoneCode,Discount,statecode,gid) values(" & temp_accid & ",'" & Me.txtParty.Text & "','" & Me.txtaddress1.Text & "','" & Me.TxtAddress2.Text & "','" & Me.TxtCity.Text & "','" & Me.txtpin.Text & "','" & Me.txtstate.Text & "','" & Me.txtemail.Text & "','" & Me.CboZone.Text & "','" & Me.txtPhone.Text & "','" & Me.txtmobile.Text & "','" & Me.txtFax.Text & "','" & Me.TxtOstCst.Text & "','" & Me.txtTin.Text & "'," & Me.txtCreditLimit.Text & "," & Me.TxtDueDay.Text & "," & Me.txtOpeningBalance.Text & ",'" & Me.cbodr_cr.Text & "'," & Me.CboZone.ItemData(Me.CboZone.ListIndex) & "," & Me.txtdiscount.Text & "," & Me.cbostate.ItemData(Me.cbostate.ListIndex) & "," & Me.txtgid.Text & ")")
            db.Execute ("insert into LedgerMaster (AccID,AccName,GroupID,Dr,Cr,TransactionType,OBalance,BalanceType,Groupname,Address1,Address2,statecode) values(" & temp_accid & ",'" & Me.txtParty.Text & "'," & rec1("GroupID") & ",'+','-','" & temp_tran_type & "'," & Me.txtOpeningBalance.Text & ",'" & Me.cbodr_cr.Text & "','Sundry Debtor','" & Me.txtaddress1.Text & "','" & Me.TxtAddress2.Text & "'," & Me.cbostate.ItemData(Me.cbostate.ListIndex) & ")")
            If formid = 100 Then
                frmInvoice.cboParty.AddItem Me.txtParty.Text
                frmInvoice.cboParty.ItemData(frmInvoice.cboParty.NewIndex) = temp_accid
                frmInvoice.cboParty.ListIndex = frmInvoice.cboParty.NewIndex
                Me.txtParty.Text = ""
                Me.txtaddress1.Text = ""
                Me.TxtAddress2.Text = ""
                Me.txtPhone.Text = ""
                Me.TxtCity.Text = ""
                Me.txtstate.Text = ""
                Me.txtpin.Text = ""
                Me.txtmobile.Text = ""
                Me.TxtDueDay.Text = 0
                Me.txtFax.Text = ""
                Me.txtemail.Text = ""
                Me.txtCreditLimit.Text = "0.00"
                Me.txtOpeningBalance.Text = "0.00"
                Me.txtgid.Text = 0
                Unload Me
            End If
            If formid = 1001 Then
                frmInvoiceR.cboParty.AddItem Me.txtParty.Text
                frmInvoiceR.cboParty.ItemData(frmInvoiceR.cboParty.NewIndex) = temp_accid
                frmInvoiceR.cboParty.ListIndex = frmInvoiceR.cboParty.NewIndex
                Me.txtParty.Text = ""
                Me.txtaddress1.Text = ""
                Me.TxtAddress2.Text = ""
                Me.txtPhone.Text = ""
                Me.TxtCity.Text = ""
                Me.txtstate.Text = ""
                Me.txtpin.Text = ""
                Me.txtmobile.Text = ""
                Me.TxtDueDay.Text = 0
                Me.txtFax.Text = ""
                Me.txtemail.Text = ""
                Me.txtCreditLimit.Text = "0.00"
                Me.txtOpeningBalance.Text = "0.00"
                Me.txtgid.Text = 0
                Unload Me
            End If

        End If

    Else
        MsgBox "Group Not Exist?", vbCritical
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 1000
    Me.Left = 1400
    Me.cbodr_cr.AddItem "Dr"
    Me.cbodr_cr.AddItem "Cr"
    Me.cbodr_cr.ListIndex = 0
    Set rec1 = db.OpenRecordset("select * from Zonemaster")
    While Not rec1.EOF
        Me.CboZone.AddItem (rec1("ZoneName"))
        Me.CboZone.ItemData(Me.CboZone.NewIndex) = rec1("SlNo")
        rec1.MoveNext
    Wend
    If Me.CboZone.ListCount > 0 Then
        Me.CboZone.ListIndex = 0
    End If
    Set rec1 = db.OpenRecordset("select * from statecode")
    While Not rec1.EOF
        Me.cbostate.AddItem (rec1("statename"))
        Me.cbostate.ItemData(Me.cbostate.NewIndex) = rec1("stcode")
        rec1.MoveNext
    Wend
    If Me.cbostate.ListCount > 0 Then
        Me.cbostate.ListIndex = 1
    End If
    Set WinHttpReq = New WinHttpRequest
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub Label20_Click()

End Sub

Private Sub lstregisteredids_Click()
itemid = Me.lstregisteredids.ListIndex + 1
Me.txtTin.Text = JSONRec(itemid).Item("GstNo")
Me.txtPhone.Text = JSONRec(itemid).Item("MobileNo")
Me.txtParty.Text = JSONRec(itemid).Item("CompanyName")
Me.txtaddress1.Text = JSONRec(itemid).Item("Address")
Me.TxtAddress2.Text = JSONRec(itemid).Item("Address1")
Me.lbldealin.Caption = "Deals In:" & JSONRec(itemid).Item("dealin")
Me.txtgid.Text = JSONRec(itemid).Item("CompanyID")
Me.txtemail.Text = JSONRec(itemid).Item("emailid")
Me.txtpin.Text = JSONRec(itemid).Item("PIN")
Me.txtstate.Text = JSONRec(itemid).Item("State")
End Sub

Private Sub lstregisteredids_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtParty.SetFocus
        Me.lstregisteredids.Visible = False
    End If
    If KeyCode = vbKeyEscape Then
        Me.txtParty.SetFocus
        Me.lstregisteredids.Visible = False
        Me.txtTin.Text = ""
        Me.txtPhone.Text = ""
        Me.txtParty.Text = ""
        Me.txtaddress1.Text = ""
        Me.TxtAddress2.Text = ""
        Me.lbldealin.Caption = ""
        Me.txtgid.Text = 0
        Me.txtemail.Text = ""
        Me.txtpin.Text = ""
        Me.txtstate.Text = ""

    End If

End Sub

Private Sub TxtAddress1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtAddress2.SetFocus
    End If
End Sub

Private Sub TxtAddress2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtCity.SetFocus
    End If
End Sub

Private Sub txtcity_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtpin.SetFocus
    End If
End Sub

Private Sub txtCreditLimit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtDueDay.SetFocus
    End If
End Sub

Private Sub txtdiscount_GotFocus()
    Me.txtdiscount.SelStart = 0
    Me.txtdiscount.SelLength = Len(Me.txtdiscount.Text)
End Sub

Private Sub txtdiscount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cbostate.SetFocus
    End If
End Sub

Private Sub TxtDueDay_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtOpeningBalance.SetFocus
    End If
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.CboZone.SetFocus
    End If
End Sub

Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtOstCst.SetFocus
    End If
End Sub

Private Sub TxtMobile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtFax.SetFocus
    End If
End Sub

Private Sub txtOpeningBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cbodr_cr.SetFocus
    End If
End Sub

Private Sub TxtOstCst_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtTin.SetFocus
    End If
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtaddress1.SetFocus
    End If
    If KeyCode = vbKeyF2 Then
        Me.lstregisteredids.Visible = True
        Me.lstregisteredids.Clear
        Dim strRes As String, reccount, countRecords
        
        
        Set rec1 = db.OpenRecordset("select * from companymaster")
        If rec1("gid") <> 0 Then
            WinHttpReq.Open "GET", _
                            "http://techspark.xp3.biz/enlite/getpartyinfo.php?searchkey=" & Me.txtParty.Text, False
            WinHttpReq.Send
            If WinHttpReq.ResponseText Like "*Not Found*" Then
                MsgBox "Not Found", vbCritical
            Else
                strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
                countRecords = UBound(Split(strRes, "CompanyName"))
                Set JSONRec = JSON.parse(strRes)
                If countRecords > 1 Then
                    reccount = 1
                    While reccount <= countRecords
                        Me.lstregisteredids.AddItem JSONRec(reccount).Item("CompanyName") & "-" & JSONRec(reccount).Item("Address")
                        reccount = reccount + 1
                    Wend
                Else
                    Me.lstregisteredids.AddItem JSONRec.Item("CompanyName") & "-" & JSONRec(reccount).Item("Address")
                End If
            End If
        End If
        
    End If
    If KeyCode = vbKeyDown Then
        Me.lstregisteredids.SetFocus
    End If
End Sub

Private Sub txtParty_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtmobile.SetFocus
    End If
End Sub

Private Sub TxtPin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtstate.SetFocus
    End If
End Sub

Private Sub txtstate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtemail.SetFocus
    End If
End Sub
Private Sub txttin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
            Me.txtCreditLimit.SetFocus
    End If
End Sub
