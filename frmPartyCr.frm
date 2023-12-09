VERSION 5.00
Begin VB.Form frmPartyCr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Party Cr Master (Sundry Creditor)"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   8550
   Begin VB.ListBox lstregisteredids 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1005
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   7005
   End
   Begin VB.TextBox txtgid 
      Height          =   315
      Left            =   4260
      TabIndex        =   19
      Text            =   "0"
      Top             =   2730
      Visible         =   0   'False
      Width           =   675
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
      TabIndex        =   16
      Top             =   2730
      Width           =   2775
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
      Left            =   6990
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtOpeningBalance 
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
      Height          =   375
      Left            =   5430
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtAddress2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5100
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   660
      Width           =   3285
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Width           =   7005
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   2835
   End
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1350
      TabIndex        =   3
      Top             =   1590
      Width           =   2775
   End
   Begin VB.TextBox txtFax 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5100
      TabIndex        =   4
      Top             =   1620
      Width           =   3255
   End
   Begin VB.TextBox txttin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Top             =   2130
      Width           =   2805
   End
   Begin VB.Label lbldealin 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   60
      TabIndex        =   18
      Top             =   3510
      Width           =   8415
   End
   Begin VB.Label Label8 
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
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2730
      Width           =   2175
   End
   Begin VB.Label Label7 
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
      Left            =   4230
      TabIndex        =   12
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label6 
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
      Left            =   4230
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Party"
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
      TabIndex        =   10
      Top             =   120
      Width           =   855
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
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
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
      Left            =   180
      TabIndex        =   8
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Left            =   4260
      TabIndex        =   7
      Top             =   1620
      Width           =   855
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2130
      Width           =   975
   End
End
Attribute VB_Name = "frmPartyCr"
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
        Me.cbostate.SetFocus
    End If
End Sub

Private Sub cbostate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        ans = MsgBox("Save This?", vbYesNo)
        If ans = 6 Then
            Set rec1 = db.OpenRecordset("select * from Groups where GroupName='Sundry Creditor'")
            If Not rec1.EOF Then
                temp_tran_type = rec1("groupnature")
                Set rec2 = db.OpenRecordset("select max(AccID) as max_slno from LedgerMaster")
                If Not IsNull(rec2!max_slno) Then
                    temp_accid = rec2!max_slno + 1
                Else
                    temp_accid = 1
                End If
                db.Execute ("insert into PartyCr(AccId,Party,Address,Address2,Phone,Fax,TIN,statecode,gid) values(" & temp_accid & ",'" & Me.txtParty.Text & "','" & Me.txtaddress.Text & "','" & Me.TxtAddress2.Text & "','" & Me.txtPhone.Text & "','" & Me.txtFax.Text & "','" & Me.txtTin.Text & "'," & Me.cbostate.ItemData(Me.cbostate.ListIndex) & "," & Me.txtgid.Text & ")")
                db.Execute ("insert into LedgerMaster (AccID,AccName,GroupID,Dr,Cr,TransactionType,OBalance,BalanceType,Groupname,TIN,statecode) values(" & temp_accid & ",'" & Me.txtParty.Text & "'," & rec1("GroupID") & ",'-','+','" & temp_tran_type & "'," & Me.txtOpeningBalance.Text & ",'" & Me.cbodr_cr.Text & "','Sundry Creditor','" & Me.txtTin.Text & "'," & Me.cbostate.ItemData(Me.cbostate.ListIndex) & ")")
                If FORMNAME = "Purchase" Then
                    frmStockin.cboSupplier.AddItem Me.txtParty.Text
                    frmStockin.cboSupplier.ItemData(frmStockin.cboSupplier.NewIndex) = temp_accid
                    frmStockin.cboSupplier.ListIndex = frmStockin.cboSupplier.NewIndex
                    Me.txtParty.Text = ""
                    Me.txtaddress.Text = ""
                    Me.TxtAddress2.Text = ""
                    Me.txtPhone.Text = ""
                    Me.txtFax.Text = ""
                    Me.txtTin.Text = ""
                    Unload Me
                End If
                If FORMNAME = "Purchase_Barcode" Then
                    frmStockint.cboSupplier.AddItem Me.txtParty.Text
                    frmStockint.cboSupplier.ItemData(frmStockint.cboSupplier.NewIndex) = temp_accid
                    frmStockint.cboSupplier.ListIndex = frmStockint.cboSupplier.NewIndex
                    Me.txtParty.Text = ""
                    Me.txtaddress.Text = ""
                    Me.TxtAddress2.Text = ""
                    Me.txtPhone.Text = ""
                    Me.txtFax.Text = ""
                    Me.txtTin.Text = ""
                    Unload Me
                End If
            Else
                MsgBox "Group Not Exist?", vbCritical
            End If

        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Me.cbodr_cr.AddItem "Cr"
    Me.cbodr_cr.AddItem "Dr"
    Me.cbodr_cr.ListIndex = 0
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

Private Sub Label9_Click()

End Sub

Private Sub lstregisteredids_Click()
itemid = Me.lstregisteredids.ListIndex + 1
Me.txtTin.Text = JSONRec(itemid).Item("GstNo")
Me.txtPhone.Text = JSONRec(itemid).Item("MobileNo")
Me.txtParty.Text = JSONRec(itemid).Item("CompanyName")
Me.txtaddress.Text = JSONRec(itemid).Item("Address")
Me.TxtAddress2.Text = JSONRec(itemid).Item("Address1")
Me.lbldealin.Caption = "Deals In:" & JSONRec(itemid).Item("dealin")
Me.txtgid.Text = JSONRec(itemid).Item("CompanyID")

End Sub


Private Sub lstregisteredids_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtOpeningBalance.SetFocus
    Me.lstregisteredids.Visible = False
End If
If KeyCode = vbKeyEscape Then
    Me.txtParty.SetFocus
    Me.lstregisteredids.Visible = False
    Me.txtaddress.Text = ""
    Me.TxtAddress2.Text = ""
    Me.txtParty.Text = ""
    Me.txtPhone.Text = ""
    Me.txtTin.Text = ""
    Me.txtgid.Text = 0
End If
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.TxtAddress2.SetFocus
    End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtAddress2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtPhone.SetFocus
    End If
End Sub

Private Sub txtAddress2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCompany_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtTin.SetFocus
    End If
End Sub

Private Sub txtOpeningBalance_GotFocus()
    Me.txtOpeningBalance.SelStart = 0
    Me.txtOpeningBalance.SelLength = Len(Me.txtOpeningBalance.Text)
End Sub

Private Sub txtOpeningBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cbostate.SetFocus
        
    End If
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtaddress.SetFocus
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
        Me.txtFax.SetFocus
    End If
End Sub

Private Sub txttin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtOpeningBalance.SetFocus
    End If
End Sub

Private Sub txttin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
