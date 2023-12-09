VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmcreditnote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Note"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6990
   Begin VB.TextBox txtinvno 
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
      Height          =   315
      Left            =   180
      TabIndex        =   31
      Text            =   "0"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton CMDDELETE 
      BackColor       =   &H0080FF80&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4055
      TabIndex        =   30
      Top             =   4050
      Width           =   1000
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2760
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.TextBox txtfinal 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "0"
      Top             =   3510
      Width           =   1575
   End
   Begin VB.TextBox txtroundv 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   3510
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "0"
      Top             =   3510
      Width           =   1575
   End
   Begin VB.TextBox txttaxamount 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0"
      Top             =   3510
      Width           =   1575
   End
   Begin VB.TextBox txtcgst 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0"
      Top             =   2940
      Width           =   1575
   End
   Begin VB.TextBox txtsgst 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   2940
      Width           =   1575
   End
   Begin VB.TextBox txtgst 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   240
      TabIndex        =   18
      Text            =   "18"
      Top             =   2940
      Width           =   1575
   End
   Begin VB.TextBox txtparticulars 
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
      Height          =   315
      Left            =   240
      TabIndex        =   16
      Top             =   2310
      Width           =   4635
   End
   Begin VB.TextBox txtnet 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "0"
      Top             =   3510
      Width           =   1575
   End
   Begin VB.TextBox txtigst 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   5130
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   2940
      Width           =   1575
   End
   Begin VB.TextBox txthsn 
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
      Height          =   315
      Left            =   5130
      TabIndex        =   10
      Top             =   2310
      Width           =   1575
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H0080FF80&
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2985
      TabIndex        =   9
      Top             =   4050
      Width           =   1000
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0080FF80&
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
      Height          =   435
      Left            =   1935
      TabIndex        =   8
      Top             =   4050
      Width           =   1000
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   5130
      TabIndex        =   6
      Text            =   "0"
      Top             =   1650
      Width           =   1575
   End
   Begin VB.ComboBox cboSupplier 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1620
      Width           =   4695
   End
   Begin VB.TextBox txtslno 
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
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Text            =   "0"
      Top             =   390
      Width           =   1575
   End
   Begin MSMask.MaskEdBox txtStockindate 
      Height          =   315
      Left            =   5130
      TabIndex        =   0
      Top             =   390
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtinvdate 
      Height          =   315
      Left            =   5160
      TabIndex        =   32
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Inv No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   34
      Top             =   750
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Inv Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5190
      TabIndex        =   33
      Top             =   750
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5190
      TabIndex        =   29
      Top             =   3300
      Width           =   1545
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Round"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3510
      TabIndex        =   27
      Top             =   3300
      Width           =   1545
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   3300
      Width           =   1545
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "CGST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   23
      Top             =   2730
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "SGST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1860
      TabIndex        =   21
      Top             =   2730
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "GST%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   2730
      Width           =   1515
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   2100
      Width           =   1665
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Gross"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1860
      TabIndex        =   15
      Top             =   3300
      Width           =   1545
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "IGST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5130
      TabIndex        =   13
      Top             =   2730
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "HSN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5130
      TabIndex        =   11
      Top             =   2100
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5130
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   5
      Top             =   1410
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sl No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   2
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "frmcreditnote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec3 As Recordset, rec4 As Recordset, DELETEDEBIT As Boolean
Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtamount.SetFocus
End If
End Sub

Private Sub CmdDelete_Click()
DELETEDEBIT = True
Me.txtslno.SetFocus
End Sub

Private Sub cmdprint_Click()
frmcrnoteprint.Show 0
End Sub

Private Sub CmdSave_Click()
    ans = MsgBox("Save the Credit Note?", vbYesNo)
    If ans = 6 Then
        db.Execute ("insert into creditnote (slno,ddate,partyname,accid,amount,gst,sgst,cgst,igst,totaltax,gross,roundoff,netamount,particulars,hsn,invno,invdate) values(" & Me.txtslno.Text & ",'" & Me.txtStockindate.Text & "','" & Me.cboSupplier.Text & "'," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & "," & Me.txtamount.Text & "," & Me.txtgst.Text & "," & Me.txtsgst.Text & "," & Me.txtcgst.Text & "," & Me.txtigst.Text & "," & Me.txttaxamount.Text & "," & Me.txtnet.Text & "," & Me.txtroundv.Text & "," & Me.txtfinal.Text & ",'" & Me.txtparticulars.Text & "','" & Me.txthsn.Text & "','" & Me.txtInvno.Text & "','" & Me.txtinvdate.Text & "')")
        Set rec1 = db.OpenRecordset("select * from ledgermaster where accname='DISCOUNT ON SALES'")
        If Not rec1.EOF Then
            cr_accname = rec1("accname")
            cr_groupid = rec1("groupid")
            cr_accid = rec1("accid")
            Set rec3 = db.OpenRecordset("select max(slno) as max_slno from ledgertran")
            If Not rec3!max_slno Then
                tempslno = rec3!max_slno + 1
            Else
                tempslno = 1
            End If
            Set rec4 = db.OpenRecordset("select * from ledgermaster where accid=" & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex))
            If Not rec4.EOF Then
                party_groupid = rec4("groupid")
            End If
            db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & tempslno & ",'" & Me.txtStockindate.Text & "','By DISCOUNT ON SALE',0," & Me.txtfinal.Text & ",0," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ",'" & Trim(Me.txtparticulars.Text) & "','Credit Note'," & Me.txtslno.Text & "," & party_groupid & "," & cr_accid & ")")
            Set rec4 = db.OpenRecordset("select max(slno) as max_slno from ledgertran")
            If Not rec4!max_slno Then
                tempslno = rec4!max_slno + 1
            Else
                tempslno = 1
            End If
            db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & tempslno & ",'" & Me.txtStockindate.Text & "','By " & Me.cboSupplier.Text & "',0," & Me.txtamount.Text & ",0," & cr_accid & ",'" & Trim(Me.txtparticulars.Text) & "','Credit Note'," & Me.txtslno.Text & "," & cr_groupid & "," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ")")
            If Val(Me.txtsgst.Text) > 0 Then
                Set rec3 = db.OpenRecordset("select * from ledgermaster where accname='SGST'")
                If Not rec3.EOF Then
                    sgst_accid = rec3("accid")
                    sgst_groupid = rec3("accid")
                    Set rec4 = db.OpenRecordset("select max(slno) as max_slno from ledgertran")
                    If Not rec4!max_slno Then
                        tempslno = rec4!max_slno + 1
                    Else
                        tempslno = 1
                    End If
                    db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & tempslno & ",'" & Me.txtStockindate.Text & "','To " & Me.cboSupplier.Text & "'," & Me.txtsgst.Text & ",0,0," & sgst_accid & ",'" & Trim(Me.txtparticulars.Text) & "','Credit Note'," & Me.txtslno.Text & "," & sgst_groupid & "," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ")")
                End If
                Set rec3 = db.OpenRecordset("select * from ledgermaster where accname='CGST'")
                If Not rec3.EOF Then
                    cgst_accid = rec3("accid")
                    cgst_groupid = rec3("accid")
                    Set rec4 = db.OpenRecordset("select max(slno) as max_slno from ledgertran")
                    If Not rec4!max_slno Then
                        tempslno = rec4!max_slno + 1
                    Else
                        tempslno = 1
                    End If
                    db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & tempslno & ",'" & Me.txtStockindate.Text & "','To " & Me.cboSupplier.Text & "'," & Me.txtcgst.Text & ",0,0," & cgst_accid & ",'" & Trim(Me.txtparticulars.Text) & "','Credit Note'," & Me.txtslno.Text & "," & cgst_groupid & "," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ")")
                End If
            End If
            If Val(Me.txtigst.Text) > 0 Then
                Set rec3 = db.OpenRecordset("select * from ledgermaster where accname='IGST'")
                If Not rec3.EOF Then
                    igst_accid = rec3("accid")
                    igst_groupid = rec3("accid")
                    Set rec4 = db.OpenRecordset("select max(slno) as max_slno from ledgertran")
                    If Not rec4!max_slno Then
                        tempslno = rec4!max_slno + 1
                    Else
                        tempslno = 1
                    End If
                    db.Execute ("insert into LedgerTran (Slno,TDate,Particulars,Dr,Cr,Balance,AccId,Remarks,VoucherType,VoucherSlno,GroupId,TranAccId) values(" & tempslno & ",'" & Me.txtStockindate.Text & "','To " & Me.cboSupplier.Text & "'," & Me.txtigst.Text & ",0,0," & igst_accid & ",'" & Trim(Me.txtparticulars.Text) & "','Credit Note'," & Me.txtslno.Text & "," & igst_groupid & "," & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex) & ")")
                End If
            End If
        End If
        Me.txtamount.Text = "0.00"
        Me.txtparticulars.Text = ""
        Me.txthsn.Text = ""
        Me.txtgst.Text = "0"
        Me.txtigst.Text = 0
        Me.txtcgst.Text = 0
        Me.txtsgst.Text = 0
        Me.txttaxamount.Text = 0
        Me.txtfinal.Text = 0
        Me.txtnet.Text = 0
        Me.txtroundv.Text = 0
        Me.txtslno.Text = Val(Me.txtslno.Text) + 1
        Me.txtStockindate.SetFocus
    End If
End Sub



Private Sub Form_Load()
    Set rec1 = db.OpenRecordset("select max(slno) as maxslno from creditnote")
    If Not IsNull(rec1!maxslno) Then
        Me.txtslno.Text = rec1!maxslno + 1
    Else
        Me.txtslno.Text = 1
    End If
    Set rec1 = db.OpenRecordset("select * from ledgermaster where groupname='Sundry Debtor'")
    If Not rec1.EOF Then
        While Not rec1.EOF
            Me.cboSupplier.AddItem (rec1("accname"))
            Me.cboSupplier.ItemData(Me.cboSupplier.NewIndex) = rec1("AccId")
            rec1.MoveNext
        Wend
    End If
    If Me.cboSupplier.ListCount > 0 Then
        Me.cboSupplier.ListIndex = 0
    End If
    Me.txtStockindate.Text = Format(Date, "dd/mm/yyyy")
    
    DELETEDEBIT = False
End Sub

Private Sub txtAmount_GotFocus()
Me.txtamount.SelStart = 0
Me.txtamount.SelLength = Len(Me.txtamount.Text)
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtparticulars.SetFocus
End If
End Sub

Private Sub txtgst_GotFocus()
Me.txtgst.SelStart = 0
Me.txtgst.SelLength = Len(Me.txtgst.Text)
End Sub

Private Sub txtgst_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Set rec1 = db.OpenRecordset("select * from partydr where accid=" & Me.cboSupplier.ItemData(Me.cboSupplier.ListIndex))
    If Not rec1.EOF Then
        If rec1("statecode") = 21 Then
            Me.txtcgst.Text = Round(Val(Me.txtamount.Text) * ((Val(Me.txtgst.Text) / 2) / 100), 2)
            Me.txtsgst.Text = Round(Val(Me.txtamount.Text) * ((Val(Me.txtgst.Text) / 2) / 100), 2)
            Me.txttaxamount.Text = Round(Val(Me.txtcgst.Text) + Val(Me.txtsgst.Text), 2)
            Me.txtnet.Text = Val(Me.txtamount.Text) + Val(Me.txttaxamount.Text)
            Me.txtfinal.Text = Round(Val(Me.txtnet.Text), 2)
            Me.txtroundv.Text = Round(Val(Me.txtfinal.Text) - Round(Val(Me.txtnet.Text), 2), 2)
        Else
            Me.txtigst.Text = Round(Val(Me.txtamount.Text) * (Val(Me.txtgst.Text) / 100), 2)
            Me.txttaxamount.Text = Me.txtigst.Text
            Me.txtnet.Text = Val(Me.txtamount.Text) + Val(Me.txttaxamount.Text)
            Me.txtfinal.Text = Round(Val(Me.txtnet.Text))
            Me.txtroundv.Text = Round(Val(Me.txtfinal.Text) - Round(Val(Me.txtnet.Text), 2), 2)
        End If
        Me.txtnet.SetFocus
    End If
    
End If
End Sub

Private Sub txthsn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtgst.SetFocus
End If
End Sub

Private Sub txtInvdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cboSupplier.SetFocus
End If
End Sub

Private Sub txtInvno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtinvdate.SetFocus
End If
End Sub

Private Sub txtnet_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.cmdsave.SetFocus
End If
End Sub

Private Sub txtparticulars_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txthsn.SetFocus
End If
End Sub

Private Sub txtslno_GotFocus()
Me.txtslno.SelStart = 0
Me.txtslno.SelLength = Len(Me.txtslno.Text)
End Sub

Private Sub txtSlno_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errtrap
If KeyCode = 13 Then
    If DELETEDEBIT = True Then
        ans = MsgBox("DELETE THE CREDIT NOTE?", vbYesNo)
        If ans = 6 Then
            db.Execute ("DELETE * FROM LEDGERTRAN WHERE VoucherType='Credit Note' and VoucherSlno=" & Me.txtslno.Text)
            db.Execute ("DELETE * FROM CREDITnote WHERE Slno=" & Me.txtslno.Text)
            MsgBox "Credit Note Deleted Successfuly", vbOKOnly
        End If
    End If
End If
Exit Sub
errtrap:
MsgBox Err.Description, vbCritical
End Sub

Private Sub txtStockindate_GotFocus()
Me.txtStockindate.SelStart = 0
Me.txtStockindate.SelLength = Len(Me.txtStockindate.Text)
End Sub

Private Sub txtStockindate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.txtInvno.SetFocus
End If
End Sub
