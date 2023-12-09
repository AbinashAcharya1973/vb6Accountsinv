VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmPartyoutstanding 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sundry Debtor OutStanding"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5175
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4200
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "View/Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   351
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox CboZone 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
      Begin VB.ComboBox cboParty 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label12 
         Caption         =   "Area"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPartyoutstanding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As Recordset, rec2 As Recordset
Private Sub cboparty_Click()
'cboParty_Change
End Sub

Private Sub CboZone_Change()
Me.cboParty.Clear
Set rec2 = db.OpenRecordset("select * from PartyDr where ZoneCode =" & Me.CboZone.ItemData(Me.CboZone.ListIndex))
While Not rec2.EOF
Me.cboParty.AddItem (rec2("Party"))
Me.cboParty.ItemData(Me.cboParty.NewIndex) = rec2("AccId")
rec2.MoveNext
Wend
If Me.cboParty.ListCount > 0 Then
Me.cboParty.ListIndex = 0
End If

End Sub

Private Sub CboZone_Click()
CboZone_Change
End Sub


Private Sub cmdprint_Click()
    frmprinter.Show vbModal
    db.Execute ("delete * from PartyDrOutsatanding")
    Set rec1 = db.OpenRecordset("select * from partydr where ZoneCode=" & Me.CboZone.ItemData(Me.CboZone.ListIndex))
    While Not rec1.EOF
        Set rs = db.OpenRecordset("select sum(dr) as max_dr from ledgertran where accid=" & rec1("accid"))
        If Not IsNull(rs!max_dr) Then
            temp_dr = rs!max_dr
        End If
        Set rs = db.OpenRecordset("select sum(cr) as max_cr from ledgertran where accid=" & rec1("accid"))
        If Not IsNull(rs!max_cr) Then
            temp_cr = rs!max_cr
        End If
        Set rs = db.OpenRecordset("select * from ledgermaster where accid=" & rec1("accid"))
        If Not rs.EOF Then
            If rs("BalanceType") = "Dr" Then
                temp_dr = temp_dr + rs("OBalance")
            End If
            If rs("BalanceType") = "Cr" Then
                temp_cr = temp_cr + rs("OBalance")
            End If

        End If
        If temp_cr > temp_dr Then
            db.Execute ("insert into PartyDrOutsatanding (Party,AccId,dr,cr) values('" & rec1("Party") & "'," & rec1("accid") & ",0," & temp_cr - temp_dr & ")")
        End If
        If temp_dr > temp_cr Then

            db.Execute ("insert into PartyDrOutsatanding (Party,AccId,dr,cr) values('" & rec1("Party") & "'," & rec1("accid") & "," & temp_dr - temp_cr & ",0)")
        End If
        temp_cr = 0
        temp_dr = 0
        rec1.MoveNext
    Wend
    tempamt = 1000000000
    While tempamt < 1
        tempamt = tempamt + 1
    Wend
    tempamt = 0
    db.Close
    Me.CrystalReport1.PrinterName = Printer.DeviceName
    Me.CrystalReport1.PrinterDriver = Printer.DriverName
    Me.CrystalReport1.PrinterPort = Printer.Port
    Me.CrystalReport1.PrintReport
End Sub


Private Sub Form_Load()
On Error GoTo errtrap
Me.Top = 0
Me.Left = 0
Me.CrystalReport1.ReportFileName = App.Path & "\overalloutstanding.rpt"

Set rec1 = db.OpenRecordset("select * from Zonemaster")
While Not rec1.EOF
Me.CboZone.AddItem (rec1("ZoneName"))
Me.CboZone.ItemData(Me.CboZone.NewIndex) = rec1("SlNo")
rec1.MoveNext
Wend
If Me.CboZone.ListCount > 0 Then
Me.CboZone.ListIndex = 0
End If
Exit Sub
errtrap:
MsgBox Err.Description, vbOKOnly
End Sub

