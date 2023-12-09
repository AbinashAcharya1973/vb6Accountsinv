VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmotheroutstanding 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Other Group Out Standing"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5445
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4440
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H0000C000&
      Caption         =   "Print"
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cbogroup 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Account group"
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
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmotheroutstanding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As Recordset, rec3 As Recordset
Attribute rec2.VB_VarUserMemId = 1073938432
Attribute rec3.VB_VarUserMemId = 1073938432
Private Sub cmdprint_Click()
    frmprinter.Show vbModal
    db.Execute ("delete * from Otheroutstanding")
    Set rec1 = db.OpenRecordset("select * from ledgermaster") ' where groupid=" & Me.cboGroup.ItemData(Me.cboGroup.ListIndex))
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
            db.Execute ("insert into Otheroutstanding (Party,AccId,dr,cr) values('" & rec1("AccName") & "'," & rec1("accid") & ",0," & temp_cr - temp_dr & ")")
        End If
        If temp_dr > temp_cr Then

            db.Execute ("insert into Otheroutstanding (Party,AccId,dr,cr) values('" & rec1("AccName") & "'," & rec1("accid") & "," & temp_dr - temp_cr & ",0)")
        End If

        temp_cr = 0
        temp_dr = 0
        rec1.MoveNext
    Wend
    db.Execute ("update Otheroutstanding set GroupName='" & Me.cboGroup.Text & "'")
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
    Set db = OpenDatabase(dbname)
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Me.Top = 0
    Me.Left = 0
    Me.CrystalReport1.ReportFileName = App.Path & "\otheroutstanding.rpt"
    Set rec1 = db.OpenRecordset("select * from Groups")
    While Not rec1.EOF
        Me.cboGroup.AddItem (rec1("GroupName"))
        Me.cboGroup.ItemData(Me.cboGroup.NewIndex) = rec1("GroupID")
        rec1.MoveNext
    Wend
    If Me.cboGroup.ListCount > 0 Then
        Me.cboGroup.ListIndex = 0
        Exit Sub
errtrap:
        MsgBox Err.Description, vbOKOnly
    End If
End Sub
