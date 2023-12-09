VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmdchallanreport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maydays Report/Delivery Challan"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6135
   Begin VB.CommandButton cmdprint 
      Caption         =   "Generate Report"
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
      Left            =   2190
      TabIndex        =   6
      Top             =   1620
      Width           =   1755
   End
   Begin VB.ComboBox cboparty 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   750
      Width           =   4275
   End
   Begin MSMask.MaskEdBox fdate 
      Height          =   345
      Left            =   1380
      TabIndex        =   2
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox tdate 
      Height          =   345
      Left            =   4140
      TabIndex        =   3
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
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
      Height          =   285
      Left            =   330
      TabIndex        =   4
      Top             =   750
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "To"
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
      Left            =   3300
      TabIndex        =   1
      Top             =   270
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "From"
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
      Left            =   360
      TabIndex        =   0
      Top             =   270
      Width           =   975
   End
End
Attribute VB_Name = "frmdchallanreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As Recordset, rec1 As Recordset

Private Sub cmdprint_Click()
    Dim strLine As String
    Dim fso As New FileSystemObject
    Dim fsoStream As TextStream
    Dim headings() As String
    Me.MousePointer = vbHourglass
    Dim from_dt() As String
    Dim to_dt() As String
    Dim tempfilename



    Set excelApp = CreateObject("Excel.Application")
    Set excelWB = excelApp.Workbooks.Add
    Set excelWS = excelWB.Worksheets(1)
    excelWS.Cells(1, 1).Value = "DATE"
    Set rec = db.OpenRecordset("select distinct productname from deliverychallandetails")
    tcol = 1
    While Not rec.EOF
        tcol = tcol + 1
        excelWS.Cells(1, tcol).Value = UCase(rec("productname"))
        rec.MoveNext
    Wend
    'excelWS.Cells(RowCount, 2).Value = rec1("AccName")
    xcol = 2
    crow = 2

    Set rec = db.OpenRecordset("select distinct productname from deliverychallandetails")
    idx = 0
    While Not rec.EOF
        idx = idx + 1
        rec.MoveNext
    Wend
    ReDim headings(idx)
    rec.MoveFirst
    tempid = 0
    While Not rec.EOF
        excelWS.Cells(1, xcol).Value = UCase(rec("productname"))
        headings(tempid) = rec("productname")
        tempid = tempid + 1
        xcol = xcol + 1
        rec.MoveNext
    Wend
    Set rec = db.OpenRecordset("select distinct challandaate from deliverychallan where accid=" & Me.cboParty.ItemData(Me.cboParty.ListIndex) & " and challandaate between#" & Format(CDate(Me.fdate.Text), "mm/dd/yyyy") & "# and #" & Format(CDate(Me.tdate.Text), "mm/dd/yyyy") & "#")
    While Not rec.EOF
        tempid = 0
        excelWS.Cells(crow, 1).Value = "'" & Format(rec("challandaate"), "dd/mm/yyyy")
        While tempid < idx
            Set rec1 = db.OpenRecordset("select sum(deliverychallandetails.qty) as tqty from deliverychallan inner join deliverychallandetails on deliverychallan.challanno=deliverychallandetails.challanno where deliverychallandetails.productname='" & headings(tempid) & "' and deliverychallan.challandaate=#" & Format(rec("challandaate"), "mm/dd/yyyy") & "#")
            If Not IsNull(rec1!tqty) Then
                excelWS.Cells(crow, tempid + 2).Value = rec1!tqty
            End If
            tempid = tempid + 1
        Wend
        crow = crow + 1
        rec.MoveNext
    Wend

    excelApp.Visible = True
    Me.MousePointer = 0
    'excelApp.Workbooks.Close

    Set ExcelSheet = Nothing
    Set excelWB = Nothing
    Set excelWS = Nothing

End Sub

Private Sub fdate_GotFocus()
Me.fdate.SelStart = 0
Me.fdate.SelLength = Len(Me.fdate.Text)
End Sub

Private Sub Form_Load()
Set rec = db.OpenRecordset("select * from partydr")
While Not rec.EOF
    Me.cboParty.AddItem rec("party")
    Me.cboParty.ItemData(Me.cboParty.NewIndex) = rec("accid")
    rec.MoveNext
Wend
If Me.cboParty.ListCount > 0 Then
    Me.cboParty.ListIndex = 0
End If
Me.fdate.Text = Format(Date, "dd/mm/yyyy")
Me.tdate.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub MaskEdBox1_Change()

End Sub

Private Sub MaskEdBox1_GotFocus()

End Sub

Private Sub tdate_GotFocus()
Me.tdate.SelStart = 0
Me.tdate.SelLength = Len(Me.tdate.Text)
End Sub
