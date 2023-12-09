VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmmarginsheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Margin Sheet"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5970
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
      Height          =   345
      Left            =   2205
      TabIndex        =   0
      Top             =   1200
      Width           =   1725
   End
   Begin MSMask.MaskEdBox txtfrom 
      Height          =   375
      Left            =   1185
      TabIndex        =   1
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtto 
      Height          =   375
      Left            =   3945
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
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
      Left            =   570
      TabIndex        =   4
      Top             =   390
      Width           =   735
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
      Height          =   315
      Left            =   3585
      TabIndex        =   3
      Top             =   390
      Width           =   735
   End
End
Attribute VB_Name = "frmmarginsheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As Recordset, rec1 As Recordset, rec2 As Recordset
Private Sub cmdprint_Click()
    Set excelApp = CreateObject("Excel.Application")
    Set excelWB = excelApp.Workbooks.Add
    Set excelWS = excelWB.Worksheets(1)
    firstdate = CDate(Me.txtfrom.Text)
    lastdate = CDate(Me.txtto.Text)
    Set rec1 = db.OpenRecordset("select * from itemmaster where item like 'BREAK*'")
    If Not rec1.EOF Then
        tempbreakfastrate = rec1("salerate")
    Else
        tempbreakfastrate = 0
    End If
    Set rec1 = db.OpenRecordset("select * from itemmaster where item like 'LUNCH*'")
    If Not rec1.EOF Then
        templunch = rec1("salerate")
    Else
        templunch = 0
    End If
    Set rec1 = db.OpenRecordset("select * from itemmaster where item like 'DINNER*'")
    If Not rec1.EOF Then
        tempdinner = rec1("salerate")
    Else
        tempdinner = 0
    End If
    excelWS.Cells(1, 1).Value = "MARGIN FROM " & Me.txtfrom.Text & " TO " & Me.txtto.Text
    excelWS.Cells(2, 1).Value = "DATE"
    excelWS.Cells(2, 2).Value = "BREAKFAST TOKENS"
    excelWS.Cells(2, 3).Value = "BREAKFAST AMOUNT"
    excelWS.Cells(2, 4).Value = "LUNCH TOKENS"
    excelWS.Cells(2, 5).Value = "LUNCH AMOUNT"
    excelWS.Cells(2, 6).Value = "DINNER TOKENS"
    excelWS.Cells(2, 7).Value = "DINNER AMOUNT"
    excelWS.Cells(2, 8).Value = "MATERIAL CONSUMPTION"
    excelWS.Cells(2, 9).Value = "MARGIN"
    
    RC = 3
    While DateDiff("d", firstdate, lastdate) <> -1
        Set rec1 = db.OpenRecordset("select sum(deliverychallandetails.qty) as tqty from deliverychallan inner join deliverychallandetails on deliverychallan.challanno=deliverychallandetails.challanno where deliverychallan.challandaate=#" & Format(firstdate, "mm/dd/yyyy") & "# and deliverychallandetails.productname like 'BREAK*'")
        If Not IsNull(rec1!tqty) Then
            tempbfqty = rec1!tqty
        Else
            tempbfqty = 0
        End If
        Set rec1 = db.OpenRecordset("select sum(deliverychallandetails.qty) as tqty from deliverychallan inner join deliverychallandetails on deliverychallan.challanno=deliverychallandetails.challanno where deliverychallan.challandaate=#" & Format(firstdate, "mm/dd/yyyy") & "# and deliverychallandetails.productname like 'LUNCH*'")
        If Not IsNull(rec1!tqty) Then
            templqty = rec1!tqty
        Else
            tempLfqty = 0
        End If
        Set rec1 = db.OpenRecordset("select sum(deliverychallandetails.qty) as tqty from deliverychallan inner join deliverychallandetails on deliverychallan.challanno=deliverychallandetails.challanno where deliverychallan.challandaate=#" & Format(firstdate, "mm/dd/yyyy") & "# and deliverychallandetails.productname like 'DINNER*'")
        If Not IsNull(rec1!tqty) Then
            tempdqty = rec1!tqty
        Else
            tempdqty = 0
        End If
        Set rec1 = db.OpenRecordset("select sum(OUTWARDchallandetails.AMOUNT) as tAMOUNT from OUTWARDchallanHEAD inner join OUTWARDchallandetails on outwardchallanhead.challanno=outwardchallandetails.challanno where outwardchallanhead.challandaate=#" & Format(firstdate, "mm/dd/yyyy") & "#")
        If Not IsNull(rec1!tamount) Then
            consume = rec1!tamount
        Else
            consume = 0
        End If
        excelWS.Cells(RC, 1).Value = "'" & Format(firstdate, "dd/mm/yyyy")
        excelWS.Cells(RC, 2).Value = tempbfqty
        excelWS.Cells(RC, 3).Value = tempbfqty * tempbreakfastrate
        excelWS.Cells(RC, 4).Value = templqty
        excelWS.Cells(RC, 5).Value = templqty * templunch
        excelWS.Cells(RC, 6).Value = tempdqty
        excelWS.Cells(RC, 7).Value = tempdqty * tempdinner
        excelWS.Cells(RC, 8).Value = consume
        excelWS.Cells(RC, 9).Value = ((tempdqty * tempdinner) + (templqty * templunch) + (tempbfqty * tempbreakfastrate)) - consume
        firstdate = firstdate + 1
        RC = RC + 1
    Wend


    excelApp.Visible = True
    Me.MousePointer = 0
    'excelApp.Workbooks.Close

    Set ExcelSheet = Nothing
    Set excelWB = Nothing
    Set excelWS = Nothing

End Sub

Private Sub Form_Load()
Me.txtfrom.Text = Format(Date, "dd/mm/yyyy")
Me.txtto.Text = Format(Date, "dd/mm/yyyy")
End Sub
