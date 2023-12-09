VERSION 5.00
Begin VB.Form frmbackup_restore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup and Restore"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   4680
   Begin VB.Frame Frame2 
      Caption         =   "Account Year Ending Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   840
      TabIndex        =   6
      Top             =   3840
      Width           =   2775
      Begin VB.TextBox txtfinancialyear 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaxLength       =   9
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "New Financial Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Avilable Accounting Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   2775
      Begin VB.ComboBox cbodate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Select Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdrestore_curr 
      Caption         =   "Restore Current Database"
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
      Left            =   840
      TabIndex        =   2
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton cmdyearending 
      Caption         =   "Activate Year Ending"
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
      Left            =   840
      TabIndex        =   0
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label lblcaption 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please Close All Programe Before Year End"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   4695
   End
End
Attribute VB_Name = "frmbackup_restore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As Recordset
Private Sub cmdact_main_Click()
'temp_date = Format(Date, "dd") & "-" & Format(Date, "mm") & "-" & Format(Date, "yy")
'fname = "InvAccJal" & temp_date & ".mdb"
    des = Dir("d:\InvAccJal\current\InvAccJal.mdb")
    If des = "" Then
        FileCopy "d:\InvAccJal\InvAccJal.mdb", "d:\InvAccJal\current\InvAccJal.mdb"
        Kill "d:\InvAccJal\InvAccJal.mdb"
        FileCopy "d:\InvAccJal\TranVault\InvAccJal.mdb", "d:\InvAccJal\InvAccJal.mdb"
    Else
        Kill "d:\InvAccJal\current\InvAccJal.mdb"
        FileCopy "d:\InvAccJal\InvAccJal.mdb", "d:\InvAccJal\current\InvAccJal.mdb"
        Kill "d:\InvAccJal\InvAccJal.mdb"
        FileCopy "d:\InvAccJal\TranVault\InvAccJal.mdb", "d:\InvAccJal\InvAccJal.mdb"
    End If
    current_db_back = "y"
End Sub

Private Sub cmdrestore_curr_Click()
'If current_db_back = "y" Then
'    Kill "d:\InvAccJal\InvAccJal.mdb"
'    FileCopy "d:\InvAccJal\current\InvAccJal.mdb", "d:\InvAccJal\InvAccJal.mdb"
'    current_db_back = "n"
'End If
End Sub
Private Sub cmdyearending_Click()
    Dim FileInQuestion As String, fname As String, sFilename As String, driveinquestion
    ans = MsgBox("Are You Sure ?", vbYesNo)
    If ans = 6 Then
        db.Close
        Set db = Nothing
        If Me.txtfinancialyear.Text = "" Or Len(Trim(Me.txtfinancialyear.Text)) < 9 Then
            MsgBox "Invalid FinancialYear or Blank", vbCritical
        Else
            FileInQuestion = Dir(App.Path & "\Data\" & dbname & ".mdb")
            If FileInQuestion = "" Then
                FileCopy dbname, App.Path & "\data\" & dbname & ".mdb"
            Else
                Kill App.Path & "\data\" & dbname & ".mdb"
                FileCopy dbname, App.Path & "\data\" & dbname & ".mdb"
            End If
            dbname = ""
            Set Db1 = OpenDatabase(dbname)
            Db1.Execute ("delete * from ledgertran")
            Db1.Execute ("update ledgermaster set OBalance=0")
            Db1.Execute ("delete * from invoicehead")
            Db1.Execute ("delete * from invoicedetails")
            Db1.Execute ("delete * from purchasehead")
            Db1.Execute ("delete * from purchasedetails")
            Db1.Execute ("delete * from PaymentHead")
            Db1.Execute ("delete * from PaymentDetails")
            Db1.Execute ("delete * from ReceiptHead")
            Db1.Execute ("delete * from ReceiptDetails")
            Db1.Execute ("delete * from JournalHead")
            Db1.Execute ("delete * from JournalDetails")
            Db1.Execute ("delete * from PurchaseReturnHead")
            Db1.Execute ("delete * from PurchaseReturnDetails")
            Db1.Execute ("delete * from ContraHead")
            Db1.Execute ("delete * from ContraDetails")
            Db1.Execute ("delete * from CreditNoteHead")
            Db1.Execute ("delete * from CreditnoteDetails")
            Db1.Close
            Set Db1 = Nothing
            dbname = Trim(Me.txtfinancialyear.Text)
            FileCopy dbname, App.Path & "\data\" & dbname & ".mdb"
            Kill dbname

            Kill App.Path & "\test.txt"
            sFilename = App.Path & "\test.txt"
            nFileNum = FreeFile
            Open sFilename For Binary Lock Read Write As nFileNum
            Put #nFileNum, , "0"
            Close #nFileNum
            MsgBox "The Programe will shutdown Restart Again", vbOKOnly
            End
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Dim pathLog As New pathLogger
    Dim total As Long
    Dim fileNames As New Collection
    Screen.MousePointer = vbHourglass

    With pathLog
        .PathLogInit
        .fullFileNamesOnly = True
        .LogPath App.Path & "\DATA"
        total& = .Count
        Set fileNames = .fullFileNames
        .Terminate
    End With

    count_str = Len(App.Path & "\DATA\")

    Me.cbodate.Clear
    For Each FileName In fileNames
        dbfilename = Mid(FileName, count_str + 1, Len(FileName))
        withoutextention = Mid(dbfilename, 1, InStr(1, dbfilename, ".mdb", vbTextCompare) - 1)
        Me.cbodate.AddItem (withoutextention)
    Next
    If Me.cbodate.ListCount > 0 Then
        Me.cbodate.ListIndex = 0
    End If
    Screen.MousePointer = vbDefault
    If current_db_back <> "y" Then
        Me.cmdrestore_curr.Enabled = False
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Set db1 = OpenDatabase(App.Path & "\InvAccJal.mdb")
End Sub
