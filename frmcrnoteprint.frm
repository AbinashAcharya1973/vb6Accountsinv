VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmcrnoteprint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Note Print"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6870
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "creditnote"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFC0FF&
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3010
      TabIndex        =   3
      Top             =   3870
      Width           =   850
   End
   Begin VB.TextBox txtsearch 
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
      Left            =   1170
      TabIndex        =   1
      Top             =   150
      Width           =   5655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmcrnoteprint.frx":0000
      Height          =   3045
      Left            =   -150
      OleObjectBlob   =   "frmcrnoteprint.frx":0014
      TabIndex        =   0
      Top             =   660
      Width           =   6975
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Party"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   2
      Top             =   210
      Width           =   1455
   End
End
Attribute VB_Name = "frmcrnoteprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
Me.CrystalReport1.SelectionFormula = "{debitnote.slno}=" & DBGrid1.Columns(0)
Me.CrystalReport1.PrintReport
End Sub

Private Sub Form_Load()
Me.Data1.databasename = db.Name
Me.CrystalReport1.ReportFileName = App.Path & "\CREDITNOTE.RPT"
End Sub



Private Sub txtsearch_Change()
Data1.RecordSource = "select * from CREDITnote where partyname like '" & Me.txtsearch.Text & "*'"
Data1.Refresh
End Sub
