VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmprintstocktran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Stock Transfer"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6930
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   780
      Top             =   2670
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2550
      TabIndex        =   1
      Top             =   2820
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmprintstocktran.frx":0000
      Height          =   2385
      Left            =   60
      OleObjectBlob   =   "frmprintstocktran.frx":0014
      TabIndex        =   0
      Top             =   360
      Width           =   6795
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   5190
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "StocktranHead"
      Top             =   2790
      Visible         =   0   'False
      Width           =   1545
   End
End
Attribute VB_Name = "frmprintstocktran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
Me.CrystalReport1.ReportFileName = App.Path & "\sttransfer.rpt"
Me.CrystalReport1.SelectionFormula = "{StocktranHead.ChallanNO}=" & Me.DBGrid1.Columns(0)
Me.CrystalReport1.PrintReport
End Sub

Private Sub Form_Load()
Me.Data1.databasename = db.Name
End Sub
