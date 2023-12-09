VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmdamageprint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Damage List"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   600
      Top             =   2220
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton CMDPRINT 
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmdamageprint.frx":0000
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "frmdamageprint.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DamageHead"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmdamageprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
    frmprinter.Show vbModal
    Me.CrystalReport1.ReportFileName = App.Path & "\damage.rpt"
    Me.CrystalReport1.PrinterName = Printer.DeviceName
    Me.CrystalReport1.PrinterDriver = Printer.DriverName
    Me.CrystalReport1.PrinterPort = Printer.Port
    Me.CrystalReport1.SelectionFormula = "{DamageHead.slno} = " & Val(Me.DBGrid1.Columns(0))
    Me.CrystalReport1.WindowState = crptMaximized
    Me.CrystalReport1.PrintReport
End Sub

Private Sub Form_Load()
    Me.Data1.databasename = db.Name
End Sub
