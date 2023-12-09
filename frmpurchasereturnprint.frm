VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmpurchasereturnprint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Return List"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1920
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton cmdprint 
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
      Left            =   2880
      Picture         =   "frmpurchasereturnprint.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   495
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmpurchasereturnprint.frx":014A
      Height          =   2175
      Left            =   120
      OleObjectBlob   =   "frmpurchasereturnprint.frx":015E
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "D:\BISINABAR_FMCG\FMCG\DATA\2011-2012\FMCG.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PurchaseReturnHead"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmpurchasereturnprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
    frmprinter.Show vbModal
    Me.CrystalReport1.ReportFileName = App.Path & "\purchasereturn.rpt"
    Me.CrystalReport1.PrinterName = Printer.DeviceName
    Me.CrystalReport1.PrinterDriver = Printer.DriverName
    Me.CrystalReport1.PrinterPort = Printer.Port
    Me.CrystalReport1.SelectionFormula = "{PurchaseReturnHead.Slno} = " & Val(Me.DBGrid1.Columns(0))
    Me.CrystalReport1.WindowState = crptMaximized
    Me.CrystalReport1.PrintReport
End Sub

Private Sub Form_Load()
Me.Data1.databasename = db.Name
End Sub
