VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FRMSALESRETURNPRINT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Return List"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FRMSALESRETURNPRINT.frx":0000
   ScaleHeight     =   2745
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2400
      Top             =   2280
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
      Left            =   2880
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FRMSALESRETURNPRINT.frx":D4E3
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "FRMSALESRETURNPRINT.frx":D4F7
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      RecordSource    =   "Salesreturnhead"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "FRMSALESRETURNPRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
    frmprinter.Show vbModal
    Me.CrystalReport1.PrinterName = Printer.DeviceName
    Me.CrystalReport1.PrinterDriver = Printer.DriverName
    Me.CrystalReport1.PrinterPort = Printer.Port
    Me.CrystalReport1.SelectionFormula = "{Salesreturnhead.InvNo} =" & Val(Me.DBGrid1.Columns(0))
    Me.CrystalReport1.WindowState = crptMaximized
    Me.CrystalReport1.PrintReport
End Sub

Private Sub Form_Load()
    Me.Top = 3300
    Me.Left = 2500
    Me.Data1.databasename = dbname
    Me.CrystalReport1.ReportFileName = App.Path & "\salesreturn.rpt"
    Data1.RecordSource = "SELECT * FROM SALESRETURNHEAD ORDER BY InvNo"
    Data1.Refresh
End Sub
