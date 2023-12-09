VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmFlagMaster 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flag Master"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   3720
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmFlagMaster.frx":0000
         Height          =   6135
         Left            =   120
         OleObjectBlob   =   "frmFlagMaster.frx":0014
         TabIndex        =   0
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Shankar_Textiles\Shankar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FlagMaster"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmFlagMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
Me.DBGrid1.AllowAddNew = True
Me.DBGrid1.AllowUpdate = True
End If
End Sub
