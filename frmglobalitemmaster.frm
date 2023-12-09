VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmglobalitemmaster 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Global Item Master"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   14670
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6780
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tempitemmaster"
      Top             =   3900
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   7680
      TabIndex        =   4
      Top             =   7560
      Width           =   1635
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete All"
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
      Left            =   5940
      TabIndex        =   3
      Top             =   7560
      Width           =   1635
   End
   Begin VB.CommandButton cmdcategorydownload 
      Caption         =   "Download Category Items"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   900
      Width           =   2775
   End
   Begin VB.CommandButton cmddownloadbrand 
      Caption         =   "Download Brand Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6060
      TabIndex        =   1
      Top             =   900
      Width           =   2715
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Height          =   6135
      Left            =   180
      TabIndex        =   0
      Top             =   1320
      Width           =   14355
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmglobalitemmaster.frx":0000
         Height          =   5775
         Left            =   60
         OleObjectBlob   =   "frmglobalitemmaster.frx":0014
         TabIndex        =   13
         Top             =   240
         Width           =   14235
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   14475
      Begin VB.ComboBox cboproducttype 
         BackColor       =   &H8000000E&
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
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2955
      End
      Begin VB.ComboBox cboitemtype 
         BackColor       =   &H8000000E&
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
         Left            =   10680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   3255
      End
      Begin VB.ComboBox cbobrandname 
         BackColor       =   &H8000000E&
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
         Left            =   6000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmdimport 
         Caption         =   "Download SubCategory Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8760
         TabIndex        =   6
         Top             =   780
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   2175
      End
      Begin VB.Label itemtype 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   9420
         TabIndex        =   11
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   300
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmglobalitemmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WinHttpReq As WinHttp.WinHttpRequest
Private Const LOCALE_SSHORTDATE = &H1F
Private Const WM_SETTINGCHANGE = &H1A
'same as the old WM_WININICHANGE
Private Const HWND_BROADCAST = &HFFFF&

Private Declare Function SetLocaleInfo Lib "kernel32" Alias _
                                       "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As _
                                                                               Long, ByVal lpLCData As String) As Boolean
Private Declare Function PostMessage Lib "user32" Alias _
                                     "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
                                                     ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" _
                                              () As Long
Dim JSONRec As Object, rec2 As DAO.Recordset

Private Sub cbobrandname_Change()
On Error GoTo errtrap
    Dim strRes As String, reccount, countRecords
    Me.MousePointer = vbHourglass
    Me.cboitemtype.Clear
    WinHttpReq.Open "GET", _
                    "http://techspark.xp3.biz/enlite/getbrand_item.php?ptype=" & Me.cboproducttype.Text & "&brand=" & Me.cbobrandname.Text, False
    WinHttpReq.Send
    If WinHttpReq.ResponseText Like "*Not Found*" Then
        MsgBox "Not Found", vbCritical
    Else
        strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
        countRecords = UBound(Split(strRes, "ItemType"))
        Set JSONRec = JSON.parse(strRes)
        If countRecords > 1 Then
            reccount = 1
            While reccount <= countRecords
                Me.cboitemtype.AddItem JSONRec(reccount).Item("ItemType")
                'JSONRec(reccount).Item("Pid")
                reccount = reccount + 1
            Wend
        Else
            Me.cboitemtype.AddItem JSONRec(1).Item("ItemType")
        End If
        Me.cboitemtype.ListIndex = 0
    End If
    Me.MousePointer = vbDefault
    Exit Sub
errtrap:
MsgBox Err.Description, vbOKOnly
End Sub

Private Sub cbobrandname_Click()
cbobrandname_Change
End Sub

Private Sub cboproducttype_Change()
On Error GoTo errtrap
    Dim strRes As String, reccount, countRecords
    Me.cbobrandname.Clear
    Me.cboitemtype.Clear
    Me.MousePointer = vbHourglass
    WinHttpReq.Open "GET", _
                    "http://techspark.xp3.biz/enlite/getp_brand.php?ptype=" & Me.cboproducttype.Text, False
    WinHttpReq.Send
    If WinHttpReq.ResponseText Like "*Not Found*" Then
        MsgBox "Not Found", vbCritical
    Else
        strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
        countRecords = UBound(Split(strRes, "Brand"))
        Set JSONRec = JSON.parse(strRes)
        If countRecords > 1 Then
            reccount = 1
            While reccount <= countRecords
                Me.cbobrandname.AddItem JSONRec(reccount).Item("Brand")
                'JSONRec(reccount).Item("Pid")
                reccount = reccount + 1
            Wend
        Else
            Me.cbobrandname.AddItem JSONRec(1).Item("Brand")
        End If
        MsgBox "Brand Names Loaded", vbOKOnly
    End If
    Me.MousePointer = vbDefault
Exit Sub
errtrap:
MsgBox Err.Description, vbOKOnly
Resume Next
End Sub

Private Sub cboproducttype_Click()
cboproducttype_Change
End Sub

Private Sub cmdcategorydownload_Click()
Me.MousePointer = vbHourglass
    Dim strRes As String, reccount, countRecords


    WinHttpReq.Open "GET", _
                    "http://techspark.xp3.biz/enlite/getitemmaster_c.php?ptype=" & Me.cboproducttype.Text, False
    WinHttpReq.Send
    If WinHttpReq.ResponseText Like "*Not Found*" Then
        MsgBox "Not Found", vbCritical
    Else
        strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
        countRecords = UBound(Split(strRes, "ProductType"))
        Set JSONRec = JSON.parse(strRes)
        If countRecords > 1 Then
            reccount = 1
            While reccount <= countRecords
                'Me.cboproducttype.AddItem JSONRec(reccount).Item("Productname")
                tempptype = JSONRec(reccount).Item("ProductType")
                tempitype = JSONRec(reccount).Item("ItemType")
                tempbrand = JSONRec(reccount).Item("Brand")
                tempitem = JSONRec(reccount).Item("Item")
                tempbarcode = JSONRec(reccount).Item("Barcode")
                TempSize = JSONRec(reccount).Item("Size")
                temputype = JSONRec(reccount).Item("UnitType")
                tempmrp = JSONRec(reccount).Item("MRP")
                temptax = JSONRec(reccount).Item("Tax")
                temphsn = JSONRec(reccount).Item("HSN")
                temptaxtype = JSONRec(reccount).Item("tax_type")
                temptaxslab = JSONRec(reccount).Item("Taxslab")
                db.Execute ("insert into tempItemMaster (ProductType,ItemType,Brand,Item,Barcode,Size,UnitType,Lose,MRP,Tax,hsn,tax_type,taxslab) values('" & tempptype & "','" & tempitype & "','" & tempbrand & "','" & tempitem & "','" & tempbarcode & "','" & TempSize & "','" & temputype & "',1," & tempmrp & "," & temptax & ",'" & temphsn & "','" & temptaxtype & "','" & temptaxslab & "')")
                'JSONRec(reccount).Item("Pid")
                reccount = reccount + 1
                Me.Data1.Refresh
                Me.DBGrid1.Refresh
            Wend
        Else
            'Me.cboproducttype.AddItem JSONRec(1).Item("Productname")
        End If
        Me.Data1.Refresh
        Me.DBGrid1.Refresh
    End If
Me.MousePointer = vbDefault

End Sub

Private Sub CmdDelete_Click()
db.Execute "delete * from tempitemmaster"
Me.Data1.Refresh
Me.DBGrid1.Refresh
End Sub

Private Sub cmddownloadbrand_Click()
On Error GoTo errtrap
Me.MousePointer = vbHourglass
    Dim strRes As String, reccount, countRecords


    WinHttpReq.Open "GET", _
                    "http://techspark.xp3.biz/enlite/getitemmaster_b.php?ptype=" & Me.cboproducttype.Text & "&brand=" & Me.cbobrandname.Text, False
    WinHttpReq.Send
    If WinHttpReq.ResponseText Like "*Not Found*" Then
        MsgBox "Not Found", vbCritical
    Else
        strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
        countRecords = UBound(Split(strRes, "ProductType"))
        Set JSONRec = JSON.parse(strRes)
        If countRecords > 1 Then
            reccount = 1
            While reccount <= countRecords
                'Me.cboproducttype.AddItem JSONRec(reccount).Item("Productname")
                tempptype = JSONRec(reccount).Item("ProductType")
                tempitype = JSONRec(reccount).Item("ItemType")
                tempbrand = JSONRec(reccount).Item("Brand")
                tempitem = JSONRec(reccount).Item("Item")
                tempbarcode = JSONRec(reccount).Item("Barcode")
                TempSize = JSONRec(reccount).Item("Size")
                temputype = JSONRec(reccount).Item("UnitType")
                tempmrp = JSONRec(reccount).Item("MRP")
                temptax = JSONRec(reccount).Item("Tax")
                temphsn = JSONRec(reccount).Item("HSN")
                temptaxtype = JSONRec(reccount).Item("tax_type")
                temptaxslab = JSONRec(reccount).Item("Taxslab")
                db.Execute ("insert into tempItemMaster (ProductType,ItemType,Brand,Item,Barcode,Size,UnitType,Lose,MRP,Tax,hsn,tax_type,taxslab) values('" & tempptype & "','" & tempitype & "','" & tempbrand & "','" & tempitem & "','" & tempbarcode & "','" & TempSize & "','" & temputype & "',1," & tempmrp & "," & temptax & ",'" & temphsn & "','" & temptaxtype & "','" & temptaxslab & "')")
                'JSONRec(reccount).Item("Pid")
                reccount = reccount + 1
                Me.Data1.Refresh
                Me.DBGrid1.Refresh
            Wend
        Else
           Me.cboproducttype.AddItem JSONRec(1).Item("Productname")
        End If
        Me.Data1.Refresh
        Me.DBGrid1.Refresh
    End If
Me.MousePointer = vbDefault
Exit Sub
errtrap:
MsgBox Err.Description, vbOKOnly
Resume Next
End Sub

Private Sub cmdimport_Click()
Me.MousePointer = vbHourglass
    Dim strRes As String, reccount, countRecords


    WinHttpReq.Open "GET", _
                    "http://techspark.xp3.biz/enlite/getitemmaster.php?ptype=" & Me.cboproducttype.Text & "&brand=" & Me.cbobrandname.Text & "&itype=" & Me.cboitemtype.Text, False
    WinHttpReq.Send
    If WinHttpReq.ResponseText Like "*Not Found*" Then
        MsgBox "Not Found", vbCritical
    Else
        strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
        countRecords = UBound(Split(strRes, "ProductType"))
        Set JSONRec = JSON.parse(strRes)
        If countRecords > 1 Then
            reccount = 1
            While reccount <= countRecords
                'Me.cboproducttype.AddItem JSONRec(reccount).Item("Productname")
                tempptype = JSONRec(reccount).Item("ProductType")
                tempitype = JSONRec(reccount).Item("ItemType")
                tempbrand = JSONRec(reccount).Item("Brand")
                tempitem = JSONRec(reccount).Item("Item")
                tempbarcode = JSONRec(reccount).Item("Barcode")
                TempSize = JSONRec(reccount).Item("Size")
                temputype = JSONRec(reccount).Item("UnitType")
                tempmrp = JSONRec(reccount).Item("MRP")
                temptax = JSONRec(reccount).Item("Tax")
                temphsn = JSONRec(reccount).Item("HSN")
                temptaxtype = JSONRec(reccount).Item("tax_type")
                temptaxslab = JSONRec(reccount).Item("Taxslab")
                db.Execute ("insert into tempItemMaster (ProductType,ItemType,Brand,Item,Barcode,Size,UnitType,Lose,MRP,Tax,hsn,tax_type,taxslab) values('" & tempptype & "','" & tempitype & "','" & tempbrand & "','" & tempitem & "','" & tempbarcode & "','" & TempSize & "','" & temputype & "',1," & tempmrp & "," & temptax & ",'" & temphsn & "','" & temptaxtype & "','" & temptaxslab & "')")
                'JSONRec(reccount).Item("Pid")
                reccount = reccount + 1
                Me.Data1.Refresh
                Me.DBGrid1.Refresh
            Wend
        Else
            Me.cboproducttype.AddItem JSONRec(1).Item("Productname")
        End If
        Me.Data1.Refresh
        Me.DBGrid1.Refresh
    End If
Me.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()

End Sub

Private Sub CmdSave_Click()
On Error GoTo errtrap
    ans = MsgBox("Do You Want to Save the Downloaded Product Names?", vbYesNo)
    If ans = 6 Then
        Set rec1 = db.OpenRecordset("select * from tempitemmaster")
        While Not rec1.EOF
            Set rec2 = db.OpenRecordset("select * from product where productname='" & rec1("producttype") & "'")
            If rec2.EOF Then
                db.Execute "insert into product values('" & rec1("producttype") & "',0)"
            End If
            Set rec2 = db.OpenRecordset("select * from brandmaster where brand='" & rec1("producttype") & "'")
            If rec2.EOF Then
                db.Execute "insert into brandmaster  (brand,brandid) values('" & rec1("brand") & "',0)"
            End If
            Set rec2 = db.OpenRecordset("select * from itemtype where item_type='" & rec1("itemtype") & "'")
            If rec2.EOF Then
                db.Execute "insert into itemtype (item_type,producttype) values('" & rec1("itemtype") & "','" & rec1("producttype") & "')"
            End If
            Set rec2 = db.OpenRecordset("SELECT MAX(PRODUCTCODE) AS MAXCODE FROM ITEMMASTER")
            If Not IsNull(rec2!MAXCODE) Then
                NEXTCODE = rec2!MAXCODE + 1
            Else
                NEXTCODE = 1001
            End If
            db.Execute ("insert into ItemMaster (ProductType,ItemType,Brand,Item,Barcode,Size,UnitType,Lose,MRP,Tax,hsn,PRODUCTCODE,tax_type,taxslab) values('" & rec1("ProductType") & "','" & rec1("ItemType") & "','" & rec1("Brand") & "','" & rec1("Item") & "','" & rec1("Barcode") & "','" & rec1("size") & "','" & rec1("unittype") & "'," & rec1("lose") & "," & rec1("mrp") & "," & rec1("tax") & ",'" & rec1("hsn") & "'," & NEXTCODE & ",'" & rec1("tax_type") & "','" & rec1("taxslab") & "')")
            db.Execute ("insert into STOCK (ProductType,ItemType,Brand,ItemNAME,Barcode,Size,UniYType,Lose,MRP,VAT,hsn,PRODUCTCODE,taxslab) values('" & rec1("ProductType") & "','" & rec1("ItemType") & "','" & rec1("Brand") & "','" & rec1("Item") & "','" & rec1("Barcode") & "',1,'" & rec1("UnitType") & "'," & rec1("Lose") & "," & rec1("MRP") & "," & rec1("tax") & ",'" & rec1("hsn") & "'," & NEXTCODE & ",'" & rec1("taxslab") & "')")
            'db.Execute ("insert into ItemMaster (ProductType,ItemType,Brand,Item,Barcode,Size,UnitType,Lose,MRP,Tax,hsn) values('" & tempptype & "','" & tempitype & "','" & tempbrand & "','" & tempitem & "','" & tempbarcode & "','" & TempSize & "','" & temputype & "',1," & tempmrp & "," & temptax & ",'" & temphsn & "')")
            rec1.MoveNext
        Wend
        db.Execute "delete * from tempitemmaster"
        Data1.Refresh
        MsgBox "Product Names has been Saved", vbOKOnly
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
    Resume Next
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
On Error GoTo errtrap
    Dim strRes As String, reccount, countRecords
    Set WinHttpReq = New WinHttpRequest
    Me.Data1.databasename = db.Name
    WinHttpReq.Open "GET", _
                    "http://techspark.xp3.biz/enlite/getproducttype.php", False
    WinHttpReq.Send
    If WinHttpReq.ResponseText Like "*Not Found*" Then
        MsgBox "Not Found", vbCritical
    Else
        strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
        countRecords = UBound(Split(strRes, "Productname"))
        Set JSONRec = JSON.parse(strRes)
        If countRecords > 1 Then
            reccount = 1
            While reccount <= countRecords
                Me.cboproducttype.AddItem JSONRec(reccount).Item("Productname")
                'JSONRec(reccount).Item("Pid")
                reccount = reccount + 1
            Wend
        Else
            Me.cboproducttype.AddItem JSONRec.Item("Productname")
        End If
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub

