VERSION 5.00
Begin VB.Form frmstartingscreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Accounting Period"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmstartingscreen.frx":0000
   ScaleHeight     =   2295
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbodate 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2760
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox txtusername 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtpassword 
      BackColor       =   &H00FFFFFF&
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
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select A/c Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmstartingscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WinHttpReq As WinHttp.WinHttpRequest, rec1 As DAO.Recordset
Attribute rec1.VB_VarUserMemId = 1073938432
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

Dim nohit
Attribute nohit.VB_VarUserMemId = 1073938434
Private Sub cbodate_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Integer
    Dim nFileNum As Integer
    Dim sFilename As String
    Dim databasename As String
    Dim sMyString As String
    Dim de As String
    Dim mystr As String
    If KeyCode = 13 Then

        Me.Label2.Visible = True
        Me.Label3.Visible = True
        Me.txtusername.Visible = True
        Me.txtpassword.Visible = True
        Me.txtusername.SetFocus
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo errtrap
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set fo = fs.GetFolder(App.Path & "\DATA\")
    For Each X In fo.SubFolders
        Me.cbodate.AddItem (X.Name)
    Next
    If Me.cbodate.ListCount > 0 Then
        Me.cbodate.ListIndex = 0
    End If
    Screen.MousePointer = vbDefault
    MsgBox "Please Connect the Internet to Login", vbOKOnly
    Exit Sub
errtrap:
    MsgBox Err.Description, vbOKOnly
End Sub
Private Sub txtpassword_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errtrap
    If KeyCode = 13 Then
        Dim strRes As String, reccount, countRecords, JSONRec As Object
        Dim fname, des, temp_date
        temp_date = Format(Date, "dd") & "-" & Format(Date, "mm") & "-" & Format(Date, "yy")
        fname = Me.cbodate.Text & "_" & temp_date & ".mdb"
        des = Dir(App.Path & "\backup\" & fname)
        If des = "" Then
            FileCopy App.Path & "\DATA\" & Me.cbodate.Text & "\FMCG.mdb", App.Path & "\backup\" & fname
        Else
            Kill App.Path & "\backup\" & fname
            FileCopy App.Path & "\DATA\" & Me.cbodate.Text & "\FMCG.mdb", App.Path & "\backup\" & fname
        End If
        Set db = OpenDatabase(App.Path & "\DATA\" & Me.cbodate.Text & "\FMCG.mdb")
        dbname = App.Path & "\DATA\" & Me.cbodate.Text & "\FMCG.mdb"
        CreateAccessODBC App.Path & "\DATA\" & Me.cbodate.Text & "\FMCG.mdb", "FMCG", "FMCG"

        Set WinHttpReq = New WinHttpRequest

        Set rec1 = db.OpenRecordset("Select * from companymaster")
        If Not rec1.EOF Then
            'Me.lblmessage.Caption = "Please Wait, Verifing Registration!"
            Customer = Trim(rec1("company"))
            Adress = Trim(rec1("address"))
            City = Trim(rec1("pin"))
            Email = Trim(rec1("email"))
            Mobile = Trim(rec1("phone"))
            gstino = Trim(rec1("taxno"))
            Address = Trim(rec1("address1"))
            dealin = Trim(rec1("dealin"))

            'WinHttpReq.Open "GET", _
                            "https://techspark.xp3.biz/enlite/checklogin_new.php?Customer=" & Customer & "&PIN=" & City & "&Email=" & Email & "&Mobile=" & Mobile & "&gstno=" & gstino & "&pwd=" & Me.txtpassword.Text, False

            'WinHttpReq.Send

            'strRes = Mid(WinHttpReq.ResponseText, 2, Len(WinHttpReq.ResponseText) - 1)
            'countRecords = UBound(Split(strRes, "r_url"))
            'countRecords = 1
'            If countRecords > 0 Then
'                Set JSONRec = JSON.parse(strRes)
'                tempmessage = JSONRec(1).Item("message")
'                tempurl = JSONRec(1).Item("r_url")
'                'tempmessage = "ok"
'                If tempmessage Like "*ok*" Then
'                    usname = Trim(Me.txtusername.Text)
'                    SoftwareVersion = "Registered"
'                    ACYEAR = Me.cbodate.Text
'                    usertype = "Admin"
'                    frmMain.Show
'                    Unload Me
'                Else
'                    ans1 = MsgBox(tempmessage, vbOKCancel)
'                    If ans1 = 1 And Len(tempurl) > 0 Then
'                        db.Close
'                        Call Shell("rundll32.exe url.dll,FileProtocolHandler " & tempurl & "?id=" & Mobile, vbMaximizedFocus)
'                    Else
'                        ans = MsgBox("Run As Demo?", vbYesNo)
'                        If ans = 6 Then
'                            usname = Trim(Me.txtusername.Text)
'                            SoftwareVersion = "Demo"
'                            usertype = "Admin"
'                            ACYEAR = Me.cbodate.Text
'                            frmMain.Show
'                            Unload Me
'                        Else
'                            db.Close
'                        End If
'                    End If
'
'                End If
'            Else
'                MsgBox "Invalid User Name or Password", vbCritical
'                ans = MsgBox("Run As Demo?", vbYesNo)
'                If ans = 6 Then
'                    usname = Trim(Me.txtusername.Text)
'                    SoftwareVersion = "Demo"
'                    usertype = "Admin"
'                    ACYEAR = Me.cbodate.Text
'                    frmMain.Show
'                    Unload Me
'                Else
'                    db.Close
'                End If
'            End If
usname = Trim(Me.txtusername.Text)
                    SoftwareVersion = "Registered"
                    ACYEAR = Me.cbodate.Text
                    usertype = "Admin"
                    frmMain.Show
                    Unload Me
        End If

        '        Set rec1 = db.OpenRecordset("select * from usertable where username='" & Me.txtusername.Text & "' and Password='" & Me.txtpassword.Text & "'")
        '        If Not rec1.EOF Then
        '            usertype = Trim(rec1("usertype"))
        '            usname = Trim(Me.txtusername.Text)
        '            acyear = Me.cbodate.Text
        '            frmMain.Show
        '            Unload Me
        '        Else
        '            MsgBox "invalid user name or password!", vbCritical
        '        End If

    End If
    If KeyCode = 27 Then
        End
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbCritical
End Sub
Private Sub txtusername_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtpassword.SetFocus
    End If
End Sub
