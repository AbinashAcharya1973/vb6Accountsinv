VERSION 5.00
Begin VB.Form FrmTools 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tools"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command5 
      Caption         =   "For Journal Voucher"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "For Payment Voucher"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "For Receipt Voucher"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "For Purchase"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "For Invoice"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec1 As DAO.Recordset, rec2 As DAO.Recordset
Private Sub Command1_Click()
Set rec1 = Db.OpenRecordset("select AccId from LedgerMaster where AccName like 'SALES ACCOUNT'")
If Not rec1.EOF Then
SalesAccId = rec1("AccId")
End If
Set rec1 = Db.OpenRecordset("select * from ledgertran where AccId=" & SalesAccId & " and VoucherType='RETAIL Invoice'")
If Not rec1.EOF Then
    While Not rec1.EOF
        Set rec2 = Db.OpenRecordset("select * from Invoicehead where InvNo=" & rec1("VoucherSlno"))
        If Not rec1.EOF Then
        Db.Execute ("update LedgerTran set TranAccId=" & rec2("AccId") & " where AccId=" & SalesAccId & " and VoucherType='RETAIL Invoice' and VoucherSlno=" & rec1("VoucherSlno"))
        End If
    rec1.MoveNext
    Wend
End If
Db.Execute ("update LedgerTran set TranAccId=" & SalesAccId & " where VoucherType='RETAIL Invoice' and AccId<>" & SalesAccId)
MsgBox "Complete Updation", vbOKOnly
End Sub
Private Sub Command2_Click()
Set rec1 = Db.OpenRecordset("select AccId from LedgerMaster where AccName like 'PURCHASE ACCOUNT'")
If Not rec1.EOF Then
PurchaseAccId = rec1("AccId")
End If
Set rec1 = Db.OpenRecordset("select * from LedgerTran where AccId=" & PurchaseAccId & " and VoucherType='Purchase'")
If Not rec1.EOF Then
    While Not rec1.EOF
        Set rec2 = Db.OpenRecordset("select * from PurchaseHead where Slno=" & rec1("VoucherSlno"))
        If Not rec2.EOF Then
        Db.Execute ("update LedgerTran set TranAccId=" & rec2("AccId") & " where AccId=" & PurchaseAccId & " and VoucherType='Purchase' and VoucherSlno=" & rec1("VoucherSlno"))
        End If
        rec1.MoveNext
    Wend
End If
Db.Execute ("update LedgerTran set TranAccId=" & PurchaseAccId & " where VoucherType='Purchase' and AccId<>" & PurchaseAccId)
MsgBox "Complete", vbOKOnly
End Sub
Private Sub Command3_Click()
Set rec1 = Db.OpenRecordset("select * from  ledgerTran where VoucherType='Receipt' and Cr>0")
If Not rec1.EOF Then
    While Not rec1.EOF
        temp_from_date = Mid(rec1("TDate"), 1, 2)
        temp_from_month = Mid(rec1("TDate"), 4, 2)
        temp_from_year = Mid(rec1("TDate"), 7, 4)
        temp_from = temp_from_month & "/" & temp_from_date & "/" & temp_from_year
        
        Set rec2 = Db.OpenRecordset("select * from ReceiptHead where ReceiptDate=#" & temp_from & "# and ReceiptNo=" & rec1("VoucherSlNo"))
        If Not rec2.EOF Then
        Db.Execute ("update LedgerTran set TranAccId=" & rec2("AccId") & " where TDate=#" & temp_from & "# and VoucherSlNo=" & rec2("ReceiptNo") & " and VoucherType='Receipt' and Cr>0")
        End If
    rec1.MoveNext
    Wend
End If
'---For MainAccount----------
Set rec1 = Db.OpenRecordset("select * from  ledgerTran where VoucherType='Receipt' and Dr>0")
If Not rec1.EOF Then
    While Not rec1.EOF
        temp_from_date = Mid(rec1("TDate"), 1, 2)
        temp_from_month = Mid(rec1("TDate"), 4, 2)
        temp_from_year = Mid(rec1("TDate"), 7, 4)
        temp_from = temp_from_month & "/" & temp_from_date & "/" & temp_from_year
        Set rec2 = Db.OpenRecordset("select *  from ReceiptDetails where ReceiptDate=#" & temp_from & "# and ReceiptNo=" & rec1("VoucherSlNo"))
        If Not rec2.EOF Then
        Db.Execute ("update LedgerTran set TranAccId=" & rec2("AccId") & ",Particulars='To " & rec2("AccName") & "' where TDate=#" & temp_from & "# and VoucherType='Receipt' and VoucherSlno=" & rec2("ReceiptNo") & " and Dr>0")
        End If
    rec1.MoveNext
    Wend
End If
MsgBox "Complete Update", vbOKOnly
End Sub
Private Sub Command4_Click()
Set rec1 = Db.OpenRecordset("select * from ledgerTran where VoucherType='Payment' and Cr>0")
If Not rec1.EOF Then
    While Not rec1.EOF
        temp_from_date = Mid(rec1("TDate"), 1, 2)
        temp_from_month = Mid(rec1("TDate"), 4, 2)
        temp_from_year = Mid(rec1("TDate"), 7, 4)
        temp_from = temp_from_month & "/" & temp_from_date & "/" & temp_from_year
        Set rec2 = Db.OpenRecordset("select * from PaymentDetails where PDate=#" & temp_from & "# and SlNo=" & rec1("VoucherSlNo"))
        If Not rec2.EOF Then
        Db.Execute ("update LedgerTran set TranAccId=" & rec2("AccId") & ",Particulars='By " & rec2("AccName") & "' where TDate=#" & temp_from & "# and VoucherType='Payment' and VoucherSlno=" & rec2("SlNo") & " and Cr>0")
        End If
    rec1.MoveNext
    Wend

End If
'----------PartyAccount Update------------
Set rec1 = Db.OpenRecordset("select * from LedgerTran where VoucherType='Payment' and Dr>0")
If Not rec1.EOF Then
    While Not rec1.EOF
        temp_from_date = Mid(rec1("TDate"), 1, 2)
        temp_from_month = Mid(rec1("TDate"), 4, 2)
        temp_from_year = Mid(rec1("TDate"), 7, 4)
        temp_from = temp_from_month & "/" & temp_from_date & "/" & temp_from_year
        Set rec2 = Db.OpenRecordset("select * from PaymentHead where PDate=#" & temp_from & "# and Slno=" & rec1("VoucherSlno"))
        If Not rec2.EOF Then
        Db.Execute ("update LedgerTran set TranAccId=" & rec2("AccId") & " where Tdate=#" & temp_from & "# and VoucherType='Payment' and VoucherSlno=" & rec2("Slno") & " and Dr>0")
        End If
      rec1.MoveNext
    Wend
  
End If
MsgBox "Complet updation", vbOKOnly
End Sub
