VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000D&
   Caption         =   "Enlite - SAS"
   ClientHeight    =   3495
   ClientLeft      =   165
   ClientTop       =   1110
   ClientWidth     =   5250
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0464
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0718
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0872
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C80
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   1217
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add Stock"
            ImageIndex      =   1
            Style           =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add Purchase Order"
            ImageIndex      =   2
            Style           =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Invoice"
            ImageIndex      =   3
            Style           =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Party Outstanding"
            ImageIndex      =   4
            Style           =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stock"
            ImageIndex      =   5
            Style           =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add Payment"
            ImageIndex      =   6
            Style           =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add Receipt"
            ImageIndex      =   7
            Style           =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "View Ledger"
            ImageIndex      =   8
            Style           =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Day Book"
            ImageIndex      =   9
            Style           =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Online Support"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnumaster 
      Caption         =   "Masters"
      Begin VB.Menu mnucompanymaster 
         Caption         =   "Company Master"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuregistration 
         Caption         =   "Registration"
      End
      Begin VB.Menu mnusp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnugroup 
         Caption         =   "Group"
      End
      Begin VB.Menu mnuledgermaster 
         Caption         =   "Ledger Master"
      End
      Begin VB.Menu mnusp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnunewunit 
         Caption         =   "Unit Master"
      End
      Begin VB.Menu mnuproducttype 
         Caption         =   "Product Type"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuitemtype 
         Caption         =   "Item Type"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnugitems 
         Caption         =   "Global Item Master"
      End
      Begin VB.Menu mnunewitemmaster 
         Caption         =   "New Item Master"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuitemslab 
         Caption         =   "Item Slab"
         Enabled         =   0   'False
         Begin VB.Menu mnuadd 
            Caption         =   "Add"
         End
      End
      Begin VB.Menu mnusp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnupartycr 
         Caption         =   "Sundry Creditor"
         Visible         =   0   'False
      End
      Begin VB.Menu mnupartydr 
         Caption         =   "Sundry Debtor"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusp7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuzonemaster 
         Caption         =   "Zone Master"
      End
      Begin VB.Menu mnulrmaster 
         Caption         =   "LR Master"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuflagmaster 
         Caption         =   "Flag Master"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuscheme 
         Caption         =   "Scheme"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnumailsetup1 
         Caption         =   "Mail Setup"
      End
      Begin VB.Menu mnubranch 
         Caption         =   "New StockPoint/Branch"
      End
      Begin VB.Menu mnufing 
         Caption         =   "Food Ingrediants"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnutransaction 
      Caption         =   "Transactions"
      Begin VB.Menu mnustockin 
         Caption         =   "Stock In / Purchase"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnustockikn_b 
         Caption         =   "Stock In/Purchase [Barcode In]"
      End
      Begin VB.Menu mnustocktran 
         Caption         =   "Stock Transfer"
      End
      Begin VB.Menu mnuoutwardchallan 
         Caption         =   "Outward Challan"
      End
      Begin VB.Menu mnudeliverychallan 
         Caption         =   "Delivery Challan"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinvoice 
         Caption         =   "Invoice"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuretailinvoice 
         Caption         =   "Retail Invoice"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnulandbill 
         Caption         =   "Land Bill"
      End
      Begin VB.Menu mnusp4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuorder 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu mnusaleorder 
         Caption         =   "Sale Order"
      End
      Begin VB.Menu mnuquotation 
         Caption         =   "Quotation"
      End
      Begin VB.Menu mnusp6 
         Caption         =   "-"
      End
      Begin VB.Menu mnudamageen 
         Caption         =   "Damage Entry"
      End
      Begin VB.Menu mnusalesreturn 
         Caption         =   "SalesReturn"
      End
      Begin VB.Menu mnupurchasereturn 
         Caption         =   "Purchase Return"
      End
      Begin VB.Menu mnudamagereturn 
         Caption         =   "Damage Return"
      End
      Begin VB.Menu mnusp5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuvoucher 
         Caption         =   "Voucher"
         Begin VB.Menu mnureceipt 
            Caption         =   "Receipt"
         End
         Begin VB.Menu mnupayment 
            Caption         =   "Payment"
         End
         Begin VB.Menu mnucontra 
            Caption         =   "Contra"
         End
         Begin VB.Menu mnujournal 
            Caption         =   "Journal"
            Visible         =   0   'False
         End
         Begin VB.Menu mnunewjournal 
            Caption         =   "New Journal"
         End
         Begin VB.Menu mnudrnote 
            Caption         =   "Debit Note"
         End
         Begin VB.Menu mnucrnote 
            Caption         =   "Credit Note"
         End
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit or View"
      Begin VB.Menu mnupurchaseorderview 
         Caption         =   "PurchaseOrder_View"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusalenoteview 
         Caption         =   "Sale Note View"
         Visible         =   0   'False
      End
      Begin VB.Menu mnupl5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnustockinedit 
         Caption         =   "Stock In / Purchase List"
      End
      Begin VB.Menu mnuinvoiceedit 
         Caption         =   "Invoice List"
      End
      Begin VB.Menu mnucreditnoteedit 
         Caption         =   "Sales Return List"
      End
      Begin VB.Menu mnudebitnoteedit 
         Caption         =   "Purchase Return  List"
      End
      Begin VB.Menu mnudamagestockvi 
         Caption         =   "Damage Stock View"
      End
      Begin VB.Menu mnustockview 
         Caption         =   "Stock View"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnupl6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuopstockedit 
         Caption         =   "Item Master / Op Stock Edit"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuitemedit 
         Caption         =   "Item Edit"
         Visible         =   0   'False
      End
      Begin VB.Menu mnumasteredit 
         Caption         =   "-"
      End
      Begin VB.Menu mnusize 
         Caption         =   "Size"
      End
      Begin VB.Menu mnubrandname 
         Caption         =   "Brand Name"
      End
   End
   Begin VB.Menu mnuaccount 
      Caption         =   "Account"
      Begin VB.Menu mnuvoucheredit 
         Caption         =   "Voucher Edit"
         Begin VB.Menu mnujournalvoucher 
            Caption         =   "Journal Voucher"
            Visible         =   0   'False
         End
         Begin VB.Menu mnunewjournalvoucher 
            Caption         =   "New Journal Voucher"
         End
         Begin VB.Menu mnureceiptvoucher 
            Caption         =   "Receipt Voucher"
         End
         Begin VB.Menu mnupaymentvoucher 
            Caption         =   "Payment Voucher"
         End
         Begin VB.Menu mnucontravoucher 
            Caption         =   "Contra Voucher"
         End
      End
      Begin VB.Menu mnupartydredit 
         Caption         =   "Party Debtor Edit"
      End
      Begin VB.Menu mnupartycredit 
         Caption         =   "Party Creditor edit"
      End
      Begin VB.Menu mnupl7 
         Caption         =   "-"
      End
      Begin VB.Menu mnumasterview 
         Caption         =   "Master"
         Begin VB.Menu mnuopbalance 
            Caption         =   "Ledger Op Balance"
         End
         Begin VB.Menu mnugroupview 
            Caption         =   "Group View"
         End
      End
      Begin VB.Menu mnuledgerview 
         Caption         =   "Ledger View"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnutrialbalance 
         Caption         =   "Trial Balance"
      End
      Begin VB.Menu mnutemptradingaccount 
         Caption         =   "Trading Account"
      End
      Begin VB.Menu mnubalancesheet 
         Caption         =   "Balance Sheet"
      End
   End
   Begin VB.Menu mnurepore 
      Caption         =   "Reports"
      Begin VB.Menu mnuinvprofit 
         Caption         =   "Invoice-wise Profit"
      End
      Begin VB.Menu mnugstreport 
         Caption         =   "GST Report"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuhsnsales 
         Caption         =   "HSN-Wise Sales"
      End
      Begin VB.Menu mnuststate 
         Caption         =   "Stock Statement"
      End
      Begin VB.Menu mnumargin 
         Caption         =   "Margin Sheet"
      End
      Begin VB.Menu mnuinprofit 
         Caption         =   "Invoice-wise Profit"
      End
      Begin VB.Menu mnuitemmovement 
         Caption         =   "Item Movement Report"
      End
      Begin VB.Menu mnustockstatement 
         Caption         =   "Stock Statement"
         Visible         =   0   'False
      End
      Begin VB.Menu mnucashbookprint 
         Caption         =   "Cashbook Print"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuoutstanding 
         Caption         =   "Out Standing "
      End
      Begin VB.Menu mnuareawiseoutstanding 
         Caption         =   "Area Wise Party Outstanding"
      End
      Begin VB.Menu mnupurchasereport 
         Caption         =   "Purchase Report"
      End
      Begin VB.Menu mnusalesreport 
         Caption         =   "Sales Report"
      End
      Begin VB.Menu mnuinputtax 
         Caption         =   "Input Tax"
      End
      Begin VB.Menu mnuoutputtax 
         Caption         =   "Output Tax"
      End
   End
   Begin VB.Menu mnutracking 
      Caption         =   "Tracking"
   End
   Begin VB.Menu mnutools 
      Caption         =   "Tools"
      Begin VB.Menu mnubackup 
         Caption         =   "BackUp"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnupasswordchange 
         Caption         =   "Password Change"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuutility 
         Caption         =   "Utilitu"
      End
      Begin VB.Menu mnumailsetup 
         Caption         =   "Mail Setup"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MDIForm_Load()
'Dim FileInQuestion As String, fname As String, driveinquestion
'fname = "FMCG" & Format(Date, "DD-MM-YYYY")
'    FileInQuestion = Dir("E:\databackup\" & fname & ".mdb")
'    If FileInQuestion = "" Then
'    'FileCopy dbname, "E:\databackup\" & fName & ".mdb"
'    Else
'    'Kill "E:\databackup\" & fName & ".mdb"
'    'FileCopy dbname, "E:\databackup\" & fName & ".mdb"
'    End If
'
'AccountingPeriod = "01/04/2008"
'Set db = OpenDatabase(dbname)
'CreateAccessODBC dbname, "FMCG", "FMCG"

End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
'Db.Close
'Set Db = Nothing
'Dim FileInQuestion As String, fName As String, driveinquestion
'Dim nFileNum As Integer
'Dim sFilename As String
'Kill App.Path & "\test.txt"
'sFilename = App.Path & "\test.txt"
'nFileNum = FreeFile
'Open sFilename For Binary Lock Read Write As nFileNum
'Put #nFileNum, , "0"
'Close #nFileNum
'      FileInQuestion = Dir(App.Path & "\Data\" & dbname & ".mdb")
'      If FileInQuestion = "" Then
'      FileCopy dbname, "D:\FMCG\DATA\" & dbname & ".mdb"
'      Kill dbname
'      Else
'
'        Kill App.Path & "\data\" & dbname & ".mdb"
'        FileCopy dbname, "D:\FMCG\DATA\" & dbname & ".mdb"
'        Kill dbname
'
'      End If

End Sub

Private Sub mnuadd_Click()
    frmitemslabmaster.Show 0
End Sub

Private Sub mnuareawiseoutstanding_Click()
frmPartyoutstanding.Show 0
End Sub

Private Sub mnubackup_Click()
    frmbackup_restore.Show 0
End Sub

Private Sub mnubalancesheet_Click()
    frmbalancesheet.Show 0
End Sub

Private Sub mnubarcode_Click()
    frmbarcode.Show 0
End Sub

Private Sub mnubranch_Click()
frmstockpoints.Show 0
End Sub

Private Sub mnubrandname_Click()
    frmbrandname.Show 0
End Sub

Private Sub mnucashbookprint_Click()
    frmCashBookView.Show 0
End Sub

Private Sub mnucompanymaster_Click()
    frmCompany.Show 0
End Sub
Private Sub mnucontra_Click()
    FrmContraVoucher.Show 0
End Sub
Private Sub mnucontravoucher_Click()
    frmContravoucherEdit.Show 0
End Sub
Private Sub mnucreditnoteedit_Click()
    frmCreditNoteEdit.Show 0
End Sub

Private Sub mnudamageentry_Click()
    frmdamageentry.Show 0
End Sub

Private Sub mnudanageitemview_Click()
frmdamageitemsview.Show 0
End Sub

Private Sub mnudamageen_Click()
frmdamageentry.Show 0
End Sub

Private Sub mnudamagereturn_Click()
frmdamagereturn.Show 0
End Sub

Private Sub mnudamagestockvi_Click()
frmDamageStock.Show 0
End Sub

Private Sub mnudebitnoteedit_Click()
    frmPurchaseReturnEdit.Show 0
End Sub
Private Sub mnudebtor_Click()
    frmPartyoutstanding.Show 0
End Sub

Private Sub mnudeliverychallan_Click()
frmdeliverychallan.Show 0
End Sub

Private Sub mnudrnote_Click()
frmdebitnote.Show 0
End Sub

Private Sub mnufing_Click()
frmFoodingrediant.Show 0
End Sub

Private Sub mnuflagmaster_Click()
    frmFlagMaster.Show 0
End Sub

Private Sub mnugitems_Click()
    frmglobalitemmaster.Show 0
End Sub

Private Sub mnugroup_Click()
    frmgroup.Show 0
End Sub
Private Sub mnugroupview_Click()
    FrmGroupView.Show 0
End Sub

Private Sub mnugstreport_Click()
   frmGSTReports.Show 0
End Sub

Private Sub mnuhsnsales_Click()
frmhsnwisesales.Show 0
End Sub

Private Sub mnuinprofit_Click()
frminvoicewiseprofit.Show 0
End Sub

Private Sub mnuinputtax_Click()
frmtaxinput.Show 0
End Sub

Private Sub mnuinvoice_Click()
    frmInvoice.Show 0
End Sub
Private Sub mnuinvoiceedit_Click()
    frmInvoiceEdit.Show 0
End Sub

Private Sub mnuinvprofit_Click()
frminvoicewiseprofit.Show 0
End Sub

Private Sub mnuitemedit_Click()
    frmItemEdit.Show 0
End Sub
Private Sub mnuitemmaster_Click()
    frmItemsmaster.Show 0
End Sub

Private Sub mnuitemmovement_Click()
frmitemmovementreport.Show 0
End Sub

Private Sub mnuitemsearch_Click()
    
End Sub

Private Sub mnuitemtype_Click()
    frmitemtype.Show 0
End Sub

Private Sub mnujournal_Click()
    frmJournal.Show 0
End Sub
Private Sub mnujournalvoucher_Click()
    frmJournalVoucherView.Show 0
End Sub

Private Sub mnulandbill_Click()
frmInvoiceLand.Show 0
End Sub

Private Sub mnuledgermaster_Click()
    frmledger.Show 0
End Sub
Private Sub mnuledgerview_Click()
    frmLedgerView.Show 0
End Sub
Private Sub mnulrmaster_Click()
    frmLrMaster.Show 0
End Sub

Private Sub mnumailsetup1_Click()
frmmailsetup.Show 0
End Sub

Private Sub mnumargin_Click()
frmmarginsheet.Show 0
End Sub

Private Sub mnunewitemmaster_Click()
    frmnewitemmaster.Show vbModal
End Sub

Private Sub mnunewjournal_Click()
    frmNewJournal.Show 0
End Sub
Private Sub mnunewjournalvoucher_Click()
    frmNewJournalView.Show 0
End Sub

Private Sub mnunewunit_Click()
    FrmUnitMaster.Show 0
End Sub

Private Sub mnuopbalance_Click()
    FrmOpBalance.Show 0
End Sub

Private Sub mnuopeningstock_Click()
    frmopeningstock.Show 0
End Sub

Private Sub mnuopstockedit_Click()
    frmopstockedit.Show 0
End Sub

Private Sub mnuorder_Click()
    'frmOrder.Show 0
End Sub


Private Sub mnuoutputtax_Click()
    frmTaxutput.Show 0
End Sub

Private Sub mnuoutstanding_Click()
    frmotheroutstanding.Show 0
End Sub

Private Sub mnuoutwardchallan_Click()
frmoutwardchallan.Show 0
End Sub

Private Sub mnupartycr_Click()
    frmPartyCr.Show 0
End Sub
Private Sub mnupartycredit_Click()
    frmPartyCrEdit.Show 0
End Sub
Private Sub mnupartycrledger_Click()
    frmPartyCrLedger.Show 0
End Sub
Private Sub mnupartydr_Click()
    frmPartyDr.Show 0
End Sub
Private Sub mnupartydredit_Click()
    FrmPartyDrEdit.Show 0
End Sub

Private Sub mnupasswordchange_Click()
    frmpasswordchange.Show 0
End Sub

Private Sub mnupayment_Click()
    frmPaymentVoucher.Show 0
End Sub
Private Sub mnupaymentvoucher_Click()
    frmPaymentvoucheredit.Show 0
End Sub
Private Sub mnupurchaseledger_Click()
    frmPurchaseledger.Show 0
End Sub

Private Sub mnupl8_Click()

End Sub

Private Sub mnuproducttype_Click()
    frmproducttype.Show 0
End Sub

Private Sub mnupurchaseorderview_Click()
    frmPurchaseorderview.Show 0
End Sub

Private Sub mnupurchasereport_Click()
frmpurchasereport.Show 0
End Sub

Private Sub mnupurchasereturn_Click()
    frmpurchasereturn.Show 0
End Sub

Private Sub mnureceipt_Click()
    FrmReceiptVoucher.Show 0
End Sub
Private Sub mnureceiptvoucher_Click()
    frmReceiptVoucherEdit.Show 0
End Sub

Private Sub mnuregistration_Click()
frmRegistration.Show vbModal
End Sub

Private Sub mnuretailinvoice_Click()
frmInvoiceR.Show 0
End Sub

Private Sub mnusalenote_Click()
    frmSalenote.Show 0
End Sub
Private Sub mnusalenoteview_Click()
    frmSalenoteview.Show 0
End Sub
Private Sub mnusalesledger_Click()
    frmSalesLedger.Show 0
End Sub

Private Sub mnusaleorder_Click()
    'frmsalesorder.Show 0
End Sub

Private Sub mnusalesreport_Click()
    frmsalesreport.Show 0
End Sub

Private Sub mnusalesreturn_Click()
    frmSalesReturn.Show 0
End Sub
Private Sub mnushademaster_Click()
    frmShadeMaster.Show 0
End Sub

Private Sub mnuscheme_Click()
frmscheme.Show 0
End Sub

Private Sub mnusize_Click()
    frmsize.Show 0
End Sub

Private Sub mnusortedit_Click()
    frmSortedit.Show 0
End Sub
Private Sub mnusortmaster_Click()
    frmQuality.Show vbModal
End Sub
Private Sub mnusortview_Click()
    frmSortView.Show 0
End Sub

Private Sub mnustockikn_b_Click()
frmStockint.Show 0
End Sub

Private Sub mnustockin_Click()
    frmStockin.Show 0
End Sub
Private Sub mnustockinedit_Click()
    frmStockinview.Show 0
End Sub
Private Sub mnustockstatement_Click()
    frmStockstatement.Show 0
End Sub

Private Sub mnustocktran_Click()
frmstocktransfer.Show 0
End Sub

Private Sub mnustockview_Click()
    'frmStockview.Show 0
    frmstock1.Show 0
End Sub
Private Sub mnusundrydebtor_Click()
    frmPartyDrLedger.Show 0
End Sub

Private Sub mnuststate_Click()
frmnewstockst.Show 0
End Sub

Private Sub mnutemptradingaccount_Click()
    frmTradingAcc.Show 0
End Sub

Private Sub mnutracking_Click()
frmtracking.Show 0
End Sub

Private Sub mnutrialbalance_Click()
    frmTrailBalance.Show 0
End Sub
Private Sub mnuunitcode_Click()
    frmUnitcode.Show 0
End Sub
Private Sub mnuvatledger_Click()
    frmVatLedger.Show 0
End Sub

Private Sub mnuutility_Click()
    frmutility.Show 0
End Sub

Private Sub mnuzonemaster_Click()
    FrmZoneMaster.Show 0
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.index = 1 Then
    frmStockin.Show 0
End If
If Button.index = 2 Then
    'frmOrder.Show 0
End If
If Button.index = 3 Then
    frmInvoice.Show 0
End If
If Button.index = 4 Then
    
End If
If Button.index = 5 Then
    frmstock.Show 0
End If
If Button.index = 8 Then
    frmLedgerView.Show 0
End If

If Button.index = 10 Then
    Shell (App.Path & "\tools\AMMYY_Admin.exe")
End If
End Sub

