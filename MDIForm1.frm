VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "MEDICINE MANAGEMENT SYSTEM FOR MAA MANGLA PHARMA"
   ClientHeight    =   8400
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   16200
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu PRODUCT 
      Caption         =   "PRODUCT"
   End
   Begin VB.Menu SUPPLIERS 
      Caption         =   "SUPPLIERS"
   End
   Begin VB.Menu ORDER 
      Caption         =   "ORDER"
   End
   Begin VB.Menu PURCHASE 
      Caption         =   "PURCHASE"
   End
   Begin VB.Menu STOCK 
      Caption         =   "STOCK"
   End
   Begin VB.Menu CUSTOMER_DETAILS 
      Caption         =   "CUSTOMER_DETAILS"
   End
   Begin VB.Menu CUSTOMER_ORDER_DETAILS 
      Caption         =   "CUSTOMER_ORDER_DETAILS"
   End
   Begin VB.Menu SALES_BILL 
      Caption         =   "SALES_BILL"
   End
   Begin VB.Menu CALCULATOR 
      Caption         =   "CALCULATOR"
   End
   Begin VB.Menu EXIT 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CALCULATOR_Click()
Unload SUP_DTL
Unload ORD_DTL
Unload PUR_DTL
Unload PRD_DTL
Unload STK_DTL
Unload CUS_DTL
Unload CUS_ORD_DTL
Unload INV_DTL
Unload CAL
CAL.Show
End Sub
Private Sub CHEQUE_DETAILS_Click()
CHQ_DTL.Show
End Sub
Private Sub CUSTOMER_DETAILS_Click()
Unload SUP_DTL
Unload ORD_DTL
Unload PUR_DTL
Unload PRD_DTL
Unload STK_DTL
Unload CUS_ORD_DTL
Unload INV_DTL
Unload CAL
CUS_DTL.Show
End Sub
Private Sub CUSTOMER_ORDER_DETAILS_Click()
Unload SUP_DTL
Unload ORD_DTL
Unload PUR_DTL
Unload PRD_DTL
Unload STK_DTL
Unload CUS_DTL
Unload INV_DTL
Unload CAL
CUS_ORD_DTL.Show
End Sub
Private Sub EXIT_Click()
Unload Me
LOGIN_FORM.Show
End Sub
Private Sub SALES_BILL_Click()
Unload SUP_DTL
Unload ORD_DTL
Unload PUR_DTL
Unload PRD_DTL
Unload STK_DTL
Unload CUS_DTL
Unload CUS_ORD_DTL
Unload CAL
INV_DTL.Show
End Sub
Private Sub STOCK_Click()
Unload SUP_DTL
Unload ORD_DTL
Unload PUR_DTL
Unload PRD_DTL
Unload CUS_DTL
Unload CUS_ORD_DTL
Unload INV_DTL
Unload CAL
STK_DTL.Show
End Sub
Private Sub CUSTOMER_Click()
Unload SUP_DTL
Unload ORD_DTL
Unload PUR_DTL
Unload PRD_DTL
Unload STK_DTL
Unload CUS_ORD_DTL
Unload INV_DTL
Unload CAL
CUS_DTL.Show
End Sub
Private Sub ORDER_Click()
Unload SUP_DTL
Unload PUR_DTL
Unload PRD_DTL
Unload STK_DTL
Unload CUS_DTL
Unload CUS_ORD_DTL
Unload INV_DTL
Unload CAL
ORD_DTL.Show
End Sub
Private Sub PRODUCT_Click()
Unload SUP_DTL
Unload ORD_DTL
Unload PUR_DTL
Unload STK_DTL
Unload CUS_DTL
Unload CUS_ORD_DTL
Unload INV_DTL
Unload CAL
PRD_DTL.Show
End Sub
Private Sub PURCHASE_Click()
Unload SUP_DTL
Unload ORD_DTL
Unload PRD_DTL
Unload STK_DTL
Unload CUS_DTL
Unload CUS_ORD_DTL
Unload INV_DTL
Unload CAL
PUR_DTL.Show
End Sub
Private Sub SUPPLIERS_Click()
Unload PRD_DTL
Unload ORD_DTL
Unload PUR_DTL
Unload STK_DTL
Unload CUS_DTL
Unload CUS_ORD_DTL
Unload INV_DTL
Unload CAL
SUP_DTL.Show
End Sub
