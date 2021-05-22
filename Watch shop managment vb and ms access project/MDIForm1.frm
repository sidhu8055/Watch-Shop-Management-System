VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   4635
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14760
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuadd 
      Caption         =   "Add"
      Begin VB.Menu mnuCustomer 
         Caption         =   "Customer"
      End
      Begin VB.Menu mnuWatchDetails 
         Caption         =   "Watch Details"
      End
      Begin VB.Menu mnuDealerDetails 
         Caption         =   "Dealer Details"
      End
      Begin VB.Menu mnuPurchase 
         Caption         =   "Purchase"
      End
      Begin VB.Menu mnuSales 
         Caption         =   "Sales"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "Report"
      Begin VB.Menu mnuCustomer1 
         Caption         =   "Customer"
      End
      Begin VB.Menu mnuWatchDetails1 
         Caption         =   "Watch Details"
      End
      Begin VB.Menu mnuDealerDetails1 
         Caption         =   "Dealer Details"
      End
      Begin VB.Menu mnuPurchase1 
         Caption         =   "Purchase"
      End
      Begin VB.Menu mnuSales1 
         Caption         =   "Sales"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuCustomer_Click()
frmCustomer.Show
End Sub

Private Sub mnuCustomer1_Click()
DataReport3.Show
End Sub

Private Sub mnuDealerDetails_Click()
frmDealerDetails.Show
End Sub

Private Sub mnuDealerDetails1_Click()
DataReport2.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuPurchase_Click()
frmPurchase.Show
End Sub

Private Sub mnuPurchase1_Click()
DataReport5.Show
End Sub

Private Sub mnuSales_Click()
frmSales.Show
End Sub

Private Sub mnuSales1_Click()
DataReport4.Show
End Sub

Private Sub mnuWatchDetails_Click()
frmWatchDetails.Show
End Sub

Private Sub mnuWatchDetails1_Click()
DataReport1.Show
End Sub
