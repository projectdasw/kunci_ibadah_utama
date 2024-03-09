VERSION 5.00
Begin VB.Form frmSales_Master 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCustomer 
      BackColor       =   &H000080FF&
      Caption         =   "Data Pelanggan"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdTrans_Sales 
      BackColor       =   &H000080FF&
      Caption         =   "Transaksi Penjualan"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdProduct 
      BackColor       =   &H000080FF&
      Caption         =   "Data Produk"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   120
      Picture         =   "frmSales_Master.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4695
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   5400
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSales_Master.frx":45265
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1935
   End
End
Attribute VB_Name = "frmSales_Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    Me.Hide
    frmHome.Show
End Sub

Private Sub cmdCustomer_Click()
    Me.Hide
    frmCustomer.Show
End Sub

Private Sub cmdProduct_Click()
    Me.Hide
    frmProduct.Show
End Sub

Private Sub cmdTrans_Sales_Click()
    Me.Hide
    frmSales.Show
End Sub
