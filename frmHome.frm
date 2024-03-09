VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "KIU Home"
   ClientHeight    =   7005
   ClientLeft      =   -90
   ClientTop       =   2280
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdKalkulasi_Master 
      BackColor       =   &H000080FF&
      Caption         =   "Kalkulasi Bahan"
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
      Left            =   1920
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdJournal 
      BackColor       =   &H000080FF&
      Caption         =   "Jurnal Umum"
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
      Left            =   7440
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H000080FF&
      Caption         =   "Laporan Keuangan"
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
      Left            =   4680
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdMaster 
      BackColor       =   &H000080FF&
      Caption         =   "Data Master"
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
      Left            =   1920
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdProduction_Master 
      BackColor       =   &H000080FF&
      Caption         =   "Produksi"
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
      Left            =   4680
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdPurchase_Master 
      BackColor       =   &H000080FF&
      Caption         =   "Pembelian"
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
      Left            =   7440
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdSales_Master 
      BackColor       =   &H000080FF&
      Caption         =   "Penjualan"
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
      Left            =   1920
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Image btnLogout 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmHome.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   48
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6480
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   48
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image Image1 
      DragMode        =   1  'Automatic
      Height          =   6795
      Left            =   120
      Picture         =   "frmHome.frx":4F12
      Stretch         =   -1  'True
      Top             =   120
      Width           =   10065
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLogout_Click()
    Me.Hide
    frmLogin.Show
End Sub

Private Sub cmdJournal_Click()
    Me.Hide
    frmJournal.Show
End Sub

Private Sub cmdKalkulasi_Master_Click()
    Me.Hide
    frmKalkulasi.Show
End Sub

Private Sub cmdMaster_Click()
    Me.Hide
    frmMaster_Data.Show
End Sub

Private Sub cmdProduction_Master_Click()
    Me.Hide
    frmProduction_Master.Show
End Sub

Private Sub cmdPurchase_Master_Click()
    Me.Hide
    frmPurchase_Master.Show
End Sub

Private Sub cmdReport_Click()
    Me.Hide
    frmReport.Show
End Sub

Private Sub cmdSales_Master_Click()
    Me.Hide
    frmSales_Master.Show
End Sub
