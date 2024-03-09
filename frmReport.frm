VERSION 5.00
Begin VB.Form frmReport 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   13830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnReport_BahanBaku 
      BackColor       =   &H000080FF&
      Caption         =   "Harga Pokok Pembelian"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10200
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmReport.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      Left            =   9840
      TabIndex        =   1
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT"
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
      Left            =   7920
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Image btnTutupBuku 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   7080
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmReport.frx":6977
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Image btnNeraca 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   3000
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmReport.frx":CA46
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Image btnLR 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   7080
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmReport.frx":131CC
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Image btnJournal 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   3000
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmReport.frx":1928C
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Image btnHPP 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   7080
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmReport.frx":1FB98
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Image btnBukuBesar 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   3000
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmReport.frx":255D7
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   8775
      Left            =   120
      Picture         =   "frmReport.frx":2B9C6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   13575
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    Me.Hide
    frmHome.Show
End Sub

Private Sub btnBukuBesar_Click()
    Report_BBHPP "Kunci_Ibadah_Utama.xlsx", "Buku Besar"
End Sub

Private Sub btnHPP_Click()
    Report_BBHPP "Kunci_Ibadah_Utama.xlsx", "HPP"
End Sub

Private Sub btnJournal_Click()
    Report_Jurnal "Kunci_Ibadah_Utama.xlsx", "Jurnal"
End Sub

Private Sub btnLR_Click()
    Report_LR "Kunci_Ibadah_Utama.xlsx", "Laba_Rugi", "Account_Table", "saldo_akhir"
End Sub

Private Sub btnNeraca_Click()
    Report_Neraca "Kunci_Ibadah_Utama.xlsx", "Neraca", "Account_Table", "saldo_akhir"
End Sub

Private Sub btnReport_BahanBaku_Click()
    Report_JurnalBahanBaku "Kunci_Ibadah_Utama.xlsx", "Harga Pokok Pembelian"
End Sub

Private Sub btnTutupBuku_Click()
    Report_TutupBuku "Kunci_Ibadah_Utama.xlsx", "Tutup Buku"
End Sub
