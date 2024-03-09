VERSION 5.00
Begin VB.Form frmMaster_Data 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEmployee 
      BackColor       =   &H000080FF&
      Caption         =   "Data Pegawai"
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
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdAccount 
      BackColor       =   &H000080FF&
      Caption         =   "Data Akun"
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
   Begin VB.CommandButton cmdPeriod 
      BackColor       =   &H000080FF&
      Caption         =   "Data Periode"
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
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   5400
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmMaster_Data.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   120
      Picture         =   "frmMaster_Data.frx":6977
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmMaster_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    Me.Hide
    frmHome.Show
End Sub

Private Sub cmdAccount_Click()
    Me.Hide
    frmAccount.Show
End Sub

Private Sub cmdEmployee_Click()
    Me.Hide
    frmEmployee.Show
End Sub

Private Sub cmdPeriod_Click()
    Me.Hide
    frmPeriod.Show
End Sub
