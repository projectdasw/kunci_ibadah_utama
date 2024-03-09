VERSION 5.00
Begin VB.Form frmProduction_Master 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProduction 
      BackColor       =   &H000080FF&
      Caption         =   "Data Produksi"
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
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdItems 
      BackColor       =   &H000080FF&
      Caption         =   "Data Bahan Baku"
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
      Left            =   2640
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   2880
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduction_Master.frx":0000
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   120
      Picture         =   "frmProduction_Master.frx":6977
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmProduction_Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
    Me.Hide
    frmHome.Show
End Sub

Private Sub cmdItems_Click()
    Me.Hide
    frmItems.Show
End Sub

Private Sub cmdProduction_Click()
    Me.Hide
    frmProduction.Show
End Sub
