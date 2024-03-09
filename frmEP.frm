VERSION 5.00
Begin VB.Form frmEP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image btnYes 
      Height          =   855
      Left            =   3960
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmEP.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   120
      Picture         =   "frmEP.frx":48D8
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnYes_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    frmEP.BackColor = RGB(195, 175, 166)
End Sub
