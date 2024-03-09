VERSION 5.00
Begin VB.Form frmLogin_Alert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image btnYes 
      Height          =   855
      Left            =   3960
      Picture         =   "frmLogin_Alert.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   120
      Picture         =   "frmLogin_Alert.frx":48D8
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmLogin_Alert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnYes_Click()
    frmHome.Show
    frmLogin.Hide
    Me.Hide
End Sub

Private Sub Form_Load()
    frmLogin_Alert.BackColor = RGB(195, 175, 166)
End Sub
