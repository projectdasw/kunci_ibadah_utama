VERSION 5.00
Begin VB.Form frmSave_Alert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image btnYes 
      Height          =   855
      Left            =   3960
      Picture         =   "frmSave_Alert.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   120
      Picture         =   "frmSave_Alert.frx":48D8
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmSave_Alert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnYes_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    frmSave_Alert.BackColor = RGB(195, 175, 166)
End Sub
