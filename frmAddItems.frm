VERSION 5.00
Begin VB.Form frmAddItems 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image btnCancel 
      Height          =   855
      Left            =   4440
      Picture         =   "frmAddItems.frx":0000
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Image Image3 
      Height          =   3375
      Left            =   120
      Picture         =   "frmAddItems.frx":6B7A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7095
   End
   Begin VB.Image btnBahanPembantu 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   3240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmAddItems.frx":4BDDF
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   3975
   End
   Begin VB.Image btnBahanBaku 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   120
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmAddItems.frx":52DE6
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   3015
   End
End
Attribute VB_Name = "frmAddItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBahanBaku_Click()
    frmItems.AN_ITEMS
    frmAddItems.Hide
    frmItems.txtID.Enabled = False
    frmItems.btnCancel.Visible = True
    frmItems.btnAdd.Enabled = False
    frmItems.btnDelete.Enabled = False
    frmItems.btnSave.Enabled = True
End Sub

Private Sub btnBahanPembantu_Click()
    frmItems.AN_PEMBANTU
    frmAddItems.Hide
    frmItems.txtID.Enabled = False
    frmItems.btnCancel.Visible = True
    frmItems.btnAdd.Enabled = False
    frmItems.btnDelete.Enabled = False
    frmItems.btnSave.Enabled = True
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub
