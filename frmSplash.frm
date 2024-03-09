VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Loading_Timer 
      Interval        =   150
      Left            =   5760
      Top             =   3240
   End
   Begin ComctlLib.ProgressBar Loading_Bar 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
      Max             =   101
   End
   Begin VB.Label Load_Percent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading ..."
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   120
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Loading_Timer_Timer()
    Loading_Bar.Value = Loading_Bar.Value + 1
    Load_Percent.Caption = Loading_Bar.Value & "%"
    If (Loading_Bar.Value = Loading_Bar.Max) Then
        Loading_Timer.Enabled = False
        frmLogin.Show
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Loading_Timer.Enabled = True
End Sub
