VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "KIU LOGIN"
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6840
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   6840
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Image btnLogin 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   8520
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   4455
      Left            =   240
      Picture         =   "frmLogin.frx":4EA2
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   24
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6840
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Image btnExit 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   9720
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmLogin.frx":4A107
      Stretch         =   -1  'True
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   24
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6840
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4695
      Left            =   120
      Picture         =   "frmLogin.frx":503EA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExit_Click()
    End
End Sub

Private Sub btnLogin_Click()
    Call OpenDB
    If txtUser = "" Then
        frmEU.Show
    ElseIf txtPass = "" Then
        frmEP.Show
    End If
    
    rsLogin.Open "Select * From Login_Table Where user = '" & txtUser & "' and pass = '" & txtPass & "'", ConnectDB, adOpenDynamic, adLockOptimistic
    If rsLogin.EOF Then
        frmInvalid.Show
    Else
        If Trim(txtPass) = Trim(rsLogin.Fields("pass")) Then
            txtUser.Text = ""
            txtPass.Text = ""
            frmLogin_Alert.Show
        Else
            frmInvalid.Show
        End If
    End If
End Sub
