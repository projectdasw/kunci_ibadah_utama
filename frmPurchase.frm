VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPurchase 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10695
   ClientLeft      =   -60
   ClientTop       =   -60
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   10695
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotal_Lama 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   35
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox txtHarga_Lama 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   34
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox txtKuan_Lama 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   33
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox txtNota 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5640
      TabIndex        =   31
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtGrand 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   6960
      Width           =   2775
   End
   Begin VB.TextBox txtBayar 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   26
      Top             =   6960
      Width           =   2775
   End
   Begin VB.TextBox txtKembali 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   25
      Top             =   6960
      Width           =   2775
   End
   Begin VB.TextBox txtTotal 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   5880
      Width           =   2175
   End
   Begin VB.ComboBox cmbPeriode 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   20
      Top             =   5880
      Width           =   1935
   End
   Begin VB.TextBox txtID_Supp 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2520
      TabIndex        =   18
      Top             =   5880
      Width           =   2175
   End
   Begin VB.ComboBox cmbNama_Supp 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   17
      Top             =   5880
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dateTrans 
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   3600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   50921473
      CurrentDate     =   44284
   End
   Begin VB.TextBox txtTrans 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5640
      TabIndex        =   13
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FORM PEMBELIAN"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   5175
      Begin VB.ComboBox cmbNama 
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2040
         TabIndex        =   23
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtKuan 
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox txtSatuan 
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox txtHarga 
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtID_Bahan 
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Kuantitas"
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Bahan"
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bahan"
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Bahan"
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView Purchase_ListView 
      Height          =   2055
      Left            =   240
      TabIndex        =   30
      Top             =   7440
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   10
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Bahan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Bahan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Harga Bahan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Satuan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Kuantitas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Lama"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   38
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Lama"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   37
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Kuantitas Lama"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   36
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Nota"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   32
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Image btnPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   7200
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPurchase.frx":0000
      Stretch         =   -1  'True
      Top             =   9840
      Width           =   1815
   End
   Begin VB.Image btnHitung 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   1200
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPurchase.frx":54CD
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   855
   End
   Begin VB.Image btnCancel 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   5280
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPurchase.frx":B30E
      Stretch         =   -1  'True
      Top             =   9840
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Bayar"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   27
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6240
      TabIndex        =   24
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Image btnDelete 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   3120
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPurchase.frx":11E88
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   855
   End
   Begin VB.Image btnSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   2160
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPurchase.frx":1702D
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   855
   End
   Begin VB.Image btnAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPurchase.frx":1CA88
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Harga"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   21
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Periode"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Transaksi"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PEMBELIAN"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   27.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSAKSI"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   27.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPurchase.frx":21B32
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   10455
      Left            =   120
      Picture         =   "frmPurchase.frx":284A9
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ClearText()
    Purchase_ListView.ListItems.Clear
    txtID_Bahan.Text = ""
    cmbNama = ""
    txtHarga.Text = ""
    txtSatuan.Text = ""
    txtKuan.Text = ""
    txtTotal.Text = ""
    dateTrans.Value = Format(Now)
    cmbNama_Supp = ""
    txtID_Supp.Text = ""
    cmbPeriode = ""
    txtGrand.Text = ""
    txtBayar.Text = ""
    txtKembali.Text = ""
    txtKuan_Lama.Text = ""
    txtHarga_Lama.Text = ""
    txtTotal_Lama.Text = ""
    txtNota.Text = ""
End Sub

Sub AN_TRANS()
    Call OpenDB
    rsPurchase.Open ("SELECT * FROM Purchase_Table WHERE id_trans in(select max(id_trans) from Purchase_Table)order by id_trans desc"), ConnectDB
    rsPurchase.Requery
    Dim Urut As String * 8
    Dim Hitung As Long
    With rsPurchase
        If .EOF Then
            Urut = "IPB-" + "0001"
            txtTrans.Text = Urut
        Else
            Hitung = Right(!id_trans, 4) + 1
            Urut = "IPB-" + Right("0000" & Hitung, 4)
        End If
        txtTrans.Text = Urut
    End With
End Sub



Private Sub btnAdd_Click()
    Dim lv As MSComctlLib.ListItem
    If txtID_Bahan.Text = "" Or cmbNama.Text = "" Or txtHarga.Text = "" Or txtSatuan.Text = "" Or txtKuan.Text = "" Or txtTotal.Text = "" Then
        frmEmpty.Show
    Else
        Set lv = Purchase_ListView.ListItems.Add(, , txtID_Bahan.Text)
        With lv
            .SubItems(1) = cmbNama.Text
            .SubItems(2) = txtHarga.Text
            .SubItems(3) = txtSatuan.Text
            .SubItems(4) = txtKuan.Text
            .SubItems(5) = txtTotal.Text
        End With
        
        btnSave.Enabled = True
        btnDelete.Enabled = True
        btnHitung.Enabled = True
        txtID_Bahan.Text = ""
        cmbNama = ""
        txtHarga.Text = ""
        txtSatuan.Text = ""
        txtKuan.Text = ""
        txtTotal.Text = ""
    End If
End Sub

Private Sub btnBack_Click()
    Me.Hide
    frmHome.Show
End Sub

Private Sub btnDelete_Click()
    If Purchase_ListView.ListItems.Count = 0 Then
        frmEmpty.Show
    Else
        Purchase_ListView.ListItems.Remove (Purchase_ListView.SelectedItem.Index)
        txtGrand.Text = ""
    End If
End Sub

Private Sub btnHitung_Click()
    If Purchase_ListView.ListItems.Count > 0 Then
        Dim GrandTotal_list As Double
        Dim i As Integer
        
        For i = 1 To Purchase_ListView.ListItems.Count
            GrandTotal_list = GrandTotal_list + Int(Purchase_ListView.ListItems(i).ListSubItems(5).Text)
        Next i
        txtGrand.Text = GrandTotal_list
    End If
End Sub

Private Sub btnPrint_Click()
    Export_Excel "Kunci_Ibadah_Utama.xlsx", "Data Pembelian"
End Sub

Private Sub btnSave_Click()
    Call OpenDB
    Dim i As Integer
    
    If Purchase_ListView.ListItems.Count = 0 Then
        frmNone.Show
    ElseIf txtID_Supp.Text = "" Or cmbNama_Supp.Text = "" Or cmbPeriode.Text = "" Or txtGrand.Text = "" Or txtBayar.Text = "" Or txtKembali.Text = "" Then
        frmEmpty.Show
    Else
        For i = 1 To Purchase_ListView.ListItems.Count
            With rsPurchase
                .Open "select * from PurchaseData_Table", ConnectDB, 1, 2
                .AddNew
                !id_bahan = Purchase_ListView.ListItems(i).Text
                !nama_bahan = Purchase_ListView.ListItems(i).ListSubItems(1).Text
                !harga_bahan = Purchase_ListView.ListItems(i).ListSubItems(2).Text
                !satuan = Purchase_ListView.ListItems(i).ListSubItems(3).Text
                !kuantitas = Purchase_ListView.ListItems(i).ListSubItems(4).Text
                !total_bahan = Purchase_ListView.ListItems(i).ListSubItems(5).Text
                !id_trans = txtTrans.Text
                !no_nota = txtNota.Text
                .Update
                .Close
            End With
            With rsPurchase
            .Open "select * from ItemsData_Table", ConnectDB
            .AddNew
            !id_trans = txtTrans.Text
            !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
            !keterangan = "Pembelian Bahan"
            !id_bahan = Purchase_ListView.ListItems(i).Text
            !kuan_lama = txtKuan_Lama.Text
            !harga_lama = txtHarga_Lama.Text
            !total_lama = txtTotal_Lama.Text
            !kuan_beli = Purchase_ListView.ListItems(i).ListSubItems(4).Text
            !harga_beli = Purchase_ListView.ListItems(i).ListSubItems(2).Text
            !total_beli = Purchase_ListView.ListItems(i).ListSubItems(5).Text
            !kuan_baru = Val(txtKuan_Lama.Text) + Val(Purchase_ListView.ListItems(i).ListSubItems(4).Text)
            !harga_baru = (Val(txtTotal_Lama.Text) + Val(Purchase_ListView.ListItems(i).ListSubItems(5).Text)) / (Val(txtKuan_Lama.Text) + Val(Purchase_ListView.ListItems(i).ListSubItems(4).Text))
            !total_baru = Val(txtTotal_Lama.Text) + Val(Purchase_ListView.ListItems(i).ListSubItems(5).Text)
            .Update
            .Close
        End With
        Next i
        With rsPurchase
            .Open "select * from Purchase_Table", ConnectDB
            .AddNew
            !id_trans = txtTrans.Text
            !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
            !id_supplier = txtID_Supp.Text
            !nama_supplier = cmbNama_Supp.Text
            !periode = cmbPeriode.Text
            !grandtotal = txtGrand.Text
            !bayar = txtBayar.Text
            !kembali = txtKembali.Text
            .Update
            .Close
        End With
        With rsPurchase
            .Open "select * from Journal_Data", ConnectDB
            .AddNew
            !id_trans = txtTrans.Text
            !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
            !periode = cmbPeriode.Text
            !uraian = "Pembelian Bahan"
            !nama_akun1 = "PERSEDIAAN BAHAN BAKU"
            !id_akun1 = "1.1.2.1"
            !debit = txtGrand.Text
            !nama_akun2 = "KAS"
            !id_akun2 = "1.1.1.1"
            !kredit = txtGrand.Text
            .Update
            .Close
        End With
        
        'Script Penambahan Persediaan Stok Barang
        i = 1
        Do While i <= Purchase_ListView.ListItems.Count
            kode = Purchase_ListView.ListItems(i)
            jmlbeli = Purchase_ListView.ListItems(i).ListSubItems(4).Text
            
            Buka_TBahan kode
            stok = rsItems!kuantitas
            
            updateStok = Val(stok) + Val(jmlbeli)
            Ubah_TBahan kode, updateStok
            i = i + 1
        Loop
        
        'Script Update Harga & Stok Barang
        i = 1
        Do While i <= Purchase_ListView.ListItems.Count
            kode = Purchase_ListView.ListItems(i).Text
            hargabahan = Purchase_ListView.ListItems(i).ListSubItems(5).Text
            
            Buka_TBahan kode
            stok = rsItems!kuantitas
            hargabahan2 = rsItems!harga_bahan
            t_bahan = rsItems!total_bahan
            
            hargabahan2 = (Val(t_bahan) + Val(hargabahan)) / Val(stok)
            Ubah_HargaBahan kode, hargabahan2
            i = i + 1
        Loop
        
        'Script Kalkulasi Songkok
        i = 1
        Do While i <= Purchase_ListView.ListItems.Count
            kode = Purchase_ListView.ListItems(i).Text
            hargabahan = Purchase_ListView.ListItems(i).ListSubItems(5).Text
            
            If kode = "IB-0001" Then
                Buka_Kalsoba kode
                stok = rsItems!kuantitas
                total_kalsoba = rsKalsoba!harga
                t_bahan = rsItems!total_bahan
                
                total_kalsoba = (Val(t_bahan) + Val(hargabahan)) / Val(stok)
                Ubah_TotalKalsoba kode, total_kalsoba
                i = i + 1
                
                Buka_Kalsoba kode
                stok_kalsoba = rsKalsoba!kuantitas
                harga_kalsoba = rsKalsoba!harga
                total_kalso = rsKalsoba!Total
                
                total_kalso = Val(harga_kalsoba) * Val(stok_kalsoba)
                Ubah_TBKalso "Kalkulasi_SongkokBagus", kode, total_kalso
                i = i + 1
            
            ElseIf kode = "IB-0002" Then
                Buka_Kalsobi kode
                stok = rsItems!kuantitas
                total_kalsobi = rsKalsobi!harga
                t_bahan = rsItems!total_bahan
                
                total_kalsobi = (Val(t_bahan) + Val(hargabahan)) / Val(stok)
                Ubah_TotalKalsobi kode, total_kalsobi
                i = i + 1
                
                Buka_Kalsobi kode
                stok_kalsobi = rsKalsobi!kuantitas
                harga_kalsobi = rsKalsobi!harga
                total_kalsi = rsKalsobi!Total
                
                total_kalsi = Val(harga_kalsobi) * Val(stok_kalsobi)
                Ubah_TBKalso "Kalkulasi_SongkokBiasa", kode, total_kalsi
                i = i + 1
            Else
                Buka_Kalsoba kode
                stok = rsItems!kuantitas
                total_kalsoba = rsKalsoba!harga
                t_bahan = rsItems!total_bahan
                
                total_kalsoba = (Val(t_bahan) + Val(hargabahan)) / Val(stok)
                Ubah_TotalKalsoba kode, total_kalsoba
                i = i + 1
                
                Buka_Kalsoba kode
                stok_kalsoba = rsKalsoba!kuantitas
                harga_kalsoba = rsKalsoba!harga
                total_kalso = rsKalsoba!Total
                
                total_kalso = Val(harga_kalsoba) * Val(stok_kalsoba)
                Ubah_TBKalso "Kalkulasi_SongkokBagus", kode, total_kalso
                i = i + 1
            
                Buka_Kalsobi kode
                stok = rsItems!kuantitas
                total_kalsobi = rsKalsobi!harga
                t_bahan = rsItems!total_bahan
                
                total_kalsobi = (Val(t_bahan) + Val(hargabahan)) / Val(stok)
                Ubah_TotalKalsobi kode, total_kalsobi
                i = i + 1
                
                Buka_Kalsobi kode
                stok_kalsobi = rsKalsobi!kuantitas
                harga_kalsobi = rsKalsobi!harga
                total_kalsi = rsKalsobi!Total
                
                total_kalsi = Val(harga_kalsobi) * Val(stok_kalsobi)
                Ubah_TBKalso "Kalkulasi_SongkokBiasa", kode, total_kalsi
                i = i + 1
            End If
        Loop
        
        'Script Update Total Harga Barang
        i = 1
        Do While i <= Purchase_ListView.ListItems.Count
            kode = Purchase_ListView.ListItems(i).Text
            
            Buka_TBahan kode
            stok = rsItems!kuantitas
            harga = rsItems!harga_bahan
            t_bahan = rsItems!total_bahan
            
            t_bahan = Val(harga) * Val(stok)
            Ubah_TotalBahan kode, t_bahan
            i = i + 1
        Loop
        
        'Script Perselisihan & Penulisan Jurnal Pembelian
        i = 1
        Do While i <= Purchase_ListView.ListItems.Count
            'PERSEDIAAN BAHAN BAKU
            Buka_Jurnal "1.1.2.1"
            SAP = rsJournal!saldo_akhir
            
            updateSAP = Val(SAP) + Val(txtGrand.Text)
            Ubah_Jurnal "1.1.2.1", updateSAP
            i = i + 1
            
            'KAS
            Buka_Jurnal "1.1.1.1"
            SAK = rsJournal!saldo_akhir
            
            updateSAK = Val(SAK) - Val(txtGrand.Text)
            Ubah_Jurnal "1.1.1.1", updateSAK
            i = i + 1
        Loop
        
        frmSave_Alert.Show
        ClearText
        AN_TRANS
        btnAdd.Enabled = True
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnHitung.Enabled = True
        btnCancel.Visible = False
    End If
End Sub

Private Sub cmbNama_Click()
    txtID_Bahan = Tampil_IDBahan("Items_Table", "id_bahan", cmbNama)
    txtSatuan = Tampil_IDBahan("Items_Table", "satuan", cmbNama)
    txtKuan_Lama.Text = Tampil_IDBahan("Items_Table", "kuantitas", cmbNama)
    txtHarga_Lama.Text = Tampil_IDBahan("Items_Table", "harga_bahan", cmbNama)
    txtTotal_Lama.Text = Tampil_IDBahan("Items_Table", "total_bahan", cmbNama)
    'txtHarga_Lama.Text = Format(txtHarga_Lama.Text, "#,##0")
    'txtTotal_Lama.Text = Format(txtTotal_Lama.Text, "#,##0")
    txtID_Bahan.Enabled = False
    txtSatuan.Enabled = False
    btnCancel.Visible = True
End Sub

Private Sub cmbNama_Supp_Click()
    txtID_Supp = Tampil_IDSupplier("Supplier_Table", "id_supplier", cmbNama_Supp)
    txtID_Supp.Enabled = False
End Sub

Private Sub Form_Load()
    Call OpenDB
    AN_TRANS
    txtTrans.Enabled = False
    txtGrand.Enabled = False
    txtKembali.Enabled = False
    btnCancel.Visible = False
    btnSave.Enabled = True
    btnDelete.Enabled = False
    btnHitung.Enabled = True
    dateTrans.Value = Format(Now)
    rsItems.Open "select * from Items_Table", ConnectDB
    rsSupp.Open "select * from Supplier_Table", ConnectDB
    rsPer.Open "select * from Periode_Table", ConnectDB
    cmbNama.Clear
    cmbNama_Supp.Clear
    cmbPeriode.Clear
    Do While Not rsItems.EOF
        cmbNama.AddItem rsItems!nama_bahan
        rsItems.MoveNext
    Loop
    Do While Not rsSupp.EOF
        cmbNama_Supp.AddItem rsSupp!nama_supplier
        rsSupp.MoveNext
    Loop
    Do While Not rsPer.EOF
        cmbPeriode.AddItem rsPer!masa
        rsPer.MoveNext
    Loop
End Sub

Private Sub btnCancel_Click()
    Purchase_ListView.ListItems.Clear
    txtID_Bahan.Text = ""
    cmbNama = ""
    txtHarga.Text = ""
    txtSatuan.Text = ""
    txtKuan.Text = ""
    txtTotal.Text = ""
    dateTrans.Value = Format(Now)
    cmbNama_Supp = ""
    txtID_Supp.Text = ""
    cmbPeriode = ""
    txtGrand.Text = ""
    txtBayar.Text = ""
    txtKembali.Text = ""
    txtKuan_Lama.Text = ""
    txtHarga_Lama.Text = ""
    txtTotal_Lama.Text = ""
    btnCancel.Visible = False
End Sub

Private Sub txtBayar_Change()
    txtKembali.Text = Val(txtBayar.Text) - Val(txtGrand.Text)
End Sub

Private Sub txtKuan_Change()
    txtTotal.Text = Val(txtHarga.Text) * Val(txtKuan.Text)
End Sub
