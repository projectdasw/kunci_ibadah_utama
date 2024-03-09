VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSales 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10305
   ClientLeft      =   -60
   ClientTop       =   -60
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHPP 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6840
      TabIndex        =   32
      Top             =   3840
      Width           =   2175
   End
   Begin MSComctlLib.ListView Sales_ListView 
      Height          =   3495
      Left            =   240
      TabIndex        =   30
      Top             =   5520
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6165
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
         Text            =   "ID Produk"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Produk"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Harga Produk"
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
   Begin VB.TextBox txtGrand 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker dateTrans 
      Height          =   495
      Left            =   4440
      TabIndex        =   28
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   127008769
      CurrentDate     =   44287
   End
   Begin VB.TextBox txtKembali 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6240
      TabIndex        =   24
      Top             =   4920
      Width           =   2775
   End
   Begin VB.TextBox txtBayar 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      TabIndex        =   22
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FORM PENJUALAN"
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
      TabIndex        =   6
      Top             =   960
      Width           =   4095
      Begin VB.ComboBox cmbNama 
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
         Left            =   2160
         TabIndex        =   27
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtID 
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
         Left            =   2160
         TabIndex        =   10
         Top             =   600
         Width           =   1695
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
         Left            =   2160
         TabIndex        =   9
         Top             =   1560
         Width           =   1695
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
         Left            =   2160
         TabIndex        =   8
         Top             =   2040
         Width           =   1695
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
         Left            =   2160
         TabIndex        =   7
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Produk"
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
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Produk"
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
         TabIndex        =   14
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Produk"
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
         TabIndex        =   13
         Top             =   1560
         Width           =   1815
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
         Left            =   1080
         TabIndex        =   12
         Top             =   2040
         Width           =   975
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
         Left            =   720
         TabIndex        =   11
         Top             =   2520
         Width           =   1335
      End
   End
   Begin VB.TextBox txtTrans 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox cmbNama_Pel 
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
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtID_Pel 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6840
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox cmbPeriode 
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
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtTotal 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   0
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "HPP Per Unit"
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
      Left            =   6960
      TabIndex        =   31
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Image btnPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   6960
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSales.frx":0000
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   2055
   End
   Begin VB.Image btnHitung 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   1320
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSales.frx":54CD
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   975
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
      TabIndex        =   29
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "PENJUALAN"
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
      Left            =   5400
      TabIndex        =   26
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label13 
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
      Left            =   1920
      TabIndex        =   25
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      TabIndex        =   23
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bayar"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3240
      TabIndex        =   21
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Left            =   4440
      TabIndex        =   20
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSales.frx":B30E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1575
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
      Left            =   4440
      TabIndex        =   19
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Pelanggan"
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
      Left            =   7080
      TabIndex        =   18
      Top             =   840
      Width           =   1695
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
      Left            =   7320
      TabIndex        =   17
      Top             =   2400
      Width           =   1335
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
      Left            =   4440
      TabIndex        =   16
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Image btnAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSales.frx":11C85
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   975
   End
   Begin VB.Image btnSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   2400
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSales.frx":16D2F
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   975
   End
   Begin VB.Image btnDelete 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   3480
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSales.frx":1C78A
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   975
   End
   Begin VB.Image btnCancel 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   4800
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSales.frx":2192F
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   10095
      Left            =   120
      Picture         =   "frmSales.frx":284A9
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ClearText()
    Sales_ListView.ListItems.Clear
    txtID.Text = ""
    cmbNama = ""
    txtHarga.Text = ""
    txtSatuan.Text = ""
    txtKuan.Text = ""
    txtTotal.Text = ""
    dateTrans.Value = Format(Now)
    cmbNama_Pel.Text = ""
    txtID_Pel.Text = ""
    cmbPeriode = ""
    txtGrand.Text = ""
    txtBayar.Text = ""
    txtKembali.Text = ""
    txtHPP.Text = ""
End Sub

Sub AN_TRANS()
    Call OpenDB
    rsSales.Open ("SELECT * FROM Sales_Table WHERE id_trans in(select max(id_trans) from Sales_Table)order by id_trans desc"), ConnectDB
    rsSales.Requery
    Dim Urut As String * 8
    Dim Hitung As Long
    With rsSales
        If .EOF Then
            Urut = "IPJ-" + "0001"
            txtTrans.Text = Urut
        Else
            Hitung = Right(!id_trans, 4) + 1
            Urut = "IPJ-" + Right("0000" & Hitung, 4)
        End If
        txtTrans.Text = Urut
    End With
End Sub

Private Sub btnAdd_Click()
    Dim lv As MSComctlLib.ListItem
    If txtID.Text = "" Or cmbNama.Text = "" Or txtHarga.Text = "" Or txtSatuan.Text = "" Or txtKuan.Text = "" Or txtTotal.Text = "" Then
        frmEmpty.Show
    Else
        Set lv = Sales_ListView.ListItems.Add(, , txtID.Text)
        With lv
            .SubItems(1) = cmbNama.Text
            .SubItems(2) = txtHarga.Text
            .SubItems(3) = txtSatuan.Text
            .SubItems(4) = txtKuan.Text
            .SubItems(5) = txtTotal.Text
        End With
    End If
    
    txtID.Text = ""
    cmbNama = ""
    txtHarga.Text = ""
    txtSatuan.Text = ""
    txtKuan.Text = ""
    txtTotal.Text = ""
End Sub

Private Sub btnBack_Click()
    Me.Hide
    frmHome.Show
End Sub

Private Sub btnCancel_Click()
    Sales_ListView.ListItems.Clear
    txtID.Text = ""
    cmbNama = ""
    txtHarga.Text = ""
    txtSatuan.Text = ""
    txtKuan.Text = ""
    txtTotal.Text = ""
    dateTrans.Value = Format(Now)
    cmbNama_Pel = ""
    txtID_Pel.Text = ""
    cmbPeriode = ""
    txtGrand.Text = ""
    txtBayar.Text = ""
    txtKembali.Text = ""
    btnCancel.Visible = False
End Sub

Private Sub btnDelete_Click()
    If Sales_ListView.ListItems.Count = 0 Then
        frmEmpty.Show
    Else
        Sales_ListView.ListItems.Remove (Sales_ListView.SelectedItem.Index)
        txtGrand.Text = ""
    End If
End Sub

Private Sub btnHitung_Click()
    If Sales_ListView.ListItems.Count > 0 Then
        Dim GrandTotal_list As Double
        Dim i As Integer
        
        For i = 1 To Sales_ListView.ListItems.Count
            GrandTotal_list = GrandTotal_list + Int(Sales_ListView.ListItems(i).ListSubItems(5).Text)
        Next i
        txtGrand.Text = GrandTotal_list
    End If
End Sub

Private Sub btnPrint_Click()
    Export_Excel "Kunci_Ibadah_Utama.xlsx", "Data Penjualan"
End Sub

Private Sub btnSave_Click()
    Call OpenDB
    Dim i As Integer
    'Dim harga_asli As Integer
    
    If Sales_ListView.ListItems.Count = 0 Then
        frmNone.Show
    ElseIf txtID_Pel.Text = "" Or cmbNama_Pel.Text = "" Or cmbPeriode.Text = "" Or txtGrand.Text = "" Or txtBayar.Text = "" Or txtKembali.Text = "" Then
        frmEmpty.Show
    Else
        For i = 1 To Sales_ListView.ListItems.Count
            With rsSales
                .Open "select * from SalesData_Table", ConnectDB, 1, 2
                .AddNew
                !id_produk = Sales_ListView.ListItems(i).Text
                !nama_produk = Sales_ListView.ListItems(i).ListSubItems(1).Text
                !harga_produk = Sales_ListView.ListItems(i).ListSubItems(2).Text
                !satuan = Sales_ListView.ListItems(i).ListSubItems(3).Text
                !kuantitas = Sales_ListView.ListItems(i).ListSubItems(4).Text
                !Total = Sales_ListView.ListItems(i).ListSubItems(5).Text
                !id_trans = txtTrans.Text
                .Update
                .Close
            End With
            With rsSales
                .Open "select * from Journal_Data", ConnectDB
                .AddNew
                !id_trans = txtTrans.Text
                !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
                !periode = cmbPeriode.Text
                !uraian = "Saldo HPP"
                !nama_akun1 = "HPP"
                !id_akun1 = "5.1.1.1"
                !debit = Val(txtHPP.Text) * Val(Sales_ListView.ListItems(i).ListSubItems(4).Text)
                !nama_akun2 = "PERSEDIAAN BARANG JADI"
                !id_akun2 = "1.1.2.4"
                !kredit = Val(txtHPP.Text) * Val(Sales_ListView.ListItems(i).ListSubItems(4).Text)
                .Update
                .Close
            End With
        Next i
        With rsSales
            .Open "select * from Sales_Table", ConnectDB
            .AddNew
            !id_trans = txtTrans.Text
            !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
            !id_customer = txtID_Pel.Text
            !nama_customer = cmbNama_Pel.Text
            !periode = cmbPeriode.Text
            !grandtotal = txtGrand.Text
            !bayar = txtBayar.Text
            !kembali = txtKembali.Text
            .Update
            .Close
        End With
        With rsSales
            .Open "select * from Journal_Data", ConnectDB
            .AddNew
            !id_trans = txtTrans.Text
            !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
            !periode = cmbPeriode.Text
            !uraian = "Penjualan Produk"
            !nama_akun1 = "KAS"
            !id_akun1 = "1.1.1.1"
            !debit = txtGrand.Text
            !nama_akun2 = "PENJUALAN"
            !id_akun2 = "4.1.1.1"
            !kredit = txtGrand.Text
            .Update
            .Close
        End With
        
        'Script Perselisihan & Penulisan Jurnal Penjualan
        i = 1
        Do While i <= Sales_ListView.ListItems.Count
            jual = Val(txtGrand.Text)
            kuan = Sales_ListView.ListItems(i).ListSubItems(4).Text
            kode = Sales_ListView.ListItems(i).Text
            hpp_perproduk = Val(txtHPP.Text) * Val(kuan)
            
            'KUANTITAS
            Buka_Produk kode
            KB = rsProduct!jumlah_unit
            
            updateKB = Val(KB) - Val(kuan)
            Ubah_KuanProduk kode, updateKB
            i = i + 1
            
            'KAS
            Buka_Jurnal "1.1.1.1"
            SA = rsJournal!saldo_akhir
            
            updateSA = Val(SA) + Val(jual)
            Ubah_Jurnal "1.1.1.1", updateSA
            i = i + 1
            
            'PENJUALAN
            Buka_Jurnal "4.1.1.1"
            SA = rsJournal!saldo_akhir
            
            updateSA = Val(SA) + Val(jual)
            Ubah_Jurnal "4.1.1.1", updateSA
            i = i + 1
            
            'HPP
            Buka_Jurnal "5.1.1.1"
            laba = rsJournal!saldo_akhir
                
            updateLaba = Val(laba) + Val(hpp_perproduk)
            
            Ubah_Jurnal "5.1.1.1", updateLaba
            i = i + 1
            
            'PERSEDIAAN BARANG JADI
            Buka_Jurnal "1.1.2.4"
            PBJ = rsJournal!saldo_akhir
            
            updatePBJ = Val(PBJ) - Val(hpp_perproduk)
            
            Ubah_Jurnal "1.1.2.4", updatePBJ
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
    txtID = Tampil_IDProduk("Product_Table", "id_produk", cmbNama)
    txtHarga = Tampil_IDProduk("Product_Table", "harga_jual", cmbNama)
    txtSatuan = Tampil_IDProduk("Product_Table", "satuan", cmbNama)
    txtHPP = Tampil_HPPProduk("HPP_Production", "hpp_produk", cmbNama)
    txtID.Enabled = False
    txtSatuan.Enabled = False
    txtTotal.Enabled = False
    btnCancel.Visible = True
End Sub

Private Sub cmbNama_Pel_Click()
    txtID_Pel = Tampil_IDPelanggan("Customer_Table", "id_customer", cmbNama_Pel)
    txtID_Pel.Enabled = False
End Sub

Private Sub Form_Load()
    Call OpenDB
    AN_TRANS
    txtTrans.Enabled = False
    txtGrand.Enabled = False
    txtKembali.Enabled = False
    btnCancel.Visible = False
    dateTrans.Value = Format(Now)
    rsPro.Open "select * from Product_Table", ConnectDB
    rsCus.Open "select * from Customer_Table", ConnectDB
    rsPer.Open "select * from Periode_Table", ConnectDB
    cmbNama.Clear
    cmbNama_Pel.Clear
    cmbPeriode.Clear
    Do While Not rsPro.EOF
        cmbNama.AddItem rsPro!nama_produk
        rsPro.MoveNext
    Loop
    Do While Not rsCus.EOF
        cmbNama_Pel.AddItem rsCus!nama_customer
        rsCus.MoveNext
    Loop
    Do While Not rsPer.EOF
        cmbPeriode.AddItem rsPer!masa
        rsPer.MoveNext
    Loop
End Sub

Private Sub txtBayar_Change()
    txtKembali.Text = Val(txtBayar.Text) - Val(txtGrand.Text)
End Sub

Private Sub txtKuan_Change()
    txtTotal.Text = Val(txtHarga.Text) * Val(txtKuan.Text)
End Sub
