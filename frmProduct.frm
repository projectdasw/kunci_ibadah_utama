VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmProduct 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H000040C0&
      Caption         =   "PENCARIAN DATA"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   5160
      TabIndex        =   11
      Top             =   1080
      Width           =   2415
      Begin VB.ComboBox cmbKriteria 
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
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtKeyword 
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
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Image btnSearch 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   240
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmProduct.frx":0000
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   2280
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid Product_DataGrid 
      Height          =   1695
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FORM PRODUK"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4815
      Begin VB.ComboBox cmbSatuan 
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
         Left            =   2160
         TabIndex        =   16
         Top             =   1920
         Width           =   2535
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
         Height          =   390
         Left            =   2160
         TabIndex        =   14
         Top             =   1440
         Width           =   2535
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
         Height          =   390
         Left            =   2160
         TabIndex        =   9
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtHPP 
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   5
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox txtNama 
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label2 
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
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label7 
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
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label6 
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
         TabIndex        =   7
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label5 
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
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Image btnPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   4320
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduct.frx":5442
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image btnRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   6000
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduct.frx":A90F
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image btnCancel 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   5760
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduct.frx":105DA
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Image btnAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduct.frx":17154
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   855
   End
   Begin VB.Image btnSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   1200
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduct.frx":1C1FE
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   855
   End
   Begin VB.Image btnDelete 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   2160
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduct.frx":21C59
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "( Barang Jadi )"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   20.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA PRODUK"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduct.frx":26DFE
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   7455
      Left            =   120
      Picture         =   "frmProduct.frx":2D775
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub TutupText()
    txtID.Enabled = False
    txtNama.Enabled = False
    txtHPP.Enabled = False
    txtKuan.Enabled = False
    cmbSatuan.Enabled = False
End Sub

Sub BukaText()
    txtID.Enabled = True
    txtNama.Enabled = True
    txtHPP.Enabled = True
    txtKuan.Enabled = True
    cmbSatuan.Enabled = True
End Sub

Sub ClearText()
    txtID.Text = ""
    txtNama.Text = ""
    txtHPP.Text = ""
    txtKuan.Text = ""
    cmbSatuan.Text = ""
End Sub

Sub TAMPILKAN_KE_TEXTBOX()
    txtID.Text = Product_DataGrid.Columns(0).Value
    txtNama.Text = Product_DataGrid.Columns(1).Value
    txtKuan.Text = Product_DataGrid.Columns(2).Value
    cmbSatuan.Text = Product_DataGrid.Columns(3).Value
    txtHPP.Text = Product_DataGrid.Columns(4).Value
End Sub

Sub AN()
    Call OpenDB
    rsProduct.Open ("SELECT * FROM Product_Table WHERE id_produk in(select max(id_produk) from Product_Table)order by id_produk desc"), ConnectDB
    rsProduct.Requery
    Dim Urut As String * 7
    Dim Hitung As Long
    With rsProduct
        If .EOF Then
            Urut = "IP-" + "0001"
            txtID = Urut
        Else
            Hitung = Right(!id_produk, 4) + 1
            Urut = "IP-" + Right("0000" & Hitung, 4)
        End If
        txtID = Urut
    End With
End Sub

Private Sub btnAdd_Click()
    BukaText
    ClearText
    AN
    txtID.Enabled = False
    btnCancel.Visible = True
    btnSave.Enabled = True
End Sub

Private Sub btnBack_Click()
    Me.Hide
    frmHome.Show
End Sub

Private Sub btnCancel_Click()
    TutupText
    ClearText
    btnCancel.Visible = False
    btnDelete.Enabled = False
    btnSave.Enabled = False
End Sub

Private Sub btnDelete_Click()
    If txtID.Text = "" Then
        frmSelected.Show
    Else
        frmDelete_alert.Show
        btnAdd.Enabled = True
        btnSave.Enabled = False
        btnDelete.Enabled = False
        btnCancel.Visible = False
    End If
End Sub

Private Sub btnPrint_Click()
    Export_Excel "Kunci_Ibadah_Utama.xlsx", "Data Produk"
End Sub

Private Sub btnRefresh_Click()
    cmbKriteria.Text = ""
    txtKeyword.Text = ""
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Product_Table"
    Adodc1.Refresh
End Sub

Private Sub btnSave_Click()
    If txtID.Text = "" Or txtNama.Text = "" Or txtHPP.Text = "" Or cmbSatuan.Text = "" Or txtKuan.Text = "" Then
        frmEmpty.Show
    Else
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!id_produk = txtID.Text
        Adodc1.Recordset!nama_produk = txtNama.Text
        Adodc1.Recordset!jumlah_unit = txtKuan.Text
        Adodc1.Recordset!satuan = cmbSatuan.Text
        Adodc1.Recordset!harga_jual = txtHPP.Text
        
        frmSave_Alert.Show
        Adodc1.Recordset.MoveFirst
        ClearText
        TutupText
        btnSave.Enabled = False
        btnDelete.Enabled = False
        btnCancel.Visible = False
    End If
End Sub

Private Sub btnSearch_Click()
    Call OpenDB
    If cmbKriteria.Text = "ID" Then
        Adodc1.RecordSource = "select * from Product_Table where id_produk like'" & txtKeyword.Text & "'"
        Adodc1.Refresh
        Set Product_DataGrid.DataSource = Adodc1
    ElseIf cmbKriteria.Text = "NAMA" Then
        Adodc1.RecordSource = "select * from Product_Table where nama_produk like'" & txtKeyword.Text & "'"
        Adodc1.Refresh
        Set Product_DataGrid.DataSource = Adodc1
    Else
        frmNo_Record.Show
    End If
End Sub

Private Sub Form_Load()
    TutupText
    ClearText
    cmbKriteria.AddItem "ID"
    cmbKriteria.AddItem "NAMA"
    cmbSatuan.AddItem "Pcs"
    btnDelete.Enabled = False
    btnSave.Enabled = False
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Product_Table"
    Adodc1.Refresh
    Set Product_DataGrid.DataSource = Adodc1
    
    btnCancel.Visible = False
End Sub

Private Sub Product_DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    TAMPILKAN_KE_TEXTBOX
    btnCancel.Visible = True
    btnDelete.Enabled = True
End Sub
