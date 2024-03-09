VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmItems 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H000040C0&
      Caption         =   "PENCARIAN DATA"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   4920
      TabIndex        =   10
      Top             =   1080
      Width           =   2655
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
         TabIndex        =   12
         Top             =   480
         Width           =   2415
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
         Height          =   405
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Image btnSearch 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   360
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmItems.frx":0000
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3720
      Top             =   5280
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin MSDataGridLib.DataGrid Items_DataGrid 
      Height          =   1695
      Left            =   240
      TabIndex        =   9
      Top             =   6000
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
      Caption         =   "FORM BAHAN BAKU"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   4575
      Begin VB.TextBox txtKuan 
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   14
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox cmbSatuan 
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
         Left            =   1920
         TabIndex        =   13
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtHarga 
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   8
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtNama 
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   7
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "Montserrat SemiBold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Left            =   480
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Image btnCancel 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   2760
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmItems.frx":5442
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1695
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
         Left            =   840
         TabIndex        =   5
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   3
         Top             =   960
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
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Image btnRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   5640
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmItems.frx":BFBC
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Image btnPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   5880
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmItems.frx":11C87
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Image btnDelete 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   2400
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmItems.frx":17154
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   975
   End
   Begin VB.Image btnSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   1320
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmItems.frx":1C2F9
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   975
   End
   Begin VB.Image btnAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmItems.frx":21D54
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA BAHAN BAKU"
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
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmItems.frx":26DFE
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   7695
      Left            =   120
      Picture         =   "frmItems.frx":2D775
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub TutupText()
    txtID.Enabled = False
    txtNama.Enabled = False
    txtHarga.Enabled = False
    cmbSatuan.Enabled = False
    txtKuan.Enabled = False
End Sub

Sub BukaText()
    txtID.Enabled = True
    txtNama.Enabled = True
    txtHarga.Enabled = True
    cmbSatuan.Enabled = True
    txtKuan.Enabled = True
End Sub

Sub ClearText()
    txtID.Text = ""
    txtNama.Text = ""
    txtHarga.Text = ""
    cmbSatuan.Text = ""
    txtKuan.Text = ""
End Sub

Sub TAMPILKAN_KE_TEXTBOX()
    txtID.Text = Items_DataGrid.Columns(0).Value
    txtNama.Text = Items_DataGrid.Columns(1).Value
    txtKuan.Text = Items_DataGrid.Columns(2).Value
    cmbSatuan.Text = Items_DataGrid.Columns(3).Value
    txtHarga.Text = Items_DataGrid.Columns(4).Value
End Sub

Sub AN_ITEMS()
    Call OpenDB
    rsBB.Open ("SELECT * FROM Items_Table WHERE id_bahan in(select max(id_bahan) from Items_Table)order by id_bahan desc"), ConnectDB
    rsBB.Requery
    Dim Urut As String * 7
    Dim Hitung As Long
    With rsBB
        If .EOF Then
            Urut = "IB-" + "0001"
            txtID = Urut
        Else
            Hitung = Right(!id_bahan, 4) + 1
            Urut = "IB-" + Right("0000" & Hitung, 4)
        End If
        txtID = Urut
    End With
End Sub

Sub AN_PEMBANTU()
    Call OpenDB
    'rsBB.Open ("SELECT * FROM Items_Table WHERE id_bahan in(select max(id_bahan) from Items_Table)order by id_bahan desc"), ConnectDB
    rsBB.Open ("SELECT id_bahan FROM Items_Table order by id_bahan asc"), ConnectDB
    rsBB.Requery
    Dim Urut As String * 7
    Dim Hitung As Long
    With rsBB
        If .EOF Then
            Urut = "BP-" + "0001"
            txtID = Urut
        Else
            Hitung = Right(!id_bahan, 4) + 1
            Urut = "BP-" + Right("0000" & Hitung, 4)
        End If
        txtID = Urut
    End With
End Sub

Private Sub btnAdd_Click()
    ClearText
    BukaText
    frmAddItems.Show
End Sub

Private Sub btnBack_Click()
    frmHome.Show
    frmItems.Hide
End Sub

Private Sub btnCancel_Click()
    TutupText
    ClearText
    btnAdd.Enabled = True
    btnDelete.Enabled = True
    btnSave.Enabled = False
    btnCancel.Visible = False
End Sub

Private Sub btnDelete_Click()
    If txtID.Text = "" Then
        frmSelected.Show
    Else
        frmDelete_alert.Show
        btnCancel.Visible = False
    End If
End Sub

Private Sub btnPrint_Click()
    Export_Excel "Kunci_Ibadah_Utama.xlsx", "Data Barang"
End Sub

Private Sub btnRefresh_Click()
    cmbKriteria.Text = ""
    txtKeyword.Text = ""
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Items_Table"
    Adodc1.Refresh
End Sub

Private Sub btnSave_Click()
    If txtID.Text = "" Or txtNama.Text = "" Or txtKuan.Text = "" Or cmbSatuan.Text = "" Or txtHarga.Text = "" Then
        frmEmpty.Show
        btnAdd.Enabled = False
        btnDelete.Enabled = False
        btnSave.Enabled = True
    Else
        STB = Val(txtHarga.Text) * Val(txtKuan.Text)
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!id_bahan = txtID.Text
        Adodc1.Recordset!nama_bahan = txtNama.Text
        Adodc1.Recordset!kuantitas = txtKuan.Text
        Adodc1.Recordset!satuan = cmbSatuan.Text
        Adodc1.Recordset!harga_bahan = txtHarga.Text
        Adodc1.Recordset!total_bahan = STB
        
        frmSave_Alert.Show
        Adodc1.Recordset.MoveFirst
        
        ClearText
        TutupText
        btnCancel.Visible = False
        btnAdd.Enabled = True
        btnDelete.Enabled = True
        btnSave.Enabled = False
    End If
End Sub

Private Sub btnSearch_Click()
    Call OpenDB
    If cmbKriteria.Text = "ID" Then
        Adodc1.RecordSource = "select * from Items_Table where id_bahan like'" & txtKeyword.Text & "'"
        Adodc1.Refresh
        Set Items_DataGrid.DataSource = Adodc1
    ElseIf cmbKriteria.Text = "NAMA" Then
        Adodc1.RecordSource = "select * from Items_Table where nama_bahan like'" & txtKeyword.Text & "'"
        Adodc1.Refresh
        Set Items_DataGrid.DataSource = Adodc1
    Else
        frmNo_Record.Show
    End If
End Sub

Private Sub Form_Load()
    TutupText
    ClearText
    cmbKriteria.AddItem "ID"
    cmbKriteria.AddItem "NAMA"
    cmbSatuan.AddItem "Yard"
    cmbSatuan.AddItem "Pcs"
    cmbSatuan.AddItem "Pack"
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Items_Table"
    Adodc1.Refresh
    Set Items_DataGrid.DataSource = Adodc1
    
    btnCancel.Visible = False
    btnSave.Enabled = False
End Sub

Private Sub Items_DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    TAMPILKAN_KE_TEXTBOX
    btnCancel.Visible = True
End Sub
