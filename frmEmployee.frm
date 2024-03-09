VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEmployee 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   7200
      Top             =   360
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
   Begin MSDataGridLib.DataGrid Employee_DataGrid 
      Height          =   2175
      Left            =   5280
      TabIndex        =   12
      Top             =   1080
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3836
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
      Height          =   1815
      Left            =   7200
      TabIndex        =   9
      Top             =   3360
      Width           =   5415
      Begin VB.ComboBox cmbKriteria 
         BeginProperty Font 
            Name            =   "Montserrat ExtraBold"
            Size            =   12
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   3135
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
         TabIndex        =   10
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Image btnSearch 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3360
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmEmployee.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FORM DATA KARYAWAN"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
      Begin VB.ComboBox cmbJabatan 
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
         Left            =   2400
         TabIndex        =   14
         Top             =   1440
         Width           =   2415
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
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   960
         Width           =   2415
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
         Height          =   420
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
      Begin RichTextLib.RichTextBox rtbAlamat 
         Height          =   1335
         Left            =   2400
         TabIndex        =   13
         Top             =   1920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   2355
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmEmployee.frx":5442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Montserrat SemiBold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image btnUpdate 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   1080
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmEmployee.frx":54CA
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Image btnCancel 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   3000
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmEmployee.frx":A55A
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Montserrat ExtraBold"
            Size            =   12
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         BeginProperty Font 
            Name            =   "Montserrat ExtraBold"
            Size            =   12
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label txtxNama 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Karyawan"
         BeginProperty Font 
            Name            =   "Montserrat ExtraBold"
            Size            =   12
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Karyawan"
         BeginProperty Font 
            Name            =   "Montserrat ExtraBold"
            Size            =   12
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Image btnAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   5280
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmEmployee.frx":110D4
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   855
   End
   Begin VB.Image btnSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   6240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmEmployee.frx":1617E
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   855
   End
   Begin VB.Image btnEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   5280
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmEmployee.frx":1BBD9
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   855
   End
   Begin VB.Image btnDelete 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   6240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmEmployee.frx":20B9B
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   855
   End
   Begin VB.Image btnRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   8880
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmEmployee.frx":25D40
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1815
   End
   Begin VB.Image btnPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   10800
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmEmployee.frx":2BA0B
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1815
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmEmployee.frx":30ED8
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PEGAWAI"
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
      Height          =   735
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA"
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
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   5175
      Left            =   120
      Picture         =   "frmEmployee.frx":3784F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   12615
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub TutupText()
    txtID.Enabled = False
    txtNama.Enabled = False
    cmbJabatan.Enabled = False
    rtbAlamat.Enabled = False
End Sub

Sub BukaText()
    txtID.Enabled = True
    txtNama.Enabled = True
    cmbJabatan.Enabled = True
    rtbAlamat.Enabled = True
End Sub

Sub ClearText()
    txtID.Text = ""
    txtNama.Text = ""
    cmbJabatan.Text = ""
    rtbAlamat.Text = ""
End Sub

Sub TAMPILKAN_KE_TEXTBOX()
    txtID.Text = Employee_DataGrid.Columns(0).Value
    txtNama.Text = Employee_DataGrid.Columns(1).Value
    cmbJabatan.Text = Employee_DataGrid.Columns(2).Value
    rtbAlamat.Text = Employee_DataGrid.Columns(3).Value
End Sub

Sub AN()
    Call OpenDB
    rsPeg.Open ("SELECT * FROM Employee_Table WHERE id_karyawan in(select max(id_karyawan) from Employee_Table)order by id_karyawan desc"), ConnectDB
    rsPeg.Requery
    Dim Urut As String * 5
    Dim Hitung As Long
    With rsPeg
        If .EOF Then
            Urut = "K0" + "001"
            txtID = Urut
        Else
            Hitung = Right(!id_karyawan, 3) + 1
            Urut = "K0" + Right("000" & Hitung, 3)
        End If
        txtID = Urut
    End With
End Sub

Private Sub btnAdd_Click()
    BukaText
    ClearText
    AN
    txtID.Enabled = False
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    btnSave.Enabled = True
    btnCancel.Visible = True
End Sub

Private Sub btnBack_Click()
    Me.Hide
    frmHome.Show
End Sub

Private Sub btnCancel_Click()
    TutupText
    ClearText
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    btnSave.Enabled = False
    btnUpdate.Visible = False
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

Private Sub btnEdit_Click()
    If txtID.Text = "" Then
        frmSelected.Show
    Else
        BukaText
        txtID.Enabled = False
        btnAdd.Enabled = False
        btnEdit.Enabled = False
        btnDelete.Enabled = False
        btnSave.Enabled = False
        btnUpdate.Visible = True
        btnCancel.Visible = True
    End If
End Sub

Private Sub btnPrint_Click()
    Export_Excel "Kunci_Ibadah_Utama.xlsx", "Data Pegawai"
End Sub

Private Sub btnRefresh_Click()
    cmbKriteria.Text = ""
    txtKeyword.Text = ""
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Employee_Table"
    Adodc1.Refresh
End Sub

Private Sub btnSave_Click()
    If txtID.Text = "" Or txtNama.Text = "" Or cmbJabatan.Text = "" Or rtbAlamat.Text = "" Then
        frmEmpty.Show
        btnAdd.Enabled = False
        btnEdit.Enabled = False
        btnDelete.Enabled = False
        btnSave.Enabled = True
    Else
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!id_karyawan = txtID.Text
        Adodc1.Recordset!nama_karyawan = txtNama.Text
        Adodc1.Recordset!jabatan = cmbJabatan.Text
        Adodc1.Recordset!alamat = rtbAlamat.Text
        
        frmSave_Alert.Show
        Adodc1.Recordset.MoveFirst
        ClearText
        TutupText
        btnAdd.Enabled = True
        btnEdit.Enabled = True
        btnDelete.Enabled = True
        btnSave.Enabled = False
        btnCancel.Visible = False
    End If
End Sub

Private Sub btnSearch_Click()
    Call OpenDB
    If cmbKriteria.Text = "ID" Then
        Adodc1.RecordSource = "select * from Employee_Table where id_karyawan like'" & txtKeyword.Text & "'"
        Adodc1.Refresh
        Set Employee_DataGrid.DataSource = Adodc1
    ElseIf cmbKriteria.Text = "NAMA" Then
        Adodc1.RecordSource = "select * from Employee_Table where nama_karyawan like'" & txtKeyword.Text & "'"
        Adodc1.Refresh
        Set Employee_DataGrid.DataSource = Adodc1
    Else
        frmNo_Record.Show
    End If
End Sub

Private Sub btnUpdate_Click()
    If txtID.Text = "" Or txtNama.Text = "" Or cmbJabatan.Text = "" Or rtbAlamat.Text = "" Then
        frmEmpty.Show
        btnAdd.Enabled = False
        btnEdit.Enabled = False
        btnDelete.Enabled = False
        btnSave.Enabled = False
    Else
        Adodc1.Recordset.Update
        Adodc1.Recordset!id_karyawan = txtID.Text
        Adodc1.Recordset!nama_karyawan = txtNama.Text
        Adodc1.Recordset!jabatan = cmbJabatan.Text
        Adodc1.Recordset!alamat = rtbAlamat.Text
        
        frmUpdate_Alert.Show
        Adodc1.Recordset.MoveFirst
        ClearText
        TutupText
        btnAdd.Enabled = True
        btnEdit.Enabled = True
        btnDelete.Enabled = True
        btnSave.Enabled = False
        btnUpdate.Visible = False
        btnCancel.Visible = False
    End If
End Sub

Private Sub Employee_DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    TAMPILKAN_KE_TEXTBOX
    btnCancel.Visible = True
End Sub

Private Sub Form_Load()
    TutupText
    ClearText
    cmbKriteria.AddItem "ID"
    cmbKriteria.AddItem "NAMA"
    cmbJabatan.AddItem "Pemilik"
    cmbJabatan.AddItem "Produksi"
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Employee_Table"
    Adodc1.Refresh
    Set Employee_DataGrid.DataSource = Adodc1
    
    btnUpdate.Visible = False
    btnCancel.Visible = False
    btnSave.Enabled = False
End Sub
