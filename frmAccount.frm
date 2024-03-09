VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAccount 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15015
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   15015
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
      Height          =   2655
      Left            =   10200
      TabIndex        =   16
      Top             =   5640
      Width           =   4575
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
         TabIndex        =   18
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox txtID_Key 
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
         TabIndex        =   17
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Image btnSearch 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1080
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmAccount.frx":0000
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2175
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   9600
      Top             =   4920
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DAFTAR AKUN SALDO"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   1
      Top             =   5640
      Width           =   9855
      Begin VB.TextBox txtSaldo 
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
         Left            =   6960
         TabIndex        =   15
         Top             =   2040
         Width           =   2775
      End
      Begin VB.ComboBox cmbKelompok 
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
         Left            =   6960
         TabIndex        =   13
         Top             =   1440
         Width           =   2775
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
         Left            =   6960
         TabIndex        =   11
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox cmbDK 
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
         TabIndex        =   9
         Top             =   1440
         Width           =   2775
      End
      Begin VB.ComboBox cmbKategori 
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
         TabIndex        =   8
         Top             =   960
         Width           =   2775
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
         Left            =   6960
         TabIndex        =   7
         Top             =   480
         Width           =   2775
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
         Left            =   2160
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin VB.Image btnCancel 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   3240
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmAccount.frx":5442
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Image btnUpdate 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1560
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmAccount.frx":BFBC
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
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
         Left            =   6120
         TabIndex        =   14
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Kelompok"
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
         Left            =   5520
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
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
         Left            =   5760
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Debit/Kredit"
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
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori"
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
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Akun"
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
         Left            =   5280
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Akun"
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
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid Account_DataGrid 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   6588
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA AKUN"
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
      Left            =   11160
      TabIndex        =   19
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image btnPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   12960
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmAccount.frx":1104C
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Image btnRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   11040
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmAccount.frx":16519
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Image btnDelete 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   3120
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmAccount.frx":1C1E4
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   855
   End
   Begin VB.Image btnEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   2160
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmAccount.frx":21389
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   855
   End
   Begin VB.Image btnSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   1200
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmAccount.frx":2634B
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   855
   End
   Begin VB.Image btnAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmAccount.frx":2BDA6
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   855
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmAccount.frx":30E50
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   8295
      Left            =   120
      Picture         =   "frmAccount.frx":377C7
      Stretch         =   -1  'True
      Top             =   120
      Width           =   14775
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub TutupText()
    txtID.Enabled = False
    txtNama.Enabled = False
    cmbKategori.Enabled = False
    cmbDK.Enabled = False
    cmbPeriode.Enabled = False
    cmbKelompok.Enabled = False
    txtSaldo.Enabled = False
End Sub

Sub BukaText()
    txtID.Enabled = True
    txtNama.Enabled = True
    cmbKategori.Enabled = True
    cmbDK.Enabled = True
    cmbPeriode.Enabled = True
    cmbKelompok.Enabled = True
    txtSaldo.Enabled = True
End Sub

Sub ClearText()
    txtID.Text = ""
    txtNama.Text = ""
    cmbKategori.Text = ""
    cmbDK.Text = ""
    cmbPeriode.Text = ""
    cmbKelompok.Text = ""
    txtSaldo.Text = ""
End Sub

Sub TAMPILKAN_KE_TEXTBOX()
    txtID.Text = Account_DataGrid.Columns(0).Value
    txtNama.Text = Account_DataGrid.Columns(1).Value
    cmbKategori.Text = Account_DataGrid.Columns(2).Value
    cmbDK.Text = Account_DataGrid.Columns(3).Value
    cmbPeriode.Text = Account_DataGrid.Columns(4).Value
    cmbKelompok.Text = Account_DataGrid.Columns(5).Value
    txtSaldo.Text = Account_DataGrid.Columns(6).Value
End Sub

Sub AN()
    Dim AN As Integer
    Dim sr As String
    Randomize
    AN1 = Int((8 * Rnd) + 1)
    AN2 = Int((7 * Rnd) + 2)
    AN3 = Int((6 * Rnd) + 3)
    AN4 = Int((5 * Rnd) + 4)
    txtID.Text = AN1 & "." & AN2 & "." & AN3 & "." & AN4
End Sub

Function Tampil_IDAkun(Tabel, Field, KD)
    Dim rs As New ADODB.Recordset
    Q = "select " & Field & " from " & Tabel & " where nama_akun='" & KD & "'"
    rs.Open Q, ConnectDB
    Tampil_IDAkun = rs.Fields(Field)
    rs.Close
End Function

Private Sub Account_DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    TAMPILKAN_KE_TEXTBOX
    btnCancel.Visible = True
End Sub

Private Sub btnAdd_Click()
    BukaText
    AN
    btnAdd.Enabled = False
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
    btnAdd.Enabled = True
    btnCancel.Visible = False
    btnUpdate.Visible = False
    btnSave.Enabled = False
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
    Export_Excel "Kunci_Ibadah_Utama.xlsx", "Data Akun"
End Sub

Private Sub btnRefresh_Click()
    cmbKriteria.Text = ""
    txtID_Key.Text = ""
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Account_Table"
    Adodc1.Refresh
End Sub

Private Sub btnSave_Click()
    Call OpenDB
    If txtID.Text = "" Or txtNama.Text = "" Or cmbKategori.Text = "" Or cmbDK.Text = "" Or cmbPeriode.Text = "" Or cmbKelompok.Text = "" Or txtSaldo.Text = "" Then
        frmEmpty.Show
        btnAdd.Enabled = False
        btnEdit.Enabled = False
        btnDelete.Enabled = False
    Else
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!id_akun = txtID.Text
        Adodc1.Recordset!nama_akun = txtNama.Text
        Adodc1.Recordset!kategori = cmbKategori.Text
        Adodc1.Recordset!Posisi = cmbDK.Text
        Adodc1.Recordset!periode = cmbPeriode.Text
        Adodc1.Recordset!kelompok = cmbKelompok.Text
        Adodc1.Recordset!saldo_awal = txtSaldo.Text
        Adodc1.Recordset!saldo_akhir = txtSaldo.Text
        
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
    Adodc1.RecordSource = "select * from Account_Table where nama_akun like'" & cmbKriteria.Text & "'"
    Adodc1.Refresh
    Set Account_DataGrid.DataSource = Adodc1
End Sub

Private Sub btnUpdate_Click()
    Call OpenDB
    If txtID.Text = "" Or txtNama.Text = "" Or cmbKategori.Text = "" Or cmbDK.Text = "" Or cmbPeriode.Text = "" Or cmbKelompok.Text = "" Or txtSaldo.Text = "" Then
        frmEmpty2.Show
        btnAdd.Enabled = False
        btnEdit.Enabled = False
        btnDelete.Enabled = False
    Else
        Adodc1.Recordset.Update
        Adodc1.Recordset!id_akun = txtID.Text
        Adodc1.Recordset!nama_akun = txtNama.Text
        Adodc1.Recordset!kategori = cmbKategori.Text
        Adodc1.Recordset!Posisi = cmbDK.Text
        Adodc1.Recordset!periode = cmbPeriode.Text
        Adodc1.Recordset!kelompok = cmbKelompok.Text
        Adodc1.Recordset!saldo_awal = txtSaldo.Text
        Adodc1.Recordset!saldo_akhir = txtSaldo.Text
        
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

Private Sub cmbKriteria_Click()
    txtID_Key.Text = Tampil_IDAkun("Account_Table", "id_akun", cmbKriteria)
End Sub

Private Sub Form_Load()
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Account_Table"
    Adodc1.Refresh
    Set Account_DataGrid.DataSource = Adodc1
    
    cmbKelompok.AddItem "ASET"
    cmbKelompok.AddItem "KEWAJIBAN"
    cmbKelompok.AddItem "MODAL"
    cmbKelompok.AddItem "PENDAPATAN"
    cmbKelompok.AddItem "HPP"
    cmbKelompok.AddItem "BEBAN"
    cmbKategori.AddItem "NERACA"
    cmbKategori.AddItem "LABA/RUGI"
    cmbDK.AddItem "DEBIT"
    cmbDK.AddItem "KREDIT"
    
    TutupText
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    btnSave.Enabled = False
    btnCancel.Visible = False
    btnUpdate.Visible = False
    rsAcc.Open "select * from Account_Table", ConnectDB
    rsPer.Open "select * from Periode_Table", ConnectDB
    cmbKriteria.Clear
    cmbPeriode.Clear
    Do While Not rsAcc.EOF
        cmbKriteria.AddItem rsAcc!nama_akun
        rsAcc.MoveNext
    Loop
    Do While Not rsPer.EOF
        cmbPeriode.AddItem rsPer!masa
        rsPer.MoveNext
    Loop
End Sub
