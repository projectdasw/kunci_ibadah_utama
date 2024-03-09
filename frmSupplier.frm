VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSupplier 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   -60
   ClientTop       =   -60
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   12705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   7320
      Top             =   240
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
      Height          =   2175
      Left            =   7320
      TabIndex        =   11
      Top             =   3240
      Width           =   5175
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
         Height          =   435
         Left            =   2640
         TabIndex        =   13
         Top             =   480
         Width           =   2415
      End
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
      Begin VB.Image btnSearch 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   1560
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmSupplier.frx":0000
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   2055
      End
   End
   Begin MSDataGridLib.DataGrid Supplier_DataGrid 
      Height          =   2175
      Left            =   5160
      TabIndex        =   10
      Top             =   960
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FORM SUPPLIER"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   4815
      Begin VB.TextBox txtTelp 
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
         Left            =   2160
         TabIndex        =   7
         Top             =   3360
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
         Height          =   375
         Left            =   2160
         TabIndex        =   6
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
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin RichTextLib.RichTextBox rtbAlamat 
         Height          =   1815
         Left            =   2160
         TabIndex        =   14
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3201
         _Version        =   393217
         TextRTF         =   $"frmSupplier.frx":5442
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
      Begin VB.Image btnCancel 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   2880
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmSupplier.frx":54CA
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Image btnUpdate 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   960
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmSupplier.frx":C044
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telp"
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
         TabIndex        =   9
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Supplier"
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
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Supplier"
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
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Image btnAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   5160
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSupplier.frx":110D4
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   975
   End
   Begin VB.Image btnSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   6240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSupplier.frx":1617E
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   975
   End
   Begin VB.Image btnEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   5160
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSupplier.frx":1BBD9
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   975
   End
   Begin VB.Image btnDelete 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   6240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSupplier.frx":20B9B
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   975
   End
   Begin VB.Image btnPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   10800
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSupplier.frx":25D40
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image btnRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   9000
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmSupplier.frx":2B20D
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      Picture         =   "frmSupplier.frx":30ED8
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   120
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
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   5415
      Left            =   120
      Picture         =   "frmSupplier.frx":3784F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   12495
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub TutupText()
    txtID.Enabled = False
    txtNama.Enabled = False
    rtbAlamat.Enabled = False
    txtTelp.Enabled = False
End Sub

Sub BukaText()
    txtID.Enabled = True
    txtNama.Enabled = True
    rtbAlamat.Enabled = True
    txtTelp.Enabled = True
End Sub

Sub ClearText()
    txtID.Text = ""
    txtNama.Text = ""
    rtbAlamat.Text = ""
    txtTelp.Text = ""
End Sub

Sub TAMPILKAN_KE_TEXTBOX()
    txtID.Text = Supplier_DataGrid.Columns(0).Value
    txtNama.Text = Supplier_DataGrid.Columns(1).Value
    rtbAlamat.Text = Supplier_DataGrid.Columns(2).Value
    txtTelp.Text = Supplier_DataGrid.Columns(3).Value
End Sub

Sub AN()
    Call OpenDB
    rsSupp.Open ("SELECT * FROM Supplier_Table WHERE id_supplier in(select max(id_supplier) from Supplier_Table)order by id_supplier desc"), ConnectDB
    rsSupp.Requery
    Dim Urut As String * 5
    Dim Hitung As Long
    With rsSupp
        If .EOF Then
            Urut = "IS" + "001"
            txtID = Urut
        Else
            Hitung = Right(!id_supplier, 3) + 1
            Urut = "IS" + Right("000" & Hitung, 3)
        End If
        txtID = Urut
    End With
End Sub

Private Sub btnAdd_Click()
    BukaText
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
    ClearText
    TutupText
    btnUpdate.Visible = False
    btnCancel.Visible = False
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    btnSave.Enabled = False
End Sub

Private Sub btnDelete_Click()
    If txtID.Text = "" Then
        frmEmpty.Show
    Else
        frmDelete_alert.Show
        btnCancel.Visible = True
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
    Export_Excel "Kunci_Ibadah_Utama.xlsx", "Data Supplier"
End Sub

Private Sub btnRefresh_Click()
    cmbKriteria.Text = ""
    txtKeyword.Text = ""
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Supplier_Table"
    Adodc1.Refresh
End Sub

Private Sub btnSave_Click()
    If txtID.Text = "" Or txtNama.Text = "" Or rtbAlamat.Text = "" Or txtTelp.Text = "" Then
        frmEmpty.Show
    Else
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!id_supplier = txtID.Text
        Adodc1.Recordset!nama_supplier = txtNama.Text
        Adodc1.Recordset!alamat = rtbAlamat.Text
        Adodc1.Recordset!cp = txtTelp.Text
        
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
    
    ClearText
    TutupText
    btnCancel.Visible = False
End Sub

Private Sub btnSearch_Click()
    Call OpenDB
    If cmbKriteria.Text = "ID" Then
        Adodc1.RecordSource = "select * from Supplier_Table where id_supplier like'" & txtKeyword.Text & "'"
        Adodc1.Refresh
        Set Supplier_DataGrid.DataSource = Adodc1
    ElseIf cmbKriteria.Text = "NAMA" Then
        Adodc1.RecordSource = "select * from Supplier_Table where nama_supplier like'" & txtKeyword.Text & "'"
        Adodc1.Refresh
        Set Supplier_DataGrid.DataSource = Adodc1
    Else
        frmNo_Record.Show
    End If
End Sub

Private Sub btnUpdate_Click()
    If txtID.Text = "" Or txtNama.Text = "" Or rtbAlamat.Text = "" Or txtTelp.Text = "" Then
        frmEmpty2.Show
        btnAdd.Enabled = False
        btnEdit.Enabled = False
        btnDelete.Enabled = False
        btnSave.Enabled = False
    Else
        Adodc1.Recordset.Update
        Adodc1.Recordset!id_supplier = txtID.Text
        Adodc1.Recordset!nama_supplier = txtNama.Text
        Adodc1.Recordset!alamat = rtbAlamat.Text
        Adodc1.Recordset!cp = txtTelp.Text
        
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
    
    ClearText
    TutupText
    btnCancel.Visible = False
    btnUpdate.Visible = False
End Sub

Private Sub Form_Load()
    TutupText
    ClearText
    cmbKriteria.AddItem "ID"
    cmbKriteria.AddItem "NAMA"
    btnUpdate.Visible = False
    btnCancel.Visible = False
    btnSave.Enabled = False
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Supplier_Table"
    Adodc1.Refresh
    Set Supplier_DataGrid.DataSource = Adodc1
End Sub

Private Sub Supplier_DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    TAMPILKAN_KE_TEXTBOX
    btnCancel.Visible = True
End Sub
