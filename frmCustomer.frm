VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   13350
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
      Height          =   1815
      Left            =   7560
      TabIndex        =   11
      Top             =   3120
      Width           =   5535
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
         ItemData        =   "frmCustomer.frx":0000
         Left            =   120
         List            =   "frmCustomer.frx":0002
         TabIndex        =   13
         Top             =   360
         Width           =   2535
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
         Left            =   2880
         TabIndex        =   12
         Top             =   360
         Width           =   2535
      End
      Begin VB.Image btnSearch 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   1800
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmCustomer.frx":0004
         Stretch         =   -1  'True
         Top             =   960
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   120
      Top             =   5160
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "FORM DATA PELANGGAN"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
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
         Height          =   405
         Left            =   2520
         TabIndex        =   14
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtCP 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   2520
         TabIndex        =   7
         Top             =   2640
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
         Height          =   405
         Left            =   2520
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin RichTextLib.RichTextBox rtbAlamat 
         Height          =   1095
         Left            =   2520
         TabIndex        =   10
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmCustomer.frx":5446
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
         Height          =   615
         Left            =   3120
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmCustomer.frx":54CE
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Image btnUpdate 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   960
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmCustomer.frx":C048
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "CP"
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
         Left            =   2040
         TabIndex        =   8
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label5 
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
         Left            =   1440
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pelanggan"
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
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Pelanggan"
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
         Width           =   1815
      End
   End
   Begin MSDataGridLib.DataGrid Customer_DataGrid 
      Height          =   1935
      Left            =   5520
      TabIndex        =   9
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3413
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
   Begin VB.Image btnSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   6600
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmCustomer.frx":110D8
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image btnEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   5520
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmCustomer.frx":16B33
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   855
   End
   Begin VB.Image btnAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   5520
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmCustomer.frx":1BAF5
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image btnDelete 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   6600
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmCustomer.frx":20B9F
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PELANGGAN"
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
      Left            =   4440
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.Image btnPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   11280
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmCustomer.frx":25D44
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1815
   End
   Begin VB.Image btnRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   9480
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmCustomer.frx":2B211
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmCustomer.frx":30EDC
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   120
      Picture         =   "frmCustomer.frx":37853
      Stretch         =   -1  'True
      Top             =   120
      Width           =   13095
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub TutupText()
    txtID.Enabled = False
    txtNama.Enabled = False
    rtbAlamat.Enabled = False
    txtCP.Enabled = False
End Sub

Sub BukaText()
    txtID.Enabled = True
    txtNama.Enabled = True
    rtbAlamat.Enabled = True
    txtCP.Enabled = True
End Sub

Sub ClearText()
    txtID.Text = ""
    txtNama.Text = ""
    rtbAlamat.Text = ""
    txtCP.Text = ""
End Sub

Sub TAMPILKAN_KE_TEXTBOX()
    txtID.Text = Customer_DataGrid.Columns(0).Value
    txtNama.Text = Customer_DataGrid.Columns(1).Value
    rtbAlamat.Text = Customer_DataGrid.Columns(2).Value
    txtCP.Text = Customer_DataGrid.Columns(3).Value
End Sub

Sub AN()
    Call OpenDB
    rsPel.Open ("SELECT * FROM Customer_Table WHERE id_customer in(select max(id_customer) from Customer_Table)order by id_customer desc"), ConnectDB
    rsPel.Requery
    Dim Urut As String * 5
    Dim Hitung As Long
    With rsPel
        If .EOF Then
            Urut = "IC" + "001"
            txtID.Text = Urut
        Else
            Hitung = Right(!id_customer, 3) + 1
            Urut = "IC" + Right("000" & Hitung, 3)
        End If
        txtID.Text = Urut
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
    Export_Excel "Kunci_Ibadah_Utama.xlsx", "Data Pelanggan"
End Sub

Private Sub btnRefresh_Click()
    cmbKriteria.Text = ""
    txtKeyword.Text = ""
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Customer_Table"
    Adodc1.Refresh
End Sub

Private Sub btnSave_Click()
    Call OpenDB
    If txtID.Text = "" Or txtNama.Text = "" Or rtbAlamat.Text = "" Or txtCP.Text = "" Then
        frmEmpty.Show
        btnAdd.Enabled = False
        btnEdit.Enabled = False
        btnDelete.Enabled = False
    Else
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!id_customer = txtID.Text
        Adodc1.Recordset!nama_customer = txtNama.Text
        Adodc1.Recordset!alamat = rtbAlamat.Text
        Adodc1.Recordset!cp = txtCP.Text
        
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
        Adodc1.RecordSource = "select * from Customer_Table where id_customer like'" & txtKeyword.Text & "'"
        Adodc1.Refresh
        Set Customer_DataGrid.DataSource = Adodc1
    ElseIf cmbKriteria.Text = "NAMA" Then
        Adodc1.RecordSource = "select * from Customer_Table where nama_customer like'" & txtKeyword.Text & "'"
        Adodc1.Refresh
        Set Customer_DataGrid.DataSource = Adodc1
    Else
        frmNo_Record.Show
    End If
End Sub

Private Sub btnUpdate_Click()
    If txtID.Text = "" Or txtNama.Text = "" Or rtbAlamat.Text = "" Or txtCP.Text = "" Then
        frmEmpty2.Show
        btnAdd.Enabled = False
        btnEdit.Enabled = False
        btnDelete.Enabled = False
        btnSave.Enabled = False
    Else
        Adodc1.Recordset.Update
        Adodc1.Recordset!id_customer = txtID.Text
        Adodc1.Recordset!nama_customer = txtNama.Text
        Adodc1.Recordset!alamat = rtbAlamat.Text
        Adodc1.Recordset!cp = txtCP.Text
        
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

Private Sub Customer_DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    TAMPILKAN_KE_TEXTBOX
    btnCancel.Visible = True
End Sub

Private Sub Form_Load()
    TutupText
    Adodc1.Visible = False
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
    Adodc1.RecordSource = "Customer_Table"
    Adodc1.Refresh
    Set Customer_DataGrid.DataSource = Adodc1
End Sub
