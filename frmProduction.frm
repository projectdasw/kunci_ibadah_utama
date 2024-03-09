VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProduction 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10320
   ClientLeft      =   -60
   ClientTop       =   -60
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid Production_DataGrid 
      Height          =   3255
      Left            =   240
      TabIndex        =   25
      Top             =   6840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5741
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
   Begin VB.TextBox txtHPP_Unit 
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
      Left            =   9360
      TabIndex        =   23
      Top             =   3840
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   10680
      Top             =   840
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
      Caption         =   "Adodc1"
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
   Begin VB.TextBox txtOverhead 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   19
      Top             =   3840
      Width           =   2175
   End
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
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox txtTotal_HPP 
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
      Height          =   375
      Left            =   7080
      TabIndex        =   17
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox txtTKL 
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
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox txtBiaya 
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
      Left            =   3960
      TabIndex        =   11
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtID_Produk 
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
      Left            =   3960
      TabIndex        =   9
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox cmbNama_Produk 
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
      Left            =   3960
      TabIndex        =   8
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FORM PRODUKSI"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
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
         Left            =   1680
         TabIndex        =   20
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtID_Produksi 
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
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dateTrans 
         Height          =   495
         Left            =   1680
         TabIndex        =   21
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   50593793
         CurrentDate     =   44285
      End
      Begin VB.Label Label5 
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
         Left            =   600
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Produksi"
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
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView Production_ListView 
      Height          =   2415
      Left            =   240
      TabIndex        =   22
      Top             =   4320
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4260
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
      NumItems        =   8
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
         Text            =   "Biaya TKL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Biaya Overhead"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Kuantitas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Total HPP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "HPP Per Unit"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "HPP Per Unit"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9360
      TabIndex        =   24
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Image btnNew 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   6600
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduction.frx":0000
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image btnCancel 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   7680
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduction.frx":45D4
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Image btnPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   9840
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduction.frx":B14E
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Image btnHitung 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   8760
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduction.frx":1061B
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image btnDelete 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   9840
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduction.frx":1645C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image btnSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   10920
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduction.frx":1B601
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image btnAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   7680
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduction.frx":2105C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total HPP"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label10 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Overhead"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Biaya TKL"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Biaya Bahan"
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUKSI"
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
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   3255
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
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmProduction.frx":26106
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   10095
      Left            =   120
      Picture         =   "frmProduction.frx":2CA7D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   11895
   End
End
Attribute VB_Name = "frmProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub AN()
    Call OpenDB
    rsProduction.Open ("SELECT * FROM ProductionData_Table WHERE id_produksi in(select max(id_produksi) from ProductionData_Table)order by id_produksi desc"), ConnectDB
    rsProduction.Requery
    Dim Urut As String * 6
    Dim Hitung As Long
    With rsProduction
        If .EOF Then
            Urut = "IPN" + "001"
            txtID_Produki = Urut
        Else
            Hitung = Right(!id_produksi, 3) + 1
            Urut = "IPN" + Right("000" & Hitung, 3)
        End If
        txtID_Produksi = Urut
    End With
End Sub

Sub Enabled_Text()
    txtID_Produksi.Enabled = True
    dateTrans.Enabled = True
    cmbPeriode.Enabled = True
    txtID_Produk.Enabled = True
    cmbNama_Produk.Enabled = True
    txtBiaya.Enabled = True
    txtTKL.Enabled = True
    txtOverhead.Enabled = True
    txtTotal_HPP.Enabled = True
    txtKuan.Enabled = True
    txtHPP_Unit.Enabled = True
End Sub

Sub Disabled_Text()
    txtID_Produksi.Enabled = False
    cmbPeriode.Enabled = False
    txtID_Produk.Enabled = False
    cmbNama_Produk.Enabled = False
    txtBiaya.Enabled = False
    txtTKL.Enabled = False
    txtOverhead.Enabled = False
    txtTotal_HPP.Enabled = False
    txtKuan.Enabled = False
    txtHPP_Unit.Enabled = False
End Sub

Sub ClearText()
    txtID_Produksi.Text = ""
    dateTrans.Value = Format(Now)
    cmbPeriode.Text = ""
    txtID_Produk.Text = ""
    cmbNama_Produk.Text = ""
    txtBiaya.Text = ""
    txtTKL.Text = ""
    txtOverhead.Text = ""
    txtTotal_HPP.Text = ""
    txtKuan.Text = ""
    txtHPP_Unit.Text = ""
End Sub

Sub TAMPILKAN_KE_TEXTBOX()
    txtID_Produksi.Text = Production_DataGrid.Columns(0).Value
    dateTrans.Value = Production_DataGrid.Columns(1).Value
    cmbPeriode.Text = Production_DataGrid.Columns(2).Value
    txtID_Produk.Text = Production_DataGrid.Columns(3).Value
    cmbNama_Produk.Text = Production_DataGrid.Columns(4).Value
    txtBiaya.Text = Production_DataGrid.Columns(5).Value
    txtTKL.Text = Production_DataGrid.Columns(6).Value
    txtOverhead.Text = Production_DataGrid.Columns(7).Value
    txtTotal_HPP.Text = Production_DataGrid.Columns(8).Value
    txtKuan.Text = Production_DataGrid.Columns(9).Value
    txtHPP_Unit.Text = Production_DataGrid.Columns(10).Value
End Sub

Private Sub btnAdd_Click()
    Dim lv As MSComctlLib.ListItem
    
    If txtID_Produk.Text = "" Or cmbNama_Produk.Text = "" Or txtTotal_HPP.Text = "" Or txtHPP_Unit.Text = "" Or txtKuan.Text = "" Or txtBiaya.Text = "" Or txtTKL.Text = "" Or txtOverhead.Text = "" Then
        frmEmpty.Show
    Else
        Set lv = Production_ListView.ListItems.Add(, , txtID_Produk)
        With lv
            .SubItems(1) = cmbNama_Produk.Text
            .SubItems(2) = txtBiaya.Text
            .SubItems(3) = txtTKL.Text
            .SubItems(4) = txtOverhead.Text
            .SubItems(5) = txtKuan.Text
            .SubItems(6) = txtTotal_HPP.Text
            .SubItems(7) = txtHPP_Unit.Text
        End With
        
        txtID_Produk.Text = ""
        cmbNama_Produk.Text = ""
        txtBiaya.Text = ""
        txtTKL.Text = ""
        txtOverhead.Text = ""
        txtKuan.Text = ""
        txtTotal_HPP.Text = ""
        txtHPP_Unit.Text = ""
    End If
End Sub

Private Sub btnBack_Click()
    Me.Hide
    frmHome.Show
End Sub

Private Sub btnCancel_Click()
    txtID_Produksi.Text = ""
    dateTrans.Value = Format(Now)
    cmbPeriode.Text = ""
    txtID_Produk.Text = ""
    cmbNama_Produk.Text = ""
    txtTotal_HPP.Text = ""
    txtHPP_Unit.Text = ""
    txtKuan.Text = ""
    txtBiaya.Text = ""
    txtTKL.Text = ""
    txtOverhead.Text = ""
    btnCancel.Visible = False
    Disabled_Text
    btnAdd.Enabled = False
    btnHitung.Enabled = False
    btnSave.Enabled = False
    btnDelete.Enabled = True
End Sub

Private Sub btnDelete_Click()
    If txtID_Produksi.Text = "" Then
        frmSelected.Show
    ElseIf Production_ListView.ListItems.Count = 0 Then
        frmDelete_alert.Show
    Else
        Production_ListView.ListItems.Remove (Production_ListView.SelectedItem.Index)
    End If
End Sub

Private Sub btnHitung_Click()
    If txtKuan.Text = "" Then
        frmCount_Alert.Show
    Else
        If cmbNama_Produk.Text = "Songkok Bagus" Then
            txtBiaya.Text = Val(frmKalkulasi.txtHarga_Soba.Text) * Val(txtKuan.Text)
        ElseIf cmbNama_Produk.Text = "Songkok Biasa" Then
            txtBiaya.Text = Val(frmKalkulasi.txtHarga_Sobi.Text) * Val(txtKuan.Text)
        Else
            txtBiaya.Text = ""
        End If
        
        txtTKL.Text = Val(txtTKL.Text) * Val(txtKuan.Text)
        txtOverhead.Text = Val(txtOverhead.Text) * Val(txtKuan.Text)
        txtTotal_HPP.Text = Val(txtBiaya.Text) + Val(txtTKL.Text) + Val(txtOverhead.Text)
        txtHPP_Unit.Text = Val(txtTotal_HPP.Text) / Val(txtKuan.Text)
    End If
End Sub

Private Sub btnNew_Click()
    ClearText
    Enabled_Text
    AN
    txtID_Produksi.Enabled = False
    txtBiaya.Enabled = False
    txtTotal_HPP.Enabled = False
    txtHPP_Unit.Enabled = False
    btnCancel.Visible = True
    btnAdd.Enabled = True
    btnSave.Enabled = True
    btnDelete.Enabled = True
    btnHitung.Enabled = True
End Sub

Private Sub btnPrint_Click()
    Export_Excel "Kunci_Ibadah_Utama.xlsx", "Data Produksi"
End Sub

Private Sub btnSave_Click()
    Call OpenDB
    Dim i As Integer
    
    If txtID_Produksi.Text = "" Or dateTrans.Value = Format(Now) Or cmbPeriode.Text = "" Then
        frmEmpty.Show
    ElseIf Production_ListView.ListItems.Count = 0 Then
        frmNone.Show
    Else
        For i = 1 To Production_ListView.ListItems.Count
            With rsProduction
                .Open "select * from ProductionData_Table", ConnectDB, 1, 2
                .AddNew
                !id_produk = Production_ListView.ListItems(i).Text
                !nama_produk = Production_ListView.ListItems(i).ListSubItems(1).Text
                !harga_produk = Production_ListView.ListItems(i).ListSubItems(2).Text
                !biaya_tkl = Production_ListView.ListItems(i).ListSubItems(3).Text
                !overhead = Production_ListView.ListItems(i).ListSubItems(4).Text
                !kuantitas = Production_ListView.ListItems(i).ListSubItems(5).Text
                !total_hpp = Production_ListView.ListItems(i).ListSubItems(6).Text
                !hpp_unit = Production_ListView.ListItems(i).ListSubItems(7).Text
                !id_produksi = txtID_Produksi.Text
                !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
                !periode = cmbPeriode.Text
                .Update
                .Close
            End With
        Next i
        For i = 1 To Production_ListView.ListItems.Count
            With rsProduction
                .Open "select * from Production_Table", ConnectDB
                .AddNew
                !id_produksi = txtID_Produksi.Text
                !id_produk = Production_ListView.ListItems(i).Text
                !nama_produk = Production_ListView.ListItems(i).ListSubItems(1).Text
                !kuantitas = Production_ListView.ListItems(i).ListSubItems(5).Text
                !total_hpp = Production_ListView.ListItems(i).ListSubItems(6).Text
                !hpp_perunit = Production_ListView.ListItems(i).ListSubItems(7).Text
                !periode = cmbPeriode.Text
                .Update
                .Close
            End With
        Next i
                
        'Script Update Biaya TKL, OVERHEAD & Kuantitas
        i = 1
        Do While i <= Production_ListView.ListItems.Count
            kode = Production_ListView.ListItems(i).Text
            h_produk = Production_ListView.ListItems(i).ListSubItems(2).Text
            tkl = Production_ListView.ListItems(i).ListSubItems(3).Text
            overhead = Production_ListView.ListItems(i).ListSubItems(4).Text
            laba_hpp = Production_ListView.ListItems(i).ListSubItems(6).Text
            produk_hpp = Production_ListView.ListItems(i).ListSubItems(7).Text
            kuan = Production_ListView.ListItems(i).ListSubItems(5).Text
            
            'KUANTITAS
            Buka_Produk kode
            PKUAN = rsProduct!jumlah_unit
            
            updatePKUAN = Val(PKUAN) + Val(kuan)
            
            Ubah_KuanProduk kode, updatePKUAN
            i = i + 1
            
            'HPP PRODUK PER UNIT
            Buka_HPP kode
            'PHPP = rsHPP!hpp_perunit
            updatePHPP = Val(produk_hpp)
            
            Ubah_HPPProduk kode, updatePHPP
            i = i + 1
            
            '1. Persed. Brg dlm proses(D) - Persed.Bahan baku(K)
            '-------------------------------------------------
                'PERSEDIAAN BARANG DALAM PROSES
                Buka_Jurnal "1.1.2.3"
                PBDB = rsJournal!saldo_akhir
                
                updatePBDB = Val(PBDB) + Val(h_produk)
                
                Ubah_Jurnal "1.1.2.3", updatePBDB
                i = i + 1
                
                'PERSEDIAAN BAHAN BAKU
                Buka_Jurnal "1.1.2.1"
                PBB = rsJournal!saldo_akhir
                
                updatePBB = Val(PBB) - Val(h_produk)
                
                Ubah_Jurnal "1.1.2.1", updatePBB
                i = i + 1
            '-------------------------------------------------
            
            '2. Persed. Brg dlm proses(D) - Biaya Overhead(K)
            '-------------------------------------------------
                'PERSEDIAAN BARANG DALAM PROSES
                Buka_Jurnal "1.1.2.3"
                PBDB = rsJournal!saldo_akhir
                
                updatePBDB = Val(PBDB) + Val(overhead)
                
                Ubah_Jurnal "1.1.2.3", updatePBDB
                i = i + 1
                
                'OVERHEAD
                Buka_Jurnal "6.2.1.2"
                bover = rsJournal!saldo_akhir
                
                updateBOVER = Val(bover) - Val(overhead)
                
                Ubah_Jurnal "6.2.1.2", updateBOVER
                i = i + 1
            '-------------------------------------------------
            
            '3. Persed. Brg dlm proses(D) - Biaya TKL(K)
            '-------------------------------------------------
                'PERSEDIAAN BARANG DALAM PROSES
                Buka_Jurnal "1.1.2.3"
                PBDB = rsJournal!saldo_akhir
                
                updatePBDB = Val(PBDB) + Val(tkl)
                
                Ubah_Jurnal "1.1.2.3", updatePBDB
                i = i + 1
                
                'TKL
                Buka_Jurnal "6.2.1.1"
                btkl = rsJournal!saldo_akhir
                
                updateBTKL = Val(btkl) - Val(tkl)
                
                Ubah_Jurnal "6.2.1.1", updateBTKL
                i = i + 1
            '-------------------------------------------------
            
            '4. Persed.Bahan jadi(D) - Persed. Brg dlm proses(K)
            '-------------------------------------------------
                'PERSEDIAAN BARANG JADI
                Buka_Jurnal "1.1.2.4"
                PBJ = rsJournal!saldo_akhir
                
                updatePBJ = Val(PBJ) + Val(laba_hpp)
                
                Ubah_Jurnal "1.1.2.4", updatePBJ
                i = i + 1
                
                'PERSEDIAAN BARANG DALAM PROSES
                Buka_Jurnal "1.1.2.3"
                PBDB = rsJournal!saldo_akhir
                
                updatePBDB = Val(PBDB) - Val(laba_hpp)
                
                Ubah_Jurnal "1.1.2.3", updatePBDB
                i = i + 1
            '-------------------------------------------------
            
            '5. Biaya TKL(D) - KAS(K)
            '-------------------------------------------------
                'TKL
                Buka_Jurnal "6.2.1.1"
                btkl = rsJournal!saldo_akhir
                
                updateBTKL = Val(btkl) + Val(tkl)
                
                Ubah_Jurnal "6.2.1.1", updateBTKL
                i = i + 1
                
                'KAS (TKL)
                Buka_Jurnal "1.1.1.1"
                kas_tkl = rsJournal!saldo_akhir
                
                updateKTKL = Val(kas_tkl) - Val(tkl)
                
                Ubah_Jurnal "1.1.1.1", updateKTKL
                i = i + 1
            '-------------------------------------------------
            
            '6. Biaya Overhead(D) - KAS(K)
            '-------------------------------------------------
                'OVERHEAD
                Buka_Jurnal "6.2.1.2"
                bover = rsJournal!saldo_akhir
                
                updateBOVER = Val(bover) + Val(overhead)
                
                Ubah_Jurnal "6.2.1.2", updateBOVER
                i = i + 1
                
                'KAS (OVERHEAD)
                Buka_Jurnal "1.1.1.1"
                kas_bover = rsJournal!saldo_akhir
                
                updateKBOVER = Val(kas_bover) - Val(overhead)
                
                Ubah_Jurnal "1.1.1.1", updateKBOVER
                i = i + 1
            '-------------------------------------------------
            
            '1. Jurnal(Persed. Brg dlm proses(D) - Persed.Bahan baku(K))
            With rsProduction
                .Open "select * from Journal_Data", ConnectDB
                .AddNew
                !id_trans = txtID_Produksi.Text
                !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
                !periode = cmbPeriode.Text
                !uraian = "Barang Proses Produksi"
                !nama_akun1 = "PERSEDIAAN BARANG DALAM PROSES"
                !id_akun1 = "1.1.2.3"
                !debit = Val(h_produk)
                !nama_akun2 = "PERSEDIAAN BAHAN BAKU"
                !id_akun2 = "1.1.2.2"
                !kredit = Val(h_produk)
                .Update
                .Close
            End With
            
            '2. Jurnal(Persed. Brg dlm proses(D) - Biaya Overhead(K))
            With rsProduction
                .Open "select * from Journal_Data", ConnectDB
                .AddNew
                !id_trans = txtID_Produksi.Text
                !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
                !periode = cmbPeriode.Text
                !uraian = "Barang Proses Produksi & Biaya Overhead"
                !nama_akun1 = "PERSEDIAAN BARANG DALAM PROSES"
                !id_akun1 = "1.1.2.3"
                !debit = Val(overhead)
                !nama_akun2 = "BIAYA OVERHEAD"
                !id_akun2 = "6.2.1.2"
                !kredit = Val(overhead)
                .Update
                .Close
            End With
            
            '3. Jurnal(Persed. Brg dlm proses(D) - Biaya TKL(K))
            With rsProduction
                .Open "select * from Journal_Data", ConnectDB
                .AddNew
                !id_trans = txtID_Produksi.Text
                !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
                !periode = cmbPeriode.Text
                !uraian = "Barang Proses Produksi & Biaya TKL"
                !nama_akun1 = "PERSEDIAAN BARANG DALAM PROSES"
                !id_akun1 = "1.1.2.3"
                !debit = Val(tkl)
                !nama_akun2 = "BIAYA TKL"
                !id_akun2 = "6.2.1.1"
                !kredit = Val(tkl)
                .Update
                .Close
            End With
            
            '4. Jurnal(Persed.Bahan jadi(D) - Persed. Brg dlm proses(K))
            With rsProduction
                .Open "select * from Journal_Data", ConnectDB
                .AddNew
                !id_trans = txtID_Produksi.Text
                !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
                !periode = cmbPeriode.Text
                !uraian = "Barang Telah diproses"
                !nama_akun1 = "PERSEDIAAN BARANG JADI"
                !id_akun1 = "1.1.2.4"
                !debit = Val(laba_hpp)
                !nama_akun2 = "PERSEDIAAN BARANG DALAM PROSES"
                !id_akun2 = "1.1.2.3"
                !kredit = Val(laba_hpp)
                .Update
                .Close
            End With
            
            '5. Jurnal(Biaya TKL(D) - KAS(K))
            With rsProduction
                .Open "select * from Journal_Data", ConnectDB
                .AddNew
                !id_trans = txtID_Produksi.Text
                !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
                !periode = cmbPeriode.Text
                !uraian = "Biaya TKL"
                !nama_akun1 = "BIAYA TKL"
                !id_akun1 = "6.2.1.1"
                !debit = Val(tkl)
                !nama_akun2 = "KAS"
                !id_akun2 = "1.1.1.1"
                !kredit = Val(tkl)
                .Update
                .Close
            End With
            
            '6. Jurnal(Biaya Overhead(D) - KAS(K))
            With rsProduction
                .Open "select * from Journal_Data", ConnectDB
                .AddNew
                !id_trans = txtID_Produksi.Text
                !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
                !periode = cmbPeriode.Text
                !uraian = "Biaya Overhead"
                !nama_akun1 = "BIAYA OVERHEAD"
                !id_akun1 = "6.2.1.2"
                !debit = Val(overhead)
                !nama_akun2 = "KAS"
                !id_akun2 = "1.1.1.1"
                !kredit = Val(overhead)
                .Update
                .Close
            End With
        Loop
        
        frmSave_Alert.Show
        Production_ListView.ListItems.Clear
        txtID_Produksi.Text = ""
        dateTrans.Value = Format(Now)
        cmbPeriode.Text = ""
        Disabled_Text
        btnCancel.Visible = False
        btnAdd.Enabled = False
        
        Call OpenDB
        Adodc1.ConnectionString = ConnectDB
        Adodc1.RecordSource = "ProductionData_Table"
        Adodc1.Refresh
        Set Production_DataGrid.DataSource = Adodc1
    End If
End Sub

Private Sub cmbNama_Produk_Click()
    txtID_Produk = Tampil_IDProduk("Product_Table", "id_produk", cmbNama_Produk)
    'txtBiaya = Tampil_IDProduk("Product_Table", "hpp_produk", cmbNama_Produk)
    txtID_Produk.Enabled = False
End Sub

Private Sub Form_Load()
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "ProductionData_Table"
    Adodc1.Refresh
    Set Production_DataGrid.DataSource = Adodc1
    
    Disabled_Text
    dateTrans.Value = Format(Now)
    txtID_Produksi.Enabled = False
    txtTotal_HPP.Enabled = False
    txtBiaya.Enabled = False
    btnCancel.Visible = False
    btnNew.Enabled = True
    btnHitung.Enabled = False
    btnAdd.Enabled = False
    btnSave.Enabled = False
    btnDelete.Enabled = True
    
    rsPro.Open "select * from Product_Table", ConnectDB
    rsPer.Open "select * from Periode_Table", ConnectDB
    cmbNama_Produk.Clear
    cmbPeriode.Clear
    Do While Not rsPro.EOF
        cmbNama_Produk.AddItem rsPro!nama_produk
        rsPro.MoveNext
    Loop
    Do While Not rsPer.EOF
        cmbPeriode.AddItem rsPer!masa
        rsPer.MoveNext
    Loop
End Sub

Private Sub Production_DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    TAMPILKAN_KE_TEXTBOX
    Disabled_Text
    btnCancel.Visible = True
End Sub
