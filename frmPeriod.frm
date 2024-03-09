VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPeriod 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   7785
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
      Height          =   2535
      Left            =   4920
      TabIndex        =   12
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
         TabIndex        =   14
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
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Image btnSearch 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   360
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmPeriod.frx":0000
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FORM PERIODE"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   4575
      Begin VB.TextBox txtMasa 
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
         TabIndex        =   8
         Top             =   1920
         Width           =   2175
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
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dateAwal 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         Format          =   116326401
         CurrentDate     =   44306
      End
      Begin MSComCtl2.DTPicker dateAkhir 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         Format          =   116326401
         CurrentDate     =   44306
      End
      Begin VB.Image btnCancel 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   2400
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmPeriod.frx":5442
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image btnUpdate 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   360
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmPeriod.frx":BFBC
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Masa Periode"
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
         TabIndex        =   11
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Akhir"
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
         TabIndex        =   10
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Awal"
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
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Periode"
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
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   6240
      Top             =   480
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
   Begin MSDataGridLib.DataGrid Period_DataGrid 
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   5280
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3625
      _Version        =   393216
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
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPeriod.frx":1104C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image btnPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   5640
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPeriod.frx":179C3
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Image btnRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   5640
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPeriod.frx":1CE90
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Image btnDelete 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   3480
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPeriod.frx":22B5B
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image btnEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   2400
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPeriod.frx":27D00
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image btnSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   1320
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPeriod.frx":2CCC2
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image btnAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmPeriod.frx":3271D
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODE"
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
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   2775
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
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   7335
      Left            =   120
      Picture         =   "frmPeriod.frx":377C7
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Disabled_Text()
    txtID.Enabled = False
    dateAwal.Enabled = False
    dateAkhir.Enabled = False
    txtMasa.Enabled = False
End Sub

Sub Enabled_Text()
    txtID.Enabled = True
    dateAwal.Enabled = True
    dateAkhir.Enabled = True
    txtMasa.Enabled = True
End Sub

Sub Clear_Text()
    txtID.Text = ""
    dateAwal.Value = Format(Now)
    dateAkhir.Value = Format(Now)
    txtMasa.Text = ""
End Sub

Sub AN()
    Call OpenDB
    rsPer.Open ("SELECT * FROM Periode_Table WHERE id_periode in(select max(id_periode) from Periode_Table)order by id_periode desc"), ConnectDB
    rsPer.Requery
    Dim Urut As String * 4
    Dim Hitung As Long
    With rsPer
        If .EOF Then
            Urut = "prd" + "1"
            txtID = Urut
        Else
            Hitung = Right(!id_periode, 1) + 1
            Urut = "prd" + Right("0" & Hitung, 1)
        End If
        txtID = Urut
    End With
End Sub

Sub TAMPILKAN_KE_TEXTBOX()
    txtID.Text = Period_DataGrid.Columns(0).Value
    dateAwal.Value = Period_DataGrid.Columns(1).Value
    dateAkhir.Value = Period_DataGrid.Columns(2).Value
    txtMasa.Text = Period_DataGrid.Columns(3).Value
End Sub

Private Sub btnAdd_Click()
    Enabled_Text
    Clear_Text
    txtID.Enabled = False
    AN
    btnCancel.Visible = True
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    btnSave.Enabled = True
End Sub

Private Sub btnBack_Click()
    Me.Hide
    frmHome.Show
End Sub

Private Sub btnCancel_Click()
    Clear_Text
    Disabled_Text
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    btnSave.Enabled = False
    btnCancel.Visible = False
    btnUpdate.Visible = False
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
        Enabled_Text
        txtID.Enabled = False
        btnUpdate.Visible = True
        btnCancel.Visible = True
    End If
End Sub

Private Sub btnPrint_Click()
    Export_Excel "Kunci_Ibadah_Utama.xlsx", "Data Periode"
End Sub

Private Sub btnRefresh_Click()
    cmbKriteria.Text = ""
    txtKeyword.Text = ""
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Periode_Table"
    Adodc1.Refresh
End Sub

Private Sub btnSave_Click()
    If txtID.Text = "" Or dateAwal.Value = "" Or dateAkhir.Value = "" Or txtMasa.Text = "" Then
        frmEmpty.Show
        btnAdd.Enabled = False
        btnEdit.Enabled = False
        btnDelete.Enabled = False
    Else
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!id_periode = txtID.Text
        Adodc1.Recordset!tgl_awal = Format(dateAwal.Value, "mm-dd-yyyy")
        Adodc1.Recordset!tgl_akhir = Format(dateAkhir.Value, "mm-dd-yyyy")
        Adodc1.Recordset!masa = txtMasa.Text
        
        frmSave_Alert.Show
        Adodc1.Recordset.MoveFirst
        Clear_Text
        Disabled_Text
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
        Adodc1.RecordSource = "select * from Periode_Table where id_periode like'" & txtKeyword.Text & "'"
        Adodc1.Refresh
        Set Period_DataGrid.DataSource = Adodc1
    ElseIf cmbKriteria.Text = "NAMA" Then
        Adodc1.RecordSource = "select * from Periode_Table where masa like'" & txtKeyword.Text & "'"
        Adodc1.Refresh
        Set Period_DataGrid.DataSource = Adodc1
    Else
        frmNo_Record.Show
    End If
End Sub

Private Sub btnUpdate_Click()
    If txtID.Text = "" Or dateAwal.Value = "" Or dateAkhir.Value = "" Or txtMasa.Text = "" Then
        frmEmpty2.Show
        btnAdd.Enabled = False
        btnEdit.Enabled = False
        btnDelete.Enabled = False
    Else
        Adodc1.Recordset.Update
        Adodc1.Recordset!id_periode = txtID.Text
        Adodc1.Recordset!tgl_awal = dateAwal.Value
        Adodc1.Recordset!tgl_akhir = dateAkhir.Value
        Adodc1.Recordset!masa = txtMasa.Text
        
        frmUpdate_Alert.Show
        Adodc1.Recordset.MoveFirst
        Clear_Text
        Disabled_Text
        btnAdd.Enabled = True
        btnEdit.Enabled = True
        btnDelete.Enabled = True
        btnSave.Enabled = False
        btnUpdate.Visible = False
        btnCancel.Visible = False
    End If
End Sub

Private Sub Form_Load()
    cmbKriteria.AddItem "ID"
    cmbKriteria.AddItem "MASA"
    
    Disabled_Text
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    btnSave.Enabled = False
    btnUpdate.Visible = False
    btnCancel.Visible = False
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Periode_Table"
    Adodc1.Refresh
    Set Period_DataGrid.DataSource = Adodc1
End Sub

Private Sub Period_DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    TAMPILKAN_KE_TEXTBOX
    btnCancel.Visible = True
End Sub
