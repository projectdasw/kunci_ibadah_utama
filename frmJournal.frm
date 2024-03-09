VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJournal 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   225
   ClientTop       =   2475
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   14880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbPosisi2 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10080
      TabIndex        =   26
      Top             =   2520
      Width           =   1815
   End
   Begin VB.ComboBox cmbPosisi1 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10080
      TabIndex        =   25
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cmbNama_Akun2 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5400
      TabIndex        =   22
      Top             =   2520
      Width           =   4575
   End
   Begin VB.TextBox txtID_Akun2 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   20
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtNominal 
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
      Left            =   12000
      TabIndex        =   18
      Top             =   1440
      Width           =   2655
   End
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
      Height          =   1695
      Left            =   5520
      TabIndex        =   16
      Top             =   3120
      Width           =   7695
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
         TabIndex        =   24
         Top             =   360
         Width           =   5415
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
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   5415
      End
      Begin VB.Image btnRefresh 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   5640
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmJournal.frx":0000
         Stretch         =   -1  'True
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image btnSearch 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   5640
         MousePointer    =   10  'Up Arrow
         Picture         =   "frmJournal.frx":5CCB
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1935
      End
   End
   Begin RichTextLib.RichTextBox rtbKet 
      Height          =   1575
      Left            =   2280
      TabIndex        =   15
      Top             =   3240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmJournal.frx":B10D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   7680
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
   Begin MSComCtl2.DTPicker dateTrans 
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
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
      Format          =   152567809
      CurrentDate     =   44285
   End
   Begin VB.ComboBox cmbNama_Akun 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5400
      TabIndex        =   10
      Top             =   1440
      Width           =   4575
   End
   Begin VB.TextBox txtID_Akun 
      BeginProperty Font 
         Name            =   "Montserrat SemiBold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   9
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FORM JURNAL"
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
      Top             =   960
      Width           =   3255
      Begin VB.TextBox txtID_Trans 
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
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   1815
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
         Left            =   1320
         TabIndex        =   6
         Top             =   1560
         Width           =   1815
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
         Height          =   435
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
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
         Height          =   435
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Trans"
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
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid Journal_DataGrid 
      Height          =   2655
      Left            =   7920
      TabIndex        =   19
      Top             =   4920
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
   Begin MSComctlLib.ListView Journal_ListView 
      Height          =   2655
      Left            =   120
      TabIndex        =   27
      Top             =   4920
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   10
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Akun 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Akun 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Debit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ID Akun 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nama Akun 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Kredit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Keterangan"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image btnAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   960
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmJournal.frx":B195
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Image btnPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   13320
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmJournal.frx":1023F
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Akun 2"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   23
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Akun 2"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      TabIndex        =   21
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Image btnCancel 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   13320
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmJournal.frx":1570C
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Image btnDelete 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   13440
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmJournal.frx":1C286
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image btnSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   12120
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmJournal.frx":2142B
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Nominal"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12000
      TabIndex        =   13
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Akun 1"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Akun 1"
      BeginProperty Font 
         Name            =   "Montserrat ExtraBold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JURNAL"
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
      Height          =   765
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   765
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image btnBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmJournal.frx":26E86
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   7575
      Left            =   120
      Picture         =   "frmJournal.frx":2D7FD
      Stretch         =   -1  'True
      Top             =   120
      Width           =   14655
   End
End
Attribute VB_Name = "frmJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Disabled_Text()
    dateTrans.Enabled = False
    cmbPeriode.Enabled = False
    cmbNama_Akun.Enabled = False
    cmbNama_Akun2.Enabled = False
    txtNominal.Enabled = False
    rtbKet.Enabled = False
End Sub

Sub Enabled_Text()
    dateTrans.Enabled = True
    cmbPeriode.Enabled = True
    cmbNama_Akun.Enabled = True
    cmbNama_Akun2.Enabled = True
    txtNominal.Enabled = True
    rtbKet.Enabled = True
End Sub

Sub ClearText()
    AN
    txtID_Akun.Text = ""
    txtID_Akun2.Text = ""
    dateTrans.Value = Format(Now)
    cmbNama_Akun.Text = ""
    cmbNama_Akun2.Text = ""
    cmbPeriode.Text = ""
    txtNominal.Text = ""
    cmbPosisi1.Text = ""
    cmbPosisi2.Text = ""
    rtbKet.Text = ""
    btnCancel.Visible = False
End Sub

Sub AN()
    Call OpenDB
    'rsJournal.Open ("SELECT * FROM Journal_Data WHERE id_trans in(select max(id_trans) from AccountData_Table)order by id_trans desc"), ConnectDB
    'rsJournal.Requery
    rsPJ.Open ("SELECT * FROM Journal_Data WHERE id_trans in(select max(id_trans) from AccountData_Table)order by id_trans desc"), ConnectDB
    rsPJ.Requery
    Dim Urut As String * 8
    Dim Hitung As Long
    'With rsJournal
    With rsPJ
        If .EOF Then
            Urut = "ITJ-" + "0001"
            txtID_Trans = Urut
        Else
            Hitung = Right(!id_trans, 4) + 1
            Urut = "ITJ-" + Right("0000" & Hitung, 4)
        End If
        txtID_Trans = Urut
    End With
End Sub

Sub TAMPILKAN_KE_TEXTBOX()
    txtID_Trans.Text = Journal_DataGrid.Columns(0).Value
    dateTrans.Value = Journal_DataGrid.Columns(1).Value
    cmbPeriode.Text = Journal_DataGrid.Columns(2).Value
    rtbKet.Text = Journal_DataGrid.Columns(3).Value
    cmbNama_Akun.Text = Journal_DataGrid.Columns(4).Value
    txtID_Akun.Text = Journal_DataGrid.Columns(5).Value
    cmbPosisi1.Text = Tampil_PosisiAkun("Account_Table", "posisi", cmbNama_Akun)
    cmbNama_Akun2.Text = Journal_DataGrid.Columns(7).Value
    txtID_Akun2.Text = Journal_DataGrid.Columns(8).Value
    cmbPosisi2.Text = Tampil_PosisiAkun("Account_Table", "posisi", cmbNama_Akun2)
    txtNominal.Text = Tampil_NominalJurnal("AccountData_Table", "nominal_perakun", txtID_Trans)
End Sub

Private Sub btnAdd_Click()
    Dim lv As MSComctlLib.ListItem
    If cmbNama_Akun.Text = "" Or cmbNama_Akun2.Text = "" Or cmbPosisi1.Text = "" Or cmbPosisi2.Text = "" Or txtNominal.Text = "" Or rtbKet.Text = "" Then
        frmEmpty.Show
    Else
        Set lv = Journal_ListView.ListItems.Add(, , txtID_Akun.Text)
        With lv
            .SubItems(1) = cmbNama_Akun.Text
            .SubItems(2) = txtNominal.Text
            .SubItems(3) = txtID_Akun2.Text
            .SubItems(4) = cmbNama_Akun2.Text
            .SubItems(5) = txtNominal.Text
            .SubItems(6) = rtbKet.Text
        End With
        
        txtID_Akun.Text = ""
        txtID_Akun2.Text = ""
        cmbNama_Akun.Text = ""
        cmbNama_Akun2.Text = ""
        txtNominal.Text = ""
        cmbPosisi1.Text = ""
        cmbPosisi2.Text = ""
        rtbKet.Text = ""
    End If
End Sub

Private Sub btnBack_Click()
    frmHome.Show
    frmJournal.Hide
End Sub

Private Sub btnCancel_Click()
    AN
    Journal_ListView.ListItems.Clear
    Enabled_Text
    ClearText
    txtNominal.Enabled = True
    btnSave.Enabled = True
    btnDelete.Enabled = True
    
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Journal_Data"
    Adodc1.Refresh
End Sub

Private Sub btnDelete_Click()
    If txtID_Akun.Text = "" Or txtID_Akun2.Text = "" Then
        frmSelected.Show
    Else
        frmDelete_alert.Show
        AN
        btnCancel.Visible = False
    End If
End Sub

Private Sub btnPrint_Click()
    Export_Excel "Kunci_Ibadah_Utama.xlsx", "Jurnal"
End Sub

Private Sub btnRefresh_Click()
    cmbKriteria.Text = ""
    txtKeyword.Text = ""
    Call OpenDB
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Journal_Data"
    Adodc1.Refresh
End Sub

Private Sub btnSave_Click()
    Call OpenDB
    Dim i As Integer
    
    If Journal_ListView.ListItems.Count = 0 Then
        frmEmpty.Show
    ElseIf cmbPeriode.Text = "" Then
        frmEmpty_Period.Show
    Else
        For i = 1 To Journal_ListView.ListItems.Count
            'With rsJournal
            With rsPJ
                .Open "select * from Journal_Data", ConnectDB, 1, 2
                .AddNew
                !id_trans = txtID_Trans.Text
                !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
                !periode = cmbPeriode.Text
                !uraian = Journal_ListView.ListItems(i).ListSubItems(6).Text
                !nama_akun1 = Journal_ListView.ListItems(i).ListSubItems(1).Text
                !id_akun1 = Journal_ListView.ListItems(i).Text
                !debit = Journal_ListView.ListItems(i).ListSubItems(2).Text
                !nama_akun2 = Journal_ListView.ListItems(i).ListSubItems(4).Text
                !id_akun2 = Journal_ListView.ListItems(i).ListSubItems(3).Text
                !kredit = Journal_ListView.ListItems(i).ListSubItems(5).Text
                .Update
                .Close
            End With
        Next i
        
        For i = 1 To Journal_ListView.ListItems.Count
            'With rsJournal
            With rsPJ
                .Open "select * from AccountData_Table", ConnectDB, 1, 2
                .AddNew
                !id_trans = txtID_Trans.Text
                !tanggal = Format(dateTrans.Value, "mm-dd-yyyy")
                !periode = cmbPeriode.Text
                !nominal_perakun = Journal_ListView.ListItems(i).ListSubItems(2).Text
                !id_akun1 = Journal_ListView.ListItems(i).Text
                !debit = Journal_ListView.ListItems(i).ListSubItems(2).Text
                !id_akun2 = Journal_ListView.ListItems(i).ListSubItems(3).Text
                !kredit = Journal_ListView.ListItems(i).ListSubItems(5).Text
                .Update
                .Close
            End With
        Next i
                
        'Script Perselisihan Jurnal Umum
        i = 1
        Do While i <= Journal_ListView.ListItems.Count
            kode1 = Journal_ListView.ListItems(i).Text
            nm1 = Journal_ListView.ListItems(i).ListSubItems(1).Text
            nom = Journal_ListView.ListItems(i).ListSubItems(2).Text
            kode2 = Journal_ListView.ListItems(i).ListSubItems(3).Text
            nm2 = Journal_ListView.ListItems(i).ListSubItems(4).Text
            nom2 = Journal_ListView.ListItems(i).ListSubItems(5).Text
            
            'DEBIT
            If Tampil_PosisiAkun("Account_Table", "posisi", nm1) = "DEBIT" Then
                Buka_Jurnal kode1
                posisi1 = rsJournal!saldo_akhir
                
                update1 = Val(posisi1) + Val(nom)
                Ubah_Jurnal kode1, update1
                i = i + 1
            Else
                Buka_Jurnal kode1
                posisi1 = rsJournal!saldo_akhir
                
                update1 = Val(posisi1) - Val(nom)
                Ubah_Jurnal kode1, update1
                i = i + 1
            End If
            
            'Kredit
            If Tampil_PosisiAkun("Account_Table", "posisi", nm2) = "KREDIT" Then
                Buka_Jurnal kode2
                posisi2 = rsJournal!saldo_akhir
                
                update2 = Val(posisi2) + Val(nom2)
                Ubah_Jurnal kode2, update2
                i = i + 1
            Else
                Buka_Jurnal kode2
                posisi2 = rsJournal!saldo_akhir
                
                update2 = Val(posisi2) - Val(nom2)
                Ubah_Jurnal kode2, update2
                i = i + 1
            End If
        Loop

        frmSave_Alert.Show
        ClearText
        Journal_ListView.ListItems.Clear
        Call OpenDB
        Adodc1.ConnectionString = ConnectDB
        Adodc1.RecordSource = "Journal_Data"
        Adodc1.Refresh
    End If
End Sub

Private Sub btnSearch_Click()
    Call OpenDB
    If cmbKriteria.Text = "" Then
        frmNo_Record.Show
    ElseIf cmbKriteria.Text = Journal_DataGrid.Columns(4).Value Then
        Adodc1.RecordSource = "select * from Journal_Data where nama_akun1 like'" & cmbKriteria.Text & "'"
        Adodc1.Refresh
        Set Journal_DataGrid.DataSource = Adodc1
    Else
        Adodc1.RecordSource = "select * from Journal_Data where nama_akun2 like'" & cmbKriteria.Text & "'"
        Adodc1.Refresh
        Set Journal_DataGrid.DataSource = Adodc1
    End If
End Sub

Private Sub cmbKriteria_Click()
    txtKeyword.Text = Tampil_Journal("Account_Table", "id_akun", cmbKriteria)
End Sub

Private Sub cmbNama_Akun_Click()
    If cmbNama_Akun.Text = "" Then
        btnCancel.Visible = False
    Else
        btnCancel.Visible = True
    End If
    
    txtID_Akun = Tampil_IDAkun("Account_Table", "id_akun", cmbNama_Akun)
    'txtPosisi = Tampil_PosisiAkun("Account_Table", "posisi", cmbNama_Akun)
End Sub

Private Sub cmbNama_Akun2_Click()
    If cmbNama_Akun2.Text = "" Then
        btnCancel.Visible = False
    Else
        btnCancel.Visible = True
    End If
    
    txtID_Akun2 = Tampil_IDAkun("Account_Table", "id_akun", cmbNama_Akun2)
    'txtPosisi2 = Tampil_PosisiAkun("Account_Table", "posisi", cmbNama_Akun2)
End Sub

Private Sub Form_Load()
    Call OpenDB
    cmbPosisi1.AddItem "DEBIT"
    cmbPosisi1.AddItem "KREDIT"
    cmbPosisi2.AddItem "DEBIT"
    cmbPosisi2.AddItem "KREDIT"
    txtID_Trans.Enabled = False
    txtID_Akun.Enabled = False
    txtID_Akun2.Enabled = False
    dateTrans.Value = Format(Now)
    rsAcc.Open "select * from Account_Table", ConnectDB
    rsPer.Open "select * from Periode_Table", ConnectDB
    cmbNama_Akun.Clear
    cmbNama_Akun2.Clear
    cmbKriteria.Clear
    cmbPeriode.Clear
    Do While Not rsAcc.EOF
        cmbNama_Akun.AddItem rsAcc!nama_akun
        cmbNama_Akun2.AddItem rsAcc!nama_akun
        cmbKriteria.AddItem rsAcc!nama_akun
        rsAcc.MoveNext
    Loop
    Do While Not rsPer.EOF
        cmbPeriode.AddItem rsPer!masa
        rsPer.MoveNext
    Loop
    AN
    btnCancel.Visible = False
    Adodc1.ConnectionString = ConnectDB
    Adodc1.RecordSource = "Journal_Data"
    Adodc1.Refresh
    Set Journal_DataGrid.DataSource = Adodc1
End Sub

Private Sub Journal_DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    TAMPILKAN_KE_TEXTBOX
    Disabled_Text
    btnCancel.Visible = True
    btnSave.Enabled = True
    btnDelete.Enabled = True
End Sub
