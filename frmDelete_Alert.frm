VERSION 5.00
Begin VB.Form frmDelete_alert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image btnNo 
      Height          =   855
      Left            =   3720
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmDelete_Alert.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Image btnYes 
      Height          =   855
      Left            =   1800
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmDelete_Alert.frx":4E08
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   120
      Picture         =   "frmDelete_Alert.frx":912A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmDelete_alert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnNo_Click()
    Me.Hide
End Sub

Private Sub btnYes_Click()
    'Customer Data
    If frmCustomer.Visible = True Then
        frmDelete_alert.Hide
        frmCustomer.Adodc1.Recordset.Delete
        frmDelete_Complete.Show
        frmCustomer.txtID.Text = ""
        frmCustomer.txtNama.Text = ""
        frmCustomer.rtbAlamat.Text = ""
        frmCustomer.txtCP.Text = ""
        
    'Employee Data
    ElseIf frmEmployee.Visible = True Then
        frmDelete_alert.Hide
        frmEmployee.Adodc1.Recordset.Delete
        frmEmployee.txtID.Text = ""
        frmEmployee.txtNama.Text = ""
        frmEmployee.cmbJabatan.Text = ""
        frmEmployee.rtbAlamat.Text = ""
        frmDelete_Complete.Show
        
    'Items Data
    ElseIf frmItems.Visible = True Then
        frmDelete_alert.Hide
        frmItems.Adodc1.Recordset.Delete
        frmItems.txtID.Text = ""
        frmItems.txtNama.Text = ""
        frmItems.txtHarga.Text = ""
        frmItems.cmbSatuan.Text = ""
        frmItems.txtKuan.Text = ""
        frmDelete_Complete.Show
        
    'Journal Data
    ElseIf frmJournal.Visible = True Then
        frmDelete_alert.Hide
        frmJournal.Adodc1.Recordset.Delete
        frmJournal.txtID_Akun.Text = ""
        frmJournal.txtID_Akun2.Text = ""
        frmJournal.dateTrans.Value = Format(Now)
        frmJournal.cmbNama_Akun.Text = ""
        frmJournal.cmbNama_Akun2.Text = ""
        frmJournal.cmbPeriode.Text = ""
        frmJournal.txtNominal.Text = ""
        frmJournal.cmbPosisi1.Text = ""
        frmJournal.cmbPosisi2.Text = ""
        frmJournal.rtbKet.Text = ""
        frmJournal.btnCancel.Visible = False
        frmDelete_Complete.Show
    
    'Period Data
    ElseIf frmPeriod.Visible = True Then
        frmDelete_alert.Hide
        frmPeriod.Adodc1.Recordset.Delete
        frmPeriod.txtID.Text = ""
        frmPeriod.dateAwal.Value = Format(Now)
        frmPeriod.dateAkhir.Value = Format(Now)
        frmPeriod.txtMasa.Text = ""
        frmDelete_Complete.Show
        
    'Product Data
    ElseIf frmProduct.Visible = True Then
        frmDelete_alert.Hide
        frmProduct.Adodc1.Recordset.Delete
        frmProduct.txtID.Text = ""
        frmProduct.txtNama.Text = ""
        frmProduct.txtHPP.Text = ""
        frmProduct.cmbSatuan.Text = ""
        frmProduct.txtKuan.Text = ""
        frmDelete_Complete.Show
        
    'Production Data
    ElseIf frmProduction.Visible = True Then
        frmDelete_alert.Hide
        frmProduction.Adodc1.Recordset.Delete
        frmProduction.ClearText
        frmProduction.btnCancel.Visible = False
        frmDelete_Complete.Show
        
    'Supplier Data
    ElseIf frmSupplier.Visible = True Then
        frmDelete_alert.Hide
        frmSupplier.Adodc1.Recordset.Delete
        frmSupplier.txtID.Text = ""
        frmSupplier.txtNama.Text = ""
        frmSupplier.rtbAlamat.Text = ""
        frmSupplier.txtTelp.Text = ""
        frmDelete_Complete.Show
    
    'Account Data
    ElseIf frmAccount.Visible = True Then
        frmDelete_alert.Hide
        frmAccount.Adodc1.Recordset.Delete
        frmAccount.txtID.Text = ""
        frmAccount.txtNama.Text = ""
        frmAccount.cmbKategori.Text = ""
        frmAccount.cmbDK.Text = ""
        frmAccount.cmbPeriode.Text = ""
        frmAccount.cmbKelompok.Text = ""
        frmAccount.txtSaldo.Text = ""
        frmDelete_Complete.Show
    End If
End Sub

Private Sub Form_Load()
    frmDelete_alert.BackColor = RGB(195, 175, 166)
End Sub
