Attribute VB_Name = "functions"
Function Buka_Produk(KD)
  Q = "SELECT Product_Table.* FROM Product_Table WHERE id_produk= '" & KD & "'"
    If rsProduct.State = adStateOpen Then
       rsProduct.Close
    End If
  rsProduct.Open Q, ConnectDB, adOpenDynamic, adLockOptimistic
End Function

Function Buka_Jurnal(KD)
  Q = "SELECT Account_Table.* FROM Account_Table WHERE id_akun= '" & KD & "'"
    If rsJournal.State = adStateOpen Then
       rsJournal.Close
    End If
  rsJournal.Open Q, ConnectDB, adOpenDynamic, adLockOptimistic
End Function

Function Buka_TBahan(KD)
  Q = "SELECT Items_Table.* FROM Items_Table WHERE id_bahan= '" & KD & "'"
    If rsItems.State = adStateOpen Then
       rsItems.Close
    End If
  rsItems.Open Q, ConnectDB, adOpenDynamic, adLockOptimistic
End Function

Function Buka_HPP(KD)
  Q = "SELECT HPP_Production.* FROM HPP_Production WHERE id_produk= '" & KD & "'"
    If rsHPP.State = adStateOpen Then
       rsHPP.Close
    End If
  rsHPP.Open Q, ConnectDB, adOpenDynamic, adLockOptimistic
End Function

Function Buka_Kalsoba(KD)
  Q = "SELECT Kalkulasi_SongkokBagus.* FROM Kalkulasi_SongkokBagus WHERE id_kal= '" & KD & "'"
    If rsKalsoba.State = adStateOpen Then
       rsKalsoba.Close
    End If
  rsKalsoba.Open Q, ConnectDB, adOpenDynamic, adLockOptimistic
End Function

Function Buka_Kalsobi(KD)
  Q = "SELECT Kalkulasi_SongkokBiasa.* FROM Kalkulasi_SongkokBiasa WHERE id_kal= '" & KD & "'"
    If rsKalsobi.State = adStateOpen Then
       rsKalsobi.Close
    End If
  rsKalsobi.Open Q, ConnectDB, adOpenDynamic, adLockOptimistic
End Function

Function Ubah_Jurnal(kp, ku)
  Q = "update Account_Table set saldo_akhir = " & ku & " where id_akun='" & kp & "'"
  ConnectDB.Execute Q
End Function

Function Ubah_TBahan(kp, ku)
  Q = "update Items_Table set kuantitas = " & ku & " where id_bahan='" & kp & "'"
  ConnectDB.Execute Q
End Function

Function Ubah_TotalBahan(kp, ku)
  Q = "update Items_Table set total_bahan = " & ku & " where id_bahan='" & kp & "'"
  ConnectDB.Execute Q
End Function

Function Ubah_TotalKalsoba(kp, ku)
  Q = "update Kalkulasi_SongkokBagus set harga = " & ku & " where id_kal='" & kp & "'"
  ConnectDB.Execute Q
End Function

Function Ubah_TotalKalsobi(kp, ku)
  Q = "update Kalkulasi_SongkokBiasa set harga = " & ku & " where id_kal='" & kp & "'"
  ConnectDB.Execute Q
End Function

Function Ubah_TBKalso(TKalso, kp, ku)
  Q = "update " & TKalso & " set total = " & ku & " where id_kal='" & kp & "'"
  ConnectDB.Execute Q
End Function

Function Ubah_HargaBahan(kp, ku)
  Q = "update Items_Table set harga_bahan = " & ku & " where id_bahan='" & kp & "'"
  ConnectDB.Execute Q
End Function

Function Ubah_KuanProduk(kp, ku)
  Q = "update Product_Table set jumlah_unit = " & ku & " where id_produk='" & kp & "'"
  ConnectDB.Execute Q
End Function

Function Ubah_HPPProduk(kp, ku)
  Q = "update HPP_Production set hpp_produk = " & ku & " where id_produk='" & kp & "'"
  ConnectDB.Execute Q
End Function

Function Tampil_IDAkun(Tabel, Field, KD)
    Dim rs As New ADODB.Recordset
    Q = "select " & Field & " from " & Tabel & " where nama_akun='" & KD & "'"
    rs.Open Q, ConnectDB
    Tampil_IDAkun = rs.Fields(Field)
    rs.Close
End Function

Function Tampil_Journal(Tabel, Field, KD)
    Dim rs As New ADODB.Recordset
    Q = "select " & Field & " from " & Tabel & " where nama_akun='" & KD & "'"
    rs.Open Q, ConnectDB
    Tampil_Journal = rs.Fields(Field)
    rs.Close
End Function

Function Tampil_PosisiAkun(Tabel, Field, KD)
    Dim rs As New ADODB.Recordset
    Q = "select " & Field & " from " & Tabel & " where nama_akun='" & KD & "'"
    rs.Open Q, ConnectDB
    Tampil_PosisiAkun = rs.Fields(Field)
    rs.Close
End Function

Function Tampil_IDPeriode(Tabel, Field, KD)
    Dim rs As New ADODB.Recordset
    Q = "select " & Field & " from " & Tabel & " where masa='" & KD & "'"
    rs.Open Q, ConnectDB
    Tampil_IDPeriode = rs.Fields(Field)
    rs.Close
End Function

Function Tampil_PeriodeJurnal(Tabel, Field, KD)
    Dim rs As New ADODB.Recordset
    Q = "select " & Field & " from " & Tabel & " where id_trans='" & KD & "'"
    rs.Open Q, ConnectDB
    Tampil_PeriodeJurnal = rs.Fields(Field)
    rs.Close
End Function

Function Tampil_NominalJurnal(Tabel, Field, KD)
    Dim rs As New ADODB.Recordset
    Q = "select " & Field & " from " & Tabel & " where id_trans='" & KD & "'"
    rs.Open Q, ConnectDB
    Tampil_NominalJurnal = rs.Fields(Field)
    rs.Close
End Function

Function Tampil_IDProduk(nmTabel, nmField, KD)
    Dim rs As New ADODB.Recordset
    Q = "select " & nmField & " from " & nmTabel & " where nama_produk='" & KD & "'"
    rs.Open Q, ConnectDB
    Tampil_IDProduk = rs.Fields(nmField)
    rs.Close
End Function

Function Tampil_HPPProduk(nmTabel, nmField, KD)
    Dim rs As New ADODB.Recordset
    Q = "select " & nmField & " from " & nmTabel & " where nama_produk='" & KD & "'"
    rs.Open Q, ConnectDB
    Tampil_HPPProduk = rs.Fields(nmField)
    rs.Close
End Function

Function Tampil_IDBahan(nmTabel, nmField, KD)
    Dim rs As New ADODB.Recordset
    Q = "select " & nmField & " from " & nmTabel & " where nama_bahan='" & KD & "'"
    rs.Open Q, ConnectDB
    Tampil_IDBahan = rs.Fields(nmField)
    rs.Close
End Function

Function Tampil_IDSupplier(Tabel, Field, KD)
    Dim rs As New ADODB.Recordset
    Q = "select " & Field & " from " & Tabel & " where nama_supplier='" & KD & "'"
    rs.Open Q, ConnectDB
    Tampil_IDSupplier = rs.Fields(Field)
    rs.Close
End Function

Function Tampil_IDPelanggan(Tabel, Field, KD)
    Dim rs As New ADODB.Recordset
    Q = "select " & Field & " from " & Tabel & " where nama_customer='" & KD & "'"
    rs.Open Q, ConnectDB
    Tampil_IDPelanggan = rs.Fields(Field)
    rs.Close
End Function
