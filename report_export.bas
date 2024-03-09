Attribute VB_Name = "Report_Export"
Function Export_Excel(file, nama)
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim i As Integer
    Dim j As Integer
    
    Call OpenDB
    Set oExcel = CreateObject("Excel.application")
    Set oBook = oExcel.Workbooks.Open(App.Path & "\" & file)
    Set oSheet = oBook.Worksheets(nama)
    oExcel.Visible = False
    
    'Supplier Table
    If frmAccount.Visible = True Then
        rsAccount.Open "select * from Account_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("E5").CopyFromRecordset rsAccount
        rsAccount.Close
    
    'Customer Table
    ElseIf frmCustomer.Visible = True Then
        rsPel.Open "select * from Customer_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("E5").CopyFromRecordset rsPel
        rsPel.Close
    
    'Employee Table
    ElseIf frmEmployee.Visible = True Then
        rsPeg.Open "select * from Employee_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("E5").CopyFromRecordset rsPeg
        rsPeg.Close
        
    'Items Table
    ElseIf frmItems.Visible = True Then
        rsBB.Open "select * from Items_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("E5").CopyFromRecordset rsBB
        rsBB.Close
    
    'Journal Table
    ElseIf frmJournal.Visible = True Then
        rsJournal.Open "select * from Journal_Data", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("J7").CopyFromRecordset rsJournal
        rsJournal.Close
        
    'Periode Table
    ElseIf frmPeriod.Visible = True Then
        rsPer.Open "select * from Periode_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("E5").CopyFromRecordset rsPer
        rsPer.Close
    
    'Product Table
    ElseIf frmProduct.Visible = True Then
        rsProduct.Open "select * from Product_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("E5").CopyFromRecordset rsProduct
        rsProduct.Close
    
    'Production Table
    ElseIf frmProduction.Visible = True Then
        rsProduction.Open "select * from Production_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("E5").CopyFromRecordset rsProduction
        rsProduction.Close
    
    'Purchase Table
    ElseIf frmPurchase.Visible = True Then
        rsPurchase.Open "select * from Purchase_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("A5").CopyFromRecordset rsPurchase
        rsPurchase.Close
        
        rsPurchase.Open "select * from PurchaseData_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("N5").CopyFromRecordset rsPurchase
        rsPurchase.Close
    
    'Sales Table
    ElseIf frmSales.Visible = True Then
        rsSales.Open "select * from Sales_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("A5").CopyFromRecordset rsSales
        rsSales.Close
        
        rsSales.Open "select * from SalesData_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("N5").CopyFromRecordset rsSales
        rsSales.Close
        
    'Supplier Table
    ElseIf frmSupplier.Visible = True Then
        rsSupp.Open "select * from Supplier_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
        'oSheet.Range("A1:Z100000").ClearContents
        oSheet.Range("E5").CopyFromRecordset rsSupp
        rsSupp.Close
    End If
    
    oSheet.Select
    oBook.Save
    oExcel.Visible = True
    Set oExcel = Nothing
End Function

Function Report_Neraca(file, nama, Tabel, Field)
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim i As Integer
    Dim j As Integer
    
    Call OpenDB
    Set oExcel = CreateObject("Excel.application")
    Set oBook = oExcel.Workbooks.Open(App.Path & "\" & file)
    Set oSheet = oBook.Worksheets(nama)
    oExcel.Visible = False
    oSheet.Select
    oBook.Save
    oExcel.Visible = True
    Set oExcel = Nothing
End Function

Function Report_LR(file, nama, Tabel, Field)
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim i As Integer
    Dim j As Integer
    
    Call OpenDB
    Set oExcel = CreateObject("Excel.application")
    Set oBook = oExcel.Workbooks.Open(App.Path & "\" & file)
    Set oSheet = oBook.Worksheets(nama)
    oExcel.Visible = False
    oSheet.Select
    oBook.Save
    oExcel.Visible = True
    Set oExcel = Nothing
End Function
        
Function Report_Jurnal(file, nama)
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim i As Integer
    Dim j As Integer
    
    Call OpenDB
    Set oExcel = CreateObject("Excel.application")
    Set oBook = oExcel.Workbooks.Open(App.Path & "\" & file)
    Set oSheet = oBook.Worksheets(nama)
    oExcel.Visible = False
    
    rsPJ.Open "select * from Journal_Data", ConnectDB, adOpenDynamic, adLockOptimistic
    
    'oSheet.Range("A1:Z100000").ClearContents
    oSheet.Range("J7").CopyFromRecordset rsPJ
    rsPJ.Close
    
    oSheet.Select
    oBook.Save
    oExcel.Visible = True
    Set oExcel = Nothing
End Function

Function Report_JurnalBahanBaku(file, nama)
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim i As Integer
    Dim j As Integer
    
    Call OpenDB
    Set oExcel = CreateObject("Excel.application")
    Set oBook = oExcel.Workbooks.Open(App.Path & "\" & file)
    Set oSheet = oBook.Worksheets(nama)
    oExcel.Visible = False
    
    rsItems.Open "select * from ItemsData_Table", ConnectDB, adOpenDynamic, adLockOptimistic
    
    'oSheet.Range("A1:Z100000").ClearContents
    oSheet.Range("J7").CopyFromRecordset rsItems
    rsItems.Close
    
    oSheet.Select
    oBook.Save
    oExcel.Visible = True
    Set oExcel = Nothing
End Function

Function Report_BBHPP(file, nama)
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim i As Integer
    Dim j As Integer
    
    Call OpenDB
    Set oExcel = CreateObject("Excel.application")
    Set oBook = oExcel.Workbooks.Open(App.Path & "\" & file)
    Set oSheet = oBook.Worksheets(nama)
    oExcel.Visible = False
    oSheet.Select
    oBook.Save
    oExcel.Visible = True
    Set oExcel = Nothing
End Function

Function Report_Akun(file, nama)
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim i As Integer
    Dim j As Integer
    
    Call OpenDB
    Set oExcel = CreateObject("Excel.application")
    Set oBook = oExcel.Workbooks.Open(App.Path & "\" & file)
    Set oSheet = oBook.Worksheets(nama)
    oExcel.Visible = False
    oSheet.Select
    oBook.Save
    oExcel.Visible = True
    Set oExcel = Nothing
End Function

Function Report_TutupBuku(file, nama)
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim i As Integer
    Dim j As Integer
    
    Call OpenDB
    Set oExcel = CreateObject("Excel.application")
    Set oBook = oExcel.Workbooks.Open(App.Path & "\" & file)
    Set oSheet = oBook.Worksheets(nama)
    oExcel.Visible = False
    oSheet.Select
    oBook.Save
    oExcel.Visible = True
    Set oExcel = Nothing
End Function
