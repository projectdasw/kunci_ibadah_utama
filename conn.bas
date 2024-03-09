Attribute VB_Name = "conn"
Public ConnectDB As New ADODB.Connection
Public rsLogin As New ADODB.Recordset
Public rsPel As New ADODB.Recordset
Public rsPeg As New ADODB.Recordset
Public rsBB As New ADODB.Recordset
Public rsJournal As New ADODB.Recordset
Public rsAcc As New ADODB.Recordset
Public rsPer As New ADODB.Recordset
Public rsProduct As New ADODB.Recordset
Public rsProduction As New ADODB.Recordset
Public rsPro As New ADODB.Recordset
Public rsPurchase As New ADODB.Recordset
Public rsItems As New ADODB.Recordset
Public rsSupp As New ADODB.Recordset
Public rsSales As New ADODB.Recordset
Public rsCus As New ADODB.Recordset
Public rsAccount As New ADODB.Recordset
Public rsPJ As New ADODB.Recordset
Public rsKalsoba As New ADODB.Recordset
Public rsKalsobi As New ADODB.Recordset
Public rsHPP As New ADODB.Recordset

Sub OpenDB()
    Set ConnectDB = New ADODB.Connection
    Set rsLogin = New ADODB.Recordset
    Set rsPel = New ADODB.Recordset
    Set rsPeg = New ADODB.Recordset
    Set rsBB = New ADODB.Recordset
    Set rsJournal = New ADODB.Recordset
    Set rsAcc = New ADODB.Recordset
    Set rsPer = New ADODB.Recordset
    Set rsProduct = New ADODB.Recordset
    Set rsProduction = New ADODB.Recordset
    Set rsPro = New ADODB.Recordset
    Set rsPurchase = New ADODB.Recordset
    Set rsItems = New ADODB.Recordset
    Set rsSupp = New ADODB.Recordset
    Set rsSales = New ADODB.Recordset
    Set rsCus = New ADODB.Recordset
    Set rsAccount = New ADODB.Recordset
    Set rsPJ = New ADODB.Recordset
    Set rsKalsoba = New ADODB.Recordset
    Set rsKalsobi = New ADODB.Recordset
    Set rsHPP = New ADODB.Recordset
    ConnectDB.Open "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & App.Path & "\KIU_Data2.mdb;"
End Sub
