Imports System.Data.OleDb
Module Lidh
    Dim path As String = My.Settings.ruajdtbpath & "tedhena.accdb;"
    Public lidhje21 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & My.Settings.dtb1 & ";")
    Public adaptori21 As OleDbDataAdapter
    Public lexuesi1 As OleDbDataReader
    Public query21 As OleDbCommand
    Public setitedhenave21 = New DataSet
    Public rreshtiaktual21 As Integer
    Public rreshtiaktual1 As Integer
End Module