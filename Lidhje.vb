Imports System.Data.OleDb
Module Lidhje
    Dim path As String = My.Settings.ruajdtbpath & "\tedhena.accdb;"
    Public lidhje2 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptori2 As OleDbDataAdapter
    Public lexuesi As OleDbDataReader
    Public query2 As OleDbCommand
    Public setitedhenave2 = New DataSet
    Public rreshtiaktual2 As Integer
    Public rreshtiaktual As Integer
    Public lidhje2y As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptori2y As OleDbDataAdapter
    Public lexuesiy As OleDbDataReader
    Public query2y As OleDbCommand
    Public setitedhenave2y = New DataSet
    Public rreshtiaktual2y As Integer
    Public lidhje2ypp As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptori2ypp As OleDbDataAdapter
    Public lexuesiypp As OleDbDataReader
    Public query2ypp As OleDbCommand
    Public setitedhenave2ypp = New DataSet
    Public rreshtiaktual2shitjet As Integer
    Public lidhje2yx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptori2yx As OleDbDataAdapter
    Public lexuesiyx As OleDbDataReader
    Public query2yx As OleDbCommand
    Public setitedhenave2yxc = New DataSet
    Public rreshtiaktual2yxc As Integer
    Public lidhje2yxy As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptori2yxy As OleDbDataAdapter
    Public lexuesiyxy As OleDbDataReader
    Public query2yxy As OleDbCommand
    Public setitedhenave2yxy = New DataSet
    Public rreshtiaktual2yxy As Integer
    Public lidhje2yxy1 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptori2yxy1 As OleDbDataAdapter
    Public lexuesiyxy1 As OleDbDataReader
    Public query2yxy1 As OleDbCommand
    Public setitedhenave2yxy1 = New DataSet
    Public rreshtiaktual2yxy1 As Integer
    Public lidhje2v As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptori2v As OleDbDataAdapter
    Public lexuesiv As OleDbDataReader
    Public query2v As OleDbCommand
    Public setitedhenave2v = New DataSet
    Public rreshtiaktual2v As Integer
    Public rreshtiaktualv As Integer
    Public lidhjedg4 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptoridg4 As OleDbDataAdapter
    Public lexuesidg4 As OleDbDataReader
    Public querydg4 As OleDbCommand
    Public setitedhenavedg4 = New DataSet
    Public rreshtiaktualdg4 As Integer
    Public lidhjedg3 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptoridg3 As OleDbDataAdapter
    Public lexuesidg3 As OleDbDataReader
    Public querydg3 As OleDbCommand
    Public setitedhenavedg3 = New DataSet
    Public rreshtiaktualdg3 As Integer
    Public lidhjedg31 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptoridg31 As OleDbDataAdapter
    Public lexuesidg31 As OleDbDataReader
    Public querydg31 As OleDbCommand
    Public setitedhenavedg31 = New DataSet
    Public rreshtiaktualdg31 As Integer
    Public lidhjedg41 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptoridg41 As OleDbDataAdapter
    Public lexuesidg41 As OleDbDataReader
    Public querydg41 As OleDbCommand
    Public setitedhenavedg41 = New DataSet
    Public rreshtiaktualdg41 As Integer
    Public lidhjedg71 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptoridg71 As OleDbDataAdapter
    Public lexuesidg71 As OleDbDataReader
    Public querydg71 As OleDbCommand
    Public setitedhenavedg71 = New DataSet
    Public rreshtiaktualdg71 As Integer
    Public lidhjedg71ofert As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptoridg71ofert As OleDbDataAdapter
    Public lexuesidg71ofert As OleDbDataReader
    Public querydg71ofert As OleDbCommand
    Public setitedhenavedg71ofert = New DataSet
    Public rreshtiaktualdg71ofert As Integer
    Public lidhjedg71ofert1 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptoridg71ofert1 As OleDbDataAdapter
    Public lexuesidg71ofert1 As OleDbDataReader
    Public querydg71ofert1 As OleDbCommand
    Public setitedhenavedg71ofert1 = New DataSet
    Public rreshtiaktualdg71ofert1 As Integer
    Public lidhjedg71ofert12 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptoridg71ofert12 As OleDbDataAdapter
    Public lexuesidg71ofert12 As OleDbDataReader
    Public querydg71ofert12 As OleDbCommand
    Public setitedhenavedg71ofert12 = New DataSet
    Public rreshtiaktualdg71ofert12 As Integer
    Public lidhjedg71ofert121 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptoridg71ofert121 As OleDbDataAdapter
    Public lexuesidg71ofert121 As OleDbDataReader
    Public querydg71ofert121 As OleDbCommand
    Public setitedhenavedg71ofert121 = New DataSet
    Public rreshtiaktualdg71ofert121 As Integer
    Public lidhjedg71ofert121klasat As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptoridg71ofert121klasat As OleDbDataAdapter
    Public lexuesidg71ofert121klasat As OleDbDataReader
    Public querydg71ofert121klasat As OleDbCommand
    Public setitedhenavedg71ofert121klasat = New DataSet
    Public rreshtiaktualdg71ofert121klasat As Integer
    Public lidhjedg71ofert121klasatkerk As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptoridg71ofert121klasatkerk As OleDbDataAdapter
    Public lexuesidg71ofert121klasatkerk As OleDbDataReader
    Public querydg71ofert121klasatkerk As OleDbCommand
    Public setitedhenavedg71ofert121klasatkerk = New DataSet
    Public rreshtiaktualdg71ofert121klasatkerk As Integer
    Public lidhjefshijklasat As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptorifshijklasat As OleDbDataAdapter
    Public lexuesifshijklasat As OleDbDataReader
    Public queryfshijklasat As OleDbCommand
    Public setitedhenavefshijklasat = New DataSet
    Public rreshtiaktualfshijklasat As Integer
    Public lidhje2autofill As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptori2autofill As OleDbDataAdapter
    Public query2autofill As OleDbCommand
    Public setitedhenave2autofill = New DataSet
    Public lidhjekompania As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptorikompania As OleDbDataAdapter
    Public lexuesifshijkompania As OleDbDataReader
    Public queryfshijkompania As OleDbCommand
    Public setitedhenavefshijkompania = New DataSet
    Public rreshtiaktualfshijkompania As Integer
    Public lidhjekompaniax As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptorikompaniax As OleDbDataAdapter
    Public lexuesifshijkompaniax As OleDbDataReader
    Public queryfshijkompaniax As OleDbCommand
    Public setitedhenavefshijkompaniax = New DataSet
    Public rreshtiaktualfshijkompaniax As Integer
    Public lidhjekompaniaxy As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptorikompaniaxy As OleDbDataAdapter
    Public lexuesifshijkompaniaxy As OleDbDataReader
    Public queryfshijkompaniaxy As OleDbCommand
    Public setitedhenavefshijkompaniaxy = New DataSet
    Public rreshtiaktualfshijkompaniaxy As Integer
End Module