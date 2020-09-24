Imports System.Data.OleDb
Public Class Statistika
    Dim path As String = My.Settings.ruajdtbpath & "\tedhena.accdb;"
    Public lidhjestatistika As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & My.Settings.dtb1 & ";")
    Public adaptoristatistika As OleDbDataAdapter
    Public lexuesistatistika As OleDbDataReader
    Public querystatistika As OleDbCommand
    Public setitedhenavestatistika = New DataSet
    Private _form_resize As clsResize
    Public Sub New()
        InitializeComponent()
        _form_resize = New clsResize(Me)
        AddHandler Me.Load, AddressOf _Load
        AddHandler Me.Resize, AddressOf _Resize
    End Sub
    Private Sub _Resize(ByVal sender As Object, ByVal e As EventArgs)
        _form_resize._resize1()
    End Sub
    Private Sub _Load(ByVal sender As Object, ByVal e As EventArgs)
        _form_resize._get_initial_size1()
    End Sub
    Private Sub Statistika_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        Farmacia.Show()
    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        TextBox1.Text = ""
        If RadioButton1.Checked = True Then
            RadioButton1.ForeColor = Color.ForestGreen
            With DataGridView1
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            DataGridView1.DataSource = Nothing
            setitedhenavestatistika.Tables.Clear()
            lidhjestatistika.Open()
            adaptoristatistika = New OleDbDataAdapter("SELECT Bleresi,COUNT(Bleresi) As Nr_Produkteve_te_Blera,SUM(Vlera_Pa_TVSH) As Totali_Shpenzuar FROM Shitjet GROUP BY Bleresi HAVING COUNT(*) > 0 ORDER BY SUM(Vlera_Pa_TVSH) DESC", lidhjestatistika)
            DataGridView1.Columns.Clear()
            DataGridView1.Columns.Add("count1", "Nr.")
            adaptoristatistika.Fill(setitedhenavestatistika, "tedhena")
            DataGridView1.DataSource = setitedhenavestatistika.Tables(0)
            lidhjestatistika.Close()
            Dim nr1 As Integer
            nr1 = DataGridView1.Rows.Count - 1
            For i = 0 To nr1
                DataGridView1.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView1.Columns(3).ValueType = GetType(Double)
            DataGridView1.Columns(3).DefaultCellStyle.Format = "N0"
        Else
            setitedhenavestatistika.Tables.Clear()
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            RadioButton1.ForeColor = DefaultForeColor
        End If
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        TextBox1.Text = ""
        If RadioButton2.Checked = True Then
            RadioButton2.ForeColor = Color.ForestGreen
            With DataGridView1
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            DataGridView1.DataSource = Nothing
            setitedhenavestatistika.Tables.Clear()
            lidhjestatistika.Open()
            adaptoristatistika = New OleDbDataAdapter("SELECT Produkti,COUNT(Produkti) As Sasia_Shitjeve,SUM(Vlera_Pa_TVSH) As Totali_Fituar_Pa_TVSH,SUM(Sasia) As Totali_Shitur FROM Shitjet GROUP BY Produkti HAVING COUNT(*) > 0 ORDER BY SUM(Vlera_Pa_TVSH) DESC", lidhjestatistika)
            DataGridView1.Columns.Clear()
            DataGridView1.Columns.Add("count1", "Nr.")
            adaptoristatistika.Fill(setitedhenavestatistika, "tedhena")
            DataGridView1.DataSource = setitedhenavestatistika.Tables(0)
            lidhjestatistika.Close()
            Dim nr1 As Integer
            nr1 = DataGridView1.Rows.Count - 1
            For i = 0 To nr1
                DataGridView1.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView1.Columns(3).ValueType = GetType(Double)
            DataGridView1.Columns(3).DefaultCellStyle.Format = "N0"
        Else
            setitedhenavestatistika.Tables.Clear()
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            RadioButton2.ForeColor = DefaultForeColor
        End If
    End Sub
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        TextBox1.Text = ""
        If RadioButton3.Checked = True Then
            RadioButton3.ForeColor = Color.ForestGreen
            With DataGridView1
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            DataGridView1.DataSource = Nothing
            setitedhenavestatistika.Tables.Clear()
            lidhjestatistika.Open()
            adaptoristatistika = New OleDbDataAdapter("SELECT Bleresi,COUNT(Zbritje_ne_perq) AS Sasia_Skontove,Zbritje_ne_perq FROM Shitjet " _
            & "WHERE Zbritje_ne_perq > '0' GROUP BY Bleresi,Zbritje_ne_perq " _
            & "ORDER BY Bleresi,COUNT(Zbritje_ne_perq) DESC;", lidhjestatistika)
            DataGridView1.Columns.Clear()
            DataGridView1.Columns.Add("count1", "Nr.")
            adaptoristatistika.Fill(setitedhenavestatistika, "tedhena")
            DataGridView1.DataSource = setitedhenavestatistika.Tables(0)
            lidhjestatistika.Close()
            Dim nr1 As Integer
            nr1 = DataGridView1.Rows.Count - 1
            For i = 0 To nr1
                DataGridView1.Rows(i).Cells(0).Value = i + 1
            Next
        Else
            setitedhenavestatistika.Tables.Clear()
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            RadioButton3.ForeColor = DefaultForeColor
        End If
    End Sub
    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        TextBox1.Text = ""
        If RadioButton4.Checked = True Then
            RadioButton4.ForeColor = Color.ForestGreen
            With DataGridView1
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            DataGridView1.DataSource = Nothing
            setitedhenavestatistika.Tables.Clear()
            lidhjestatistika.Open()
            adaptoristatistika = New OleDbDataAdapter("SELECT Bleresi,COUNT(Zbritje_ne_perq) AS Sasia_Skontove FROM Shitjet " _
            & "WHERE Zbritje_ne_perq > '0' GROUP BY Bleresi " _
            & "ORDER BY COUNT(Zbritje_ne_perq) DESC;", lidhjestatistika)
            DataGridView1.Columns.Clear()
            DataGridView1.Columns.Add("count1", "Nr.")
            adaptoristatistika.Fill(setitedhenavestatistika, "tedhena")
            DataGridView1.DataSource = setitedhenavestatistika.Tables(0)
            lidhjestatistika.Close()
            Dim nr1 As Integer
            nr1 = DataGridView1.Rows.Count - 1
            For i = 0 To nr1
                DataGridView1.Rows(i).Cells(0).Value = i + 1
            Next
        Else
            setitedhenavestatistika.Tables.Clear()
            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = Nothing
            DataGridView1.Refresh()
            RadioButton4.ForeColor = DefaultForeColor
        End If
    End Sub
    Public lidhjesearch As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & My.Settings.dtb1 & ";")
    Public adaptorisearch As OleDbDataAdapter
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If RadioButton1.Checked = True Then
            lidhjesearch.Close()
            lidhjesearch.Open()
            adaptorisearch = New OleDbDataAdapter("SELECT Bleresi,COUNT(Bleresi) As Nr_Produkteve_te_Blera,SUM(Vlera_Pa_TVSH) As Totali_Shpenzuar FROM Shitjet WHERE(Bleresi Like '" & TextBox1.Text & "%') GROUP BY Bleresi HAVING COUNT(*) > 0 ORDER BY SUM(Vlera_Pa_TVSH) DESC", lidhjesearch)
            Dim dataTables As New DataTable
            adaptorisearch.Fill(dataTables)
            DataGridView1.DataSource = dataTables
            Dim nr1 As Integer
            nr1 = DataGridView1.Rows.Count - 1
            For i = 0 To nr1
                DataGridView1.Rows(i).Cells(0).Value = i + 1
            Next
            lidhjesearch.Close()
            DataGridView1.Columns(3).ValueType = GetType(Double)
            DataGridView1.Columns(3).DefaultCellStyle.Format = "N0"
        ElseIf RadioButton2.Checked = True Then
            lidhjesearch.Close()
            lidhjesearch.Open()
            adaptorisearch = New OleDbDataAdapter("SELECT Produkti,COUNT(Produkti) As Sasia_Shitjeve,SUM(Vlera_Pa_TVSH) As Totali_Fituar_Pa_TVSH,SUM(Sasia) As Totali_Shitur FROM Shitjet WHERE(Produkti Like '" & TextBox1.Text & "%') GROUP BY Produkti HAVING COUNT(*) > 0 ORDER BY SUM(Vlera_Pa_TVSH) DESC", lidhjesearch)
            Dim dataTables As New DataTable
            adaptorisearch.Fill(dataTables)
            DataGridView1.DataSource = dataTables
            Dim nr1 As Integer
            nr1 = DataGridView1.Rows.Count - 1
            For i = 0 To nr1
                DataGridView1.Rows(i).Cells(0).Value = i + 1
            Next
            lidhjesearch.Close()
            DataGridView1.Columns(3).ValueType = GetType(Double)
            DataGridView1.Columns(3).DefaultCellStyle.Format = "N0"
        ElseIf RadioButton3.Checked = True Then
            lidhjesearch.Close()
            lidhjesearch.Open()
            adaptorisearch = New OleDbDataAdapter("SELECT Bleresi,COUNT(Zbritje_ne_perq) AS Sasia_Skontove,Zbritje_ne_perq FROM Shitjet WHERE(Bleresi Like '" & TextBox1.Text & "%') " _
            & "WHERE Zbritje_ne_perq > '0' GROUP BY Bleresi,Zbritje_ne_perq " _
            & "ORDER BY Bleresi,COUNT(Zbritje_ne_perq) DESC;", lidhjesearch)
            Dim dataTables As New DataTable
            adaptorisearch.Fill(dataTables)
            DataGridView1.DataSource = dataTables
            Dim nr1 As Integer
            nr1 = DataGridView1.Rows.Count - 1
            For i = 0 To nr1
                DataGridView1.Rows(i).Cells(0).Value = i + 1
            Next
            lidhjesearch.Close()
        ElseIf RadioButton4.Checked = True Then
            lidhjesearch.Close()
            lidhjesearch.Open()
            adaptorisearch = New OleDbDataAdapter("SELECT Bleresi,COUNT(Zbritje_ne_perq) AS Sasia_Skontove FROM Shitjet " _
                & "WHERE( Zbritje_ne_perq > '0' AND Bleresi Like '" & TextBox1.Text & "%') GROUP BY Bleresi " _
                & "ORDER BY COUNT(Zbritje_ne_perq) DESC;", lidhjesearch)
            Dim dataTables As New DataTable
            adaptorisearch.Fill(dataTables)
            DataGridView1.DataSource = dataTables
            Dim nr1 As Integer
            nr1 = DataGridView1.Rows.Count - 1
            For i = 0 To nr1
                DataGridView1.Rows(i).Cells(0).Value = i + 1
            Next
            lidhjesearch.Close()
        End If
    End Sub


End Class