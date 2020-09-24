Imports System.Data.OleDb
Imports System.Globalization
Imports System.Linq
Public Class Bli
    Dim path As String = My.Settings.ruajdtbpath & "\tedhena.accdb;"
    Private _form_resize As clsResize
    Public Sub New()
        InitializeComponent()
        _form_resize = New clsResize(Me)
        AddHandler Me.Load, AddressOf _Load
        AddHandler Me.Resize, AddressOf _Resize
    End Sub
    Private Sub _Resize(ByVal sender As Object, ByVal e As EventArgs)
        _form_resize._resize1()
        If Me.WindowState = FormWindowState.Maximized Then
            Label35.Text = "_________________________________________________________"
            Label36.Text = "_________________________________________________________"
        ElseIf Me.WindowState = FormWindowState.Normal Then
            Label35.Text = "____________________________________________"
            Label36.Text = "____________________________________________"
        End If
    End Sub
    Private Sub _Load(ByVal sender As Object, ByVal e As EventArgs)
        _form_resize._get_initial_size()
    End Sub
    Private Sub ComboBox1_DrawItem(ByVal sender As Object,
    ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox1.DrawItem
        If e.Index < 0 Then Exit Sub
        Dim rect As Rectangle = e.Bounds
        If e.State And DrawItemState.Selected Then
            e.Graphics.FillRectangle(Brushes.LightGreen, rect) 'change the selected color you like
        Else
            e.Graphics.FillRectangle(SystemBrushes.Window, rect)
        End If
        Dim colorname As String = ComboBox1.Items(e.Index)
        Dim b As New SolidBrush(Color.FromName(colorname))
        Dim b2 As Brush = Brushes.Black 'add one
        e.Graphics.DrawString(colorname, Me.ComboBox1.Font, b2, rect.X, rect.Y)
    End Sub
    Private Sub ComboBox2_DrawItem(ByVal sender As Object,
    ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox2.DrawItem
        If e.Index < 0 Then Exit Sub
        Dim rect As Rectangle = e.Bounds
        If e.State And DrawItemState.Selected Then
            e.Graphics.FillRectangle(Brushes.LightGreen, rect) 'change the selected color you like
        Else
            e.Graphics.FillRectangle(SystemBrushes.Window, rect)
        End If
        Dim colorname As String = ComboBox2.Items(e.Index)
        Dim b As New SolidBrush(Color.FromName(colorname))
        Dim b2 As Brush = Brushes.Black 'add one
        e.Graphics.DrawString(colorname, Me.ComboBox2.Font, b2, rect.X, rect.Y)
    End Sub
    Private Sub ComboBox3_DrawItem(ByVal sender As Object,
    ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox3.DrawItem
        If e.Index < 0 Then Exit Sub
        Dim rect As Rectangle = e.Bounds
        If e.State And DrawItemState.Selected Then
            e.Graphics.FillRectangle(Brushes.LightGreen, rect) 'change the selected color you like
        Else
            e.Graphics.FillRectangle(SystemBrushes.Window, rect)
        End If
        Dim colorname As String = ComboBox3.Items(e.Index)
        Dim b As New SolidBrush(Color.FromName(colorname))
        Dim b2 As Brush = Brushes.Black 'add one
        e.Graphics.DrawString(colorname, Me.ComboBox3.Font, b2, rect.X, rect.Y)
    End Sub
    Public Sub CreateReader_load_produktet_Blerje(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                ComboBox1.Items.Add(reader(1).ToString())
            End While
            reader.Close()
            If ComboBox1.Text = "" Then

            Else
                ComboBox1.SelectedIndex = 0
            End If
        End Using
    End Sub
    Public Sub CreateReader_load_fornitoret_Blerje(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                ComboBox2.Items.Add(reader(1).ToString())
            End While
            reader.Close()
            connection.Close()
            If ComboBox2.Text = "" Then

            Else
                ComboBox2.SelectedIndex = 0
            End If
        End Using
    End Sub
    Public Function GenerateRandomString(ByRef iLength As Integer) As String
        Dim rdm As New Random()
        Dim allowChrs() As Char = "0123456789".ToCharArray()
        Dim sResult As String = ""
        For i As Integer = 0 To iLength - 1
            sResult += allowChrs(rdm.Next(0, allowChrs.Length))
        Next
        Return sResult
    End Function
    Private Sub Bli_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Farmacia.Show()
        My.Settings.tvsh = NumericUpDown2.Value
    End Sub
    Private Sub Bli_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        NumericUpDown2.Value = My.Settings.tvsh
        If DataGridView1.Rows.Count = 0 Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If
        With DataGridView1
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        Timer1.Start()
        Dim regDate As DateTime = Date.Now
        Dim strDate As String = regDate.ToString("dd/MM/yyyy")
        Label25.Text = strDate
        TextBox2.Text = GenerateRandomString(8)
        Dim con2 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim que2 As String = "SELECT * FROM Shitesi ORDER BY ID"
        CreateReader3(con2, que2)
        Dim conbli As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim querybli As String = "SELECT * FROM Produktet ORDER BY ID"
        CreateReader_load_produktet_Blerje(conbli, querybli)
        Dim confornbli As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim queryforbli As String = "SELECT * FROM Fornitoret ORDER BY ID"
        CreateReader_load_fornitoret_Blerje(confornbli, queryforbli)
        NumericUpDown1.Value = 0
        ComboBox3.SelectedIndex = 0
    End Sub
    Public Sub CreateReader_Fornitoret(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                TextBox3.Text = (reader(2).ToString())
                TextBox4.Text = (reader(3).ToString())
                TextBox5.Text = (reader(4).ToString())
            End While
            reader.Close()
        End Using
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim confornitoret As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim queryfornitoret As String = "SELECT * FROM Fornitoret WHERE(Emri LIKE '" & ComboBox2.Text & "%')"
        CreateReader_Fornitoret(confornitoret, queryfornitoret)
    End Sub
    Public Sub CreateReader4(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                TextBox6.Text = (reader(2).ToString())
                TextBox7.Text = (reader(3).ToString())
                TextBox8.Text = (reader(4).ToString())
            End While
            reader.Close()
        End Using
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim con1 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim que1 As String = "SELECT * FROM Shitesi WHERE(Emri LIKE '" & ComboBox3.Text & "%')"
        CreateReader4(con1, que1)
    End Sub
    Public Sub CreateReader3(ByVal connectionString As String,
   ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                ComboBox3.Items.Add(reader(1).ToString())
            End While
            reader.Close()
            If ComboBox3.Text = "" Then
            Else
                ComboBox3.SelectedIndex = 0
            End If
        End Using
    End Sub
    Public Sub CreateReader_produktet_blerje(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                ComboBox4.Text = (reader(2).ToString())
                TextBox9.Text = (reader(3).ToString())
                ' TextBox7.Text = (reader(8).ToString())
            End While
            reader.Close()
        End Using
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim conxble As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim quexble As String = "SELECT * FROM Produktet WHERE(Emri_produktit LIKE '" & ComboBox1.Text & "%')"
        CreateReader_produktet_blerje(conxble, quexble)
        NumericUpDown1.Value = 1
        If TextBox9.Text = "" Then

        Else
            Dim vlera_shitjes = NumericUpDown1.Value * TextBox9.Text
            TextBox10.Text = vlera_shitjes.ToString("#,#", CultureInfo.InvariantCulture)
            Dim vlera_shitjes_metvsh = (NumericUpDown2.Value / 100) * vlera_shitjes
            Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + vlera_shitjes
            TextBox11.Text = vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture)
            TextBox12.Text = vlera_shitjes_metvsh1.ToString("#,#", CultureInfo.InvariantCulture)
        End If
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        If (e.RowIndex >= 0) Then
            Dim row As DataGridViewRow = DataGridView1.Rows.Item(e.RowIndex)
            Label18.Text = row.Cells.Item("ID").Value.ToString
        End If
    End Sub
    Dim Lidhjedhenat As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Dim adaptordhenat As OleDbDataAdapter
    Dim setidatavedhenat = New DataSet
    Dim querydhenat As OleDbCommand
    Dim nr1 As Integer = 0
    Dim nr11 As Integer = 0
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Or ComboBox2.Text = "" Or ComboBox3.Text = "" Then
            MsgBox("Plotesoni te gjitha fushat!", MsgBoxStyle.Information)
        Else
            Button2.Enabled = True
            Farmacia.DataGridView1.Columns("ID").Visible = True
            Dim nrekzistues As New List(Of Integer)
            For Each kolone As DataGridViewRow In Farmacia.DataGridView1.Rows
                nrekzistues.Add(CInt(kolone.Cells(1).Value))
            Next
            Dim existingNumbers As New List(Of Integer)
            For Each r As DataGridViewRow In Farmacia.DataGridView1.Rows
                existingNumbers.Add(CInt(r.Cells(1).Value))
            Next
            If existingNumbers.Count = 0 And nrekzistues.Count = 0 Then
                Dim vlera_shitjes_metvsh = (NumericUpDown2.Value / 100) * TextBox10.Text
                Dim Str1 As String
                Dim regDate As DateTime = Date.Now
                Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                Dim strDate1 As String = regDate.ToString("HH:mm tt")
                Try
                    Str1 = "insert into dhenat values("
                    Str1 += "1"
                    Str1 += ","
                    Str1 += """" & TextBox2.Text & """"
                    Str1 += ","
                    Str1 += """" & strDate.Trim & """"
                    Str1 += ","
                    Str1 += """" & strDate1.Trim & """"
                    Str1 += ","
                    Str1 += """" & ComboBox1.Text.Trim() & """"
                    Str1 += ","
                    Str1 += """" & ComboBox2.Text.Trim() & """"
                    Str1 += ","
                    Str1 += """" & ComboBox3.Text.Trim() & """"
                    Str1 += ","
                    Str1 += """" & ComboBox4.Text.Trim() & """"
                    Str1 += ","
                    Str1 += """" & NumericUpDown1.Value & """"
                    Str1 += ","
                    Str1 += """" & TextBox9.Text.Trim() & """"
                    Str1 += ","
                    Str1 += """" & TextBox10.Text.Trim & """"
                    Str1 += ","
                    Str1 += """" & TextBox11.Text.Trim() & """"
                    Str1 += ","
                    Str1 += """" & TextBox12.Text.Trim() & """"
                    Str1 += ","
                    Str1 += """" & Hyrje.TextBox1.Text.Trim() & """"
                    Str1 += ")"
                    Lidhjedhenat.Open()
                    querydhenat = New OleDbCommand(Str1, Lidhjedhenat)
                    querydhenat.ExecuteNonQuery()
                    setidatavedhenat.Clear()
                    adaptordhenat = New OleDbDataAdapter("SELECT * FROM dhenat WHERE(Kodi_Blerjes LIKE '" & TextBox2.Text & "%')", Lidhjedhenat)
                    adaptordhenat.Fill(setidatavedhenat, "tedhena")
                    DataGridView1.Columns.Add("count1", "Nr.")
                    DataGridView1.DataSource = setidatavedhenat.Tables(0)
                    DataGridView1.Columns("ID").Visible = False
                    nr11 = DataGridView1.Rows.Count - 1
                    For i = 0 To nr11
                        DataGridView1.Rows(i).Cells(0).Value = i + 1
                    Next
                    Lidhjedhenat.Close()
                    Dim nr As Integer = 0
                    nr = DataGridView1.Columns.Count - 1
                    If nr > 14 Then
                        DataGridView1.Columns.RemoveAt(nr)
                    End If
                    Dim Lidhje1xx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
                    Dim adaptor1xx As OleDbDataAdapter
                    Dim setidatave1xx = New DataSet
                    With Farmacia.DataGridView1
                        .RowsDefaultCellStyle.BackColor = Color.Bisque
                        .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                    End With
                    rreshtiaktual2 = 0
                    Lidhje1xx.Open()
                    adaptor1xx = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", Lidhje1xx)
                    adaptor1xx.Fill(setidatave1xx, "tedhena")
                    Farmacia.DataGridView1.DataSource = setidatave1xx.Tables(0)
                    Farmacia.Merrtedhenat_Shitjet(rreshtiaktual2shitjet)
                    Lidhje1xx.Close()
                    Dim total2 As Integer
                    Dim total3 As Integer
                    Dim total4 As Integer
                    For ii As Integer = 0 To DataGridView1.RowCount - 1
                        total2 = total2 + DataGridView1.Rows(ii).Cells(11).Value
                    Next
                    For iii As Integer = 0 To DataGridView1.RowCount - 1
                        total3 = total3 + DataGridView1.Rows(iii).Cells(12).Value
                    Next
                    For iiii As Integer = 0 To DataGridView1.RowCount - 1
                        total4 = total4 + DataGridView1.Rows(iiii).Cells(13).Value
                    Next
                    Label34.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                    Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                    Label30.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    Farmacia.DataGridView1.Columns("ID").Visible = False
                    nr1 = Farmacia.DataGridView1.Rows.Count - 1
                    For i = 0 To nr1
                        Farmacia.DataGridView1.Rows(i).Cells(0).Value = i + 1
                    Next
                Catch ex As Exception
                    MessageBox.Show("Nuk u shtua")
                    MsgBox(ex.Message & " -  " & ex.Source)
                End Try
            Else
                Dim missingNumbers = Enumerable.Range(existingNumbers.First, existingNumbers.Max() - existingNumbers.First + 1).Except(existingNumbers)
                Dim max = nrekzistues.Max() + 1
                If missingNumbers.Count = 0 Then
                    Dim vlera_shitjes_metvsh = (NumericUpDown2.Value / 100) * TextBox10.Text
                    Dim Str1 As String
                    Dim regDate As DateTime = Date.Now
                    Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                    Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                    Dim strDate1 As String = regDate.ToString("HH:mm tt")
                    Try
                        Str1 = "insert into dhenat values("
                        Str1 += max.ToString
                        Str1 += ","
                        Str1 += """" & TextBox2.Text & """"
                        Str1 += ","
                        Str1 += """" & strDate.Trim & """"
                        Str1 += ","
                        Str1 += """" & strDate1.Trim & """"
                        Str1 += ","
                        Str1 += """" & ComboBox1.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & ComboBox2.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & ComboBox3.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & ComboBox4.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & NumericUpDown1.Value & """"
                        Str1 += ","
                        Str1 += """" & TextBox9.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & TextBox10.Text.Trim & """"
                        Str1 += ","
                        Str1 += """" & TextBox11.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & TextBox12.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & Hyrje.TextBox1.Text.Trim() & """"
                        Str1 += ")"
                        Lidhjedhenat.Open()
                        querydhenat = New OleDbCommand(Str1, Lidhjedhenat)
                        querydhenat.ExecuteNonQuery()
                        setidatavedhenat.Clear()
                        adaptordhenat = New OleDbDataAdapter("SELECT * FROM dhenat WHERE(Kodi_Blerjes LIKE '" & TextBox2.Text & "%')", Lidhjedhenat)
                        adaptordhenat.Fill(setidatavedhenat, "tedhena")
                        DataGridView1.Columns.Add("count1", "Nr.")
                        DataGridView1.DataSource = setidatavedhenat.Tables(0)
                        DataGridView1.Columns("ID").Visible = False
                        nr11 = DataGridView1.Rows.Count - 1
                        For i = 0 To nr11
                            DataGridView1.Rows(i).Cells(0).Value = i + 1
                        Next
                        Lidhjedhenat.Close()
                        Dim nr As Integer = 0
                        nr = DataGridView1.Columns.Count - 1
                        If nr > 14 Then
                            DataGridView1.Columns.RemoveAt(nr)
                        End If
                        Dim Lidhje1xx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
                        Dim adaptor1xx As OleDbDataAdapter
                        Dim setidatave1xx = New DataSet
                        With Farmacia.DataGridView1
                            .RowsDefaultCellStyle.BackColor = Color.Bisque
                            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                        End With
                        rreshtiaktual2 = 0
                        Lidhje1xx.Open()
                        adaptor1xx = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", Lidhje1xx)
                        adaptor1xx.Fill(setidatave1xx, "tedhena")
                        Farmacia.DataGridView1.DataSource = setidatave1xx.Tables(0)
                        Farmacia.Merrtedhenat_Shitjet(rreshtiaktual2shitjet)
                        Lidhje1xx.Close()
                        Dim total2 As Integer
                        Dim total3 As Integer
                        Dim total4 As Integer
                        For ii As Integer = 0 To DataGridView1.RowCount - 1
                            total2 = total2 + DataGridView1.Rows(ii).Cells(11).Value
                        Next
                        For iii As Integer = 0 To DataGridView1.RowCount - 1
                            total3 = total3 + DataGridView1.Rows(iii).Cells(12).Value
                        Next
                        For iiii As Integer = 0 To DataGridView1.RowCount - 1
                            total4 = total4 + DataGridView1.Rows(iiii).Cells(13).Value
                        Next
                        Label34.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                        Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                        Label30.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                        MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                        Farmacia.DataGridView1.Columns("ID").Visible = False
                        nr1 = Farmacia.DataGridView1.Rows.Count - 1
                        For i = 0 To nr1
                            Farmacia.DataGridView1.Rows(i).Cells(0).Value = i + 1
                        Next
                    Catch ex As Exception
                        MessageBox.Show("Nuk u shtua")
                        MsgBox(ex.Message & " -  " & ex.Source)
                    End Try
                Else
                    Farmacia.DataGridView1.Columns("ID").Visible = True
                    Dim vlera_shitjes_metvsh = (NumericUpDown2.Value / 100) * TextBox10.Text
                    Dim Str1 As String
                    Dim regDate As DateTime = Date.Now
                    Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                    Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                    Dim strDate1 As String = regDate.ToString("HH:mm tt")
                    Try
                        Str1 = "insert into dhenat values("
                        Str1 += missingNumbers.First.ToString
                        Str1 += ","
                        Str1 += """" & TextBox2.Text & """"
                        Str1 += ","
                        Str1 += """" & strDate.Trim & """"
                        Str1 += ","
                        Str1 += """" & strDate1.Trim & """"
                        Str1 += ","
                        Str1 += """" & ComboBox1.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & ComboBox2.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & ComboBox3.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & ComboBox4.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & NumericUpDown1.Value & """"
                        Str1 += ","
                        Str1 += """" & TextBox9.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & TextBox10.Text.Trim & """"
                        Str1 += ","
                        Str1 += """" & TextBox11.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & TextBox12.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & Hyrje.TextBox1.Text.Trim() & """"
                        Str1 += ")"
                        Lidhjedhenat.Open()
                        querydhenat = New OleDbCommand(Str1, Lidhjedhenat)
                        querydhenat.ExecuteNonQuery()
                        setidatavedhenat.Clear()
                        adaptordhenat = New OleDbDataAdapter("SELECT * FROM dhenat WHERE(Kodi_Blerjes LIKE '" & TextBox2.Text & "%')", Lidhjedhenat)
                        adaptordhenat.Fill(setidatavedhenat, "tedhena")
                        'DataGridView1.Columns.Clear()
                        DataGridView1.Columns.Add("count1", "Nr.")
                        DataGridView1.DataSource = setidatavedhenat.Tables(0)
                        DataGridView1.Columns("ID").Visible = False
                        nr11 = DataGridView1.Rows.Count - 1
                        For i = 0 To nr11
                            DataGridView1.Rows(i).Cells(0).Value = i + 1
                        Next
                        Lidhjedhenat.Close()
                        Dim nr As Integer = 0
                        nr = DataGridView1.Columns.Count - 1
                        If nr > 14 Then
                            DataGridView1.Columns.RemoveAt(nr)
                        End If
                        Dim Lidhje1xx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
                        Dim adaptor1xx As OleDbDataAdapter
                        Dim setidatave1xx = New DataSet
                        With Farmacia.DataGridView1
                            .RowsDefaultCellStyle.BackColor = Color.Bisque
                            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                        End With
                        rreshtiaktual2 = 0
                        Lidhje1xx.Open()
                        adaptor1xx = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", Lidhje1xx)
                        adaptor1xx.Fill(setidatave1xx, "tedhena")
                        Farmacia.DataGridView1.DataSource = setidatave1xx.Tables(0)
                        Farmacia.Merrtedhenat_Shitjet(rreshtiaktual2shitjet)
                        Lidhje1xx.Close()
                        Dim total2 As Integer
                        Dim total3 As Integer
                        Dim total4 As Integer
                        For ii As Integer = 0 To DataGridView1.RowCount - 1
                            total2 = total2 + DataGridView1.Rows(ii).Cells(11).Value
                        Next
                        For iii As Integer = 0 To DataGridView1.RowCount - 1
                            total3 = total3 + DataGridView1.Rows(iii).Cells(12).Value
                        Next
                        For iiii As Integer = 0 To DataGridView1.RowCount - 1
                            total4 = total4 + DataGridView1.Rows(iiii).Cells(13).Value
                        Next
                        Label34.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                        Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                        Label30.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                        MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                        Farmacia.DataGridView1.Columns("ID").Visible = False
                        nr1 = Farmacia.DataGridView1.Rows.Count - 1
                        For i = 0 To nr1
                            Farmacia.DataGridView1.Rows(i).Cells(0).Value = i + 1
                        Next
                    Catch ex As Exception
                        MessageBox.Show("Nuk u shtua")
                        MsgBox(ex.Message & " -  " & ex.Source)
                    End Try
                End If
            End If
        End If
        If Not Farmacia.DataGridView1.Rows.Count > 0 Then
            Farmacia.Button2.Enabled = False
            Farmacia.Button3.Enabled = False
        Else
            Farmacia.Button2.Enabled = True
            Farmacia.Button3.Enabled = True
        End If
    End Sub
    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        If TextBox9.Text = "" Then

        Else
            Dim vlera_shitjes = NumericUpDown1.Value * TextBox9.Text
            TextBox10.Text = vlera_shitjes.ToString("#,#", CultureInfo.InvariantCulture)
            Dim vlera_shitjes_metvsh = (NumericUpDown2.Value / 100) * vlera_shitjes
            Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + vlera_shitjes
            TextBox11.Text = vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture)
            TextBox12.Text = vlera_shitjes_metvsh1.ToString("#,#", CultureInfo.InvariantCulture)
        End If
    End Sub
    Private Sub DataGridView1_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If (e.RowIndex >= 0) Then
            Dim row As DataGridViewRow = DataGridView1.Rows.Item(e.RowIndex)
            Label18.Text = row.Cells.Item("ID").Value.ToString
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label26.Text = TimeOfDay.ToString("h:mm:ss tt")
    End Sub
    Dim Lidhjedhenathiq As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Dim adaptordhenathiq As OleDbDataAdapter
    Dim setidatavedhenathiq = New DataSet
    Dim querydhenathiq As OleDbCommand
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.Rows.Count = 0 Then
        Else
            Dim Str As String
            Try
                With DataGridView1
                    .RowsDefaultCellStyle.BackColor = Color.Bisque
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                End With
                Str = "delete from dhenat where ID="
                Str += DataGridView1.CurrentRow.Cells(1).Value.ToString
                Lidhjedhenathiq.Open()
                querydhenathiq = New OleDbCommand(Str, Lidhjedhenathiq)
                querydhenathiq.ExecuteNonQuery()
                setidatavedhenathiq.clear()
                adaptordhenathiq = New OleDbDataAdapter("SELECT * FROM dhenat WHERE(Kodi_Blerjes LIKE '" & TextBox2.Text & "%')", Lidhjedhenathiq)
                adaptordhenathiq.Fill(setidatavedhenathiq, "tedhena")
                DataGridView1.DataSource = setidatavedhenathiq.Tables(0)
                Dim nrd As Integer = 0
                nrd = DataGridView1.Rows.Count - 1
                For i = 0 To nrd
                    DataGridView1.Rows(i).Cells(0).Value = i + 1
                Next
                Lidhjedhenathiq.Close()
                Dim nr As Integer = 0
                nr = DataGridView1.Columns.Count - 1
                If nr > 13 Then
                    DataGridView1.Columns.RemoveAt(nr)
                End If
                Dim Lidhje1xx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
                Dim adaptor1xx As OleDbDataAdapter
                Dim setidatave1xx = New DataSet
                With Farmacia.DataGridView1
                    .RowsDefaultCellStyle.BackColor = Color.Bisque
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                End With
                rreshtiaktual2 = 0
                Lidhje1xx.Open()
                adaptor1xx = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", Lidhje1xx)
                adaptor1xx.Fill(setidatave1xx, "tedhena")
                Farmacia.DataGridView1.DataSource = setidatave1xx.Tables(0)
                Farmacia.Merrtedhenat_Shitjet(rreshtiaktual2shitjet)
                Lidhje1xx.Close()
                Dim total2 As Integer
                Dim total3 As Integer
                Dim total4 As Integer
                For ii As Integer = 0 To DataGridView1.RowCount - 1
                    total2 = total2 + DataGridView1.Rows(ii).Cells(11).Value
                Next
                For iii As Integer = 0 To DataGridView1.RowCount - 1
                    total3 = total3 + DataGridView1.Rows(iii).Cells(12).Value
                Next
                For iiii As Integer = 0 To DataGridView1.RowCount - 1
                    total4 = total4 + DataGridView1.Rows(iiii).Cells(13).Value
                Next
                Label34.Text = total2
                Label29.Text = total3
                Label30.Text = total4
                Dim nrd1 As Integer = 0
                nrd1 = Farmacia.DataGridView1.Rows.Count - 1
                For i = 0 To nrd1
                    Farmacia.DataGridView1.Rows(i).Cells(0).Value = i + 1
                Next
                MsgBox("U fshi me sukses", MsgBoxStyle.Information)
            Catch ex As Exception
                MessageBox.Show("Nuk u fshi me sukses!")
                MsgBox(ex.Message & " -  " & ex.Source)
                lidhje2.Close()
            End Try
        End If
        If Not Farmacia.DataGridView1.Rows.Count > 0 Then
            Farmacia.Button2.Enabled = False
            Farmacia.Button3.Enabled = False
        Else
            Farmacia.Button2.Enabled = True
            Farmacia.Button3.Enabled = True
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox2.Text = GenerateRandomString(8)
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
        NumericUpDown1.Value = 1
        TextBox12.Text = 0
        TextBox10.Text = 0
        TextBox11.Text = 0
        Label34.Text = 0
        Label29.Text = 0
        Label30.Text = 0
        DataGridView1.DataSource = Nothing
        MsgBox("Fatura u ruajt me sukses!", MsgBoxStyle.Information)
    End Sub
    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        If TextBox9.Text = "" Then

        Else
            Dim vlera_shitjes = NumericUpDown1.Value * TextBox9.Text
            TextBox10.Text = vlera_shitjes
            Dim vlera_shitjes_metvsh = (NumericUpDown2.Value / 100) * TextBox10.Text
            Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + TextBox10.Text
            TextBox11.Text = vlera_shitjes_metvsh
            TextBox12.Text = vlera_shitjes_metvsh1
        End If
    End Sub
End Class