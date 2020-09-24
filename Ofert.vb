Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Public Class Ofert
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
            Label38.Text = "_________________________________________________________"
            Label39.Text = "_________________________________________________________"
        ElseIf Me.WindowState = FormWindowState.Normal Then
            Label38.Text = "______________________________________________"
            Label39.Text = "______________________________________________"
        End If
    End Sub
    Private Sub _Load(ByVal sender As Object, ByVal e As EventArgs)
        _form_resize._get_initial_size()
    End Sub
    Dim totale As Integer = 0
    Dim totale1 As Integer = 0
    Dim totalefund As Integer = 0
    'Gjeneron kodin e ofertes
    Public Function GenerateRandomString(ByRef iLength As Integer) As String
        Dim rdm As New Random()
        Dim allowChrs() As Char = "0123456789".ToCharArray()
        Dim sResult As String = ""
        For i As Integer = 0 To iLength - 1
            sResult += allowChrs(rdm.Next(0, allowChrs.Length))
        Next
        Return sResult
    End Function
    'Ngjyros ComboBox1 = Bleresi
    Private Sub ComboBox1_DrawItem(ByVal sender As Object,
    ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox1.DrawItem
        If e.Index < 0 Then Exit Sub
        Dim rect As System.Drawing.Rectangle = e.Bounds
        If e.State And DrawItemState.Selected Then
            e.Graphics.FillRectangle(Brushes.LightGreen, rect)
        Else
            e.Graphics.FillRectangle(SystemBrushes.Window, rect)
        End If
        Dim colorname As String = ComboBox1.Items(e.Index)
        Dim b As New SolidBrush(Color.FromName(colorname))
        Dim b2 As Brush = Brushes.Black 'add one
        e.Graphics.DrawString(colorname, Me.ComboBox1.Font, b2, rect.X, rect.Y)
    End Sub
    'Ngjyros ComboBox2 = Shitesi
    Private Sub ComboBox2_DrawItem(ByVal sender As Object,
    ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox2.DrawItem
        If e.Index < 0 Then Exit Sub
        Dim rect As System.Drawing.Rectangle = e.Bounds
        If e.State And DrawItemState.Selected Then
            e.Graphics.FillRectangle(Brushes.LightGreen, rect)
        Else
            e.Graphics.FillRectangle(SystemBrushes.Window, rect)
        End If
        Dim colorname As String = ComboBox2.Items(e.Index)
        Dim b As New SolidBrush(Color.FromName(colorname))
        Dim b2 As Brush = Brushes.Black 'add one
        e.Graphics.DrawString(colorname, Me.ComboBox2.Font, b2, rect.X, rect.Y)
    End Sub
    'Ngjyros ComboBox3 = Produkti
    Private Sub ComboBox3_DrawItem(ByVal sender As Object,
    ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox3.DrawItem
        If e.Index < 0 Then Exit Sub
        Dim rect As System.Drawing.Rectangle = e.Bounds
        If e.State And DrawItemState.Selected Then
            e.Graphics.FillRectangle(Brushes.LightGreen, rect)
        Else
            e.Graphics.FillRectangle(SystemBrushes.Window, rect)
        End If
        Dim colorname As String = ComboBox3.Items(e.Index)
        Dim b As New SolidBrush(Color.FromName(colorname))
        Dim b2 As Brush = Brushes.Black 'add one
        e.Graphics.DrawString(colorname, Me.ComboBox3.Font, b2, rect.X, rect.Y)
    End Sub
    'Lexon Bleresit ne ComboBox1
    Public Sub Kap_Bleresit(ByVal connectionString As String,
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
    'Lexon Shitesit ne ComboBox2
    Public Sub Kap_Shitesit(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                ComboBox2.Items.Add(reader(1).ToString())
            End While
            reader.Close()
            If ComboBox2.Text = "" Then

            Else
                ComboBox2.SelectedIndex = 0
            End If
        End Using
    End Sub
    'Lexon Produktet ne ComboBox3
    Public Sub Kap_Produkte(ByVal connectionString As String,
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
    Private Sub Ofert_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Farmacia.Show()
        My.Settings.tvsh = NumericUpDown3.Value
    End Sub
    Public Sub Nrofertes(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                If reader(0).ToString() = Nothing Then
                Else
                    ListBox1.Items.Add(reader(0).ToString())
                End If
            End While
            reader.Close()
            For i = 0 To ListBox1.Items.Count
                Label31.Text = i.ToString.Max().ToString + 1
            Next
        End Using
    End Sub
    Private Sub Ofert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        NumericUpDown3.Value = My.Settings.tvsh
        Konfigurime.TextBox1.Text = My.Settings.backup
        Konfigurime.TextBox2.Text = My.Settings.logo
        Konfigurime.TextBox3.Text = My.Settings.ruajraportet
        Konfigurime.TextBox4.Text = My.Settings.faturatofert
        If DataGridView1.Rows.Count = 0 Then
            Button3.Enabled = False
        Else
            Button3.Enabled = True
        End If
        Dim regDate2 As DateTime = Date.Now
        Dim strDate2 As String = regDate2.ToString("dd/MM/yyyy")
        Dim connrfat As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim queynrfat As String = "SELECT DISTINCT Kodi_Ofertes FROM Ofertat WHERE(Data Like '" & strDate2 & "%')"
        Nrofertes(connrfat, queynrfat)
        'Afishon daten ne label18
        Dim regDate As DateTime = Date.Now
        Dim strDate As String = regDate.ToString("dd/MM/yyyy")
        Label18.Text = strDate
        'Ngjyros celizat e datagridview1
        With DataGridView1
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        'Therret oren
        Timer1.Start()
        'Gjeneron kodin e ofertes ne textbox8
        TextBox8.Text = GenerateRandomString(8)
        'Kap bleresit nga tabela Klientet dhe i eksporton ne ComboBox1
        Dim con_bleresit As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim query_bleresit As String = "SELECT * FROM Bleresit ORDER BY ID"
        Kap_Bleresit(con_bleresit, query_bleresit)
        'Kap bleresit nga tabela SHitesi dhe i eksporton ne ComboBox1
        Dim con_shitesit As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim query_shitesit As String = "SELECT * FROM Shitesi ORDER BY ID"
        Kap_Shitesit(con_shitesit, query_shitesit)
        'Kap produktet nga tabela Produket dhe i eksporton ne ComboBox3
        Dim conprod As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim queyryprod As String = "SELECT * FROM Produktet ORDER BY ID"
        Kap_Produkte(conprod, queyryprod)
        NumericUpDown1.Value = 1
        Button5.Text = "Ruaj Faturen(" & TextBox8.Text & ")"
        Button6.Text = "Printo Faturen(" & TextBox8.Text & ")"
        NumericUpDown1.Value = 0
        ComboBox2.SelectedIndex = 0
    End Sub
    'Cekon nqs bleresi eshte klient ekzistues apo jo
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        For Each cbItem As String In Me.ComboBox1.Items
            If TextBox6.Text = "" Then
                Label24.Text = "-"
                Label24.ForeColor = DefaultForeColor
            Else
                If TextBox6.Text = ComboBox1.Text Then
                    Label24.Text = "PO"
                    Label24.ForeColor = Color.ForestGreen
                Else
                    Label24.Text = "JO"
                    Label24.ForeColor = Color.Red
                End If
            End If
        Next
    End Sub
    'Shfaq listen e klienteve ekzistues
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            ComboBox1.Visible = True
            Label6.Visible = True
            Label7.Visible = True
            Label8.Visible = True
            TextBox1.Visible = True
            TextBox2.Visible = True
            TextBox3.Visible = True
            TextBox6.Text = ComboBox1.Text
        Else
            ComboBox1.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Label8.Visible = False
            TextBox1.Visible = False
            TextBox2.Visible = False
            TextBox3.Visible = False
            TextBox6.Text = ""
        End If
    End Sub
    'Shfaq oren ne label20
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label20.Text = TimeOfDay.ToString("h:mm:ss tt")
    End Sub
    'Kap te dhenat e shitesit si cel,nipt,adrese
    Public Sub Kap_tedhenateshitesit(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                TextBox12.Text = (reader(2).ToString())
                TextBox13.Text = (reader(3).ToString())
                TextBox14.Text = (reader(4).ToString())
            End While
            reader.Close()
        End Using
    End Sub
    'Shfaq te dhenat e shitesit neper textbox
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim contedhenateshitesit As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim querytedhenateshitesit As String = "SELECT * FROM Shitesi WHERE(Emri LIKE '" & ComboBox2.Text & "%')"
        Kap_tedhenateshitesit(contedhenateshitesit, querytedhenateshitesit)
    End Sub
    'kap te dhenat e bleresit si adrese,nipt,cel neper textbox
    Public Sub Kap_Bleresit1(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                TextBox1.Text = (reader(2).ToString())
                TextBox2.Text = (reader(3).ToString())
                TextBox3.Text = (reader(4).ToString())
            End While
            reader.Close()
        End Using
    End Sub
    'lexon te dhenat e bleresit ne textboxe
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim conbleresit As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim querybleresit As String = "SELECT * FROM Bleresit WHERE(Emri LIKE '" & ComboBox1.Text & "%')"
        Kap_Bleresit1(conbleresit, querybleresit)
        If CheckBox2.Checked = True Then
            TextBox6.Text = ComboBox1.Text
        Else
            TextBox6.Text = ""
        End If
    End Sub
    'Lexon te dhenat e produktetve si cmim dhe njesi
    Public Sub CreateReader_produktet(ByVal connectionString As String,
   ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                ComboBox4.Text = (reader(2).ToString())
                TextBox9.Text = (reader(4).ToString())
            End While
            reader.Close()
        End Using
    End Sub
    'shfaq te dhenat e produkteve te texboxe
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        Dim conproduktet As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim queryproduktet As String = "SELECT * FROM Produktet WHERE(Emri_produktit LIKE '" & ComboBox3.Text & "%')"
        CreateReader_produktet(conproduktet, queryproduktet)
        ' NumericUpDown1.Value = 1
        Dim conxzz As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim quexzz As String = "SELECT * FROM Shitjet WHERE(Produkti Like '" &
                                           ComboBox3.Text & "%') ORDER BY Kodi_Shitjes"
        CreateReader_sasieshitur(conxzz, quexzz)
        Application.DoEvents()
        Dim conxzz1 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim quexzz1 As String = "SELECT * FROM dhenat WHERE(Produkti Like '" &
                                           ComboBox3.Text & "%') ORDER BY Kodi_Blerjes"
        sasiefutur(conxzz1, quexzz1)

        Application.DoEvents()
        totalefund = 0
        totalefund = totale1 - totale
        If totalefund = 0 Then
            Label25.ForeColor = Color.Red
            Label25.Text = totalefund
            NumericUpDown1.Value = totalefund
        Else
            Label25.ForeColor = Color.ForestGreen
            Label25.Text = totalefund
            NumericUpDown1.Value = 0
            NumericUpDown1.Value = 1
        End If
        totale = 0
        totale1 = 0
    End Sub
    Public Sub CreateReader_sasieshitur(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                totale = totale + reader(8).ToString()
            End While
            reader.Close()
        End Using
    End Sub
    Public Sub sasiefutur(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                totale1 = totale1 + reader(8).ToString()
            End While
            reader.Close()
        End Using
    End Sub
    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        Dim conxzz As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim quexzz As String = "SELECT * FROM Shitjet WHERE(Produkti Like '" &
                                           ComboBox3.Text & "%') ORDER BY Kodi_Shitjes"
        CreateReader_sasieshitur(conxzz, quexzz)
        Application.DoEvents()
        Dim conxzz1 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim quexzz1 As String = "SELECT * FROM dhenat WHERE(Produkti Like '" &
                                           ComboBox3.Text & "%') ORDER BY Kodi_Blerjes"
        sasiefutur(conxzz1, quexzz1)
        totalefund = 0
        totalefund = totale1 - totale
        Application.DoEvents()
        If NumericUpDown1.Value > totalefund Then
            MsgBox("Gjendja e disponueshme e " & ComboBox3.Text & " eshte = " & totalefund, MsgBoxStyle.Information)
            NumericUpDown1.Value = totalefund
        Else
            '  NumericUpDown1.Value = totalefund
            If CheckBox1.Checked = True Then
                Dim zbritje = NumericUpDown1.Value * TextBox9.Text
                Dim total1 As Decimal = zbritje * (1 - NumericUpDown2.Value / 100)
                Dim vlera_shitjes = total1
                TextBox4.Text = vlera_shitjes.ToString("#,#", CultureInfo.InvariantCulture)
                Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * vlera_shitjes
                Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + vlera_shitjes
                TextBox23.Text = vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture)
                TextBox24.Text = vlera_shitjes_metvsh1.ToString("#,#", CultureInfo.InvariantCulture)
            Else
                If TextBox9.Text = "" Then
                Else
                    Dim vlera_shitjes = NumericUpDown1.Value * TextBox9.Text
                    TextBox4.Text = vlera_shitjes.ToString("#,#", CultureInfo.InvariantCulture)
                    Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * vlera_shitjes
                    Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + vlera_shitjes
                    TextBox23.Text = vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture)
                    TextBox24.Text = vlera_shitjes_metvsh1.ToString("#,#", CultureInfo.InvariantCulture)
                End If
            End If
            totale = 0
            totale1 = 0
            'totalefund = 0
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox1.ForeColor = Color.ForestGreen
            NumericUpDown2.Enabled = True
            NumericUpDown2.BackColor = Color.LightGreen
        Else
            NumericUpDown2.Enabled = False
            CheckBox1.ForeColor = Color.Red
            NumericUpDown2.BackColor = DefaultBackColor
        End If
    End Sub
    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        If CheckBox1.Checked = True Then
            If TextBox9.Text = "" Then
            Else
                Dim zbritje = NumericUpDown1.Value * TextBox9.Text
                Dim total1 As Decimal = zbritje * (1 - NumericUpDown2.Value / 100)
                Dim totalizbriturcmimit As Decimal = TextBox9.Text * (1 - NumericUpDown2.Value / 100)
                If NumericUpDown2.Value > 0 Then
                    Label26.Text = totalizbriturcmimit.ToString("#,#", CultureInfo.InvariantCulture)
                    Label26.ForeColor = Color.Red
                Else
                    Label26.Text = 0
                    Label26.ForeColor = DefaultForeColor
                End If
                Dim vlera_shitjes = total1
                TextBox4.Text = vlera_shitjes.ToString("#,#", CultureInfo.InvariantCulture)
                Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * vlera_shitjes
                Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + vlera_shitjes
                TextBox23.Text = vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture)
                TextBox24.Text = vlera_shitjes_metvsh1.ToString("#,#", CultureInfo.InvariantCulture)
            End If
        Else
            If TextBox9.Text = "" Then
            Else
                Dim vlera_shitjes = NumericUpDown1.Value * TextBox9.Text
                TextBox4.Text = vlera_shitjes.ToString("#,#", CultureInfo.InvariantCulture)
                Label26.Text = 0
                Label26.ForeColor = DefaultForeColor
                Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * vlera_shitjes
                Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + vlera_shitjes
                TextBox23.Text = vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture)
                TextBox24.Text = vlera_shitjes_metvsh1.ToString("#,#", CultureInfo.InvariantCulture)
            End If
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox8.Text = GenerateRandomString(8)
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
        NumericUpDown1.Value = 1
        NumericUpDown2.Value = 0
        TextBox4.Text = 0
        TextBox23.Text = 0
        TextBox24.Text = 0
        Label37.Text = 0
        Label38.Text = 0
        Label23.Text = 0
        Label26.Text = 0
        DataGridView1.DataSource = Nothing
        ListBox1.Items.Clear()
        Button5.Text = "Ruaj Faturen(" & TextBox8.Text & ")"
        Button6.Text = "Printo Faturen(" & TextBox8.Text & ")"
        Dim regDate2 As DateTime = Date.Now
        Dim strDate2 As String = regDate2.ToString("dd/MM/yyyy")
        Dim connrfat As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim queynrfat As String = "SELECT DISTINCT Kodi_Ofertes FROM Ofertat WHERE(Data Like '" & strDate2 & "%')"
        Nrofertes(connrfat, queynrfat)
        MsgBox("Fatura u ruajt me sukses!", MsgBoxStyle.Information)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox6.Text = "" Then
            MsgBox("Plotesoni fushen e bleresit!", MsgBoxStyle.Information)
        Else
            Button3.Enabled = True
            Farmacia.DataGridView7.Columns("ID").Visible = True
            Dim zbritje = NumericUpDown1.Value * TextBox9.Text
            If CheckBox1.Checked = True Then
                Dim total1 As Decimal = zbritje * (1 - NumericUpDown2.Value / 100)
                ' TextBox9.Text = total1
                Dim nrekzistues As New List(Of Integer)
                For Each kolone As DataGridViewRow In Farmacia.DataGridView7.Rows
                    nrekzistues.Add(CInt(kolone.Cells(1).Value))
                Next
                Dim existingNumbers As New List(Of Integer)
                For Each r As DataGridViewRow In Farmacia.DataGridView7.Rows
                    existingNumbers.Add(CInt(r.Cells(1).Value))
                Next
                If existingNumbers.Count = 0 And nrekzistues.Count = 0 Then
                    Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox4.Text
                    Dim Str1 As String
                    Dim regDate As DateTime = Date.Now
                    Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                    Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                    Dim strDate1 As String = regDate.ToString("HH:mm tt")
                    Try
                        Str1 = "insert into Ofertat values("
                        Str1 += "1"
                        Str1 += ","
                        Str1 += """" & TextBox8.Text & """"
                        Str1 += ","
                        Str1 += """" & strDate.Trim & """"
                        Str1 += ","
                        Str1 += """" & strDate1.Trim & """"
                        Str1 += ","
                        Str1 += """" & ComboBox2.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & TextBox6.Text & """"
                        Str1 += ","
                        Str1 += """" & Label24.Text & """"
                        Str1 += ","
                        Str1 += """" & ComboBox3.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & ComboBox4.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & NumericUpDown1.Value & """"
                        Str1 += ","
                        Str1 += """" & total1.ToString("#,#", CultureInfo.InvariantCulture) & """"
                        Str1 += ","
                        Str1 += """" & TextBox4.Text.Trim & """"
                        Str1 += ","
                        Str1 += """" & vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture) & """"
                        Str1 += ","
                        Str1 += """" & TextBox24.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & NumericUpDown2.Value & """"
                        Str1 += ","
                        Str1 += """" & Hyrje.TextBox1.Text.Trim() & """"
                        Str1 += ")"
                        lidhjedg71ofert.Open()
                        querydg71ofert = New OleDbCommand(Str1, lidhjedg71ofert)
                        querydg71ofert.ExecuteNonQuery()
                        setitedhenavedg71ofert.Clear()
                        adaptoridg71ofert = New OleDbDataAdapter("SELECT * FROM Ofertat WHERE(Kodi_Ofertes LIKE '" & TextBox8.Text & "%')", lidhjedg71ofert)
                        adaptoridg71ofert.Fill(setitedhenavedg71ofert, "tedhena")
                        Dim nrshit As Integer = 0
                        DataGridView1.Columns.Add("count1", "Nr.")
                        DataGridView1.DataSource = setitedhenavedg71ofert.Tables(0)
                        DataGridView1.Columns("ID").Visible = False
                        nrshit = DataGridView1.Rows.Count - 1
                        For i = 0 To nrshit
                            DataGridView1.Rows(i).Cells(0).Value = i + 1
                        Next
                        lidhjedg71ofert.Close()
                        Dim nr As Integer = 0
                        nr = DataGridView1.Columns.Count - 1
                        If nr > 16 Then
                            DataGridView1.Columns.RemoveAt(nr)
                        End If
                        With Farmacia.DataGridView7
                            .RowsDefaultCellStyle.BackColor = Color.Bisque
                            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                        End With
                        rreshtiaktualdg71ofert = 0
                        lidhjedg71ofert1.Open()
                        adaptoridg71ofert1 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhjedg71ofert1)
                        adaptoridg71ofert1.Fill(setitedhenavedg71ofert1, "tedhena")
                        Farmacia.DataGridView7.DataSource = setitedhenavedg71ofert1.Tables(0)
                        Farmacia.Merrtedhenat_Ofertat(rreshtiaktualdg71ofert)
                        lidhjedg71ofert1.Close()
                        Dim total2 As Integer
                        Dim total3 As Integer
                        Dim total4 As Integer
                        For ii As Integer = 0 To DataGridView1.RowCount - 1
                            total2 = total2 + DataGridView1.Rows(ii).Cells(12).Value
                        Next
                        For iii As Integer = 0 To DataGridView1.RowCount - 1
                            total3 = total3 + DataGridView1.Rows(iii).Cells(13).Value
                        Next
                        For iiii As Integer = 0 To DataGridView1.RowCount - 1
                            total4 = total4 + DataGridView1.Rows(iiii).Cells(14).Value
                        Next
                        Label37.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                        Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                        Label23.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                        MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                        Farmacia.DataGridView7.Columns("ID").Visible = False
                        Dim nr1 As Integer = 0
                        nr1 = Farmacia.DataGridView7.Rows.Count - 1
                        For i = 0 To nr1
                            Farmacia.DataGridView7.Rows(i).Cells(0).Value = i + 1
                        Next
                    Catch ex As Exception
                        MsgBox("Nuk u shtua me sukses", MsgBoxStyle.Information)
                        lidhjedg71ofert1.Close()
                    End Try
                Else
                    Dim missingNumbers = Enumerable.Range(existingNumbers.First, existingNumbers.Max() - existingNumbers.First + 1).Except(existingNumbers)
                    Dim max = nrekzistues.Max() + 1
                    If missingNumbers.Count = 0 Then
                        Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox4.Text
                        Dim Str1 As String
                        Dim regDate As DateTime = Date.Now
                        Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                        Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                        Dim strDate1 As String = regDate.ToString("HH:mm tt")
                        Try
                            Str1 = "insert into Ofertat values("
                            Str1 += max.ToString
                            Str1 += ","
                            Str1 += """" & TextBox8.Text & """"
                            Str1 += ","
                            Str1 += """" & strDate.Trim & """"
                            Str1 += ","
                            Str1 += """" & strDate1.Trim & """"
                            Str1 += ","
                            Str1 += """" & ComboBox2.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & TextBox6.Text & """"
                            Str1 += ","
                            Str1 += """" & Label24.Text & """"
                            Str1 += ","
                            Str1 += """" & ComboBox3.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & ComboBox4.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & NumericUpDown1.Value & """"
                            Str1 += ","
                            Str1 += """" & total1.ToString("#,#", CultureInfo.InvariantCulture) & """"
                            Str1 += ","
                            Str1 += """" & TextBox4.Text.Trim & """"
                            Str1 += ","
                            Str1 += """" & vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture) & """"
                            Str1 += ","
                            Str1 += """" & TextBox24.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & NumericUpDown2.Value & """"
                            Str1 += ","
                            Str1 += """" & Hyrje.TextBox1.Text.Trim() & """"
                            Str1 += ")"
                            lidhjedg71ofert.Open()
                            querydg71ofert = New OleDbCommand(Str1, lidhjedg71ofert)
                            querydg71ofert.ExecuteNonQuery()
                            setitedhenavedg71ofert.Clear()
                            adaptoridg71ofert = New OleDbDataAdapter("SELECT * FROM Ofertat WHERE(Kodi_Ofertes LIKE '" & TextBox8.Text & "%')", lidhjedg71ofert)
                            adaptoridg71ofert.Fill(setitedhenavedg71ofert, "tedhena")
                            Dim nrshit As Integer = 0
                            DataGridView1.Columns.Add("count1", "Nr.")
                            DataGridView1.DataSource = setitedhenavedg71ofert.Tables(0)
                            DataGridView1.Columns("ID").Visible = False
                            nrshit = DataGridView1.Rows.Count - 1
                            For i = 0 To nrshit
                                DataGridView1.Rows(i).Cells(0).Value = i + 1
                            Next
                            lidhjedg71ofert.Close()
                            Dim nr As Integer = 0
                            nr = DataGridView1.Columns.Count - 1
                            If nr > 16 Then
                                DataGridView1.Columns.RemoveAt(nr)
                            End If
                            With Farmacia.DataGridView7
                                .RowsDefaultCellStyle.BackColor = Color.Bisque
                                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                            End With
                            rreshtiaktualdg71ofert = 0
                            lidhjedg71ofert1.Open()
                            adaptoridg71ofert1 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhjedg71ofert1)
                            adaptoridg71ofert1.Fill(setitedhenavedg71ofert1, "tedhena")
                            Farmacia.DataGridView7.DataSource = setitedhenavedg71ofert1.Tables(0)
                            Farmacia.Merrtedhenat_Ofertat(rreshtiaktualdg71ofert)
                            lidhjedg71ofert1.Close()
                            Dim total2 As Integer
                            Dim total3 As Integer
                            Dim total4 As Integer
                            For ii As Integer = 0 To DataGridView1.RowCount - 1
                                total2 = total2 + DataGridView1.Rows(ii).Cells(12).Value
                            Next
                            For iii As Integer = 0 To DataGridView1.RowCount - 1
                                total3 = total3 + DataGridView1.Rows(iii).Cells(13).Value
                            Next
                            For iiii As Integer = 0 To DataGridView1.RowCount - 1
                                total4 = total4 + DataGridView1.Rows(iiii).Cells(14).Value
                            Next
                            Label37.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                            Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                            Label23.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                            MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                            Farmacia.DataGridView7.Columns("ID").Visible = False
                            Dim nr1 As Integer = 0
                            nr1 = Farmacia.DataGridView7.Rows.Count - 1
                            For i = 0 To nr1
                                Farmacia.DataGridView7.Rows(i).Cells(0).Value = i + 1
                            Next
                        Catch ex As Exception
                            MessageBox.Show("Nuk u shtua")
                            MsgBox(ex.Message & " -  " & ex.Source)
                            lidhjedg71ofert1.Close()
                        End Try
                    Else
                        Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox4.Text
                        Farmacia.DataGridView7.Columns("ID").Visible = True
                        Dim Str1 As String
                        Dim regDate As DateTime = Date.Now
                        Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                        Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                        Dim strDate1 As String = regDate.ToString("HH:mm tt")
                        Try
                            Str1 = "insert into Ofertat values("
                            Str1 += missingNumbers.First.ToString
                            Str1 += ","
                            Str1 += """" & TextBox8.Text & """"
                            Str1 += ","
                            Str1 += """" & strDate.Trim & """"
                            Str1 += ","
                            Str1 += """" & strDate1.Trim & """"
                            Str1 += ","
                            Str1 += """" & ComboBox2.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & TextBox6.Text & """"
                            Str1 += ","
                            Str1 += """" & Label24.Text & """"
                            Str1 += ","
                            Str1 += """" & ComboBox3.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & ComboBox4.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & NumericUpDown1.Value & """"
                            Str1 += ","
                            Str1 += """" & total1.ToString("#,#", CultureInfo.InvariantCulture) & """"
                            Str1 += ","
                            Str1 += """" & TextBox4.Text.Trim & """"
                            Str1 += ","
                            Str1 += """" & vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture) & """"
                            Str1 += ","
                            Str1 += """" & TextBox24.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & NumericUpDown2.Value & """"
                            Str1 += ","
                            Str1 += """" & Hyrje.TextBox1.Text.Trim() & """"
                            Str1 += ")"
                            lidhjedg71ofert.Open()
                            querydg71ofert = New OleDbCommand(Str1, lidhjedg71ofert)
                            querydg71ofert.ExecuteNonQuery()
                            setitedhenavedg71ofert.Clear()
                            adaptoridg71ofert = New OleDbDataAdapter("SELECT * FROM Ofertat WHERE(Kodi_Ofertes LIKE '" & TextBox8.Text & "%')", lidhjedg71ofert)
                            adaptoridg71ofert.Fill(setitedhenavedg71ofert, "tedhena")
                            DataGridView1.Columns.Add("count1", "Nr.")
                            DataGridView1.DataSource = setitedhenavedg71ofert.Tables(0)
                            DataGridView1.Columns("ID").Visible = False
                            Dim nr11 As Integer = 0
                            nr11 = DataGridView1.Rows.Count - 1
                            For i = 0 To nr11
                                DataGridView1.Rows(i).Cells(0).Value = i + 1
                            Next
                            lidhjedg71ofert.Close()
                            Dim nr As Integer = 0
                            nr = DataGridView1.Columns.Count - 1
                            If nr > 16 Then
                                DataGridView1.Columns.RemoveAt(nr)
                            End If
                            With Farmacia.DataGridView7
                                .RowsDefaultCellStyle.BackColor = Color.Bisque
                                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                            End With
                            rreshtiaktualdg71ofert = 0
                            lidhjedg71ofert1.Open()
                            adaptoridg71ofert1 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhjedg71ofert1)
                            adaptoridg71ofert1.Fill(setitedhenavedg71ofert1, "tedhena")
                            Farmacia.DataGridView7.DataSource = setitedhenavedg71ofert1.Tables(0)
                            Farmacia.Merrtedhenat_Ofertat(rreshtiaktualdg71ofert)
                            lidhjedg71ofert1.Close()
                            Dim total2 As Integer
                            Dim total3 As Integer
                            Dim total4 As Integer
                            For ii As Integer = 0 To DataGridView1.RowCount - 1
                                total2 = total2 + DataGridView1.Rows(ii).Cells(12).Value
                            Next
                            For iii As Integer = 0 To DataGridView1.RowCount - 1
                                total3 = total3 + DataGridView1.Rows(iii).Cells(13).Value
                            Next
                            For iiii As Integer = 0 To DataGridView1.RowCount - 1
                                total4 = total4 + DataGridView1.Rows(iiii).Cells(14).Value
                            Next
                            Label37.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                            Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                            Label23.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                            MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                            Dim nr1 As Integer = 0
                            Farmacia.DataGridView7.Columns("ID").Visible = False
                            nr1 = Farmacia.DataGridView7.Rows.Count - 1
                            For i = 0 To nr1
                                Farmacia.DataGridView7.Rows(i).Cells(0).Value = i + 1
                            Next
                        Catch ex As Exception
                            MsgBox("Nuk u shtua me sukses", MsgBoxStyle.Information)
                            lidhjedg71ofert1.Close()
                        End Try
                    End If
                End If
            Else
                Farmacia.DataGridView2.Columns("ID").Visible = True
                Dim nrekzistues As New List(Of Integer)
                For Each kolone As DataGridViewRow In Farmacia.DataGridView7.Rows
                    nrekzistues.Add(CInt(kolone.Cells(1).Value))
                Next
                Dim existingNumbers As New List(Of Integer)
                For Each r As DataGridViewRow In Farmacia.DataGridView7.Rows
                    existingNumbers.Add(CInt(r.Cells(1).Value))
                Next
                If existingNumbers.Count = 0 And nrekzistues.Count = 0 Then
                    Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox4.Text
                    Dim Str1 As String
                    Dim regDate As DateTime = Date.Now
                    Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                    Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                    Dim strDate1 As String = regDate.ToString("HH:mm tt")
                    Try
                        Str1 = "insert into Ofertat values("
                        Str1 += "1"
                        Str1 += ","
                        Str1 += """" & TextBox8.Text & """"
                        Str1 += ","
                        Str1 += """" & strDate.Trim & """"
                        Str1 += ","
                        Str1 += """" & strDate1.Trim & """"
                        Str1 += ","
                        Str1 += """" & ComboBox2.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & TextBox6.Text & """"
                        Str1 += ","
                        Str1 += """" & Label24.Text & """"
                        Str1 += ","
                        Str1 += """" & ComboBox3.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & ComboBox4.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & NumericUpDown1.Value & """"
                        Str1 += ","
                        Str1 += """" & TextBox9.Text & """"
                        Str1 += ","
                        Str1 += """" & TextBox4.Text.Trim & """"
                        Str1 += ","
                        Str1 += """" & vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture) & """"
                        Str1 += ","
                        Str1 += """" & TextBox24.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & NumericUpDown2.Value & """"
                        Str1 += ","
                        Str1 += """" & Hyrje.TextBox1.Text.Trim() & """"
                        Str1 += ")"
                        lidhjedg71ofert.Open()
                        querydg71ofert = New OleDbCommand(Str1, lidhjedg71ofert)
                        querydg71ofert.ExecuteNonQuery()
                        setitedhenavedg71ofert.Clear()
                        adaptoridg71ofert = New OleDbDataAdapter("SELECT * FROM Ofertat WHERE(Kodi_Ofertes LIKE '" & TextBox8.Text & "%')", lidhjedg71ofert)
                        adaptoridg71ofert.Fill(setitedhenavedg71ofert, "tedhena")
                        DataGridView1.Columns.Add("count1", "Nr.")
                        DataGridView1.DataSource = setitedhenavedg71ofert.Tables(0)
                        DataGridView1.Columns("ID").Visible = False
                        Dim nr11 As Integer = 0
                        nr11 = DataGridView1.Rows.Count - 1
                        For i = 0 To nr11
                            DataGridView1.Rows(i).Cells(0).Value = i + 1
                        Next
                        lidhjedg71ofert.Close()
                        Dim nr As Integer = 0
                        nr = DataGridView1.Columns.Count - 1
                        If nr > 16 Then
                            DataGridView1.Columns.RemoveAt(nr)
                        End If
                        With Farmacia.DataGridView7
                            .RowsDefaultCellStyle.BackColor = Color.Bisque
                            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                        End With
                        rreshtiaktualdg71ofert = 0
                        lidhjedg71ofert1.Open()
                        adaptoridg71ofert1 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhjedg71ofert1)
                        adaptoridg71ofert1.Fill(setitedhenavedg71ofert1, "tedhena")
                        Farmacia.DataGridView7.DataSource = setitedhenavedg71ofert1.Tables(0)
                        Farmacia.Merrtedhenat_Ofertat(rreshtiaktualdg71ofert)
                        lidhjedg71ofert1.Close()
                        Dim total2 As Integer
                        Dim total3 As Integer
                        Dim total4 As Integer
                        For ii As Integer = 0 To DataGridView1.RowCount - 1
                            total2 = total2 + DataGridView1.Rows(ii).Cells(12).Value
                        Next
                        For iii As Integer = 0 To DataGridView1.RowCount - 1
                            total3 = total3 + DataGridView1.Rows(iii).Cells(13).Value
                        Next
                        For iiii As Integer = 0 To DataGridView1.RowCount - 1
                            total4 = total4 + DataGridView1.Rows(iiii).Cells(14).Value
                        Next
                        Label37.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                        Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                        Label23.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                        MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                        Dim nr1 As Integer = 0
                        Farmacia.DataGridView7.Columns("ID").Visible = False
                        nr1 = Farmacia.DataGridView7.Rows.Count - 1
                        For i = 0 To nr1
                            Farmacia.DataGridView7.Rows(i).Cells(0).Value = i + 1
                        Next
                    Catch ex As Exception
                        MessageBox.Show("Nuk u shtua")
                        MsgBox(ex.Message & " -  " & ex.Source)
                        lidhjedg71ofert1.Close()
                    End Try
                Else
                    Dim missingNumbers = Enumerable.Range(existingNumbers.First, existingNumbers.Max() - existingNumbers.First + 1).Except(existingNumbers)
                    Dim max = nrekzistues.Max() + 1
                    If missingNumbers.Count = 0 Then
                        Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox4.Text
                        Dim Str1 As String
                        Dim regDate As DateTime = Date.Now
                        Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                        Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                        Dim strDate1 As String = regDate.ToString("HH:mm tt")
                        Try
                            Str1 = "insert into Ofertat values("
                            Str1 += max.ToString
                            Str1 += ","
                            Str1 += """" & TextBox8.Text & """"
                            Str1 += ","
                            Str1 += """" & strDate.Trim & """"
                            Str1 += ","
                            Str1 += """" & strDate1.Trim & """"
                            Str1 += ","
                            Str1 += """" & ComboBox2.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & TextBox6.Text & """"
                            Str1 += ","
                            Str1 += """" & Label24.Text & """"
                            Str1 += ","
                            Str1 += """" & ComboBox3.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & ComboBox4.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & NumericUpDown1.Value & """"
                            Str1 += ","
                            Str1 += """" & TextBox9.Text & """"
                            Str1 += ","
                            Str1 += """" & TextBox4.Text.Trim & """"
                            Str1 += ","
                            Str1 += """" & vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture) & """"
                            Str1 += ","
                            Str1 += """" & TextBox24.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & NumericUpDown2.Value & """"
                            Str1 += ","
                            Str1 += """" & Hyrje.TextBox1.Text.Trim() & """"
                            Str1 += ")"
                            lidhjedg71ofert.Open()
                            querydg71ofert = New OleDbCommand(Str1, lidhjedg71ofert)
                            querydg71ofert.ExecuteNonQuery()
                            setitedhenavedg71ofert.Clear()
                            adaptoridg71ofert = New OleDbDataAdapter("SELECT * FROM Ofertat WHERE(Kodi_Ofertes LIKE '" & TextBox8.Text & "%')", lidhjedg71ofert)
                            adaptoridg71ofert.Fill(setitedhenavedg71ofert, "tedhena")
                            DataGridView1.Columns.Add("count1", "Nr.")
                            DataGridView1.DataSource = setitedhenavedg71ofert.Tables(0)
                            DataGridView1.Columns("ID").Visible = False
                            Dim nr11 As Integer = 0
                            nr11 = DataGridView1.Rows.Count - 1
                            For i = 0 To nr11
                                DataGridView1.Rows(i).Cells(0).Value = i + 1
                            Next
                            lidhjedg71ofert.Close()
                            Dim nr As Integer = 0
                            nr = DataGridView1.Columns.Count - 1
                            If nr > 16 Then
                                DataGridView1.Columns.RemoveAt(nr)
                            End If
                            With Farmacia.DataGridView7
                                .RowsDefaultCellStyle.BackColor = Color.Bisque
                                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                            End With
                            rreshtiaktualdg71ofert = 0
                            lidhjedg71ofert1.Open()
                            adaptoridg71ofert1 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhjedg71ofert1)
                            adaptoridg71ofert1.Fill(setitedhenavedg71ofert1, "tedhena")
                            Farmacia.DataGridView7.DataSource = setitedhenavedg71ofert1.Tables(0)
                            Farmacia.Merrtedhenat_Ofertat(rreshtiaktualdg71ofert)
                            lidhjedg71ofert1.Close()
                            Dim total2 As Integer
                            Dim total3 As Integer
                            Dim total4 As Integer
                            For ii As Integer = 0 To DataGridView1.RowCount - 1
                                total2 = total2 + DataGridView1.Rows(ii).Cells(12).Value
                            Next
                            For iii As Integer = 0 To DataGridView1.RowCount - 1
                                total3 = total3 + DataGridView1.Rows(iii).Cells(13).Value
                            Next
                            For iiii As Integer = 0 To DataGridView1.RowCount - 1
                                total4 = total4 + DataGridView1.Rows(iiii).Cells(14).Value
                            Next
                            Label37.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                            Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                            Label23.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                            MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                            Dim nr1 As Integer = 0
                            Farmacia.DataGridView7.Columns("ID").Visible = False
                            nr1 = Farmacia.DataGridView7.Rows.Count - 1
                            For i = 0 To nr1
                                Farmacia.DataGridView7.Rows(i).Cells(0).Value = i + 1
                            Next
                        Catch ex As Exception
                            MessageBox.Show("Nuk u shtua")
                            MsgBox(ex.Message & " -  " & ex.Source)
                            lidhjedg71ofert1.Close()
                        End Try
                    Else
                        Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox4.Text
                        Farmacia.DataGridView7.Columns("ID").Visible = True
                        Dim Str1 As String
                        Dim regDate As DateTime = Date.Now
                        Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                        Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                        Dim strDate1 As String = regDate.ToString("HH:mm tt")
                        Try
                            Str1 = "insert into Ofertat values("
                            Str1 += missingNumbers.First.ToString
                            Str1 += ","
                            Str1 += """" & TextBox8.Text & """"
                            Str1 += ","
                            Str1 += """" & strDate.Trim & """"
                            Str1 += ","
                            Str1 += """" & strDate1.Trim & """"
                            Str1 += ","
                            Str1 += """" & ComboBox2.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & TextBox6.Text & """"
                            Str1 += ","
                            Str1 += """" & Label24.Text & """"
                            Str1 += ","
                            Str1 += """" & ComboBox3.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & ComboBox4.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & NumericUpDown1.Value & """"
                            Str1 += ","
                            Str1 += """" & TextBox9.Text & """"
                            Str1 += ","
                            Str1 += """" & TextBox4.Text.Trim & """"
                            Str1 += ","
                            Str1 += """" & vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture) & """"
                            Str1 += ","
                            Str1 += """" & TextBox24.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & NumericUpDown2.Value & """"
                            Str1 += ","
                            Str1 += """" & Hyrje.TextBox1.Text.Trim() & """"
                            Str1 += ")"
                            lidhjedg71ofert.Open()
                            querydg71ofert = New OleDbCommand(Str1, lidhjedg71ofert)
                            querydg71ofert.ExecuteNonQuery()
                            setitedhenavedg71ofert.Clear()
                            adaptoridg71ofert = New OleDbDataAdapter("SELECT * FROM Ofertat WHERE(Kodi_Ofertes LIKE '" & TextBox8.Text & "%')", lidhjedg71ofert)
                            adaptoridg71ofert.Fill(setitedhenavedg71ofert, "tedhena")
                            DataGridView1.Columns.Add("count1", "Nr.")
                            DataGridView1.DataSource = setitedhenavedg71ofert.Tables(0)
                            DataGridView1.Columns("ID").Visible = False
                            Dim nr11 As Integer = 0
                            nr11 = DataGridView1.Rows.Count - 1
                            For i = 0 To nr11
                                DataGridView1.Rows(i).Cells(0).Value = i + 1
                            Next
                            lidhjedg71ofert.Close()
                            Dim nr As Integer = 0
                            nr = DataGridView1.Columns.Count - 1
                            If nr > 16 Then
                                DataGridView1.Columns.RemoveAt(nr)
                            End If
                            With Farmacia.DataGridView7
                                .RowsDefaultCellStyle.BackColor = Color.Bisque
                                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                            End With
                            rreshtiaktualdg71ofert = 0
                            lidhjedg71ofert1.Open()
                            adaptoridg71ofert1 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhjedg71ofert1)
                            adaptoridg71ofert1.Fill(setitedhenavedg71ofert1, "tedhena")
                            Farmacia.DataGridView7.DataSource = setitedhenavedg71ofert1.Tables(0)
                            Farmacia.Merrtedhenat_Ofertat(rreshtiaktualdg71ofert)
                            lidhjedg71ofert1.Close()
                            Dim total2 As Integer
                            Dim total3 As Integer
                            Dim total4 As Integer
                            For ii As Integer = 0 To DataGridView1.RowCount - 1
                                total2 = total2 + DataGridView1.Rows(ii).Cells(12).Value
                            Next
                            For iii As Integer = 0 To DataGridView1.RowCount - 1
                                total3 = total3 + DataGridView1.Rows(iii).Cells(13).Value
                            Next
                            For iiii As Integer = 0 To DataGridView1.RowCount - 1
                                total4 = total4 + DataGridView1.Rows(iiii).Cells(14).Value
                            Next
                            Label37.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                            Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                            Label23.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                            MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                            Dim nr1 As Integer = 0
                            Farmacia.DataGridView7.Columns("ID").Visible = False
                            nr1 = Farmacia.DataGridView7.Rows.Count - 1
                            For i = 0 To nr1
                                Farmacia.DataGridView7.Rows(i).Cells(0).Value = i + 1
                            Next
                        Catch ex As Exception
                            MessageBox.Show("Nuk u shtua")
                            MsgBox(ex.Message & " -  " & ex.Source)
                            lidhjedg71ofert1.Close()
                        End Try
                    End If
                End If
            End If
        End If
        If Not Farmacia.DataGridView7.Rows.Count > 0 Then
            Farmacia.Button19.Enabled = False
            Farmacia.Button18.Enabled = False
        Else
            Farmacia.Button19.Enabled = True
            Farmacia.Button18.Enabled = True
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView1.Rows.Count = 0 Then
        Else
            Dim Str As String
            Try
                With DataGridView1
                    .RowsDefaultCellStyle.BackColor = Color.Bisque
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                End With
                Str = "delete from Ofertat where ID="
                Str += Label33.Text.Trim
                lidhjedg71ofert12.Open()
                querydg71ofert12 = New OleDbCommand(Str, lidhjedg71ofert12)
                querydg71ofert12.ExecuteNonQuery()
                setitedhenavedg71ofert12.clear()
                adaptoridg71ofert12 = New OleDbDataAdapter("SELECT * FROM Ofertat WHERE(Kodi_Ofertes LIKE '" & TextBox8.Text & "%')", lidhjedg71ofert12)
                adaptoridg71ofert12.Fill(setitedhenavedg71ofert12, "tedhena")
                DataGridView1.DataSource = setitedhenavedg71ofert12.Tables(0)
                Dim nrd As Integer = 0
                nrd = DataGridView1.Rows.Count - 1
                For i = 0 To nrd
                    DataGridView1.Rows(i).Cells(0).Value = i + 1
                Next
                lidhjedg71ofert12.Close()
                Dim nr As Integer = 0
                nr = DataGridView1.Columns.Count - 1
                If nr > 16 Then
                    DataGridView1.Columns.RemoveAt(nr)
                End If
                With Farmacia.DataGridView7
                    .RowsDefaultCellStyle.BackColor = Color.Bisque
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                End With
                rreshtiaktualdg71ofert121 = 0
                lidhjedg71ofert121.Open()
                adaptoridg71ofert121 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhjedg71ofert121)
                adaptoridg71ofert121.Fill(setitedhenavedg71ofert121, "tedhena")
                Farmacia.DataGridView7.DataSource = setitedhenavedg71ofert121.Tables(0)
                Farmacia.Merrtedhenat_Ofertat(rreshtiaktualdg71ofert121)
                lidhjedg71ofert121.Close()
                Dim total2 As Integer
                Dim total3 As Integer
                Dim total4 As Integer
                For ii As Integer = 0 To DataGridView1.RowCount - 1
                    total2 = total2 + DataGridView1.Rows(ii).Cells(12).Value
                Next
                For iii As Integer = 0 To DataGridView1.RowCount - 1
                    total3 = total3 + DataGridView1.Rows(iii).Cells(13).Value
                Next
                For iiii As Integer = 0 To DataGridView1.RowCount - 1
                    total4 = total4 + DataGridView1.Rows(iiii).Cells(14).Value
                Next
                Label37.Text = total2
                Label29.Text = total3
                Label23.Text = total4
                MsgBox("U fshi me sukses", MsgBoxStyle.Information)
            Catch ex As Exception
                MsgBox("Nuk u fshi me sukses", MsgBoxStyle.Information)
                lidhjedg71ofert121.Close()
            End Try
        End If
        If Not Farmacia.DataGridView7.Rows.Count > 0 Then
            Farmacia.Button19.Enabled = False
            Farmacia.Button18.Enabled = False
        Else
            Farmacia.Button19.Enabled = True
            Farmacia.Button18.Enabled = True
        End If
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If (e.RowIndex >= 0) Then
            Dim row As DataGridViewRow = DataGridView1.Rows.Item(e.RowIndex)
            Label33.Text = row.Cells.Item("ID").Value.ToString
        End If
    End Sub
    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        If CheckBox1.Checked = True Then
            If TextBox9.Text = "" Then
            Else
                Dim zbritje = NumericUpDown1.Value * TextBox9.Text
                Dim total1 As Decimal = zbritje * (1 - NumericUpDown2.Value / 100)
                Dim totalizbriturcmimit As Decimal = TextBox9.Text * (1 - NumericUpDown2.Value / 100)
                Label26.Text = totalizbriturcmimit
                Label26.ForeColor = Color.Red
                Dim vlera_shitjes = total1
                TextBox4.Text = vlera_shitjes

                Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox4.Text
                Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + TextBox4.Text
                TextBox23.Text = vlera_shitjes_metvsh
                TextBox24.Text = vlera_shitjes_metvsh1
            End If
        Else
            If TextBox9.Text = "" Then
            Else
                Dim vlera_shitjes = NumericUpDown1.Value * TextBox9.Text
                TextBox4.Text = vlera_shitjes
                Label26.Text = 0
                Label26.ForeColor = DefaultForeColor
                Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox4.Text
                Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + TextBox4.Text
                TextBox23.Text = vlera_shitjes_metvsh
                TextBox24.Text = vlera_shitjes_metvsh1
            End If
        End If
    End Sub
    Public lidhje211of As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptori211of As OleDbDataAdapter
    Public lexuesi11of As OleDbDataReader
    Public query211of As OleDbCommand
    Public setitedhenave211of = New DataSet
    Public rreshtiaktual211of As Integer
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Not DataGridView1.Rows.Count > 0 Then
            MsgBox("Fatura eshte bosh!Me pare gjeneroni nje fature!", MsgBoxStyle.Information)
        Else
            Dim regDate As DateTime = Date.Now
            Dim strDate As String = regDate.ToString("dd/MM/yyyy")
            If My.Settings.logo = "" Or My.Settings.faturatofert = "" Then
                MsgBox("Ju nuk keni zgjedhur nje logo ose vendodhjen se ku do te ruhen faturat.Shko tek Konfigurimet!", MsgBoxStyle.Information)
            Else
                Try
                    DataGridView2.DataSource = Nothing
                    Konfigurime.TextBox1.Text = My.Settings.backup
                    Konfigurime.TextBox2.Text = My.Settings.logo
                    Konfigurime.TextBox3.Text = My.Settings.ruajraportet
                    Konfigurime.TextBox4.Text = My.Settings.faturatofert
                    lidhje211of.Open()
                    setitedhenave211of.Clear
                    adaptori211of = New OleDbDataAdapter("SELECT Produkti,Njesia,Sasia,Cmimi,Vlera_Pa_TVSH,TVSH,Vlera_Me_TVSH,Zbritje_ne_perq FROM Ofertat WHERE(Kodi_Ofertes Like '" & TextBox8.Text & "%')", lidhje211of)
                    adaptori211of.Fill(setitedhenave211of, "tedhena")
                    DataGridView2.Columns.Clear()
                    DataGridView2.Columns.Add("count1", "Nr.")
                    DataGridView2.DataSource = setitedhenave211of.Tables(0)
                    'DataGridView8.Columns("ID").Visible = False
                    lidhje211of.Close()
                    Dim nr9 As Integer = 0
                    nr9 = DataGridView2.Rows.Count - 1
                    For i = 0 To nr9
                        DataGridView2.Rows(i).Cells(0).Value = i + 1
                    Next
                    DataGridView2.Refresh()
                    vlerapatvsh = (From row As DataGridViewRow In DataGridView2.Rows
                                   Where row.Cells(5).FormattedValue.ToString() <> String.Empty
                                   Select Convert.ToInt32(row.Cells(5).FormattedValue)).Sum().ToString()
                    vleraetvsh = (From row As DataGridViewRow In DataGridView2.Rows
                                  Where row.Cells(6).FormattedValue.ToString() <> String.Empty
                                  Select Convert.ToInt32(row.Cells(6).FormattedValue)).Sum().ToString()
                    vlerametvsh = (From row As DataGridViewRow In DataGridView2.Rows
                                   Where row.Cells(7).FormattedValue.ToString() <> String.Empty
                                   Select Convert.ToInt32(row.Cells(7).FormattedValue)).Sum().ToString()
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(5).Value = vlerapatvsh
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(6).Value = vleraetvsh
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(7).Value = vlerametvsh
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(4).Value = "TOTALI"
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(0).Value = ""
                    Dim con_bleresit As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
                    Dim query_bleresit As String = "SELECT TOP 1 Data,Shitesi,Bleresi,Kodi_Ofertes FROM Ofertat WHERE(Kodi_Ofertes Like '" & TextBox8.Text & "%')"
                    Using connection As New OleDbConnection(con_bleresit)
                        Dim command As New OleDbCommand(query_bleresit, connection)
                        connection.Open()
                        Dim readerx As OleDbDataReader = command.ExecuteReader()
                        While readerx.Read()
                            Dim con_bleresit1 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
                            Dim query_bleresit1 As String = "SELECT * FROM Shitesi WHERE(Emri LIKE '" & readerx(1).ToString() & "%')"
                            data = readerx(0).ToString()
                            kodi_shit = readerx(3).ToString()
                            Using connection1 As New OleDbConnection(con_bleresit1)
                                Dim command1 As New OleDbCommand(query_bleresit1, connection1)
                                connection1.Open()
                                Dim reader1 As OleDbDataReader = command1.ExecuteReader()
                                While reader1.Read()
                                    shitesi = reader1(1).ToString()
                                    shita = reader1(2).ToString()
                                    shitc = reader1(3).ToString()
                                    shitn = reader1(4).ToString()
                                End While
                                reader1.Close()
                            End Using
                            Dim con_bleresit2 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
                            Dim query_bleresit2 As String = "SELECT * FROM Bleresit WHERE(Emri LIKE '" & readerx(2).ToString() & "%')"
                            Using connection2 As New OleDbConnection(con_bleresit2)
                                Dim command2 As New OleDbCommand(query_bleresit2, connection2)
                                connection2.Open()
                                Dim reader2 As OleDbDataReader = command2.ExecuteReader()
                                While reader2.Read()
                                    bleresi = reader2(1).ToString()
                                    bleresia = reader2(2).ToString()
                                    bleresic = reader2(3).ToString()
                                    bleresin = reader2(4).ToString()
                                End While
                                reader2.Close()
                            End Using
                        End While
                        readerx.Close()
                        Dim pdfTable As New PdfPTable(DataGridView2.ColumnCount)
                        pdfTable.DefaultCell.Padding = 3
                        pdfTable.WidthPercentage = 100
                        pdfTable.HorizontalAlignment = Element.ALIGN_LEFT
                        'Adding Header row
                        For Each column As DataGridViewColumn In DataGridView2.Columns
                            Dim cell As New PdfPCell(New Phrase(column.HeaderText.Replace("_", " ").Replace("ne perq", "%")))
                            cell.BorderWidthTop = 1
                            cell.BorderWidthBottom = 1
                            cell.BackgroundColor = New iTextSharp.text.BaseColor(250, 235, 215)
                            pdfTable.AddCell(cell)
                        Next
                        'Adding DataRow
                        Dim cellvalue As String = ""
                        Dim i As Integer = 0
                        For Each row As DataGridViewRow In DataGridView2.Rows
                            For Each cell As DataGridViewCell In row.Cells
                                cellvalue = cell.FormattedValue
                                pdfTable.AddCell(Convert.ToString(cellvalue))
                            Next
                        Next
                        'Exporting to PDF
                        Dim folderPath As String = Konfigurime.TextBox4.Text
                        If Not Directory.Exists(folderPath) Then
                            Directory.CreateDirectory(folderPath)
                        End If
                        Using stream As New FileStream(folderPath & "\DataGridViewExport.pdf", FileMode.Create)
                            Dim pdfDoc As Document = New Document(PageSize.A4, 10.0F, 10.0F, 380, 0F)
                            PdfWriter.GetInstance(pdfDoc, stream)
                            Dim pdfDest As PdfDestination = New PdfDestination(PdfDestination.XYZ, 0, pdfDoc.PageSize.Height, 1.0F)
                            pdfDoc.Open()
                            pdfDoc.Add(pdfTable)
                            pdfDoc.Close()
                            stream.Close()
                        End Using
                        Dim oldFile As String = Konfigurime.TextBox4.Text & "\DataGridViewExport.pdf"
                        Dim newFile As String = Konfigurime.TextBox4.Text & "\DataGridViewExport1.pdf"
                        Dim reader As New PdfReader(oldFile)
                        Dim size As Rectangle = reader.GetPageSizeWithRotation(1)
                        Dim document As New Document(size)
                        Dim fs As New FileStream(newFile, FileMode.Create, FileAccess.Write)
                        Dim writer As PdfWriter = PdfWriter.GetInstance(document, fs)
                        document.Open()
                        Dim cb As PdfContentByte = writer.DirectContent
                        Dim bf As BaseFont = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)
                        cb.SetColorFill(BaseColor.BLACK)
                        cb.SetFontAndSize(bf, 11)
                        cb.BeginText()
                        'Emri shitesit
                        Dim emri_shitesit As String = shitesi
                        cb.ShowTextAligned(1, emri_shitesit, 99, 597, 0)
                        Dim ad_shitesit As String = shita
                        cb.ShowTextAligned(1, ad_shitesit, 102, 578, 0)
                        Dim cel_shitesit As String = shitc
                        cb.ShowTextAligned(1, cel_shitesit, 95, 559, 0)
                        Dim nipt_shitesit As String = shitn
                        cb.ShowTextAligned(1, nipt_shitesit, 92, 541, 0)
                        Dim emri_bleresit As String = bleresi
                        cb.ShowTextAligned(1, emri_bleresit, 353, 597, 0)
                        Dim ad_bleres As String = bleresia
                        cb.ShowTextAligned(1, ad_bleres, 357, 577, 0)
                        Dim cel_ble As String = bleresic
                        cb.ShowTextAligned(1, cel_ble, 353, 558, 0)
                        Dim nipt_ble As String = bleresin
                        cb.ShowTextAligned(1, nipt_ble, 355, 540, 0)
                        'Data
                        Dim data1 As String = data
                        cb.ShowTextAligned(1, data1, 521, 707, 0)
                        'nr fatures
                        Dim kod As String = kodi_shit
                        cb.ShowTextAligned(1, kod, 522, 670, 0)
                        Dim ble As String = "BLERESI"
                        cb.ShowTextAligned(1, ble, 88, 40, 0)
                        Dim vij1 As String = "___________________________"
                        cb.ShowTextAligned(3, vij1, 20, 65, 0)
                        Dim shits As String = "SHITESI"
                        cb.ShowTextAligned(1, shits, 500, 40, 0)
                        Dim vij2 As String = "___________________________"
                        cb.ShowTextAligned(3, vij2, 425, 65, 0)
                        cb.EndText()
                        Dim page As PdfImportedPage = writer.GetImportedPage(reader, 1)
                        cb.AddTemplate(page, 0, 0)
                        document.Close()
                        fs.Close()
                        writer.Close()
                        reader.Close()
                        My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox4.Text & "\DataGridViewExport.pdf")
                    End Using
                    Using inputPdfStream As Stream = New FileStream(Konfigurime.TextBox4.Text & "\DataGridViewExport1.pdf", FileMode.Open, FileAccess.Read, FileShare.Read)
                        My.Resources.xxoferte.Save("C:\ProgramData\xxoferte.jpg", Drawing.Imaging.ImageFormat.Jpeg)
                        Using inputImageStream As Stream = New FileStream("C:\ProgramData\xxoferte.jpg", FileMode.Open, FileAccess.Read, FileShare.Read)

                            Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox4.Text & "\DataGridViewExport3.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
                                Dim readers = New PdfReader(inputPdfStream)
                                Dim stamper = New PdfStamper(readers, outputPdfStream)
                                Dim pdfContentByte = stamper.GetUnderContent(1)
                                Dim image As Image = Image.GetInstance(inputImageStream)
                                image.Alignment = Image.UNDERLYING
                                image.SetAbsolutePosition(0, 0)
                                image.ScaleAbsolute(iTextSharp.text.PageSize.A4.Width, iTextSharp.text.PageSize.A4.Height)
                                pdfContentByte.AddImage(image)
                                stamper.Close()
                            End Using
                        End Using
                    End Using
                    My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox4.Text & "\DataGridViewExport1.pdf")
                    Using inputPdfStream As Stream = New FileStream(Konfigurime.TextBox4.Text & "\DataGridViewExport3.pdf", FileMode.Open, FileAccess.Read, FileShare.Read)
                        Using inputImageStream As Stream = New FileStream(Konfigurime.TextBox2.Text, FileMode.Open, FileAccess.Read, FileShare.Read)
                            Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox4.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Oferte.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
                                Dim readers = New PdfReader(inputPdfStream)
                                Dim stamper = New PdfStamper(readers, outputPdfStream)
                                Dim pdfContentByte = stamper.GetOverContent(1)
                                Dim image As Image = Image.GetInstance(inputImageStream)
                                image.Alignment = Image.UNDERLYING
                                image.SetAbsolutePosition(30, 695)
                                image.ScaleAbsolute(150, 100)
                                pdfContentByte.AddImage(image)
                                stamper.Close()
                            End Using
                        End Using
                    End Using
                    My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox4.Text & "\DataGridViewExport3.pdf")
                    My.Computer.FileSystem.DeleteFile("C:\ProgramData\xxoferte.jpg")
                    System.Diagnostics.Process.Start(Konfigurime.TextBox4.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Oferte.pdf")
                Catch ex As Exception
                    MsgBox("Dokumenti eshte i hapur.Mbyll dokumentin dhe provo perseri!", MsgBoxStyle.Information)
                End Try
            End If
        End If
    End Sub
    Dim data, shitesi, bleresi, kodi_shit, shita, shitc, shitn, bleresia, bleresic, bleresin As String
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Not DataGridView1.Rows.Count > 0 Then
            MsgBox("Fatura eshte bosh!Me pare gjeneroni nje fature!", MsgBoxStyle.Information)
        Else
            Dim regDate As DateTime = Date.Now
            Dim strDate As String = regDate.ToString("dd/MM/yyyy")
            If My.Settings.logo = "" Or My.Settings.faturatofert = "" Then
                MsgBox("Ju nuk keni zgjedhur nje logo ose vendodhjen se ku do te ruhen faturat.Shko tek Konfigurimet!", MsgBoxStyle.Information)
            Else
                Try
                    DataGridView2.DataSource = Nothing
                    Konfigurime.TextBox1.Text = My.Settings.backup
                    Konfigurime.TextBox2.Text = My.Settings.logo
                    Konfigurime.TextBox3.Text = My.Settings.ruajraportet
                    Konfigurime.TextBox4.Text = My.Settings.faturatofert
                    lidhje211of.Open()
                    setitedhenave211of.Clear
                    adaptori211of = New OleDbDataAdapter("SELECT Produkti,Njesia,Sasia,Cmimi,Vlera_Pa_TVSH,TVSH,Vlera_Me_TVSH,Zbritje_ne_perq FROM Ofertat WHERE(Kodi_Ofertes Like '" & TextBox8.Text & "%')", lidhje211of)
                    adaptori211of.Fill(setitedhenave211of, "tedhena")
                    DataGridView2.Columns.Clear()
                    DataGridView2.Columns.Add("count1", "Nr.")
                    DataGridView2.DataSource = setitedhenave211of.Tables(0)
                    'DataGridView8.Columns("ID").Visible = False
                    lidhje211of.Close()
                    Dim nr9 As Integer = 0
                    nr9 = DataGridView2.Rows.Count - 1
                    For i = 0 To nr9
                        DataGridView2.Rows(i).Cells(0).Value = i + 1
                    Next
                    DataGridView2.Refresh()
                    vlerapatvsh = (From row As DataGridViewRow In DataGridView2.Rows
                                   Where row.Cells(5).FormattedValue.ToString() <> String.Empty
                                   Select Convert.ToInt32(row.Cells(5).FormattedValue)).Sum().ToString()
                    vleraetvsh = (From row As DataGridViewRow In DataGridView2.Rows
                                  Where row.Cells(6).FormattedValue.ToString() <> String.Empty
                                  Select Convert.ToInt32(row.Cells(6).FormattedValue)).Sum().ToString()
                    vlerametvsh = (From row As DataGridViewRow In DataGridView2.Rows
                                   Where row.Cells(7).FormattedValue.ToString() <> String.Empty
                                   Select Convert.ToInt32(row.Cells(7).FormattedValue)).Sum().ToString()
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(5).Value = vlerapatvsh
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(6).Value = vleraetvsh
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(7).Value = vlerametvsh
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(4).Value = "TOTALI"
                    DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(0).Value = ""
                    Dim con_bleresit As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
                    Dim query_bleresit As String = "SELECT TOP 1 Data,Shitesi,Bleresi,Kodi_Ofertes FROM Ofertat WHERE(Kodi_Ofertes Like '" & TextBox8.Text & "%')"
                    Using connection As New OleDbConnection(con_bleresit)
                        Dim command As New OleDbCommand(query_bleresit, connection)
                        connection.Open()
                        Dim readerx As OleDbDataReader = command.ExecuteReader()
                        While readerx.Read()
                            Dim con_bleresit1 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
                            Dim query_bleresit1 As String = "SELECT * FROM Shitesi WHERE(Emri LIKE '" & readerx(1).ToString() & "%')"
                            data = readerx(0).ToString()
                            kodi_shit = readerx(3).ToString()
                            Using connection1 As New OleDbConnection(con_bleresit1)
                                Dim command1 As New OleDbCommand(query_bleresit1, connection1)
                                connection1.Open()
                                Dim reader1 As OleDbDataReader = command1.ExecuteReader()
                                While reader1.Read()
                                    shitesi = reader1(1).ToString()
                                    shita = reader1(2).ToString()
                                    shitc = reader1(3).ToString()
                                    shitn = reader1(4).ToString()
                                End While
                                reader1.Close()
                            End Using
                            Dim con_bleresit2 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
                            Dim query_bleresit2 As String = "SELECT * FROM Bleresit WHERE(Emri LIKE '" & readerx(2).ToString() & "%')"
                            Using connection2 As New OleDbConnection(con_bleresit2)
                                Dim command2 As New OleDbCommand(query_bleresit2, connection2)
                                connection2.Open()
                                Dim reader2 As OleDbDataReader = command2.ExecuteReader()
                                While reader2.Read()
                                    bleresi = reader2(1).ToString()
                                    bleresia = reader2(2).ToString()
                                    bleresic = reader2(3).ToString()
                                    bleresin = reader2(4).ToString()
                                End While
                                reader2.Close()
                            End Using
                        End While
                        readerx.Close()
                        Dim pdfTable As New PdfPTable(DataGridView2.ColumnCount)
                        pdfTable.DefaultCell.Padding = 3
                        pdfTable.WidthPercentage = 100
                        pdfTable.HorizontalAlignment = Element.ALIGN_LEFT
                        'Adding Header row
                        For Each column As DataGridViewColumn In DataGridView2.Columns
                            Dim cell As New PdfPCell(New Phrase(column.HeaderText.Replace("_", " ").Replace("ne perq", "%")))
                            cell.BorderWidthTop = 1
                            cell.BorderWidthBottom = 1
                            cell.BackgroundColor = New iTextSharp.text.BaseColor(250, 235, 215)
                            pdfTable.AddCell(cell)
                        Next
                        'Adding DataRow
                        Dim cellvalue As String = ""
                        Dim i As Integer = 0
                        For Each row As DataGridViewRow In DataGridView2.Rows
                            For Each cell As DataGridViewCell In row.Cells
                                cellvalue = cell.FormattedValue
                                pdfTable.AddCell(Convert.ToString(cellvalue))
                            Next
                        Next
                        'Exporting to PDF
                        Dim folderPath As String = Konfigurime.TextBox4.Text
                        If Not Directory.Exists(folderPath) Then
                            Directory.CreateDirectory(folderPath)
                        End If
                        Using stream As New FileStream(folderPath & "\DataGridViewExport.pdf", FileMode.Create)
                            Dim pdfDoc As Document = New Document(PageSize.A4, 10.0F, 10.0F, 380, 0F)
                            PdfWriter.GetInstance(pdfDoc, stream)
                            pdfDoc.Open()
                            pdfDoc.Add(pdfTable)
                            pdfDoc.Close()
                            stream.Close()
                        End Using
                        Dim oldFile As String = Konfigurime.TextBox4.Text & "\DataGridViewExport.pdf"
                        Dim newFile As String = Konfigurime.TextBox4.Text & "\DataGridViewExport1.pdf"
                        Dim reader As New PdfReader(oldFile)
                        Dim size As Rectangle = reader.GetPageSizeWithRotation(1)
                        Dim document As New Document(size)
                        Dim fs As New FileStream(newFile, FileMode.Create, FileAccess.Write)
                        Dim writer As PdfWriter = PdfWriter.GetInstance(document, fs)
                        document.Open()
                        Dim cb As PdfContentByte = writer.DirectContent
                        Dim bf As BaseFont = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)
                        cb.SetColorFill(BaseColor.BLACK)
                        cb.SetFontAndSize(bf, 11)
                        cb.BeginText()
                        'Emri shitesit
                        Dim emri_shitesit As String = shitesi
                        cb.ShowTextAligned(1, emri_shitesit, 99, 597, 0)
                        Dim ad_shitesit As String = shita
                        cb.ShowTextAligned(1, ad_shitesit, 102, 578, 0)
                        Dim cel_shitesit As String = shitc
                        cb.ShowTextAligned(1, cel_shitesit, 95, 559, 0)
                        Dim nipt_shitesit As String = shitn
                        cb.ShowTextAligned(1, nipt_shitesit, 92, 541, 0)
                        Dim emri_bleresit As String = bleresi
                        cb.ShowTextAligned(1, emri_bleresit, 353, 597, 0)
                        Dim ad_bleres As String = bleresia
                        cb.ShowTextAligned(1, ad_bleres, 357, 577, 0)
                        Dim cel_ble As String = bleresic
                        cb.ShowTextAligned(1, cel_ble, 353, 558, 0)
                        Dim nipt_ble As String = bleresin
                        cb.ShowTextAligned(1, nipt_ble, 355, 540, 0)
                        'Data
                        Dim data1 As String = data
                        cb.ShowTextAligned(1, data1, 521, 707, 0)
                        'nr fatures
                        Dim kod As String = kodi_shit
                        cb.ShowTextAligned(1, kod, 522, 670, 0)
                        Dim ble As String = "BLERESI"
                        cb.ShowTextAligned(1, ble, 88, 40, 0)
                        Dim vij1 As String = "___________________________"
                        cb.ShowTextAligned(3, vij1, 20, 65, 0)
                        Dim shits As String = "SHITESI"
                        cb.ShowTextAligned(1, shits, 500, 40, 0)
                        Dim vij2 As String = "___________________________"
                        cb.ShowTextAligned(3, vij2, 425, 65, 0)
                        cb.EndText()
                        Dim page As PdfImportedPage = writer.GetImportedPage(reader, 1)
                        cb.AddTemplate(page, 0, 0)
                        document.Close()
                        fs.Close()
                        writer.Close()
                        reader.Close()
                        My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox4.Text & "\DataGridViewExport.pdf")
                    End Using
                    Using inputPdfStream As Stream = New FileStream(Konfigurime.TextBox4.Text & "\DataGridViewExport1.pdf", FileMode.Open, FileAccess.Read, FileShare.Read)
                        My.Resources.xxoferte.Save("C:\ProgramData\xxoferte.jpg", Drawing.Imaging.ImageFormat.Jpeg)
                        Using inputImageStream As Stream = New FileStream("C:\ProgramData\xxoferte.jpg", FileMode.Open, FileAccess.Read, FileShare.Read)
                            Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox4.Text & "\DataGridViewExport3.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
                                Dim readers = New PdfReader(inputPdfStream)
                                Dim stamper = New PdfStamper(readers, outputPdfStream)
                                Dim pdfContentByte = stamper.GetUnderContent(1)
                                Dim image As Image = Image.GetInstance(inputImageStream)
                                image.Alignment = Image.UNDERLYING
                                image.SetAbsolutePosition(0, 0)
                                image.ScaleAbsolute(iTextSharp.text.PageSize.A4.Width, iTextSharp.text.PageSize.A4.Height)
                                pdfContentByte.AddImage(image)
                                stamper.Close()
                            End Using
                        End Using
                    End Using
                    My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox4.Text & "\DataGridViewExport1.pdf")
                    Using inputPdfStream As Stream = New FileStream(Konfigurime.TextBox4.Text & "\DataGridViewExport3.pdf", FileMode.Open, FileAccess.Read, FileShare.Read)
                        Using inputImageStream As Stream = New FileStream(Konfigurime.TextBox2.Text, FileMode.Open, FileAccess.Read, FileShare.Read)
                            Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox4.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Oferte.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
                                Dim readers = New PdfReader(inputPdfStream)
                                Dim stamper = New PdfStamper(readers, outputPdfStream)
                                Dim pdfContentByte = stamper.GetOverContent(1)
                                Dim image As Image = Image.GetInstance(inputImageStream)
                                image.Alignment = Image.UNDERLYING
                                image.SetAbsolutePosition(30, 695)
                                image.ScaleAbsolute(150, 100)
                                pdfContentByte.AddImage(image)
                                stamper.Close()
                            End Using
                        End Using
                    End Using
                    My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox4.Text & "\DataGridViewExport3.pdf")
                    My.Computer.FileSystem.DeleteFile("C:\ProgramData\xxoferte.jpg")
                    Dim regDate1 As DateTime = Date.Now
                    Dim strDate1 As String = regDate1.ToString("dd/MM/yyyy")
                    Dim PrintPDF As New ProcessStartInfo
                    PrintPDF.UseShellExecute = True
                    PrintPDF.Verb = "print"
                    PrintPDF.WindowStyle = ProcessWindowStyle.Hidden
                    PrintPDF.FileName = Konfigurime.TextBox4.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate1.Replace("/", "-") & "_" & TextBox8.Text & "_Oferte.pdf"
                    Process.Start(PrintPDF)
                    Threading.Thread.Sleep(20000)
                    killProcess("Acrobat")
                    Threading.Thread.Sleep(10000)
                    My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox4.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate1.Replace("/", "-") & "_" & TextBox8.Text & "_Oferte.pdf")
                    MsgBox("Fatura u printua me sukses!", MsgBoxStyle.Information)
                    ' System.Diagnostics.Process.Start(Konfigurime.TextBox4.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Oferte.pdf")
                Catch ex As Exception
                    MsgBox("Dokumenti eshte i hapur.Mbyll dokumentin dhe provo perseri!", MsgBoxStyle.Information)
                End Try
            End If
        End If
    End Sub
    Private Sub killProcess(ByVal processName As String)
        Dim procesos As Process()
        procesos = Process.GetProcessesByName(processName) 'I used "AcroRd32" as parameter
        If procesos.Length > 0 Then
            For i = procesos.Length - 1 To 0 Step -1
                procesos(i).Kill()
            Next
        End If
    End Sub
    Dim vlerapatvsh, vleraetvsh, vlerametvsh As String
End Class