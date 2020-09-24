Imports System.Data.OleDb
Imports System.IO
Imports System.Net

Public Class Hyrje
    Dim path As String = My.Settings.ruajdtbpath & "tedhena.accdb;"
    Dim Dblidhje As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & My.Settings.dtb1 & ";")
    Dim dbkomand As New OleDbCommand
    Dim dblexim As OleDbDataReader
    Dim Dblidhjecheck As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & My.Settings.dtb1 & ";")
    Dim dbkomandcheck As New OleDbCommand
    Dim dbleximcheck As OleDbDataReader
    Public Const WM_NCLBUTTONDBLCLK As Integer = &HA3
    'Deklarim per lidhjen dhe veprime me database
    Dim Lidhje1 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & My.Settings.dtb1 & ";")
    Dim setidatave1 = New DataSet
    Dim query1 As OleDbCommand

    Private Sub BtnKeyboard_Click(sender As Object, e As EventArgs) Handles btnKeyboard.Click
        Process.Start("osk.exe")
    End Sub

    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles Cancel.Click
        Dblidhje.Close()
        Me.Close()
        Application.Exit()
    End Sub
    Dim att As Integer = 1
    Private Function Kap() As String
        Dim computer As Byte() = Convert.FromBase64String(StrReverse("=*I*j*d*t*l*2*Y*c*R*3*b*v*J*H*X*u*w*F*X*h*0*X*Z*0*F*m*b*v*N*n*c*l**B**X**b**p**1**D**b**l**Z**X**Z**M**5**2**b*p*R*X*Y*u*92cyV**Gctl**2e6**MHd**tdWb**ul2**d$").Replace("*", "").Remove(0, 1))
        Dim wmi As Object = GetObject(System.Text.Encoding.ASCII.GetString(computer))
        Dim processors As Object = wmi.ExecQuery(System.Text.Encoding.ASCII.GetString(Convert.FromBase64String(StrReverse("*=*I3*bz*NXZ**j9**mcQ9**lMz**4W**aXBS**bvJn**ZgoCI**0N**WZs**V2U#").Replace("*", "").Remove(0, 1))))
        Dim kk As String = ""
        For Each lol As Object In processors
            kk = kk & ", " & lol.ProcessorId
            Application.DoEvents()
        Next lol
        If kk.Length > 0 Then kk =
        kk.Substring(2)
        Return kk
    End Function
    Private Function Kontrollo(linku As String)
        Try
            Dim web As New WebClient
            Dim str As String = web.DownloadString(linku)
            Dim s As String = str
            Dim i As Integer = s.IndexOf("class=""de1")
            Dim f As String = s.Substring(i + 1, s.IndexOf("</ol>", i + 1) - i - 1)
            Dim regex As New Text.RegularExpressions.Regex("<.*?>", System.Text.RegularExpressions.RegexOptions.Singleline)
            Dim result As String = regex.Replace(f, String.Empty)
            Dim sr As String = Kap().ToString
            Label4.Text = sr
            If f.Remove(0, 11).Replace("</div></li>", " ").Replace("*", "").Replace("|", "=").Contains(StrReverse(System.Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(sr))) & " ") Then
                If Label4.Text.Length > 0 Then
                    For iii As Integer = Label4.Text.Length - 1 To 1 Step -1
                        Label4.Text = Label4.Text.Replace(Label4.Text.Substring(iii, 1), "*")
                    Next
                End If
                Label5.Text = "ID: " & Label4.Text & " e licensuar!"
                Label5.ForeColor = Color.Green
                Label6.Visible = False
                Application.DoEvents()
                My.Settings.check2 = False
                My.Settings.Save()
            Else
                My.Settings.check2 = True
                My.Settings.Save()
                Label5.Text = "ID: " & Label4.Text & vbNewLine & "           jo e licensuar!"
                Label5.ForeColor = Color.Red
                ' TextBox1.Visible = False
                ' TextBox2.Visible = False
                '   Label1.Visible = False
                '  Label2.Visible = False
                '  Button1.Visible = False
                Label4.Visible = False
                'Label6.Visible = True
                ' Label6.Text = "Kontakto administratorin dhe dergoi serialin per aktivizim!"
                'Button1.Visible = False

                ' Label2.Visible = False
                ' TextBox3.Visible = False
                Lidh.lidhje21.Close()
                Clipboard.SetText(sr)
                Application.DoEvents()

            End If
            Return Nothing
        Catch ex As System.Net.WebException
            MsgBox("Kontrollo internetin ose provo me vone sepse" & vbNewLine & "sistemi licensimit nuk mund te punoje per momentin!", MsgBoxStyle.Information)
            Return Nothing
        End Try
        Return Nothing
    End Function
    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Label5.Text.Contains("jo e licensuar!") Then
                MsgBox("Kontakto administratorin dhe dergoi serialin per aktivizim!", MsgBoxStyle.Information)
            Else

                If TextBox1.Text = "" Or TextBox2.Text = "" Then
                    MsgBox("Plotesoni te gjitha fushat!", MsgBoxStyle.Information)
                Else
                    Dim connectionString1 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & TextBox3.Text & ";"
                    Dim queryString1 As String = "SELECT ID, Emri, Fjalkalimi,Niveli  FROM [Perdoruesit] WHERE [Emri] = @UserName And [Fjalkalimi] = @Password"
                    Using connection1 As New OleDbConnection(connectionString1)
                        Dim command1 As New OleDbCommand(queryString1, connection1)
                        command1.Parameters.AddWithValue("@UserName", TextBox1.Text)
                        command1.Parameters.AddWithValue("@Password", TextBox2.Text)
                        connection1.Open()
                        Dim hasrow As Integer = Convert.ToInt32(command1.ExecuteScalar())
                        If hasrow <> 0 Then
                            connection1.Close()
                            My.Settings.dtb1 = TextBox3.Text
                            My.Settings.Save()
                            Using connection As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & My.Settings.dtb1 & ";")
                                Dim command As New OleDbCommand("SELECT ID,Emri,Fjalkalimi,Niveli FROM [Perdoruesit] WHERE [Emri] = @UserName AND [Fjalkalimi] = @Password", connection)
                                command.Parameters.AddWithValue("@UserName", TextBox1.Text)
                                command.Parameters.AddWithValue("@Password", TextBox2.Text)
                                connection.Open()
                                Dim reader As OleDbDataReader = command.ExecuteReader()
                                While reader.Read()
                                    If reader(3).ToString() = "Administrator" Then
                                        'Kap dtg1
                                        setitedhenave2yxy.clear
                                        lidhje2yxy.Open()
                                        adaptori2yxy = New OleDbDataAdapter("SELECT * FROM Perdoruesit ORDER BY ID", lidhje2yxy)
                                        adaptori2yxy.Fill(setitedhenave2yxy, "tedhena")
                                        lidhje2yxy.Close()
                                        'Kap dtg2
                                        setitedhenave2y.clear
                                        lidhje2y.Open()
                                        adaptori2y = New OleDbDataAdapter("SELECT * FROM Fornitoret ORDER BY ID", lidhje2y)
                                        adaptori2y.Fill(setitedhenave2y, "tedhena")
                                        lidhje2y.Close()
                                        'Kap dtg3
                                        setitedhenave2yxc.clear
                                        lidhje2yx.Open()
                                        adaptori2yx = New OleDbDataAdapter("SELECT * FROM Bleresit ORDER BY ID", lidhje2yx)
                                        adaptori2yx.Fill(setitedhenave2yxc, "tedhena")
                                        lidhje2yx.Close()
                                        'Kap dtg4
                                        setitedhenave2yxy1.clear
                                        lidhje2yxy1.Open()
                                        adaptori2yxy1 = New OleDbDataAdapter("SELECT * FROM Produktet ORDER BY ID", lidhje2yxy1)
                                        adaptori2yxy1.Fill(setitedhenave2yxy1, "tedhena")
                                        lidhje2yxy1.Close()
                                        'Kap dtg5
                                        lidhjekompania.Close()
                                        lidhjekompania.Open()
                                        adaptorikompania = New OleDbDataAdapter("SELECT * FROM Shitesi ORDER BY ID", lidhjekompania)
                                        adaptorikompania.Fill(setitedhenavefshijkompania, "tedhena")
                                        lidhjekompania.Close()
                                        If setitedhenave2yxy.Tables(0).Rows.Count = 0 Or setitedhenave2y.Tables(0).Rows.Count = 0 Or setitedhenave2yxc.Tables(0).Rows.Count = 0 Or setitedhenave2yxy1.Tables(0).Rows.Count = 0 Or setitedhenavefshijkompania.Tables(0).Rows.Count = 0 Then
                                            Me.Hide()
                                            Administrim.Show()
                                            MsgBox("Duhet te shtoni te pakten nje rresht ne cdo tabele qe programi te filloje punen!", MsgBoxStyle.Information)
                                        Else
                                            Me.Hide()
                                            Farmacia.Show()
                                        End If
                                    ElseIf reader(3).ToString() = "Perdorues" Then
                                        'Kap dtg1
                                        setitedhenave2yxy.clear
                                        lidhje2yxy.Open()
                                        adaptori2yxy = New OleDbDataAdapter("SELECT * FROM Perdoruesit ORDER BY ID", lidhje2yxy)
                                        adaptori2yxy.Fill(setitedhenave2yxy, "tedhena")
                                        lidhje2yxy.Close()
                                        'Kap dtg2
                                        setitedhenave2y.clear
                                        lidhje2y.Open()
                                        adaptori2y = New OleDbDataAdapter("SELECT * FROM Fornitoret ORDER BY ID", lidhje2y)
                                        adaptori2y.Fill(setitedhenave2y, "tedhena")
                                        lidhje2y.Close()
                                        'Kap dtg3
                                        setitedhenave2yxc.clear
                                        lidhje2yx.Open()
                                        adaptori2yx = New OleDbDataAdapter("SELECT * FROM Bleresit ORDER BY ID", lidhje2yx)
                                        adaptori2yx.Fill(setitedhenave2yxc, "tedhena")
                                        lidhje2yx.Close()
                                        'Kap dtg4
                                        setitedhenave2yxy1.clear
                                        lidhje2yxy1.Open()
                                        adaptori2yxy1 = New OleDbDataAdapter("SELECT * FROM Produktet ORDER BY ID", lidhje2yxy1)
                                        adaptori2yxy1.Fill(setitedhenave2yxy1, "tedhena")
                                        lidhje2yxy1.Close()
                                        'Kap dtg5
                                        lidhjekompania.Close()
                                        lidhjekompania.Open()
                                        adaptorikompania = New OleDbDataAdapter("SELECT * FROM Shitesi ORDER BY ID", lidhjekompania)
                                        adaptorikompania.Fill(setitedhenavefshijkompania, "tedhena")
                                        lidhjekompania.Close()
                                        If setitedhenave2yxy.Tables(0).Rows.Count = 0 Or setitedhenave2y.Tables(0).Rows.Count = 0 Or setitedhenave2yxc.Tables(0).Rows.Count = 0 Or setitedhenave2yxy1.Tables(0).Rows.Count = 0 Or setitedhenavefshijkompania.tables(0).rows.count = 0 Then
                                            MsgBox("Ju jeni nje user!Databaza eshte bosh.Kontakto administratorin!", MsgBoxStyle.Information)
                                        Else
                                            Me.Hide()
                                            'Blere,Administrim,Klasat te caktivizuara
                                            Farmacia.ContextMenuStrip1.Items(0).Enabled = False
                                            Farmacia.ContextMenuStrip1.Items(3).Enabled = False
                                            Farmacia.ContextMenuStrip1.Items(4).Enabled = False
                                            Farmacia.ContextMenuStrip1.Items(5).Enabled = False
                                            Farmacia.ContextMenuStrip1.Items(6).Enabled = False
                                            'Hyrjet
                                            Farmacia.Button2.Enabled = False
                                            Farmacia.Button3.Enabled = False
                                            Farmacia.RadioButton1.Enabled = False
                                            Farmacia.RadioButton2.Enabled = False
                                            'Daljet
                                            Farmacia.Button4.Enabled = False
                                            Farmacia.Button6.Enabled = False
                                            Farmacia.RadioButton3.Enabled = False
                                            Farmacia.RadioButton4.Enabled = False
                                            'Magazina
                                            Farmacia.Button19.Enabled = False
                                            Farmacia.Button18.Enabled = False
                                            Farmacia.RadioButton5.Enabled = False
                                            Farmacia.RadioButton6.Enabled = False
                                            Farmacia.Show()
                                            'Blerjet,Raportet
                                            Farmacia.TabControl1.SelectedIndex = 1
                                            Farmacia.TabControl1.TabPages.Remove(Farmacia.TabPage1)
                                            Farmacia.TabControl1.TabPages.Remove(Farmacia.TabPage3)
                                            Farmacia.TabControl1.TabPages.Remove(Farmacia.TabPage4)
                                            Farmacia.TabControl1.TabPages.Remove(Farmacia.TabPage5)
                                            'Caktivizim per editim
                                            Farmacia.DataGridView1.ReadOnly = True
                                            Farmacia.DataGridView2.ReadOnly = True
                                            Farmacia.DataGridView3.ReadOnly = True
                                            Farmacia.DataGridView4.ReadOnly = True
                                            Farmacia.DataGridView5.ReadOnly = True
                                            Farmacia.DataGridView6.ReadOnly = True
                                            Farmacia.DataGridView7.ReadOnly = True
                                            Shit.CheckBox1.Enabled = False
                                            Shit.CheckBox2.Enabled = False
                                            Ofert.CheckBox1.Enabled = False
                                            Bli.NumericUpDown2.Enabled = False
                                            Shit.NumericUpDown3.Enabled = False
                                            Ofert.NumericUpDown3.Enabled = False
                                        End If
                                    End If
                                End While
                                connection.Close()
                            End Using
                        Else
                            MsgBox("Emri ose Fjalkalimi i gabuar!", MsgBoxStyle.Information)
                            connection1.Close()
                            TextBox1.Text = ""
                            TextBox2.Text = ""
                        End If
                        connection1.Close()
                    End Using
                End If
            End If
        End If

    End Sub
    Public Function autofill()
        setitedhenave2autofill.clear()
        lidhje2autofill.Open()
        adaptori2autofill = New OleDbDataAdapter("SELECT Emri FROM Perdoruesit ORDER BY ID", lidhje2autofill)
        adaptori2autofill.Fill(setitedhenave2autofill, "tedhena")
        Dim col As New AutoCompleteStringCollection
        Dim i As Integer
        For i = 0 To setitedhenave2autofill.Tables(0).Rows.Count - 1
            col.Add(setitedhenave2autofill.Tables(0).Rows(i)("Emri").ToString())
        Next
        TextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox1.AutoCompleteCustomSource = col
        TextBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        lidhje2autofill.Close()
        Return Nothing
    End Function
    Private Sub Hyrje_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Label5.Text.Contains("jo e licensuar!") Then
            Lidh.lidhje21.Close()
            IO.File.Delete(My.Settings.ruajdtbpath & "tedhena.accdb")
        Else
        End If
    End Sub
    Private Sub Hyrje_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox3.Focus()
        Dim data() As Byte
        Dim dec As String
        data = Convert.FromBase64String(StrReverse("*=IUaBRHUxcFOv02b*j5ibpJWZ0NX*Yw9yL6MHc0RHa*#$").Replace("*", "").Remove(0, 2))
        dec = System.Text.ASCIIEncoding.ASCII.GetString(data)
        Application.DoEvents()
        If My.Settings.check2 Then
            TextBox2.Text = "admin123"
            TextBox1.Text = "Administratori"
            Kontrollo(dec)
            TextBox3.Text = My.Settings.dtb1
            Dim msg = "Jeni administrator?"
            Dim title = "Lloji perdoruesit"
            Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1 Or
                    MsgBoxStyle.Information
            Dim response = MsgBox(msg, style, title)
            If response = MsgBoxResult.Yes Then
                MsgBox("Hera e pare?!" & vbNewLine & "Zgjidhni ku do te ruani databazen?!" & vbNewLine & "Nderroni fjalkalimin menjehere!" & vbNewLine & "Klik me te djathten>Tabela administrimit>Perdoruesit" & vbNewLine & "Emri: " & TextBox3.Text & vbNewLine & "Fjalkalimi fillestar: " & TextBox2.Text, MsgBoxStyle.Information)
                If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                    'MsgBox(FolderBrowserDialog1.SelectedPath & "\")
                    IO.File.WriteAllBytes(FolderBrowserDialog1.SelectedPath & "\" & "tedhena.accdb", My.Resources.tedhena)
                    File.SetAttributes(FolderBrowserDialog1.SelectedPath & "\" & "tedhena.accdb", FileAttributes.Hidden)








                End If




                My.Settings.ruajdtbpath = FolderBrowserDialog1.SelectedPath & "\"
                My.Settings.Save()

                'MsgBox("Nderroni fjalkalimin menjehere!" & vbNewLine & "Klik me te djathten>Tabela administrimit>Perdoruesit" & vbNewLine & "Emri: " & TextBox3.Text & vbNewLine & "Fjalkalimi fillestar: " & TextBox2.Text, MsgBoxStyle.Information)
                With DataGridView1
                    .RowsDefaultCellStyle.BackColor = Color.Bisque
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                End With











                rreshtiaktual21 = 0
                Lidh.lidhje21.Open()
                adaptori21 = New OleDbDataAdapter("SELECT Emri,Niveli FROM Perdoruesit ORDER BY ID", Lidh.lidhje21)
                adaptori21.Fill(setitedhenave21, "tedhena")
                DataGridView1.DataSource = setitedhenave21.Tables(0)
                Lidh.lidhje21.Close()








                autofill()


            ElseIf response = MsgBoxResult.No Then
                Dim StatusDate As String
                StatusDate = InputBox("Shkruaj vendodhjen e databazes qe eshte vendosur nga administratori!", "Kerkese", " ")
                My.Settings.ruajdtbpath = StatusDate & "\"
                My.Settings.Save()
                With DataGridView1
                    .RowsDefaultCellStyle.BackColor = Color.Bisque
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                End With
                rreshtiaktual21 = 0
                Lidh.lidhje21.Open()
                adaptori21 = New OleDbDataAdapter("Select Emri, Niveli FROM Perdoruesit ORDER BY ID", Lidh.lidhje21)
                adaptori21.Fill(setitedhenave21, "tedhena")
                DataGridView1.DataSource = setitedhenave21.Tables(0)
                Lidh.lidhje21.Close()
                autofill()
            End If
        Else
            Dim path1 As String = My.Settings.ruajdtbpath & "\tedhena.accdb"
            TextBox3.Text = My.Settings.dtb1
            My.Settings.dtb1 = TextBox3.Text.Trim()
            My.Settings.Save()
            If File.Exists(path1) = True Then
                File.SetAttributes(path1, FileAttributes.Hidden)
                Label5.Text = "           ID e licensuar!"
                Label5.ForeColor = Color.Green
                Label6.Visible = False
                With DataGridView1
                    .RowsDefaultCellStyle.BackColor = Color.Bisque
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                End With
                rreshtiaktual21 = 0
                Lidh.lidhje21.Open()
                adaptori21 = New OleDbDataAdapter("Select Emri, Niveli FROM Perdoruesit ORDER BY ID", Lidh.lidhje21)
                adaptori21.Fill(setitedhenave21, "tedhena")
                DataGridView1.DataSource = setitedhenave21.Tables(0)
                Lidh.lidhje21.Close()
                autofill()
            Else
                If File.Exists(My.Settings.backup) = True Then
                    System.IO.File.Copy(My.Settings.backup, My.Settings.ruajdtbpath & "\tedhena.accdb", True)
                    File.SetAttributes(My.Settings.ruajdtbpath & "\tedhena.accdb", FileAttributes.Hidden)
                    MsgBox("Databaza nuk ekzistonte!U rikthye databaza nga mbyllja e fundit e programit!", MsgBoxStyle.Information)
                    MsgBox("Programi do te rihapet per te filluar punen!", MsgBoxStyle.Information)
                    Application.Restart()
                    Application.Exit()
                Else
                    MsgBox("Kopja e databazes nuk ekziston!Te gjitha te dhenat humben!Kontakto administratorin per t'ja nisur nga e para!")
            End If
            End If
        End If
    End Sub

    Private Sub BtnChangePassword_Click(sender As Object, e As EventArgs) Handles btnChangePassword.Click

        If TextBox1.Text = "" Then
            MsgBox("Plotesoni te gjitha fushat!", MsgBoxStyle.Information)

        Else
            Using connection As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & My.Settings.dtb1 & ";")
                Dim command As New OleDbCommand("SELECT ID,Emri,Fjalkalimi,Niveli FROM Perdoruesit WHERE(Emri Like '" &
                                             TextBox1.Text & "%')", connection)
                connection.Open()
                Dim reader As OleDbDataReader = command.ExecuteReader()
                While reader.Read()
                    If reader(3).ToString() = "Administrator" Then
                        Me.Hide()
                        Fjalkalimi_Databazes.Show()
                    ElseIf reader(3).ToString() = "Perdorues" Then
                        MsgBox("Perdorues i thjesht!Ju nuk keni te drejta!", MsgBoxStyle.Information)

                    Else
                        TextBox1.Text = ""
                        TextBox2.Text = ""
                        MsgBox("Emri i gabuar", MsgBoxStyle.Information)

                        connection.Close()
                    End If
                End While
            End Using
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Label5.Text.Contains("jo e licensuar!") Then
            MsgBox("Kontakto administratorin dhe dergoi serialin per aktivizim!", MsgBoxStyle.Information)
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Then
                MsgBox("Plotesoni te gjitha fushat!", MsgBoxStyle.Information)
            Else
                Dim connectionString1 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & TextBox3.Text & ";"
                Dim queryString1 As String = "SELECT ID, Emri, Fjalkalimi,Niveli  FROM [Perdoruesit] WHERE [Emri] = @UserName And [Fjalkalimi] = @Password"
                Using connection1 As New OleDbConnection(connectionString1)
                    Dim command1 As New OleDbCommand(queryString1, connection1)
                    command1.Parameters.AddWithValue("@UserName", TextBox1.Text)
                    command1.Parameters.AddWithValue("@Password", TextBox2.Text)
                    connection1.Open()
                    Dim hasrow As Integer = Convert.ToInt32(command1.ExecuteScalar())
                    If hasrow <> 0 Then
                        connection1.Close()
                        My.Settings.dtb1 = TextBox3.Text
                        My.Settings.Save()
                        Using connection As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & My.Settings.dtb1 & ";")
                            Dim command As New OleDbCommand("SELECT ID, Emri, Fjalkalimi,Niveli  FROM [Perdoruesit] WHERE [Emri] = @UserName And [Fjalkalimi] = @Password", connection)
                            command.Parameters.AddWithValue("@UserName", TextBox1.Text)
                            command.Parameters.AddWithValue("@Password", TextBox2.Text)
                            connection.Open()
                            Dim reader As OleDbDataReader = command.ExecuteReader()
                            While reader.Read()
                                If reader(3).ToString() = "Administrator" Then
                                    'Kap dtg1
                                    setitedhenave2yxy.clear
                                    lidhje2yxy.Open()
                                    adaptori2yxy = New OleDbDataAdapter("SELECT * FROM Perdoruesit ORDER BY ID", lidhje2yxy)
                                    adaptori2yxy.Fill(setitedhenave2yxy, "tedhena")
                                    lidhje2yxy.Close()
                                    'Kap dtg2
                                    setitedhenave2y.clear
                                    lidhje2y.Open()
                                    adaptori2y = New OleDbDataAdapter("SELECT * FROM Fornitoret ORDER BY ID", lidhje2y)
                                    adaptori2y.Fill(setitedhenave2y, "tedhena")
                                    lidhje2y.Close()
                                    'Kap dtg3
                                    setitedhenave2yxc.clear
                                    lidhje2yx.Open()
                                    adaptori2yx = New OleDbDataAdapter("SELECT * FROM Bleresit ORDER BY ID", lidhje2yx)
                                    adaptori2yx.Fill(setitedhenave2yxc, "tedhena")
                                    lidhje2yx.Close()
                                    'Kap dtg4
                                    setitedhenave2yxy1.clear
                                    lidhje2yxy1.Open()
                                    adaptori2yxy1 = New OleDbDataAdapter("SELECT * FROM Produktet ORDER BY ID", lidhje2yxy1)
                                    adaptori2yxy1.Fill(setitedhenave2yxy1, "tedhena")
                                    lidhje2yxy1.Close()
                                    'Kap dtg5
                                    lidhjekompania.Close()
                                    lidhjekompania.Open()
                                    adaptorikompania = New OleDbDataAdapter("SELECT * FROM Shitesi ORDER BY ID", lidhjekompania)
                                    adaptorikompania.Fill(setitedhenavefshijkompania, "tedhena")
                                    lidhjekompania.Close()
                                    If setitedhenave2yxy.Tables(0).Rows.Count = 0 Or setitedhenave2y.Tables(0).Rows.Count = 0 Or setitedhenave2yxc.Tables(0).Rows.Count = 0 Or setitedhenave2yxy1.Tables(0).Rows.Count = 0 Or setitedhenavefshijkompania.Tables(0).Rows.Count = 0 Then
                                        Me.Hide()
                                        Administrim.Show()
                                        MsgBox("Duhet te shtoni te pakten nje rresht ne cdo tabele qe programi te filloje punen!", MsgBoxStyle.Information)
                                    Else
                                        Me.Hide()
                                        Farmacia.Show()
                                    End If
                                ElseIf reader(3).ToString() = "Perdorues" Then
                                    'Kap dtg1
                                    setitedhenave2yxy.clear
                                    lidhje2yxy.Open()
                                    adaptori2yxy = New OleDbDataAdapter("SELECT * FROM Perdoruesit ORDER BY ID", lidhje2yxy)
                                    adaptori2yxy.Fill(setitedhenave2yxy, "tedhena")
                                    lidhje2yxy.Close()
                                    'Kap dtg2
                                    setitedhenave2y.clear
                                    lidhje2y.Open()
                                    adaptori2y = New OleDbDataAdapter("SELECT * FROM Fornitoret ORDER BY ID", lidhje2y)
                                    adaptori2y.Fill(setitedhenave2y, "tedhena")
                                    lidhje2y.Close()
                                    'Kap dtg3
                                    setitedhenave2yxc.clear
                                    lidhje2yx.Open()
                                    adaptori2yx = New OleDbDataAdapter("SELECT * FROM Bleresit ORDER BY ID", lidhje2yx)
                                    adaptori2yx.Fill(setitedhenave2yxc, "tedhena")
                                    lidhje2yx.Close()
                                    'Kap dtg4
                                    setitedhenave2yxy1.clear
                                    lidhje2yxy1.Open()
                                    adaptori2yxy1 = New OleDbDataAdapter("SELECT * FROM Produktet ORDER BY ID", lidhje2yxy1)
                                    adaptori2yxy1.Fill(setitedhenave2yxy1, "tedhena")
                                    lidhje2yxy1.Close()
                                    'Kap dtg5
                                    lidhjekompania.Close()
                                    lidhjekompania.Open()
                                    adaptorikompania = New OleDbDataAdapter("SELECT * FROM Shitesi ORDER BY ID", lidhjekompania)
                                    adaptorikompania.Fill(setitedhenavefshijkompania, "tedhena")
                                    lidhjekompania.Close()
                                    If setitedhenave2yxy.Tables(0).Rows.Count = 0 Or setitedhenave2y.Tables(0).Rows.Count = 0 Or setitedhenave2yxc.Tables(0).Rows.Count = 0 Or setitedhenave2yxy1.Tables(0).Rows.Count = 0 Or setitedhenavefshijkompania.tables(0).rows.count = 0 Then
                                        MsgBox("Ju jeni nje user!Databaza eshte bosh.Kontakto administratorin!", MsgBoxStyle.Information)
                                    Else
                                        Me.Hide()
                                        Farmacia.Show()
                                        'Blere, Administrim, Klasat te caktivizuara
                                        Farmacia.ContextMenuStrip1.Items(0).Enabled = False
                                        Farmacia.ContextMenuStrip1.Items(3).Enabled = False
                                        Farmacia.ContextMenuStrip1.Items(4).Enabled = False
                                        Farmacia.ContextMenuStrip1.Items(5).Enabled = False
                                        Farmacia.ContextMenuStrip1.Items(6).Enabled = False
                                        'Hyrjet
                                        Farmacia.Button2.Enabled = False
                                        Farmacia.Button3.Enabled = False
                                        Farmacia.RadioButton1.Enabled = False
                                        Farmacia.RadioButton2.Enabled = False
                                        'Daljet
                                        Farmacia.Button4.Enabled = False
                                        Farmacia.Button6.Enabled = False
                                        Farmacia.RadioButton3.Enabled = False
                                        Farmacia.RadioButton4.Enabled = False
                                        'Magazina
                                        Farmacia.Button19.Enabled = False
                                        Farmacia.Button18.Enabled = False
                                        Farmacia.RadioButton5.Enabled = False
                                        Farmacia.RadioButton6.Enabled = False

                                        'Blerjet,Raportet
                                        Farmacia.TabControl1.SelectedIndex = 1
                                        Farmacia.TabControl1.TabPages.Remove(Farmacia.TabPage1)
                                        Farmacia.TabControl1.TabPages.Remove(Farmacia.TabPage3)
                                        Farmacia.TabControl1.TabPages.Remove(Farmacia.TabPage4)
                                        Farmacia.TabControl1.TabPages.Remove(Farmacia.TabPage5)
                                        'Caktivizim per editim
                                        Farmacia.DataGridView1.ReadOnly = True
                                        Farmacia.DataGridView2.ReadOnly = True
                                        Farmacia.DataGridView3.ReadOnly = True
                                        Farmacia.DataGridView4.ReadOnly = True
                                        Farmacia.DataGridView5.ReadOnly = True
                                        Farmacia.DataGridView6.ReadOnly = True
                                        Farmacia.DataGridView7.ReadOnly = True
                                        Shit.CheckBox1.Enabled = False
                                        Shit.CheckBox2.Enabled = False
                                        Ofert.CheckBox1.Enabled = False
                                        Bli.NumericUpDown2.Enabled = False
                                        Shit.NumericUpDown3.Enabled = False
                                        Ofert.NumericUpDown3.Enabled = False

                                    End If

                                End If
                            End While
                            connection.Close()
                        End Using
                    Else
                        MsgBox("Emri ose Fjalkalimi i gabuar!", MsgBoxStyle.Information)
                        connection1.Close()
                        TextBox1.Text = ""
                        TextBox2.Text = ""
                    End If
                    connection1.Close()
                End Using
            End If
        End If


    End Sub


End Class