Imports System.Data.OleDb
Imports System.Globalization
Imports System.Linq
Public Class Administrim
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
    End Sub
    Private Sub _Load(ByVal sender As Object, ByVal e As EventArgs)
        _form_resize._get_initial_size()
    End Sub
    Private Sub Administrim_FormClosing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If DataGridView1.Rows.Count = 0 Or DataGridView2.Rows.Count = 0 Or DataGridView3.Rows.Count = 0 Or DataGridView4.Rows.Count = 0 Then
            MsgBox("Programi nuk do te filloje nga puna neqoftese ju nuk plotesoni te pakten nje rresht per cdo tabele!", MsgBoxStyle.Information)
            Application.Exit()
        Else
            Farmacia.Show()
        End If

    End Sub
    ' Public Const WM_NCLBUTTONDBLCLK As Integer = &HA3
    Private Sub Merrtedhenat2(ByVal rreshtiaktual2yxy)
        Try
            TextBox2.Text = setitedhenave2yxy.Tables("tedhena").Rows(rreshtiaktual2yxy)("ID")
            TextBox3.Text = setitedhenave2yxy.Tables("tedhena").Rows(rreshtiaktual2yxy)("Emri")
            TextBox4.Text = setitedhenave2yxy.Tables("tedhena").Rows(rreshtiaktual2yxy)("Fjalkalimi")
            ComboBox1.Text = setitedhenave2yxy.Tables("tedhena").Rows(rreshtiaktual2yxy)("Niveli")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Merrtedhenat_Fornitori(ByVal rreshtiaktual2y)
        Try
            TextBox5.Text = setitedhenave2y.Tables("tedhena").Rows(rreshtiaktual2y)("ID")
            TextBox6.Text = setitedhenave2y.Tables("tedhena").Rows(rreshtiaktual2y)("Emri")
            TextBox7.Text = setitedhenave2y.Tables("tedhena").Rows(rreshtiaktual2y)("Adresa")
            TextBox8.Text = setitedhenave2y.Tables("tedhena").Rows(rreshtiaktual2y)("Telefon")
            TextBox9.Text = setitedhenave2y.Tables("tedhena").Rows(rreshtiaktual2y)("NIPT")
            TextBox10.Text = setitedhenave2y.Tables("tedhena").Rows(rreshtiaktual2y)("Kodi_Fornitorit")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Merrtedhenat_produktet(ByVal rreshtiaktual2yxy1)
        Try
            TextBox19.Text = setitedhenave2yxy1.Tables("tedhena").Rows(rreshtiaktual2yxy1)("ID")
            TextBox20.Text = setitedhenave2yxy1.Tables("tedhena").Rows(rreshtiaktual2yxy1)("Emri_produktit")
            TextBox23.Text = setitedhenave2yxy1.Tables("tedhena").Rows(rreshtiaktual2yxy1)("Njesia")
            TextBox21.Text = setitedhenave2yxy1.Tables("tedhena").Rows(rreshtiaktual2yxy1)("Cmimi_blerjes")
            TextBox22.Text = setitedhenave2yxy1.Tables("tedhena").Rows(rreshtiaktual2yxy1)("Cmimi_Shitjes")
            TextBox25.Text = setitedhenave2yxy1.Tables("tedhena").Rows(rreshtiaktual2yxy1)("Kodi_produktit")
            TextBox26.Text = setitedhenave2yxy1.Tables("tedhena").Rows(rreshtiaktual2yxy1)("Sasia_limit")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Merrtedhenat_kompania(ByVal rreshtiaktualfshijkompania)
        Try
            TextBox24.Text = setitedhenavefshijkompania.Tables("tedhena").Rows(rreshtiaktualfshijkompania)("ID")
            TextBox35.Text = setitedhenavefshijkompania.Tables("tedhena").Rows(rreshtiaktualfshijkompania)("Emri")
            TextBox36.Text = setitedhenavefshijkompania.Tables("tedhena").Rows(rreshtiaktualfshijkompania)("Adresa")
            TextBox37.Text = setitedhenavefshijkompania.Tables("tedhena").Rows(rreshtiaktualfshijkompania)("Telefon")
            TextBox38.Text = setitedhenavefshijkompania.Tables("tedhena").Rows(rreshtiaktualfshijkompania)("NIPT")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Merrtedhenat_Bleresit(ByVal rreshtiaktual2yxc)
        Try
            TextBox12.Text = setitedhenave2yxc.Tables("tedhena").Rows(rreshtiaktual2yxc)("ID")
            TextBox13.Text = setitedhenave2yxc.Tables("tedhena").Rows(rreshtiaktual2yxc)("Emri")
            TextBox14.Text = setitedhenave2yxc.Tables("tedhena").Rows(rreshtiaktual2yxc)("Adresa")
            TextBox15.Text = setitedhenave2yxc.Tables("tedhena").Rows(rreshtiaktual2yxc)("Telefon")
            TextBox16.Text = setitedhenave2yxc.Tables("tedhena").Rows(rreshtiaktual2yxc)("NIPT")
            TextBox17.Text = setitedhenave2yxc.Tables("tedhena").Rows(rreshtiaktual2yxc)("Kodi_klientit")
        Catch ex As Exception

        End Try
    End Sub
    '  Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
    '  If m.Msg = WM_NCLBUTTONDBLCLK Then Return
    'MyBase.WndProc(m)
    ' End Sub
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
    Dim nr1 As Integer = 0
    Dim nr2 As Integer = 0
    Dim nr3 As Integer = 0
    Dim nr4 As Integer = 0
    Private Sub Administrim_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'rizmadho tabcontrol faqe1
        ' Me.TabControl1.Size = New Size(672, 359)
        '  Me.Width = 688
        ' Me.Height = 398
        ' Me.StartPosition = FormStartPosition.CenterScreen
        If DataGridView1.Rows.Count = 0 Or DataGridView2.Rows.Count = 0 Or DataGridView3.Rows.Count = 0 Or DataGridView4.Rows.Count = 0 Then

        Else
            autofill()
            autofillklient()
            autofillperdoruesit()
            autofillproduktet()
        End If

        'Kap dtg1
        setitedhenave2yxy.clear
        With DataGridView1
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual2yxy = 0
        lidhje2yxy.Open()
        adaptori2yxy = New OleDbDataAdapter("SELECT * FROM Perdoruesit ORDER BY ID", lidhje2yxy)
        adaptori2yxy.Fill(setitedhenave2yxy, "tedhena")
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("count1", "Nr.")
        DataGridView1.DataSource = setitedhenave2yxy.Tables(0)
        DataGridView1.Columns("ID").Visible = False
        Merrtedhenat2(rreshtiaktual2yxy)
        lidhje2yxy.Close()


        'Kap dtg2
        setitedhenave2y.clear
        With DataGridView2
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual2y = 0
        lidhje2y.Open()
        adaptori2y = New OleDbDataAdapter("SELECT * FROM Fornitoret ORDER BY ID", lidhje2y)
        adaptori2y.Fill(setitedhenave2y, "tedhena")
        DataGridView2.Columns.Clear()
        DataGridView2.Columns.Add("count1", "Nr.")
        DataGridView2.DataSource = setitedhenave2y.Tables(0)
        DataGridView2.Columns("ID").Visible = False
        nr2 = DataGridView2.Rows.Count - 1
        For i = 0 To nr2
            DataGridView2.Rows(i).Cells(0).Value = i + 1
        Next
        Merrtedhenat_Fornitori(rreshtiaktual2y)
        lidhje2y.Close()


        'Kap dtg3
        setitedhenave2yxc.clear
        With DataGridView3
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual2yxc = 0
        lidhje2yx.Open()
        adaptori2yx = New OleDbDataAdapter("SELECT * FROM Bleresit ORDER BY ID", lidhje2yx)
        adaptori2yx.Fill(setitedhenave2yxc, "tedhena")
        DataGridView3.Columns.Clear()
        DataGridView3.Columns.Add("count1", "Nr.")
        DataGridView3.DataSource = setitedhenave2yxc.Tables(0)
        DataGridView3.Columns("ID").Visible = False
        Merrtedhenat_Bleresit(rreshtiaktual2yxc)
        lidhje2yx.Close()



        'Kap dtg4
        setitedhenave2yxy1.clear

        With DataGridView4
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual2yxy1 = 0
        lidhje2yxy1.Open()
        adaptori2yxy1 = New OleDbDataAdapter("SELECT * FROM Produktet ORDER BY ID", lidhje2yxy1)
        adaptori2yxy1.Fill(setitedhenave2yxy1, "tedhena")
        DataGridView4.Columns.Clear()
        DataGridView4.Columns.Add("count1", "Nr.")
        DataGridView4.DataSource = setitedhenave2yxy1.Tables(0)
        DataGridView4.Columns("ID").Visible = False
        Merrtedhenat_produktet(rreshtiaktual2yxy1)
        lidhje2yxy1.Close()

        Dim bgw5 As System.ComponentModel.BackgroundWorker
        bgw5 = New System.ComponentModel.BackgroundWorker
        AddHandler bgw5.DoWork, AddressOf dtg5
        bgw5.RunWorkerAsync()

        If Not DataGridView2.Rows.Count > 0 Then
            Button1.Enabled = False
            Button2.Enabled = False
        Else
            Button1.Enabled = True
            Button2.Enabled = True
        End If
        If Not DataGridView4.Rows.Count > 0 Then
            Button11.Enabled = False
            Button12.Enabled = False
        Else
            Button11.Enabled = True
            Button12.Enabled = True
        End If



    End Sub
    Sub dtg5()
        'Datagridview5
        setitedhenavefshijkompania.clear
        With DataGridView5
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktualfshijkompania = 0
        lidhjekompania.Close()
        lidhjekompania.Open()
        adaptorikompania = New OleDbDataAdapter("SELECT * FROM Shitesi ORDER BY ID", lidhjekompania)
        adaptorikompania.Fill(setitedhenavefshijkompania, "tedhena")
        DataGridView5.Columns.Clear()
        DataGridView5.Columns.Add("count1", "Nr.")
        DataGridView5.DataSource = setitedhenavefshijkompania.Tables(0)
        DataGridView5.Columns("ID").Visible = False
        ' DataGridView4.CurrentCell = DataGridView4.Rows(0).Cells(0)
        Merrtedhenat_kompania(rreshtiaktualfshijkompania)
        lidhjekompania.Close()
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If (e.RowIndex >= 0) Then
            Dim row As DataGridViewRow = DataGridView1.Rows.Item(e.RowIndex)
            TextBox2.Text = row.Cells.Item("ID").Value.ToString
            TextBox3.Text = row.Cells.Item("Emri").Value.ToString
            TextBox4.Text = row.Cells.Item("Fjalkalimi").Value.ToString
            ComboBox1.Text = row.Cells.Item("Niveli").Value.ToString
            If row.Cells.Item("Niveli").Value.ToString = "Administrator" Then
                Label13.ForeColor = Color.Green
            Else
                Label13.ForeColor = Color.Red
            End If
            Label13.Text = row.Cells.Item("Niveli").Value.ToString
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim nrekzistues As New List(Of Integer)
        For Each kolone As DataGridViewRow In DataGridView1.Rows
            nrekzistues.Add(CInt(kolone.Cells(0).Value))
        Next
        Dim existingNumbers As New List(Of Integer)
        For Each r As DataGridViewRow In DataGridView1.Rows
            existingNumbers.Add(CInt(r.Cells(0).Value))
        Next
        If existingNumbers.Count = 0 And nrekzistues.Count = 0 Then
            Dim str As String
            Try
                str = "insert into Perdoruesit values("
                str += "1"
                str += ","
                str += """" & TextBox3.Text.Trim & """"
                str += ","
                str += """" & TextBox4.Text.Trim & """"
                str += ","
                str += """" & ComboBox1.Text.Trim() & """"
                str += ")"
                lidhje211.Open()
                query211 = New OleDbCommand(str, lidhje211)
                query211.ExecuteNonQuery()
                setitedhenave211.Clear()
                adaptori211 = New OleDbDataAdapter("SELECT * FROM Perdoruesit ORDER BY ID", lidhje211)
                adaptori211.Fill(setitedhenave211, "tedhena")
                DataGridView1.DataSource = setitedhenave211.Tables(0)
                Dim nr1 As Integer
                nr1 = DataGridView1.Rows.Count - 1
                For i = 0 To nr1
                    DataGridView1.Rows(i).Cells(0).Value = i + 1
                Next
                DataGridView1.Refresh()
                MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                lidhje211.Close()
                TextBox3.Text = ""
                TextBox4.Text = ""
            Catch ex As Exception
                MsgBox("Nuk u shtua", MsgBoxStyle.Information)

                lidhje211.Close()
            End Try
        Else
            Dim missingNumbers = Enumerable.Range(existingNumbers.First, existingNumbers.Max() - existingNumbers.First + 1).Except(existingNumbers)
            Dim max = nrekzistues.Max() + 1
            If missingNumbers.Count = 0 Then
                Dim str As String
                Try
                    str = "insert into Perdoruesit values("
                    str += max.ToString
                    str += ","
                    str += """" & TextBox3.Text.Trim & """"
                    str += ","
                    str += """" & TextBox4.Text.Trim & """"
                    str += ","
                    str += """" & ComboBox1.Text.Trim() & """"
                    str += ")"
                    lidhje211.Open()
                    query211 = New OleDbCommand(str, lidhje211)
                    query211.ExecuteNonQuery()
                    setitedhenave211.Clear()
                    adaptori211 = New OleDbDataAdapter("SELECT * FROM Perdoruesit ORDER BY ID", lidhje211)
                    adaptori211.Fill(setitedhenave211, "tedhena")
                    DataGridView1.DataSource = setitedhenave211.Tables(0)
                    Dim nr1 As Integer
                    nr1 = DataGridView1.Rows.Count - 1
                    For i = 0 To nr1
                        DataGridView1.Rows(i).Cells(0).Value = i + 1
                    Next
                    DataGridView1.Refresh()
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    lidhje211.Close()
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                Catch ex As Exception
                    MsgBox("Nuk u shtua", MsgBoxStyle.Information)
                    lidhje211.Close()
                End Try
            Else
                Try
                    Dim str As String
                    str = "insert into Perdoruesit values("
                    str += missingNumbers.First.ToString
                    str += ","
                    str += """" & TextBox3.Text.Trim & """"
                    str += ","
                    str += """" & TextBox4.Text.Trim & """"
                    str += ","
                    str += """" & ComboBox1.Text.Trim() & """"
                    str += ")"
                    lidhje211.Open()
                    query211 = New OleDbCommand(str, lidhje211)
                    query211.ExecuteNonQuery()
                    setitedhenave211.Clear()
                    adaptori211 = New OleDbDataAdapter("SELECT * FROM Perdoruesit ORDER BY ID", lidhje211)
                    adaptori211.Fill(setitedhenave211, "tedhena")
                    DataGridView1.DataSource = setitedhenave211.Tables(0)
                    Dim nr1 As Integer
                    nr1 = DataGridView1.Rows.Count - 1
                    For i = 0 To nr1
                        DataGridView1.Rows(i).Cells(0).Value = i + 1
                    Next
                    DataGridView1.Refresh()
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    lidhje211.Close()
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                Catch ex As Exception
                    MsgBox("Nuk u shtua", MsgBoxStyle.Information)
                    lidhje211.Close()
                End Try
            End If
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Dim Str As String
            Str = "update Perdoruesit set Emri="
            Str += """" & TextBox3.Text & """"
            Str += " where ID="
            Str += TextBox2.Text.Trim()
            lidhje211.Open()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()
            lidhje211.Open()
            Str = "update Perdoruesit set Fjalkalimi="
            Str += """" & TextBox4.Text & """"
            Str += " where ID="
            Str += TextBox2.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()
            lidhje211.Open()
            Str = "update Perdoruesit set Niveli="
            Str += """" & ComboBox1.Text & """"
            Str += " where ID="
            Str += TextBox2.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()
            lidhje211.Open()
            setitedhenave211.Clear()
            adaptori211 = New OleDbDataAdapter("SELECT * FROM Perdoruesit ORDER BY ID", lidhje211)
            adaptori211.Fill(setitedhenave211, "tedhena")
            DataGridView1.DataSource = setitedhenave211.Tables(0)
            nr1 = DataGridView1.Rows.Count - 1
            For i = 0 To nr1
                DataGridView1.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView1.Refresh()
            MsgBox("U rifreskua me sukses", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("Nuk u rifreskua me sukses", MsgBoxStyle.Information)
        End Try
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Not DataGridView1.Rows.Count > 0 Then
            MsgBox("Lista eshte bosh!", MsgBoxStyle.Information)
        Else
            Dim Str As String
            Try
                Str = "delete from Perdoruesit where ID="
                Str += DataGridView1.CurrentRow.Cells(1).Value.ToString
                lidhje211.Open()
                query211 = New OleDbCommand(Str, lidhje211)
                query211.ExecuteNonQuery()
                setitedhenave211.clear()
                adaptori211 = New OleDbDataAdapter("SELECT * FROM Perdoruesit ORDER BY ID", lidhje211)
                adaptori211.Fill(setitedhenave211, "tedhena")
                DataGridView1.DataSource = setitedhenave211.Tables(0)
                nr1 = DataGridView1.Rows.Count - 1
                For i = 0 To nr1
                    DataGridView1.Rows(i).Cells(0).Value = i + 1
                Next
                DataGridView1.Refresh()
                MsgBox("U fshi me sukses", MsgBoxStyle.Information)
                If rreshtiaktual11 > 0 Then
                    rreshtiaktual11 -= 1
                    Merrtedhenat2(rreshtiaktual11)
                End If
                lidhje211.Close()
            Catch ex As Exception
                MsgBox("Nuk u fshi", MsgBoxStyle.Information)
                lidhje211.Close()
            End Try
        End If
    End Sub
    Public Function GenerateRandomString(ByRef iLength As Integer) As String
        Dim rdm As New Random()
        Dim allowChrs() As Char = "ABCDEFGHIJKLOMNOPQRSTUVWXYZ0123456789".ToCharArray()
        Dim sResult As String = ""

        For i As Integer = 0 To iLength - 1
            sResult += allowChrs(rdm.Next(0, allowChrs.Length))
        Next

        Return sResult
    End Function
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Shto fornitor ne tabelen Fornitoret
        Dim nrekzistues As New List(Of Integer)
        For Each kolone As DataGridViewRow In DataGridView2.Rows
            nrekzistues.Add(CInt(kolone.Cells(0).Value))
        Next
        Dim existingNumbers As New List(Of Integer)
        For Each r As DataGridViewRow In DataGridView2.Rows
            existingNumbers.Add(CInt(r.Cells(0).Value))
        Next
        If existingNumbers.Count = 0 And nrekzistues.Count = 0 Then
            Dim Str As String
            Try
                Str = "insert into Fornitoret values("
                Str += "1"
                Str += ","
                Str += """" & TextBox6.Text.Trim & """"
                Str += ","
                Str += """" & TextBox7.Text.Trim & """"
                Str += ","
                Str += """" & TextBox8.Text.Trim() & """"
                Str += ","
                Str += """" & TextBox9.Text.Trim() & """"
                Str += ","
                Str += """" & GenerateRandomString(8) & """"
                Str += ")"
                lidhje2112.Open()
                query2112 = New OleDbCommand(Str, lidhje2112)
                query2112.ExecuteNonQuery()
                setitedhenave2112.Clear()
                adaptori2112 = New OleDbDataAdapter("SELECT * FROM Fornitoret ORDER BY ID", lidhje2112)
                adaptori2112.Fill(setitedhenave2112, "tedhena")
                DataGridView2.DataSource = setitedhenave2112.Tables(0)
                Dim nr2 As Integer
                nr2 = DataGridView2.Rows.Count - 1
                For i = 0 To nr2
                    DataGridView2.Rows(i).Cells(0).Value = i + 1
                Next
                MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                lidhje2112.Close()

                TextBox6.Text = ""
                TextBox7.Text = ""
                TextBox8.Text = ""
                TextBox9.Text = ""
            Catch ex As Exception
                MsgBox("Nuk u shtua", MsgBoxStyle.Information)

                lidhje2112.Close()
            End Try
        Else
            Dim missingNumbers = Enumerable.Range(existingNumbers.First, existingNumbers.Max() - existingNumbers.First + 1).Except(existingNumbers)
            If missingNumbers.Count = 0 Then
                Dim max = nrekzistues.Max() + 1
                Dim Str As String
                Try
                    Str = "insert into Fornitoret values("
                    Str += max.ToString
                    Str += ","
                    Str += """" & TextBox6.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox7.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox8.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox9.Text.Trim() & """"
                    Str += ","
                    Str += """" & GenerateRandomString(8) & """"
                    Str += ")"
                    lidhje2112.Open()
                    query2112 = New OleDbCommand(Str, lidhje2112)
                    query2112.ExecuteNonQuery()
                    setitedhenave2112.Clear()
                    adaptori2112 = New OleDbDataAdapter("SELECT * FROM Fornitoret ORDER BY ID", lidhje2112)
                    adaptori2112.Fill(setitedhenave2112, "tedhena")
                    DataGridView2.DataSource = setitedhenave2112.Tables(0)
                    Dim nr2 As Integer
                    nr2 = DataGridView2.Rows.Count - 1
                    For i = 0 To nr2
                        DataGridView2.Rows(i).Cells(0).Value = i + 1
                    Next
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    lidhje2112.Close()
                    TextBox6.Text = ""
                    TextBox7.Text = ""
                    TextBox8.Text = ""
                    TextBox9.Text = ""
                Catch ex As Exception
                    MsgBox("Nuk u shtua", MsgBoxStyle.Information)

                    lidhje2112.Close()
                End Try
            Else
                Dim Str As String
                Try
                    Str = "insert into Fornitoret values("
                    Str += missingNumbers.First.ToString
                    Str += ","
                    Str += """" & TextBox6.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox7.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox8.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox9.Text.Trim() & """"
                    Str += ","
                    Str += """" & GenerateRandomString(8) & """"
                    Str += ")"
                    lidhje2112.Open()
                    query2112 = New OleDbCommand(Str, lidhje2112)
                    query2112.ExecuteNonQuery()
                    setitedhenave2112.Clear()
                    adaptori2112 = New OleDbDataAdapter("SELECT * FROM Fornitoret ORDER BY ID", lidhje2112)
                    adaptori2112.Fill(setitedhenave2112, "tedhena")
                    DataGridView2.DataSource = setitedhenave2112.Tables(0)
                    Dim nr2 As Integer
                    nr2 = DataGridView2.Rows.Count - 1
                    For i = 0 To nr2
                        DataGridView2.Rows(i).Cells(0).Value = i + 1
                    Next
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    lidhje2112.Close()
                    TextBox6.Text = ""
                    TextBox7.Text = ""
                    TextBox8.Text = ""
                    TextBox9.Text = ""
                Catch ex As Exception
                    MsgBox("Nuk u shtua", MsgBoxStyle.Information)

                    lidhje2112.Close()
                End Try
            End If
        End If
    End Sub
    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Select Case TabControl1.SelectedIndex
            Case 0
                ' Me.TabControl1.Size = New Size(672, 359)
                ' Me.Width = 688
                ' Me.Height = 398
                Dim nr2 As Integer
                nr2 = DataGridView2.Rows.Count - 1
                For i = 0 To nr2
                    DataGridView2.Rows(i).Cells(0).Value = i + 1
                Next
            Case 1
                '   Me.TabControl1.Size = New Size(670, 350)
                '   Me.Width = 686
                '   Me.Height = 389
                Dim nr3 As Integer
                nr3 = DataGridView3.Rows.Count - 1
                For i = 0 To nr3
                    DataGridView3.Rows(i).Cells(0).Value = i + 1
                Next
            Case 2
                '  Me.TabControl1.Size = New Size(628, 657)
                '  Me.Width = 644
                '  Me.Height = 696
                Dim nr4 As Integer
                nr4 = DataGridView4.Rows.Count - 1
                For i = 0 To nr4
                    DataGridView4.Rows(i).Cells(0).Value = i + 1
                Next
            Case 3
                'Me.TabControl1.Size = New Size(666, 292)
                ' Me.Width = 685
                ' Me.Height = 334
                Dim nr5 As Integer
                nr5 = DataGridView5.Rows.Count - 1
                For i = 0 To nr5
                    DataGridView5.Rows(i).Cells(0).Value = i + 1
                Next
            Case 4
                'Me.TabControl1.Size = New Size(509, 312)
                '  Me.Width = 525
                '  Me.Height = 351
                Dim nr1 As Integer
                nr1 = DataGridView1.Rows.Count - 1
                For i = 0 To nr1
                    DataGridView1.Rows(i).Cells(0).Value = i + 1
                Next
        End Select
    End Sub
    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If (e.RowIndex >= 0) Then
            Dim row As DataGridViewRow = DataGridView2.Rows.Item(e.RowIndex)
            TextBox5.Text = row.Cells.Item("ID").Value.ToString
            TextBox6.Text = row.Cells.Item("Emri").Value.ToString
            TextBox7.Text = row.Cells.Item("Adresa").Value.ToString
            TextBox8.Text = row.Cells.Item("Telefon").Value.ToString
            TextBox9.Text = row.Cells.Item("NIPT").Value.ToString
            TextBox10.Text = row.Cells.Item("Kodi_Fornitorit").Value.ToString
        End If
    End Sub
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim Str As String
            Str = "update Fornitoret set Emri="
            Str += """" & TextBox6.Text & """"
            Str += " where ID="
            Str += TextBox5.Text.Trim()
            lidhje211.Open()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Fornitoret set Adresa="
            Str += """" & TextBox7.Text & """"
            Str += " where ID="
            Str += TextBox5.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Fornitoret set Telefon="
            Str += """" & TextBox8.Text & """"
            Str += " where ID="
            Str += TextBox5.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Fornitoret set NIPT="
            Str += """" & TextBox9.Text & """"
            Str += " where ID="
            Str += TextBox5.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            setitedhenave211.Clear()
            adaptori211 = New OleDbDataAdapter("SELECT * FROM Fornitoret ORDER BY ID", lidhje211)
            adaptori211.Fill(setitedhenave211, "tedhena")
            DataGridView2.DataSource = setitedhenave211.Tables(0)
            DataGridView2.Refresh()
            lidhje211.Close()


            setitedhenave2y.clear
            With DataGridView2
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            rreshtiaktual2y = 0
            lidhje2y.Open()
            adaptori2y = New OleDbDataAdapter("SELECT * FROM Fornitoret ORDER BY ID", lidhje2y)
            adaptori2y.Fill(setitedhenave2y, "tedhena")
            DataGridView2.DataSource = setitedhenave2y.Tables(0)
            nr2 = DataGridView2.Rows.Count - 1
            For i = 0 To nr2
                DataGridView2.Rows(i).Cells(0).Value = i + 1
            Next
            Merrtedhenat_Fornitori(rreshtiaktual2y)
            lidhje2y.Close()
            MsgBox("U rifreskua me sukses", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("Nuk u rifreskua me sukses", MsgBoxStyle.Information)
        End Try

    End Sub
    Public Function autofill()
        setitedhenave2autofill.clear()
        lidhje2autofill.Open()
        adaptori2autofill = New OleDbDataAdapter("SELECT Emri FROM Fornitoret ORDER BY ID", lidhje2autofill)
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
    Public Function autofillklient()
        setitedhenave2autofill.clear()
        lidhje2autofill.Open()
        adaptori2autofill = New OleDbDataAdapter("SELECT Emri FROM Bleresit ORDER BY ID", lidhje2autofill)
        adaptori2autofill.Fill(setitedhenave2autofill, "tedhena")
        Dim col As New AutoCompleteStringCollection
        Dim i As Integer
        For i = 0 To setitedhenave2autofill.Tables(0).Rows.Count - 1
            col.Add(setitedhenave2autofill.Tables(0).Rows(i)("Emri").ToString())
        Next
        TextBox18.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox18.AutoCompleteCustomSource = col
        TextBox18.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        lidhje2autofill.Close()
        Return Nothing
    End Function
    Public Function autofillproduktet()
        setitedhenave2autofill.clear()
        lidhje2autofill.Open()
        adaptori2autofill = New OleDbDataAdapter("SELECT Emri_produktit FROM Produktet ORDER BY ID", lidhje2autofill)
        adaptori2autofill.Fill(setitedhenave2autofill, "tedhena")
        Dim col As New AutoCompleteStringCollection
        Dim i As Integer
        For i = 0 To setitedhenave2autofill.Tables(0).Rows.Count - 1
            col.Add(setitedhenave2autofill.Tables(0).Rows(i)("Emri_produktit").ToString())
        Next
        TextBox28.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox28.AutoCompleteCustomSource = col
        TextBox28.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        lidhje2autofill.Close()
        Return Nothing
    End Function
    Public Function autofillperdoruesit()
        setitedhenave2autofill.clear()
        lidhje2autofill.Open()
        adaptori2autofill = New OleDbDataAdapter("SELECT Emri FROM Perdoruesit ORDER BY ID", lidhje2autofill)
        adaptori2autofill.Fill(setitedhenave2autofill, "tedhena")
        Dim col As New AutoCompleteStringCollection
        Dim i As Integer
        For i = 0 To setitedhenave2autofill.Tables(0).Rows.Count - 1
            col.Add(setitedhenave2autofill.Tables(0).Rows(i)("Emri").ToString())
        Next
        TextBox11.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox11.AutoCompleteCustomSource = col
        TextBox11.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        lidhje2autofill.Close()
        Return Nothing
    End Function
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        lidhje2.Close()
        lidhje2.Open()
        adaptori2 = New OleDbDataAdapter("SELECT * FROM Fornitoret WHERE(ID Like '" &
                                         TextBox1.Text & "%' OR Emri Like '" & TextBox1.Text & "%' OR Adresa Like '" &
                                         TextBox1.Text & "%' OR  Telefon Like '" & TextBox1.Text & "%' OR NIPT Like '" &
                                         TextBox1.Text & "%' OR  Kodi_Fornitorit Like  '" & TextBox1.Text & "%')", Lidhje.lidhje2)
        Dim dataTable As New DataTable
        adaptori2.Fill(dataTable)
        DataGridView2.DataSource = dataTable
        nr2 = DataGridView2.Rows.Count - 1
        For i = 0 To nr2
            DataGridView2.Rows(i).Cells(0).Value = i + 1
        Next
        Merrtedhenat_Fornitori(rreshtiaktual2)
        DataGridView2.CurrentCell = Nothing
        If DataGridView2.CurrentCell Is Nothing Then

        Else
            Dim row As DataGridViewRow = Me.DataGridView2.Rows.Item(rreshtiaktual2)
            TextBox5.Text = row.Cells.Item("ID").Value.ToString
            TextBox6.Text = row.Cells.Item("Emri").Value.ToString
            TextBox7.Text = row.Cells.Item("Adresa").Value.ToString
            TextBox8.Text = row.Cells.Item("Telefon").Value.ToString
            TextBox9.Text = row.Cells.Item("NIPT").Value.ToString
            TextBox10.Text = row.Cells.Item("Kodi_Fornitorit").Value.ToString
            lidhje2.Close()
            DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            DataGridView2.RefreshEdit()
            lidhje2.Close()
        End If
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        lidhje2.Close()
        lidhje2.Open()
        adaptori2 = New OleDbDataAdapter("SELECT ID,Emri,Fjalkalimi,Niveli FROM Perdoruesit WHERE(ID Like '" &
                                         TextBox11.Text & "%' OR Emri Like '" & TextBox11.Text & "%' OR Fjalkalimi Like '" &
                                         TextBox11.Text & "%' OR  Niveli Like '" & TextBox11.Text & "%')", Lidhje.lidhje2)
        Dim dataTable As New DataTable
        adaptori2.Fill(dataTable)
        DataGridView1.DataSource = dataTable
        nr1 = DataGridView1.Rows.Count - 1
        For i = 0 To nr1
            DataGridView1.Rows(i).Cells(0).Value = i + 1
        Next
        Merrtedhenat_Fornitori(rreshtiaktual2)
        DataGridView1.CurrentCell = Nothing
        If DataGridView1.CurrentCell Is Nothing Then

        Else
            Dim row As DataGridViewRow = Me.DataGridView1.Rows.Item(rreshtiaktual2)
            TextBox2.Text = row.Cells.Item("ID").Value.ToString
            TextBox3.Text = row.Cells.Item("Emri").Value.ToString
            TextBox4.Text = row.Cells.Item("Fjalkalimi").Value.ToString
            ComboBox1.Text = row.Cells.Item("Niveli").Value.ToString
            If row.Cells.Item("Niveli").Value.ToString = "Administrator" Then
                Label13.ForeColor = Color.Green
            Else
                Label13.ForeColor = Color.Red
            End If
            Label13.Text = row.Cells.Item("Niveli").Value.ToString
            lidhje2.Close()
            DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            DataGridView1.RefreshEdit()
            lidhje2.Close()
        End If
    End Sub
    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
        If (e.RowIndex >= 0) Then
            Dim row As DataGridViewRow = DataGridView3.Rows.Item(e.RowIndex)
            TextBox12.Text = row.Cells.Item("ID").Value.ToString
            TextBox13.Text = row.Cells.Item("Emri").Value.ToString
            TextBox14.Text = row.Cells.Item("Adresa").Value.ToString
            TextBox15.Text = row.Cells.Item("Telefon").Value.ToString
            TextBox16.Text = row.Cells.Item("NIPT").Value.ToString
            TextBox17.Text = row.Cells.Item("Kodi_klientit").Value.ToString
        End If
    End Sub
    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged
        lidhje2.Close()
        lidhje2.Open()
        adaptori2 = New OleDbDataAdapter("SELECT * FROM Bleresit WHERE(ID Like '" &
                                         TextBox18.Text & "%' OR Emri Like '" & TextBox18.Text & "%' OR Adresa Like '" &
                                          TextBox18.Text & "%' OR  Telefon Like '" & TextBox18.Text & "%' OR NIPT Like '" &
                                          TextBox18.Text & "%' OR  Kodi_klientit Like  '" & TextBox18.Text & "%')", Lidhje.lidhje2)
        Dim dataTable As New DataTable
        adaptori2.Fill(dataTable)
        DataGridView3.DataSource = dataTable
        nr3 = DataGridView3.Rows.Count - 1
        For i = 0 To nr3
            DataGridView3.Rows(i).Cells(0).Value = i + 1
        Next
        Merrtedhenat_Bleresit(rreshtiaktual2)
        DataGridView3.CurrentCell = Nothing
        If DataGridView3.CurrentCell Is Nothing Then

        Else
            Dim row As DataGridViewRow = Me.DataGridView3.Rows.Item(rreshtiaktual2)
            TextBox5.Text = row.Cells.Item("ID").Value.ToString
            TextBox6.Text = row.Cells.Item("Emri").Value.ToString
            TextBox7.Text = row.Cells.Item("Adresa").Value.ToString
            TextBox8.Text = row.Cells.Item("Telefon").Value.ToString
            TextBox9.Text = row.Cells.Item("NIPT").Value.ToString
            TextBox10.Text = row.Cells.Item("Kodi_klientit").Value.ToString
            lidhje2.Close()
            DataGridView3.Sort(DataGridView3.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            DataGridView3.RefreshEdit()
            lidhje2.Close()
        End If
    End Sub
    Public lidhje2pro As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptori2pro As OleDbDataAdapter
    Public lexuesipro As OleDbDataReader
    Public query2pro As OleDbCommand
    Public setitedhenave2pro = New DataSet
    Public rreshtiaktual2pro As Integer
    Public rreshtiaktualpro As Integer
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'Shto fornitor ne tabelen Klientet
        Dim nrekzistues As New List(Of Integer)
        For Each kolone As DataGridViewRow In DataGridView3.Rows
            nrekzistues.Add(CInt(kolone.Cells(0).Value))
        Next
        Dim existingNumbers As New List(Of Integer)
        For Each r As DataGridViewRow In DataGridView3.Rows
            existingNumbers.Add(CInt(r.Cells(0).Value))
        Next
        If existingNumbers.Count = 0 And nrekzistues.Count = 0 Then
            Dim Str As String
            Try
                Str = "insert into Bleresit values("
                Str += "1"
                Str += ","
                Str += """" & TextBox13.Text.Trim & """"
                Str += ","
                Str += """" & TextBox14.Text.Trim & """"
                Str += ","
                Str += """" & TextBox15.Text.Trim() & """"
                Str += ","
                Str += """" & TextBox16.Text.Trim() & """"
                Str += ","
                Str += """" & GenerateRandomString(8) & """"
                Str += ")"
                lidhje2pro.Open()
                query2pro = New OleDbCommand(Str, lidhje2pro)
                query2pro.ExecuteNonQuery()
                setitedhenave2pro.Clear()
                adaptori2pro = New OleDbDataAdapter("SELECT * FROM Bleresit ORDER BY ID", lidhje2pro)
                adaptori2pro.Fill(setitedhenave2pro, "tedhena")
                DataGridView3.DataSource = setitedhenave2pro.Tables(0)
                Dim nr3 As Integer
                nr3 = DataGridView3.Rows.Count - 1
                For i = 0 To nr3
                    DataGridView3.Rows(i).Cells(0).Value = i + 1
                Next
                MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                lidhje2pro.Close()
                TextBox13.Text = ""
                TextBox14.Text = ""
                TextBox15.Text = ""
                TextBox16.Text = ""
            Catch ex As Exception
                MsgBox("Nuk u shtua", MsgBoxStyle.Information)
                lidhje2pro.Close()
            End Try
        Else
            Dim missingNumbers = Enumerable.Range(existingNumbers.First, existingNumbers.Max() - existingNumbers.First + 1).Except(existingNumbers)
            Dim max = nrekzistues.Max() + 1
            If missingNumbers.Count = 0 Then
                Dim Str As String
                Try
                    Str = "insert into Bleresit values("
                    Str += max.ToString
                    Str += ","
                    Str += """" & TextBox13.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox14.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox15.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox16.Text.Trim() & """"
                    Str += ","
                    Str += """" & GenerateRandomString(8) & """"
                    Str += ")"
                    lidhje2pro.Open()
                    query2pro = New OleDbCommand(Str, lidhje2pro)
                    query2pro.ExecuteNonQuery()
                    setitedhenave2pro.Clear()
                    adaptori2pro = New OleDbDataAdapter("SELECT * FROM Bleresit ORDER BY ID", lidhje2pro)
                    adaptori2pro.Fill(setitedhenave2pro, "tedhena")
                    DataGridView3.DataSource = setitedhenave2pro.Tables(0)
                    Dim nr3 As Integer
                    nr3 = DataGridView3.Rows.Count - 1
                    For i = 0 To nr3
                        DataGridView3.Rows(i).Cells(0).Value = i + 1
                    Next
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    lidhje2pro.Close()
                    TextBox13.Text = ""
                    TextBox14.Text = ""
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                Catch ex As Exception
                    MsgBox("Nuk u shtua", MsgBoxStyle.Information)
                    lidhje2pro.Close()
                End Try
            Else
                Dim Str As String
                Try
                    Str = "insert into Bleresit values("
                    Str += missingNumbers.First.ToString
                    Str += ","
                    Str += """" & TextBox13.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox14.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox15.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox16.Text.Trim() & """"
                    Str += ","
                    Str += """" & GenerateRandomString(8) & """"
                    Str += ")"
                    lidhje2pro.Open()
                    query2pro = New OleDbCommand(Str, lidhje2pro)
                    query2pro.ExecuteNonQuery()
                    setitedhenave2pro.Clear()
                    adaptori2pro = New OleDbDataAdapter("SELECT * FROM Bleresit ORDER BY ID", lidhje2pro)
                    adaptori2pro.Fill(setitedhenave2pro, "tedhena")
                    DataGridView3.DataSource = setitedhenave2pro.Tables(0)
                    Dim nr3 As Integer
                    nr3 = DataGridView3.Rows.Count - 1
                    For i = 0 To nr3
                        DataGridView3.Rows(i).Cells(0).Value = i + 1
                    Next
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    lidhje2pro.Close()
                    TextBox13.Text = ""
                    TextBox14.Text = ""
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                Catch ex As Exception
                    MsgBox("Nuk u shtua", MsgBoxStyle.Information)
                    lidhje2pro.Close()
                End Try
            End If
        End If
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            Dim Str As String
            Str = "update Bleresit set Emri="
            Str += """" & TextBox13.Text & """"
            Str += " where ID="
            Str += TextBox12.Text.Trim()
            lidhjekompaniax.Open()
            queryfshijkompaniax = New OleDbCommand(Str, lidhjekompaniax)
            queryfshijkompaniax.ExecuteNonQuery()
            lidhjekompaniax.Close()

            lidhjekompaniax.Open()
            Str = "update Bleresit set Adresa="
            Str += """" & TextBox14.Text & """"
            Str += " where ID="
            Str += TextBox12.Text.Trim()
            queryfshijkompaniax = New OleDbCommand(Str, lidhjekompaniax)
            queryfshijkompaniax.ExecuteNonQuery()
            lidhjekompaniax.Close()

            lidhjekompaniax.Open()
            Str = "update Bleresit set Telefon="
            Str += """" & TextBox15.Text & """"
            Str += " where ID="
            Str += TextBox12.Text.Trim()
            queryfshijkompaniax = New OleDbCommand(Str, lidhjekompaniax)
            queryfshijkompaniax.ExecuteNonQuery()
            lidhjekompaniax.Close()

            lidhjekompaniax.Open()
            Str = "update Bleresit set NIPT="
            Str += """" & TextBox16.Text & """"
            Str += " where ID="
            Str += TextBox12.Text.Trim()
            queryfshijkompaniax = New OleDbCommand(Str, lidhjekompaniax)
            queryfshijkompaniax.ExecuteNonQuery()
            lidhjekompaniax.Close()

            lidhjekompaniax.Open()
            setitedhenavefshijkompaniax.Clear()
            adaptorikompaniax = New OleDbDataAdapter("SELECT * FROM Bleresit ORDER BY ID", lidhjekompaniax)
            adaptorikompaniax.Fill(setitedhenavefshijkompaniax, "tedhena")
            DataGridView3.DataSource = setitedhenavefshijkompaniax.Tables(0)
            DataGridView3.Refresh()
            lidhjekompaniax.Close()


            setitedhenavefshijkompaniaxy.clear
            With DataGridView3
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            rreshtiaktualfshijkompaniaxy = 0
            lidhjekompaniaxy.Open()
            adaptorikompaniaxy = New OleDbDataAdapter("SELECT * FROM Bleresit ORDER BY ID", lidhjekompaniaxy)
            adaptorikompaniaxy.Fill(setitedhenavefshijkompaniaxy, "tedhena")
            DataGridView3.DataSource = setitedhenavefshijkompaniaxy.Tables(0)
            Dim nr3 As Integer = 0
            nr3 = DataGridView3.Rows.Count - 1
            For i = 0 To nr3
                DataGridView3.Rows(i).Cells(0).Value = i + 1
            Next
            Merrtedhenat_Bleresit(rreshtiaktualfshijkompaniaxy)
            lidhjekompaniaxy.Close()
            MsgBox("U rifreskua me sukses", MsgBoxStyle.Information)
        Catch ex As Exception
            lidhjekompaniaxy.Close()
            MsgBox("Nuk u rifreskua me sukses", MsgBoxStyle.Information)
        End Try
    End Sub
    Sub DataGridView4_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        DataGridView4.CurrentCell = DataGridView4.Rows(1).Cells(1)
    End Sub
    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick
        If (e.RowIndex >= 0) Then
            Dim row As DataGridViewRow = DataGridView4.Rows.Item(e.RowIndex)
            TextBox19.Text = row.Cells.Item("ID").Value.ToString
            TextBox20.Text = row.Cells.Item("Emri_produktit").Value.ToString
            TextBox23.Text = row.Cells.Item("Njesia").Value.ToString
            TextBox21.Text = row.Cells.Item("Cmimi_blerjes").Value.ToString
            TextBox22.Text = row.Cells.Item("Cmimi_Shitjes").Value.ToString
        End If
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim nrekzistues As New List(Of Integer)
        For Each kolone As DataGridViewRow In DataGridView4.Rows
            nrekzistues.Add(CInt(kolone.Cells(0).Value))
        Next
        Dim existingNumbers As New List(Of Integer)
        For Each r As DataGridViewRow In DataGridView4.Rows
            existingNumbers.Add(CInt(r.Cells(0).Value))
        Next
        If existingNumbers.Count = 0 And nrekzistues.Count = 0 Then
            Dim Str As String
            Try
                Str = "insert into Produktet values("
                Str += "1"
                Str += ","
                Str += """" & TextBox20.Text & """"
                Str += ","
                Str += """" & TextBox23.Text.Trim & """"
                Str += ","
                Str += """" & CDbl(TextBox21.Text.Trim()).ToString("#,#", CultureInfo.InvariantCulture) & """"
                Str += ","
                Str += """" & CDbl(TextBox22.Text.Trim()).ToString("#,#", CultureInfo.InvariantCulture) & """"

                Str += ","
                Str += """" & TextBox25.Text.Trim() & """"
                Str += ","
                Str += """" & TextBox26.Text.Trim() & """"


                Str += ")"

                lidhje2113.Open()
                query2113 = New OleDbCommand(Str, lidhje2113)
                query2113.ExecuteNonQuery()
                setitedhenave2113.Clear()
                adaptori2113 = New OleDbDataAdapter("SELECT * FROM Produktet ORDER BY ID", lidhje2113)
                adaptori2113.Fill(setitedhenave2113, "tedhena")
                DataGridView4.DataSource = setitedhenave2113.Tables(0)
                Dim nr4 As Integer
                nr3 = DataGridView4.Rows.Count - 1
                For i = 0 To nr4
                    DataGridView4.Rows(i).Cells(0).Value = i + 1
                Next
                MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                lidhje2113.Close()
                TextBox20.Text = ""
                TextBox21.Text = ""
                TextBox22.Text = ""
                TextBox23.Text = ""
                TextBox24.Text = ""
                TextBox25.Text = ""
                TextBox26.Text = ""
            Catch ex As Exception
                MsgBox("U shtua!", MsgBoxStyle.Information)
                lidhje2113.Close()
            End Try

        Else
            Dim missingNumbers = Enumerable.Range(existingNumbers.First, existingNumbers.Max() - existingNumbers.First + 1).Except(existingNumbers)
            Dim max = nrekzistues.Max() + 1
            If missingNumbers.Count = 0 Then
                Dim Str As String
                Try
                    Str = "insert into Produktet values("
                    Str += max.ToString
                    Str += ","
                    Str += """" & TextBox20.Text & """"
                    Str += ","
                    Str += """" & TextBox23.Text.Trim & """"
                    Str += ","
                    Str += """" & CDbl(TextBox21.Text.Trim()).ToString("#,#", CultureInfo.InvariantCulture) & """"
                    Str += ","
                    Str += """" & CDbl(TextBox22.Text.Trim()).ToString("#,#", CultureInfo.InvariantCulture) & """"
                    Str += ","
                    Str += """" & TextBox25.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox26.Text.Trim() & """"
                    Str += ")"

                    lidhje2113.Open()
                    query2113 = New OleDbCommand(Str, lidhje2113)
                    query2113.ExecuteNonQuery()
                    setitedhenave2113.Clear()
                    adaptori2113 = New OleDbDataAdapter("SELECT * FROM Produktet ORDER BY ID", lidhje2113)
                    adaptori2113.Fill(setitedhenave2113, "tedhena")
                    DataGridView4.DataSource = setitedhenave2113.Tables(0)
                    Dim nr4 As Integer
                    nr4 = DataGridView4.Rows.Count - 1
                    For i = 0 To nr4
                        DataGridView4.Rows(i).Cells(0).Value = i + 1
                    Next
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    lidhje2113.Close()
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                Catch ex As Exception
                    MsgBox("Nuk u shtua me sukses", MsgBoxStyle.Information)

                    lidhje2113.Close()
                End Try
            Else
                Dim Str As String
                Try
                    Str = "insert into Produktet values("
                    Str += missingNumbers.First.ToString
                    Str += ","
                    Str += """" & TextBox20.Text & """"
                    Str += ","
                    Str += """" & TextBox23.Text.Trim & """"
                    Str += ","
                    Str += """" & CDbl(TextBox21.Text.Trim()).ToString("#,#", CultureInfo.InvariantCulture) & """"
                    Str += ","
                    Str += """" & CDbl(TextBox22.Text.Trim()).ToString("#,#", CultureInfo.InvariantCulture) & """"
                    Str += ","
                    Str += """" & TextBox25.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox26.Text.Trim() & """"
                    Str += ")"

                    lidhje2113.Open()
                    query2113 = New OleDbCommand(Str, lidhje2113)
                    query2113.ExecuteNonQuery()
                    setitedhenave2113.Clear()
                    adaptori2113 = New OleDbDataAdapter("SELECT * FROM Produktet ORDER BY ID", lidhje2113)
                    adaptori2113.Fill(setitedhenave2113, "tedhena")
                    DataGridView4.DataSource = setitedhenave2113.Tables(0)
                    Dim nr4 As Integer
                    nr4 = DataGridView4.Rows.Count - 1
                    For i = 0 To nr4
                        DataGridView4.Rows(i).Cells(0).Value = i + 1
                    Next
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    lidhje2113.Close()
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                Catch ex As Exception
                    MsgBox("Nuk u shtua", MsgBoxStyle.Information)
                    lidhje2113.Close()
                End Try
            End If
        End If
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Try
            Dim Str As String
            Str = "update Produktet set Emri_produktit="
            Str += """" & TextBox20.Text & """"
            Str += " where ID="
            Str += TextBox19.Text.Trim()
            lidhje211.Open()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Produktet set Njesia="
            Str += """" & TextBox23.Text & """"
            Str += " where ID="
            Str += TextBox19.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Produktet set Cmimi_blerjes="
            Str += """" & TextBox21.Text & """"
            Str += " where ID="
            Str += TextBox19.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Produktet set Cmimi_Shitjes="
            Str += """" & TextBox22.Text & """"
            Str += " where ID="
            Str += TextBox19.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Produktet set Kodi_produktit="
            Str += """" & TextBox25.Text & """"
            Str += " where ID="
            Str += TextBox19.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Produktet set Sasia_limit="
            Str += """" & TextBox26.Text & """"
            Str += " where ID="
            Str += TextBox19.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            setitedhenave211.Clear()
            adaptori211 = New OleDbDataAdapter("SELECT * FROM Produktet ORDER BY ID", lidhje211)
            adaptori211.Fill(setitedhenave211, "tedhena")
            DataGridView4.DataSource = setitedhenave211.Tables(0)
            Dim nr4 As Integer = 0
            nr4 = DataGridView4.Rows.Count - 1
            For i = 0 To nr4
                DataGridView4.Rows(i).Cells(0).Value = i + 1
            Next
            lidhje211.Close()
            MsgBox("U rifreskua me sukses", MsgBoxStyle.Information)
        Catch ex As Exception
            lidhje211.Close()
            MsgBox("Nuk u rifreskua me sukses", MsgBoxStyle.Information)
        End Try
    End Sub
    Public lidhjekompaniaxyx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptorikompaniaxyx As OleDbDataAdapter
    Public lexuesifshijkompaniaxyx As OleDbDataReader
    Public queryfshijkompaniaxyx As OleDbCommand
    Public setitedhenavefshijkompaniaxyx = New DataSet
    Public rreshtiaktualfshijkompaniaxyx As Integer
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If Not DataGridView3.Rows.Count > 0 Then
            MsgBox("Lista eshte bosh!", MsgBoxStyle.Information)
        Else
            Dim Str As String
            Try
                Str = "delete from Bleresit where ID="
                Str += DataGridView3.CurrentRow.Cells(1).Value.ToString
                lidhjekompaniaxyx.Open()
                queryfshijkompaniaxyx = New OleDbCommand(Str, lidhjekompaniaxyx)
                queryfshijkompaniaxyx.ExecuteNonQuery()
                setitedhenavefshijkompaniaxyx.clear()
                adaptorikompaniaxyx = New OleDbDataAdapter("SELECT * FROM Bleresit ORDER BY ID", lidhjekompaniaxyx)
                adaptorikompaniaxyx.Fill(setitedhenavefshijkompaniaxyx, "tedhena")
                DataGridView3.DataSource = setitedhenavefshijkompaniaxyx.Tables(0)
                Dim nr3 As Integer = 0
                nr3 = DataGridView3.Rows.Count - 1
                For i = 0 To nr3
                    DataGridView3.Rows(i).Cells(0).Value = i + 1
                Next
                DataGridView3.Refresh()
                MsgBox("U fshi me sukses", MsgBoxStyle.Information)
                If rreshtiaktualfshijkompaniaxyx > 0 Then
                    rreshtiaktualfshijkompaniaxyx -= 1
                    Merrtedhenat_Bleresit(rreshtiaktualfshijkompaniaxyx)
                End If
                lidhjekompaniaxyx.Close()
            Catch ex As Exception
                MsgBox("Nuk u fshi me sukses", MsgBoxStyle.Information)
                lidhjekompaniaxyx.Close()
            End Try
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Str As String
        Try
            Str = "delete from Fornitoret where ID="
            Str += DataGridView2.CurrentRow.Cells(1).Value.ToString
            lidhje211.Open()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            setitedhenave211.clear()
            adaptori211 = New OleDbDataAdapter("SELECT * FROM Fornitoret ORDER BY ID", lidhje211)
            adaptori211.Fill(setitedhenave211, "tedhena")
            DataGridView2.DataSource = setitedhenave211.Tables(0)
            nr2 = DataGridView2.Rows.Count - 1
            For i = 0 To nr2
                DataGridView2.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView2.Refresh()
            MsgBox("U fshi me sukses", MsgBoxStyle.Information)
            If rreshtiaktual11 > 0 Then
                rreshtiaktual11 -= 1
                Merrtedhenat2(rreshtiaktual11)
            End If
            lidhje211.Close()
        Catch ex As Exception
            MsgBox("Nuk u fshi", MsgBoxStyle.Information)
            lidhje211.Close()
        End Try
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim Str As String
        Try
            Str = "delete from Produktet where ID="
            Str += DataGridView4.CurrentRow.Cells(1).Value.ToString
            lidhje211.Open()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            setitedhenave211.clear()
            adaptori211 = New OleDbDataAdapter("SELECT * FROM Produktet ORDER BY ID", lidhje211)
            adaptori211.Fill(setitedhenave211, "tedhena")
            DataGridView4.DataSource = setitedhenave211.Tables(0)
            nr4 = DataGridView4.Rows.Count - 1
            For i = 0 To nr4
                DataGridView4.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView4.Refresh()
            MsgBox("U fshi me sukses", MsgBoxStyle.Information)
            If rreshtiaktual11 > 0 Then
                rreshtiaktual11 -= 1
                Merrtedhenat2(rreshtiaktual11)
            End If
            lidhje211.Close()
        Catch ex As Exception
            MsgBox("Nuk u fshi", MsgBoxStyle.Information)
            lidhje211.Close()
        End Try
    End Sub
    Private Sub TextBox28_TextChanged(sender As Object, e As EventArgs) Handles TextBox28.TextChanged
        lidhje2.Close()
        lidhje2.Open()
        adaptori2 = New OleDbDataAdapter("SELECT * FROM Produktet WHERE(ID Like '" &
                                         TextBox28.Text & "%' OR Emri_produktit Like '" & TextBox28.Text & "%' OR Njesia Like '" &
                                          TextBox28.Text & "%'  OR Cmimi_blerjes Like '" &
                                           TextBox28.Text & "%' OR Cmimi_Shitjes Like '" &
                                            TextBox28.Text & "%'" & "Or Kodi_produktit Like '" &
                                            TextBox28.Text & "Or Sasia_limit Like '" &
                                            TextBox28.Text & ")", Lidhje.lidhje2)
        Dim dataTable As New DataTable
        adaptori2.Fill(dataTable)
        DataGridView4.DataSource = dataTable
        nr4 = DataGridView4.Rows.Count - 1
        For i = 0 To nr4
            DataGridView4.Rows(i).Cells(0).Value = i + 1
        Next
        Merrtedhenat_produktet(rreshtiaktual2)
        DataGridView4.CurrentCell = Nothing
        If DataGridView4.CurrentCell Is Nothing Then

        Else
            Dim row As DataGridViewRow = Me.DataGridView4.Rows.Item(rreshtiaktual2)
            TextBox19.Text = row.Cells.Item("ID").Value.ToString
            TextBox20.Text = row.Cells.Item("Emri_produktit").Value.ToString
            TextBox23.Text = row.Cells.Item("Njesia").Value.ToString
            TextBox21.Text = row.Cells.Item("Cmimi_blerjes").Value.ToString
            TextBox22.Text = row.Cells.Item("Cmimi_Shitjes").Value.ToString
            TextBox25.Text = row.Cells.Item("Kodi_produktit").Value.ToString
            TextBox26.Text = row.Cells.Item("Sasia_limit").Value.ToString
            lidhje2.Close()
            DataGridView4.Sort(DataGridView4.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            DataGridView4.RefreshEdit()
            lidhje2.Close()
        End If
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        'Shto fornitor ne tabelen Fornitoret
        Dim nrekzistues As New List(Of Integer)
        For Each kolone As DataGridViewRow In DataGridView5.Rows
            nrekzistues.Add(CInt(kolone.Cells(0).Value))
        Next
        Dim existingNumbers As New List(Of Integer)
        For Each r As DataGridViewRow In DataGridView5.Rows
            existingNumbers.Add(CInt(r.Cells(0).Value))
        Next
        If existingNumbers.Count = 0 And nrekzistues.Count = 0 Then
            Dim Str As String
            Try
                Str = "insert into Shitesi values("
                Str += "1"
                Str += ","
                Str += """" & TextBox35.Text.Trim & """"
                Str += ","
                Str += """" & TextBox36.Text.Trim & """"
                Str += ","
                Str += """" & TextBox37.Text.Trim() & """"
                Str += ","
                Str += """" & TextBox38.Text.Trim() & """"
                Str += ")"
                lidhjekompania.Open()
                queryfshijkompania = New OleDbCommand(Str, lidhjekompania)
                queryfshijkompania.ExecuteNonQuery()
                setitedhenavefshijkompania.Clear()
                adaptorikompania = New OleDbDataAdapter("SELECT * FROM Shitesi ORDER BY ID", lidhjekompania)
                adaptorikompania.Fill(setitedhenavefshijkompania, "tedhena")
                Dim nr5 As Integer = 0
                DataGridView5.DataSource = setitedhenavefshijkompania.Tables(0)
                nr5 = DataGridView5.Rows.Count - 1
                For i = 0 To nr5
                    DataGridView5.Rows(i).Cells(0).Value = i + 1
                Next
                MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                lidhjekompania.Close()
                TextBox35.Text = ""
                TextBox36.Text = ""
                TextBox37.Text = ""
                TextBox38.Text = ""

            Catch ex As Exception
                MsgBox("Nuk u shtua", MsgBoxStyle.Information)
                lidhjekompania.Close()
            End Try
        Else
            Dim missingNumbers = Enumerable.Range(existingNumbers.First, existingNumbers.Max() - existingNumbers.First + 1).Except(existingNumbers)
            If missingNumbers.Count = 0 Then
                Dim max = nrekzistues.Max() + 1
                Dim Str As String
                Try
                    Str = "insert into Shitesi values("
                    Str += max.ToString
                    Str += ","
                    Str += """" & TextBox35.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox36.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox37.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox38.Text.Trim() & """"
                    Str += ")"
                    lidhjekompania.Open()
                    queryfshijkompania = New OleDbCommand(Str, lidhjekompania)
                    queryfshijkompania.ExecuteNonQuery()
                    setitedhenavefshijkompania.Clear()
                    adaptorikompania = New OleDbDataAdapter("SELECT * FROM Shitesi ORDER BY ID", lidhjekompania)
                    adaptorikompania.Fill(setitedhenavefshijkompania, "tedhena")
                    DataGridView5.DataSource = setitedhenavefshijkompania.Tables(0)
                    Dim nr5 As Integer
                    nr5 = DataGridView5.Rows.Count - 1
                    For i = 0 To nr5
                        DataGridView5.Rows(i).Cells(0).Value = i + 1
                    Next
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    lidhjekompania.Close()
                    TextBox35.Text = ""
                    TextBox36.Text = ""
                    TextBox37.Text = ""
                    TextBox38.Text = ""
                Catch ex As Exception
                    MsgBox("Nuk u shtua", MsgBoxStyle.Information)
                    lidhjekompania.Close()
                End Try
            Else
                Dim Str As String
                Try
                    Str = "insert into Shitesi values("
                    Str += missingNumbers.First.ToString
                    Str += ","
                    Str += """" & TextBox35.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox36.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox37.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox38.Text.Trim() & """"
                    Str += ")"
                    lidhjekompania.Open()
                    queryfshijkompania = New OleDbCommand(Str, lidhjekompania)
                    queryfshijkompania.ExecuteNonQuery()
                    setitedhenavefshijkompania.Clear()
                    adaptorikompania = New OleDbDataAdapter("SELECT * FROM Shitesi ORDER BY ID", lidhjekompania)
                    adaptorikompania.Fill(setitedhenavefshijkompania, "tedhena")
                    DataGridView5.DataSource = setitedhenavefshijkompania.Tables(0)
                    Dim nr5 As Integer
                    nr5 = DataGridView5.Rows.Count - 1
                    For i = 0 To nr5
                        DataGridView5.Rows(i).Cells(0).Value = i + 1
                    Next
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    lidhjekompania.Close()
                    TextBox35.Text = ""
                    TextBox36.Text = ""
                    TextBox37.Text = ""
                    TextBox38.Text = ""
                Catch ex As Exception
                    MsgBox("Nuk u shtua", MsgBoxStyle.Information)
                    lidhjekompania.Close()
                End Try
            End If
        End If
    End Sub
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Try
            Dim Str As String
            Str = "update Shitesi set Emri="
            Str += """" & TextBox35.Text & """"
            Str += " where ID="
            Str += TextBox24.Text.Trim()
            lidhjekompania.Open()
            queryfshijkompania = New OleDbCommand(Str, lidhjekompania)
            queryfshijkompania.ExecuteNonQuery()
            lidhjekompania.Close()

            lidhjekompania.Open()
            Str = "update Shitesi set Adresa="
            Str += """" & TextBox36.Text & """"
            Str += " where ID="
            Str += TextBox24.Text.Trim()
            queryfshijkompania = New OleDbCommand(Str, lidhjekompania)
            queryfshijkompania.ExecuteNonQuery()
            lidhjekompania.Close()

            lidhjekompania.Open()
            Str = "update Shitesi set Telefon="
            Str += """" & TextBox37.Text & """"
            Str += " where ID="
            Str += TextBox24.Text.Trim()
            queryfshijkompania = New OleDbCommand(Str, lidhjekompania)
            queryfshijkompania.ExecuteNonQuery()
            lidhjekompania.Close()

            lidhjekompania.Open()
            Str = "update Shitesi set NIPT="
            Str += """" & TextBox38.Text & """"
            Str += " where ID="
            Str += TextBox24.Text.Trim()
            queryfshijkompania = New OleDbCommand(Str, lidhjekompania)
            queryfshijkompania.ExecuteNonQuery()
            lidhjekompania.Close()

            lidhjekompania.Open()
            setitedhenavefshijkompania.Clear()
            adaptorikompania = New OleDbDataAdapter("SELECT * FROM Shitesi ORDER BY ID", lidhje211)
            adaptorikompania.Fill(setitedhenavefshijkompania, "tedhena")
            DataGridView5.DataSource = setitedhenavefshijkompania.Tables(0)
            Dim nr5 As Integer = 0
            nr5 = DataGridView5.Rows.Count - 1
            For i = 0 To nr5
                DataGridView5.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView5.Refresh()
            lidhjekompania.Close()
            MsgBox("U rifreskua me sukses", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("Nuk u rifreskua me sukses", MsgBoxStyle.Information)
        End Try
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If Not DataGridView5.Rows.Count > 0 Then
            MsgBox("Lista eshte bosh!", MsgBoxStyle.Information)
        Else
            Dim Str As String
            Try
                Str = "delete from Shitesi where ID="
                Str += DataGridView5.CurrentRow.Cells(1).Value.ToString
                lidhjekompania.Open()
                queryfshijkompania = New OleDbCommand(Str, lidhjekompania)
                queryfshijkompania.ExecuteNonQuery()
                setitedhenavefshijkompania.clear()
                adaptorikompania = New OleDbDataAdapter("SELECT * FROM Shitesi ORDER BY ID", lidhjekompania)
                adaptorikompania.Fill(setitedhenavefshijkompania, "tedhena")
                DataGridView5.DataSource = setitedhenavefshijkompania.Tables(0)
                Dim nr5 As Integer = 0
                nr5 = DataGridView5.Rows.Count - 1
                For i = 0 To nr5
                    DataGridView5.Rows(i).Cells(0).Value = i + 1
                Next
                DataGridView5.Refresh()
                MsgBox("U fshi me sukses", MsgBoxStyle.Information)
                If rreshtiaktualfshijkompania > 0 Then
                    rreshtiaktualfshijkompania -= 1
                    Merrtedhenat_kompania(rreshtiaktualfshijkompania)
                End If
                lidhjekompania.Close()
            Catch ex As Exception
                MsgBox("Nuk u fshi", MsgBoxStyle.Information)
                lidhjekompania.Close()
            End Try
        End If
    End Sub
    Private Sub DataGridView5_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellContentClick
        If (e.RowIndex >= 0) Then
            Dim row As DataGridViewRow = DataGridView5.Rows.Item(e.RowIndex)
            TextBox24.Text = row.Cells.Item("ID").Value.ToString
            TextBox35.Text = row.Cells.Item("Emri").Value.ToString
            TextBox36.Text = row.Cells.Item("Adresa").Value.ToString
            TextBox37.Text = row.Cells.Item("Telefon").Value.ToString
            TextBox38.Text = row.Cells.Item("NIPT").Value.ToString
        End If
    End Sub
    Private Sub TextBox39_TextChanged(sender As Object, e As EventArgs) Handles TextBox39.TextChanged
        lidhjekompania.Close()
        lidhjekompania.Open()
        adaptorikompania = New OleDbDataAdapter("SELECT * FROM Shitesi WHERE(ID Like '" &
                                         TextBox39.Text & "%' OR Emri Like '" & TextBox39.Text & "%' OR Adresa Like '" &
                                         TextBox39.Text & "%' OR  Telefon Like '" & TextBox39.Text & "%' OR NIPT Like '" &
                                         TextBox39.Text & "%')", lidhjekompania)
        Dim dataTable As New DataTable
        dataTable.Clear()
        adaptorikompania.Fill(dataTable)
        DataGridView5.DataSource = dataTable
        nr2 = DataGridView5.Rows.Count - 1
        For i = 0 To nr2
            DataGridView5.Rows(i).Cells(0).Value = i + 1
        Next
        Merrtedhenat_kompania(rreshtiaktualfshijkompania)
        DataGridView5.CurrentCell = Nothing
        If DataGridView5.CurrentCell Is Nothing Then

        Else
            Dim row As DataGridViewRow = Me.DataGridView5.Rows.Item(rreshtiaktual2)
            TextBox24.Text = row.Cells.Item("ID").Value.ToString
            TextBox35.Text = row.Cells.Item("Emri").Value.ToString
            TextBox36.Text = row.Cells.Item("Adresa").Value.ToString
            TextBox37.Text = row.Cells.Item("Telefon").Value.ToString
            TextBox38.Text = row.Cells.Item("NIPT").Value.ToString
            lidhjekompania.Close()
            DataGridView5.Sort(DataGridView5.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            DataGridView5.RefreshEdit()
            lidhjekompania.Close()
        End If
    End Sub
End Class