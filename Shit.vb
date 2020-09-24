Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Public Class Shit
    Private _form_resize As clsResize
    Dim path As String = My.Settings.ruajdtbpath & "\tedhena.accdb;"
    Public Sub New()
        InitializeComponent()
        _form_resize = New clsResize(Me)
        AddHandler Me.Load, AddressOf _Load
        AddHandler Me.Resize, AddressOf _Resize
    End Sub
    Private Sub _Resize(ByVal sender As Object, ByVal e As EventArgs)
        _form_resize._resize1()
        If Me.WindowState = FormWindowState.Maximized Then
            Label23.Text = "_________________________________________________________"
            Label24.Text = "_________________________________________________________"
        ElseIf Me.WindowState = FormWindowState.Normal Then
            Label23.Text = "____________________________________________"
            Label24.Text = "____________________________________________"
        End If
    End Sub
    Private Sub _Load(ByVal sender As Object, ByVal e As EventArgs)
        _form_resize._get_initial_size()
    End Sub

    Dim Lidhje1x As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Dim adaptor1x As OleDbDataAdapter
    Dim setidatave1x = New DataSet
    Dim query1x As OleDbCommand
    Dim rreshtiaktual2 As Integer
    'Public Const WM_NCLBUTTONDBLCLK As Integer = &HA3
    Dim totale As Integer = 0
    Dim totale1 As Integer = 0
    Dim totalefund As Integer = 0
    Dim Lidhje1 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Dim adaptor1 As OleDbDataAdapter
    Dim setidatave1 = New DataSet
    Dim query1 As OleDbCommand
    Private Sub ComboBox1_DrawItem(ByVal sender As Object,
    ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox1.DrawItem
        If e.Index < 0 Then Exit Sub
        Dim rect As System.Drawing.Rectangle = e.Bounds
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
        Dim rect As System.Drawing.Rectangle = e.Bounds
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
        Dim rect As System.Drawing.Rectangle = e.Bounds
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Or ComboBox2.Text = "" Or ComboBox3.Text = "" Then
            MsgBox("Plotesoni te gjitha fushat!", MsgBoxStyle.Information)
        Else
            Button3.Enabled = True
            Farmacia.DataGridView2.Columns("ID").Visible = True
            Dim zbritje = NumericUpDown1.Value * TextBox9.Text
            If CheckBox1.Checked = True Then
                Dim total1 As Decimal = zbritje * (1 - NumericUpDown2.Value / 100)
                'TextBox9.Text = total1.ToString("#,#", CultureInfo.InvariantCulture)
                Dim nrekzistues As New List(Of Integer)
                For Each kolone As DataGridViewRow In Farmacia.DataGridView2.Rows
                    nrekzistues.Add(CInt(kolone.Cells(1).Value))
                Next
                Dim existingNumbers As New List(Of Integer)
                For Each r As DataGridViewRow In Farmacia.DataGridView2.Rows
                    existingNumbers.Add(CInt(r.Cells(1).Value))
                Next
                If existingNumbers.Count = 0 And nrekzistues.Count = 0 Then
                    Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox10.Text
                    Dim Str1 As String
                    Dim regDate As DateTime = Date.Now
                    Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                    Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                    Dim strDate1 As String = regDate.ToString("HH:mm tt")
                    Try
                        Str1 = "insert into Shitjet values("
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
                        Str1 += """" & ComboBox1.Text.Trim() & """"
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
                        Str1 += """" & vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture) & """"
                        Str1 += ","
                        Str1 += """" & TextBox11.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & NumericUpDown2.Value & """"
                        Str1 += ","
                        Str1 += """" & Hyrje.TextBox1.Text.Trim() & """"
                        Str1 += ")"
                        Lidhje1.Open()
                        query1 = New OleDbCommand(Str1, Lidhje1)
                        query1.ExecuteNonQuery()
                        setidatave1.Clear()
                        adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Kodi_Shitjes LIKE '" & TextBox8.Text & "%')", Lidhje1)
                        adaptor1.Fill(setidatave1, "tedhena")
                        Dim nrshit As Integer = 0
                        DataGridView1.Columns.Add("count1", "Nr.")
                        DataGridView1.DataSource = setidatave1.Tables(0)
                        DataGridView1.Columns("ID").Visible = False
                        nrshit = DataGridView1.Rows.Count - 1
                        For i = 0 To nrshit
                            DataGridView1.Rows(i).Cells(0).Value = i + 1
                        Next
                        Lidhje1.Close()
                        Dim nr As Integer = 0
                        nr = DataGridView1.Columns.Count - 1
                        If nr > 15 Then
                            DataGridView1.Columns.RemoveAt(nr)
                        End If
                        Dim Lidhje1xx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
                        Dim adaptor1xx As OleDbDataAdapter
                        Dim setidatave1xx = New DataSet
                        With Farmacia.DataGridView2
                            .RowsDefaultCellStyle.BackColor = Color.Bisque
                            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                        End With
                        rreshtiaktual2 = 0
                        Lidhje1xx.Open()
                        adaptor1xx = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", Lidhje1)
                        adaptor1xx.Fill(setidatave1xx, "tedhena")
                        Farmacia.DataGridView2.DataSource = setidatave1xx.Tables(0)

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
                        Label21.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                        Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                        Label30.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                        MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                        Farmacia.DataGridView2.Columns("ID").Visible = False
                        Dim nr1 As Integer = 0
                        nr1 = Farmacia.DataGridView2.Rows.Count - 1
                        For i = 0 To nr1
                            Farmacia.DataGridView2.Rows(i).Cells(0).Value = i + 1
                        Next
                    Catch ex As Exception
                        Lidhje1x.Close()
                    End Try
                Else
                    Dim missingNumbers = Enumerable.Range(existingNumbers.First, existingNumbers.Max() - existingNumbers.First + 1).Except(existingNumbers)
                    Dim max = nrekzistues.Max() + 1
                    If missingNumbers.Count = 0 Then
                        Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox10.Text
                        Dim Str1 As String
                        Dim regDate As DateTime = Date.Now
                        Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                        Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                        Dim strDate1 As String = regDate.ToString("HH:mm tt")
                        Try
                            Str1 = "insert into Shitjet values("
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
                            Str1 += """" & ComboBox1.Text.Trim() & """"
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
                            Str1 += """" & vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture) & """"
                            Str1 += ","
                            Str1 += """" & TextBox11.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & NumericUpDown2.Value & """"
                            Str1 += ","
                            Str1 += """" & Hyrje.TextBox1.Text.Trim() & """"
                            Str1 += ")"
                            Lidhje1.Open()
                            query1 = New OleDbCommand(Str1, Lidhje1)
                            query1.ExecuteNonQuery()
                            setidatave1.Clear()
                            adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Kodi_Shitjes LIKE '" & TextBox8.Text & "%')", Lidhje1)
                            adaptor1.Fill(setidatave1, "tedhena")
                            Dim nrshit As Integer = 0
                            DataGridView1.Columns.Add("count1", "Nr.")
                            DataGridView1.DataSource = setidatave1.Tables(0)
                            DataGridView1.Columns("ID").Visible = False
                            nrshit = DataGridView1.Rows.Count - 1
                            For i = 0 To nrshit
                                DataGridView1.Rows(i).Cells(0).Value = i + 1
                            Next
                            Lidhje1.Close()
                            Dim nr As Integer = 0
                            nr = DataGridView1.Columns.Count - 1
                            If nr > 15 Then
                                DataGridView1.Columns.RemoveAt(nr)
                            End If
                            Dim Lidhje1xx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
                            Dim adaptor1xx As OleDbDataAdapter
                            Dim setidatave1xx = New DataSet
                            With Farmacia.DataGridView2
                                .RowsDefaultCellStyle.BackColor = Color.Bisque
                                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                            End With

                            Lidhje1xx.Open()
                            adaptor1xx = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", Lidhje1)
                            adaptor1xx.Fill(setidatave1xx, "tedhena")
                            Farmacia.DataGridView2.DataSource = setidatave1xx.Tables(0)

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
                            Label21.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                            Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                            Label30.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                            MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                            Farmacia.DataGridView2.Columns("ID").Visible = False
                            Dim nr1 As Integer = 0
                            nr1 = Farmacia.DataGridView2.Rows.Count - 1
                            For i = 0 To nr1
                                Farmacia.DataGridView2.Rows(i).Cells(0).Value = i + 1
                            Next
                        Catch ex As Exception

                            MsgBox("Nuk u shtua me sukses", MsgBoxStyle.Information)
                            Lidhje1x.Close()
                        End Try
                    Else
                        Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox10.Text
                        Farmacia.DataGridView2.Columns("ID").Visible = True
                        Dim Str1 As String
                        Dim regDate As DateTime = Date.Now
                        Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                        Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                        Dim strDate1 As String = regDate.ToString("HH:mm tt")
                        Try
                            Str1 = "insert into Shitjet values("
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
                            Str1 += """" & ComboBox1.Text.Trim() & """"
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
                            Str1 += """" & vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture) & """"
                            Str1 += ","
                            Str1 += """" & TextBox11.Text.Trim() & """"
                            Str1 += ","
                            Str1 += """" & NumericUpDown2.Value & """"
                            Str1 += ","
                            Str1 += """" & Hyrje.TextBox1.Text.Trim() & """"
                            Str1 += ")"
                            Lidhje1.Open()
                            query1 = New OleDbCommand(Str1, Lidhje1)
                            query1.ExecuteNonQuery()
                            setidatave1.Clear()
                            adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Kodi_Shitjes LIKE '" & TextBox8.Text & "%')", Lidhje1)
                            adaptor1.Fill(setidatave1, "tedhena")
                            DataGridView1.Columns.Add("count1", "Nr.")
                            DataGridView1.DataSource = setidatave1.Tables(0)
                            DataGridView1.Columns("ID").Visible = False
                            Dim nr11 As Integer = 0
                            nr11 = DataGridView1.Rows.Count - 1
                            For i = 0 To nr11
                                DataGridView1.Rows(i).Cells(0).Value = i + 1
                            Next
                            Lidhje1.Close()
                            Dim nr As Integer = 0
                            nr = DataGridView1.Columns.Count - 1
                            If nr > 15 Then
                                DataGridView1.Columns.RemoveAt(nr)
                            End If
                            Dim Lidhje1xx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
                            Dim adaptor1xx As OleDbDataAdapter
                            Dim setidatave1xx = New DataSet
                            With Farmacia.DataGridView2
                                .RowsDefaultCellStyle.BackColor = Color.Bisque
                                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                            End With
                            rreshtiaktual2 = 0
                            Lidhje1xx.Open()
                            adaptor1xx = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", Lidhje1)
                            adaptor1xx.Fill(setidatave1xx, "tedhena")
                            Farmacia.DataGridView2.DataSource = setidatave1xx.Tables(0)

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
                            Label21.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                            Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                            Label30.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                            MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                            Dim nr1 As Integer = 0
                            Farmacia.DataGridView2.Columns("ID").Visible = False
                            nr1 = Farmacia.DataGridView2.Rows.Count - 1
                            For i = 0 To nr1
                                Farmacia.DataGridView2.Rows(i).Cells(0).Value = i + 1
                            Next
                        Catch ex As Exception
                            MsgBox("Nuk u shtua me sukses", MsgBoxStyle.Information)
                            Lidhje1x.Close()
                        End Try
                    End If
                End If
            Else
                Farmacia.DataGridView2.Columns("ID").Visible = True
                Dim nrekzistues As New List(Of Integer)
                For Each kolone As DataGridViewRow In Farmacia.DataGridView2.Rows
                    nrekzistues.Add(CInt(kolone.Cells(1).Value))
                Next
                Dim existingNumbers As New List(Of Integer)
                For Each r As DataGridViewRow In Farmacia.DataGridView2.Rows
                    existingNumbers.Add(CInt(r.Cells(1).Value))
                Next
                Dim missingNumbers = Enumerable.Range(existingNumbers.First, existingNumbers.Max() - existingNumbers.First + 1).Except(existingNumbers)
                Dim max = nrekzistues.Max() + 1
                If missingNumbers.Count = 0 Then
                    Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox10.Text
                    Dim Str1 As String
                    Dim regDate As DateTime = Date.Now
                    Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                    Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                    Dim strDate1 As String = regDate.ToString("HH:mm tt")
                    Try
                        Str1 = "insert into Shitjet values("
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
                        Str1 += """" & ComboBox1.Text.Trim() & """"
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
                        Str1 += """" & vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture) & """"
                        Str1 += ","
                        Str1 += """" & TextBox11.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & NumericUpDown2.Value & """"
                        Str1 += ")"
                        Lidhje1.Open()
                        query1 = New OleDbCommand(Str1, Lidhje1)
                        query1.ExecuteNonQuery()
                        setidatave1.Clear()
                        adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Kodi_Shitjes LIKE '" & TextBox8.Text & "%')", Lidhje1)
                        adaptor1.Fill(setidatave1, "tedhena")
                        DataGridView1.Columns.Add("count1", "Nr.")
                        DataGridView1.DataSource = setidatave1.Tables(0)
                        DataGridView1.Columns("ID").Visible = False
                        Dim nr11 As Integer = 0
                        nr11 = DataGridView1.Rows.Count - 1
                        For i = 0 To nr11
                            DataGridView1.Rows(i).Cells(0).Value = i + 1
                        Next
                        Lidhje1.Close()
                        Dim nr As Integer = 0
                        nr = DataGridView1.Columns.Count - 1
                        If nr > 14 Then
                            DataGridView1.Columns.RemoveAt(nr)
                        End If
                        Dim Lidhje1xx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
                        Dim adaptor1xx As OleDbDataAdapter
                        Dim setidatave1xx = New DataSet
                        With Farmacia.DataGridView2
                            .RowsDefaultCellStyle.BackColor = Color.Bisque
                            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                        End With
                        rreshtiaktual2 = 0
                        Lidhje1xx.Open()
                        adaptor1xx = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", Lidhje1)
                        adaptor1xx.Fill(setidatave1xx, "tedhena")
                        Farmacia.DataGridView2.DataSource = setidatave1xx.Tables(0)

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
                        Label21.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                        Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                        Label30.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                        MsgBox("U shtua me sukses")
                        Dim nr1 As Integer = 0
                        Farmacia.DataGridView2.Columns("ID").Visible = False
                        nr1 = Farmacia.DataGridView2.Rows.Count - 1
                        For i = 0 To nr1
                            Farmacia.DataGridView2.Rows(i).Cells(0).Value = i + 1
                        Next
                    Catch ex As Exception
                        MsgBox("Nuk u shtua me sukses", MsgBoxStyle.Information)
                        Lidhje1x.Close()
                    End Try
                Else
                    Farmacia.DataGridView2.Columns("ID").Visible = True
                    Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox10.Text
                    Dim Str1 As String
                    Dim regDate As DateTime = Date.Now
                    Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                    Dim regDate1 As DateTime = DateTime.Now.ToShortTimeString()
                    Dim strDate1 As String = regDate.ToString("HH:mm tt")
                    Try
                        Str1 = "insert into Shitjet values("
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
                        Str1 += """" & ComboBox1.Text.Trim() & """"
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
                        Str1 += """" & vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture) & """"
                        Str1 += ","
                        Str1 += """" & TextBox11.Text.Trim() & """"
                        Str1 += ","
                        Str1 += """" & NumericUpDown2.Value & """"
                        Str1 += ")"
                        Lidhje1.Open()
                        query1 = New OleDbCommand(Str1, Lidhje1)
                        query1.ExecuteNonQuery()
                        setidatave1.Clear()
                        adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Kodi_Shitjes LIKE '" & TextBox8.Text & "%')", Lidhje1)
                        adaptor1.Fill(setidatave1, "tedhena")
                        DataGridView1.Columns.Add("count1", "Nr.")
                        DataGridView1.DataSource = setidatave1.Tables(0)
                        DataGridView1.Columns("ID").Visible = False
                        Dim nr11 As Integer = 0
                        nr11 = DataGridView1.Rows.Count - 1
                        For i = 0 To nr11
                            DataGridView1.Rows(i).Cells(0).Value = i + 1
                        Next
                        Lidhje1.Close()
                        Dim nr As Integer = 0
                        nr = DataGridView1.Columns.Count - 1
                        If nr > 14 Then
                            DataGridView1.Columns.RemoveAt(nr)
                        End If
                        Dim Lidhje1xx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
                        Dim adaptor1xx As OleDbDataAdapter
                        Dim setidatave1xx = New DataSet
                        With Farmacia.DataGridView2
                            .RowsDefaultCellStyle.BackColor = Color.Bisque
                            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                        End With
                        rreshtiaktual2 = 0
                        Lidhje1xx.Open()
                        adaptor1xx = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", Lidhje1)
                        adaptor1xx.Fill(setidatave1xx, "tedhena")
                        Farmacia.DataGridView2.DataSource = setidatave1xx.Tables(0)

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
                        Label21.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
                        Label29.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
                        Label30.Text = total4.ToString("#,#", CultureInfo.InvariantCulture)
                        MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                        Dim nr1 As Integer = 0
                        Farmacia.DataGridView2.Columns("ID").Visible = False
                        nr1 = Farmacia.DataGridView2.Rows.Count - 1
                        For i = 0 To nr1
                            Farmacia.DataGridView2.Rows(i).Cells(0).Value = i + 1
                        Next
                    Catch ex As Exception
                        MsgBox("Nuk u shtua me sukses", MsgBoxStyle.Information)
                        Lidhje1x.Close()
                    End Try
                End If
            End If
        End If
        If Not Farmacia.DataGridView2.Rows.Count > 0 Then
            Farmacia.Button4.Enabled = False
            Farmacia.Button6.Enabled = False
        Else
            Farmacia.Button4.Enabled = True
            Farmacia.Button6.Enabled = True
        End If
    End Sub

    Public Sub Nrfatures(ByVal connectionString As String,
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
                Label39.Text = i.ToString.Max().ToString + 1
            Next
        End Using
    End Sub
    Public Sub CreateReader(ByVal connectionString As String,
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
    Public Sub CreateReader3(ByVal connectionString As String,
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
    Private Sub Button2_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.MouseHover
        With Button2
            .ForeColor = Color.White
        End With
    End Sub
    Private Sub Button2_MouseDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.MouseDown
        With Button2
            .ForeColor = Color.White
        End With
    End Sub
    Private Sub Button2_Mouseleave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.MouseLeave
        With Button2
            .ForeColor = Color.Black
        End With
    End Sub
    Private Sub Button5_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.MouseHover
        With Button5
            .ForeColor = Color.White
        End With
    End Sub
    Private Sub Button5_MouseDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.MouseDown
        With Button5
            .ForeColor = Color.White
        End With
    End Sub
    Private Sub Button5_Mouseleave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.MouseLeave
        With Button5
            .ForeColor = Color.Black
        End With
    End Sub
    Private Sub Shit_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        Farmacia.Show()
        My.Settings.tvsh = NumericUpDown3.Value
    End Sub
    Private Sub Shit_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Konfigurime.TextBox1.Text = My.Settings.backup
        Konfigurime.TextBox2.Text = My.Settings.logo
        Konfigurime.TextBox3.Text = My.Settings.ruajraportet
        Konfigurime.TextBox4.Text = My.Settings.faturatofert
        NumericUpDown3.Value = My.Settings.tvsh
        If DataGridView1.Rows.Count = 0 Then
            Button3.Enabled = False
        Else
            Button3.Enabled = True
        End If
        NumericUpDown2.ReadOnly = True
        NumericUpDown2.Region = New Region(New System.Drawing.Rectangle(0, 0, NumericUpDown2.Width - NumericUpDown2.Controls(0).Width + 1, NumericUpDown2.Height))
        Dim regDate2 As DateTime = Date.Now
        Dim strDate2 As String = regDate2.ToString("dd/MM/yyyy")
        Dim connrfat As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim queynrfat As String = "SELECT DISTINCT Kodi_Shitjes FROM Shitjet WHERE(Data Like '" & strDate2 & "%')"
        Nrfatures(connrfat, queynrfat)
        'load form klasat
        With Klasat_e_produkteve.DataGridView1
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktualdg71ofert121klasat = 0
        lidhjedg71ofert121klasat.Open()
        adaptoridg71ofert121klasat = New OleDbDataAdapter("SELECT * FROM Klasat ORDER BY ID", lidhjedg71ofert121klasat)
        adaptoridg71ofert121klasat.Fill(setitedhenavedg71ofert121klasat, "tedhena")
        Klasat_e_produkteve.DataGridView1.DataSource = setitedhenavedg71ofert121klasat.Tables(0)
        If setitedhenavedg71ofert121klasat.Tables(0).Rows.Count > 0 Then
            Klasat_e_produkteve.Merrtedhenat_Klasat(rreshtiaktualdg71ofert121klasat)
        Else
        End If
        lidhjedg71ofert121klasat.Close()
        Try
            If My.Settings.YourItems IsNot Nothing Then
                For Each S As String In My.Settings.YourItems
                    Klasat_e_produkteve.ListBox1.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems1 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems1
                    Klasat_e_produkteve.ListBox2.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems2 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems2
                    Klasat_e_produkteve.ListBox3.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems3 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems3
                    Klasat_e_produkteve.ListBox4.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems4 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems4
                    Klasat_e_produkteve.ListBox5.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems5 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems5
                    Klasat_e_produkteve.ListBox6.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems6 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems6
                    Klasat_e_produkteve.ListBox7.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems7 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems7
                    Klasat_e_produkteve.ListBox8.Items.Add(S)
                Next
            End If
        Catch ex As Exception

        End Try
        Klasat_e_produkteve.ListBox1.AllowDrop = True
        Klasat_e_produkteve.ListBox2.AllowDrop = True
        Klasat_e_produkteve.ListBox3.AllowDrop = True
        Klasat_e_produkteve.ListBox4.AllowDrop = True
        Klasat_e_produkteve.ListBox5.AllowDrop = True
        Klasat_e_produkteve.ListBox6.AllowDrop = True
        Klasat_e_produkteve.ListBox7.AllowDrop = True
        Klasat_e_produkteve.ListBox8.AllowDrop = True
        Dim conyklasa As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim queyklasa As String = "SELECT * FROM Produktet ORDER BY ID"
        Klasat_e_produkteve.CreateReader_load_produktet_klasa(conyklasa, queyklasa)
        totale = 0
        totale1 = 0
        totalefund = 0
        Dim regDate As DateTime = Date.Now
        Dim strDate As String = regDate.ToString("dd/MM/yyyy")
        Label18.Text = strDate
        With DataGridView1
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        Timer1.Start()
        TextBox8.Text = GenerateRandomString(8)
        Button2.Text = "Ruaj Faturen(" & TextBox8.Text & ")"
        Button5.Text = "Printo Faturen(" & TextBox8.Text & ")"
        Dim con As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim que As String = "SELECT * FROM Bleresit ORDER BY ID"
        CreateReader(con, que)

        Dim con2 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim que2 As String = "SELECT * FROM Shitesi ORDER BY ID"
        CreateReader3(con2, que2)
        ComboBox2.SelectedIndex = 0
        Dim cony As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim quey As String = "SELECT * FROM Produktet ORDER BY ID"
        CreateReader_load_produktet(cony, quey)
        With Farmacia.DataGridView2
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual = 0
        Lidhje1x.Open()
        adaptor1x = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", Lidhje1x)
        adaptor1x.Fill(setidatave1x, "tedhena")
        Farmacia.DataGridView2.DataSource = setidatave1x.Tables(0)
        Lidhje1x.Close()
        NumericUpDown1.Value = 0
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
    Public Sub CreateReader1(ByVal connectionString As String,
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
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim con1 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim que1 As String = "SELECT * FROM Bleresit WHERE(Emri LIKE '" & ComboBox1.Text & "%')"
        CreateReader1(con1, que1)
        CheckBox1.Checked = True
        If Klasat_e_produkteve.ListBox2.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_A"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(2).Value.ToString())
                End If
            Next
        ElseIf Klasat_e_produkteve.ListBox3.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_B"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(3).Value.ToString())
                End If
            Next
        ElseIf Klasat_e_produkteve.ListBox4.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_C"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(4).Value.ToString())
                End If
            Next
        ElseIf Klasat_e_produkteve.ListBox5.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_D"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(5).Value.ToString())
                End If
            Next
        ElseIf Klasat_e_produkteve.ListBox6.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_E"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(6).Value.ToString())
                End If
            Next
        ElseIf Klasat_e_produkteve.ListBox7.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_F"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(7).Value.ToString())
                End If
            Next
        ElseIf Klasat_e_produkteve.ListBox8.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_G"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(8).Value.ToString())
                End If
            Next
        Else
            NumericUpDown2.Value = 0
        End If
    End Sub
    Public Sub CreateReader4(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                TextBox4.Text = (reader(2).ToString())
                TextBox5.Text = (reader(3).ToString())
                TextBox6.Text = (reader(4).ToString())
            End While
            reader.Close()
        End Using
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim con1 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim que1 As String = "SELECT * FROM Shitesi WHERE(Emri LIKE '" & ComboBox2.Text & "%')"
        CreateReader4(con1, que1)
    End Sub
    Public Sub CreateReader_produktet(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                ComboBox4.Text = (reader(2).ToString())
                TextBox9.Text = (reader(4).ToString())
                '  TextBox7.Text = (reader(8).ToString())
            End While
            reader.Close()
        End Using
    End Sub
    Public Sub CreateReader_load_produktet(ByVal connectionString As String,
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
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        NumericUpDown1.Value = 0
        Dim conx As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim quex As String = "SELECT * FROM Produktet WHERE(Emri_produktit LIKE '" & ComboBox3.Text & "%')"
        CreateReader_produktet(conx, quex)
        NumericUpDown1.Value = 1
        Dim conxzz As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim quexzz As String = "SELECT * FROM Shitjet WHERE(Produkti Like '" &
                                           ComboBox3.Text & "%') ORDER BY Kodi_Shitjes"
        CreateReader_sasieshitur(conxzz, quexzz)
        Application.DoEvents()
        Dim conxzz1 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim quexzz1 As String = "SELECT * FROM dhenat WHERE(Produkti Like '" &
                                           ComboBox3.Text & "%') ORDER BY Kodi_Blerjes"
        sasiefutur(conxzz1, quexzz1)




        Dim conxzz12 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim quexzz12 As String = "SELECT * FROM Produktet WHERE(Emri_produktit Like '" &
                                           ComboBox3.Text & "%') "
        sasialimit(conxzz12, quexzz12)
        NumericUpDown1.Value = 1
        Application.DoEvents()
        totalefund = 0
        totalefund = totale1 - totale
        If totalefund = 0 Then
            Label34.ForeColor = Color.Red
            Label34.Text = totalefund
            NumericUpDown1.Value = totalefund
        Else
            Label34.ForeColor = Color.ForestGreen
            Label34.Text = totalefund
            NumericUpDown1.Value = 1
        End If
        totale = 0
        totale1 = 0
        CheckBox1.Checked = True
        If Klasat_e_produkteve.ListBox2.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_A"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(2).Value.ToString())
                End If
            Next
        ElseIf Klasat_e_produkteve.ListBox3.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_B"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(3).Value.ToString())
                End If
            Next
        ElseIf Klasat_e_produkteve.ListBox4.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_C"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(4).Value.ToString())
                End If
            Next
        ElseIf Klasat_e_produkteve.ListBox5.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_D"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(5).Value.ToString())
                End If
            Next
        ElseIf Klasat_e_produkteve.ListBox6.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_E"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(6).Value.ToString())
                End If
            Next
        ElseIf Klasat_e_produkteve.ListBox7.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_F"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(7).Value.ToString())
                End If
            Next
        ElseIf Klasat_e_produkteve.ListBox8.Items.Contains(ComboBox3.Text) Then
            ' "Klasa_G"
            For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                    NumericUpDown2.Value = (row.Cells(8).Value.ToString())
                End If
            Next
        Else
            NumericUpDown2.Value = 0
        End If
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
                Dim totalizbriturcmimit As Decimal = TextBox9.Text * (1 - NumericUpDown2.Value / 100)
                If NumericUpDown2.Value > 0 Then
                    Label36.Text = totalizbriturcmimit.ToString("#,#", CultureInfo.InvariantCulture)
                    Label36.ForeColor = Color.Red
                Else
                    Label36.Text = 0
                    Label36.ForeColor = DefaultForeColor
                End If


                Dim vlera_shitjes = total1
                TextBox10.Text = vlera_shitjes.ToString("#,#", CultureInfo.InvariantCulture)
                Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * vlera_shitjes
                Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + vlera_shitjes
                TextBox7.Text = vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture)
                TextBox11.Text = vlera_shitjes_metvsh1.ToString("#,#", CultureInfo.InvariantCulture)
            Else
                If TextBox9.Text = "" Then

                Else

                    Dim vlera_shitjes = NumericUpDown1.Value * TextBox9.Text
                    TextBox10.Text = vlera_shitjes.ToString("#,#", CultureInfo.InvariantCulture)
                    Label36.Text = 0
                    Label36.ForeColor = DefaultForeColor
                    Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox10.Text
                    Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + TextBox10.Text
                    TextBox7.Text = vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture)
                    TextBox11.Text = vlera_shitjes_metvsh1.ToString("#,#", CultureInfo.InvariantCulture)
                End If
            End If
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
    Public Sub sasialimit(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                Label41.Text = reader(6).ToString()
            End While
            reader.Close()
        End Using
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Not DataGridView1.Rows.Count > 0 Then
            MsgBox("Fatura eshte bosh!Me pare gjeneroni nje fature!", MsgBoxStyle.Information)
        Else
            TextBox8.Text = GenerateRandomString(8)
            ComboBox1.SelectedIndex = 0
            ComboBox2.SelectedIndex = 0
            ComboBox3.SelectedIndex = 0
            'NumericUpDown1.Value = 1
            NumericUpDown1.Value = 0
            Label21.Text = 0
            Label29.Text = 0
            Label30.Text = 0
            Label36.Text = 0
            TextBox10.Text = 0
            TextBox11.Text = 0
            TextBox7.Text = 0
            Button2.Text = "Ruaj Faturen(" & TextBox8.Text & ")"
            Button5.Text = "Printo Faturen(" & TextBox8.Text & ")"
            DataGridView1.DataSource = Nothing
            ListBox1.Items.Clear()
            Dim regDate2 As DateTime = Date.Now
            Dim strDate2 As String = regDate2.ToString("dd/MM/yyyy")
            Dim connrfat As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
            Dim queynrfat As String = "SELECT DISTINCT Kodi_Shitjes FROM Shitjet WHERE(Data Like '" & strDate2 & "%')"
            Nrfatures(connrfat, queynrfat)
            MsgBox("Fatura u ruajt me sukses!", MsgBoxStyle.Information)
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label20.Text = TimeOfDay.ToString("h:mm:ss tt")
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox1.ForeColor = Color.ForestGreen
            NumericUpDown2.Enabled = True
            NumericUpDown2.BackColor = Color.LightGreen
            If Label36.Text = 0 Then
            Else
                Label36.ForeColor = Color.Red
            End If
            If Klasat_e_produkteve.ListBox2.Items.Contains(ComboBox3.Text) Then
                ' "Klasa_A"
                For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                    If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                        NumericUpDown2.Value = (row.Cells(2).Value.ToString())
                    End If
                Next
            ElseIf Klasat_e_produkteve.ListBox3.Items.Contains(ComboBox3.Text) Then
                ' "Klasa_B"
                For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                    If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                        NumericUpDown2.Value = (row.Cells(3).Value.ToString())
                    End If
                Next
            ElseIf Klasat_e_produkteve.ListBox4.Items.Contains(ComboBox3.Text) Then
                ' "Klasa_C"
                For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                    If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                        NumericUpDown2.Value = (row.Cells(4).Value.ToString())
                    End If
                Next
            ElseIf Klasat_e_produkteve.ListBox5.Items.Contains(ComboBox3.Text) Then
                ' "Klasa_D"
                For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                    If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                        NumericUpDown2.Value = (row.Cells(5).Value.ToString())
                    End If
                Next
            ElseIf Klasat_e_produkteve.ListBox6.Items.Contains(ComboBox3.Text) Then
                ' "Klasa_E"
                For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                    If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                        NumericUpDown2.Value = (row.Cells(6).Value.ToString())
                    End If
                Next
            ElseIf Klasat_e_produkteve.ListBox7.Items.Contains(ComboBox3.Text) Then
                ' "Klasa_F"
                For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                    If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                        NumericUpDown2.Value = (row.Cells(7).Value.ToString())
                    End If
                Next
            ElseIf Klasat_e_produkteve.ListBox8.Items.Contains(ComboBox3.Text) Then
                ' "Klasa_G"
                For Each row As DataGridViewRow In Klasat_e_produkteve.DataGridView1.Rows
                    If row.Cells(1).Value.ToString().Contains(ComboBox1.Text) Then
                        NumericUpDown2.Value = (row.Cells(8).Value.ToString())
                    End If
                Next
            Else
                NumericUpDown2.Value = 0
            End If
        Else
            NumericUpDown2.Value = 0
            NumericUpDown2.Enabled = False
            CheckBox1.ForeColor = Color.Red
            NumericUpDown2.BackColor = DefaultBackColor
            Label36.Text = 0
            Label36.ForeColor = DefaultForeColor
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
                    Label36.Text = totalizbriturcmimit.ToString("#,#", CultureInfo.InvariantCulture)
                    Label36.ForeColor = Color.Red
                Else
                    Label36.Text = 0
                    Label36.ForeColor = DefaultForeColor
                End If
                Dim vlera_shitjes = total1
                TextBox10.Text = vlera_shitjes.ToString("#,#", CultureInfo.InvariantCulture)
                Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * vlera_shitjes
                Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + vlera_shitjes
                TextBox7.Text = vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture)
                TextBox11.Text = vlera_shitjes_metvsh1.ToString("#,#", CultureInfo.InvariantCulture)
            End If
        Else
            If TextBox9.Text = "" Then
            Else
                Label36.Text = 0
                Label36.ForeColor = DefaultForeColor
                Dim vlera_shitjes = NumericUpDown1.Value * TextBox9.Text
                TextBox10.Text = vlera_shitjes.ToString("#,#", CultureInfo.InvariantCulture)
                Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * vlera_shitjes
                Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + vlera_shitjes
                TextBox7.Text = vlera_shitjes_metvsh.ToString("#,#", CultureInfo.InvariantCulture)
                TextBox11.Text = vlera_shitjes_metvsh1.ToString("#,#", CultureInfo.InvariantCulture)
            End If
        End If
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If (e.RowIndex >= 0) Then
            Dim row As DataGridViewRow = DataGridView1.Rows.Item(e.RowIndex)
            Label33.Text = row.Cells.Item("ID").Value.ToString
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
                Str = "delete from Shitjet where ID="
                Str += DataGridView1.CurrentRow.Cells(1).Value.ToString
                lidhje2.Open()
                query2 = New OleDbCommand(Str, lidhje2)
                query2.ExecuteNonQuery()
                setidatave1.clear()
                adaptori2 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Kodi_Shitjes LIKE '" & TextBox8.Text & "%')", lidhje2)
                adaptori2.Fill(setidatave1, "tedhena")
                DataGridView1.DataSource = setidatave1.Tables(0)
                Dim nrd As Integer = 0
                nrd = DataGridView1.Rows.Count - 1
                For i = 0 To nrd
                    DataGridView1.Rows(i).Cells(0).Value = i + 1
                Next
                lidhje2.Close()
                Dim nr As Integer = 0
                nr = DataGridView1.Columns.Count - 1
                If nr > 15 Then
                    DataGridView1.Columns.RemoveAt(nr)
                End If
                Dim Lidhje1xx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
                Dim adaptor1xx As OleDbDataAdapter
                Dim setidatave1xx = New DataSet
                With Farmacia.DataGridView2
                    .RowsDefaultCellStyle.BackColor = Color.Bisque
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                End With
                rreshtiaktual2 = 0
                Lidhje1xx.Open()
                adaptor1xx = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", Lidhje1)
                adaptor1xx.Fill(setidatave1xx, "tedhena")
                Farmacia.DataGridView2.DataSource = setidatave1xx.Tables(0)
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
                Label21.Text = total2
                Label29.Text = total3
                Label30.Text = total4
                MsgBox("U fshi me sukses", MsgBoxStyle.Information)
                Dim nrd1 As Integer = 0
                nrd1 = Farmacia.DataGridView2.Rows.Count - 1
                For i = 0 To nrd1
                    Farmacia.DataGridView2.Rows(i).Cells(0).Value = i + 1
                Next
            Catch ex As Exception
                MessageBox.Show("Nuk u fshi")
                MsgBox(ex.Message & " -  " & ex.Source)
                lidhje2.Close()
            End Try
        End If
        If Not Farmacia.DataGridView2.Rows.Count > 0 Then
            Farmacia.Button4.Enabled = False
            Farmacia.Button6.Enabled = False
        Else
            Farmacia.Button4.Enabled = True
            Farmacia.Button6.Enabled = True
        End If
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            NumericUpDown2.ReadOnly = False
            NumericUpDown2.Region = New Region(New System.Drawing.Rectangle(0, 0, NumericUpDown2.Width, NumericUpDown2.Height))
        Else
            NumericUpDown2.ReadOnly = True
            NumericUpDown2.Region = New Region(New System.Drawing.Rectangle(0, 0, NumericUpDown2.Width - NumericUpDown2.Controls(0).Width + 1, NumericUpDown2.Height))
        End If
    End Sub
    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        If CheckBox1.Checked = True Then
            If TextBox9.Text = "" Then
            Else
                Dim zbritje = NumericUpDown1.Value * TextBox9.Text
                Dim total1 As Decimal = zbritje * (1 - NumericUpDown2.Value / 100)
                Dim totalizbriturcmimit As Decimal = TextBox9.Text * (1 - NumericUpDown2.Value / 100)
                Label36.Text = totalizbriturcmimit
                Label36.ForeColor = Color.Red
                Dim vlera_shitjes = total1
                TextBox10.Text = vlera_shitjes
                Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox10.Text
                Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + TextBox10.Text
                TextBox7.Text = vlera_shitjes_metvsh
                TextBox11.Text = vlera_shitjes_metvsh1
            End If
        Else
            If TextBox9.Text = "" Then
            Else
                Label36.Text = 0
                Label36.ForeColor = DefaultForeColor
                Dim vlera_shitjes = NumericUpDown1.Value * TextBox9.Text
                TextBox10.Text = vlera_shitjes
                Dim vlera_shitjes_metvsh = (NumericUpDown3.Value / 100) * TextBox10.Text
                Dim vlera_shitjes_metvsh1 = vlera_shitjes_metvsh + TextBox10.Text
                TextBox7.Text = vlera_shitjes_metvsh
                TextBox11.Text = vlera_shitjes_metvsh1
            End If
        End If
    End Sub
    Public lidhjeprint As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptori2print As OleDbDataAdapter
    Public lexuesiprint As OleDbDataReader
    Public query2print As OleDbCommand
    Public setitedhenave2print = New DataSet
    Public rreshtiaktual2print As Integer
    Dim data, shitesi, bleresi, kodi_shit, shita, shitc, shitn, bleresia, bleresic, bleresin As String



    Dim vlerapatvsh, vleraetvsh, vlerametvsh As String
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim regDate As DateTime = Date.Now
        Dim strDate As String = regDate.ToString("dd/MM/yyyy")
        If Not DataGridView1.Rows.Count > 0 Then
            MsgBox("Fatura eshte bosh!Me pare gjeneroni nje fature!", MsgBoxStyle.Information)
        Else
            If My.Settings.logo = "" Or My.Settings.ruajraportet = "" Then
                MsgBox("Ju nuk keni zgjedhur nje logo ose vendodhjen se ku do te ruhen faturat.Shko tek Konfigurimet!", MsgBoxStyle.Information)
            Else
                Konfigurime.TextBox1.Text = My.Settings.backup
                Konfigurime.TextBox2.Text = My.Settings.logo
                Konfigurime.TextBox3.Text = My.Settings.ruajraportet
                Konfigurime.TextBox4.Text = My.Settings.faturatofert
                lidhjeprint.Open()
                setitedhenave2print.Clear
                adaptori2print = New OleDbDataAdapter("SELECT Produkti,Njesia,Sasia,Cmimi,Vlera_Pa_TVSH,TVSH,Vlera_Me_TVSH,Zbritje_ne_perq FROM Shitjet WHERE(Kodi_Shitjes Like '" & TextBox8.Text & "%')", lidhjeprint)
                adaptori2print.Fill(setitedhenave2print, "tedhena")
                DataGridView8.Columns.Clear()
                DataGridView8.Columns.Add("count1", "Nr.")
                DataGridView8.DataSource = setitedhenave2print.Tables(0)
                'DataGridView8.Columns("ID").Visible = False
                lidhjeprint.Close()
                Dim nr8 As Integer = 0
                nr8 = DataGridView8.Rows.Count - 1
                For i = 0 To nr8
                    DataGridView8.Rows(i).Cells(0).Value = i + 1
                Next
                DataGridView8.Refresh()
                vlerapatvsh = (From row As DataGridViewRow In DataGridView8.Rows
                               Where row.Cells(5).FormattedValue.ToString() <> String.Empty
                               Select Convert.ToInt32(row.Cells(5).FormattedValue)).Sum().ToString()
                vleraetvsh = (From row As DataGridViewRow In DataGridView8.Rows
                              Where row.Cells(6).FormattedValue.ToString() <> String.Empty
                              Select Convert.ToInt32(row.Cells(6).FormattedValue)).Sum().ToString()
                vlerametvsh = (From row As DataGridViewRow In DataGridView8.Rows
                               Where row.Cells(7).FormattedValue.ToString() <> String.Empty
                               Select Convert.ToInt32(row.Cells(7).FormattedValue)).Sum().ToString()
                DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(5).Value = vlerapatvsh
                DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(6).Value = vleraetvsh
                DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(7).Value = vlerametvsh
                DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(4).Value = "TOTALI"
                DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(0).Value = ""
                Dim con_bleresit As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
                Dim query_bleresit As String = "SELECT TOP 1 Data,Shitesi,Bleresi,Kodi_Shitjes FROM Shitjet WHERE(Kodi_Shitjes Like '" & TextBox8.Text & "%')"
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
                    Dim pdfTable As New PdfPTable(DataGridView8.ColumnCount)
                    pdfTable.DefaultCell.Padding = 3
                    pdfTable.WidthPercentage = 100
                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT
                    'Adding Header row
                    For Each column As DataGridViewColumn In DataGridView8.Columns
                        Dim cell As New PdfPCell(New Phrase(column.HeaderText.Replace("_", " ").Replace("ne perq", "%")))
                        cell.BorderWidthTop = 1
                        cell.BorderWidthBottom = 1
                        cell.BackgroundColor = New iTextSharp.text.BaseColor(250, 235, 215)
                        pdfTable.AddCell(cell)
                    Next
                    'Adding DataRow
                    Dim cellvalue As String = ""
                    Dim i As Integer = 0
                    For Each row As DataGridViewRow In DataGridView8.Rows
                        For Each cell As DataGridViewCell In row.Cells
                            cellvalue = cell.FormattedValue
                            pdfTable.AddCell(Convert.ToString(cellvalue))
                        Next
                    Next
                    'Exporting to PDF
                    Dim folderPath As String = Konfigurime.TextBox3.Text
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
                    Dim oldFile As String = Konfigurime.TextBox3.Text & "\DataGridViewExport.pdf"
                    Dim newFile As String = Konfigurime.TextBox3.Text & "\DataGridViewExport1.pdf"
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
                    cb.ShowTextAligned(1, emri_shitesit, 100, 597, 0)
                    Dim ad_shitesit As String = shita
                    cb.ShowTextAligned(1, ad_shitesit, 111, 578, 0)
                    Dim cel_shitesit As String = shitc
                    cb.ShowTextAligned(1, cel_shitesit, 97, 559, 0)
                    Dim nipt_shitesit As String = shitn
                    cb.ShowTextAligned(1, nipt_shitesit, 96, 541, 0)
                    Dim emri_bleresit As String = bleresi
                    cb.ShowTextAligned(1, emri_bleresit, 360, 597, 0)
                    Dim ad_bleres As String = bleresia
                    cb.ShowTextAligned(1, ad_bleres, 372, 577, 0)
                    Dim cel_ble As String = bleresic
                    cb.ShowTextAligned(1, cel_ble, 355, 558, 0)
                    Dim nipt_ble As String = bleresin
                    cb.ShowTextAligned(1, nipt_ble, 356, 540, 0)
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
                    My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox3.Text & "\DataGridViewExport.pdf")
                End Using
                Using inputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\DataGridViewExport1.pdf", FileMode.Open, FileAccess.Read, FileShare.Read)
                    My.Resources.xx.Save("C:\ProgramData\xx.jpg", Drawing.Imaging.ImageFormat.Jpeg)
                    Using inputImageStream As Stream = New FileStream("C:\ProgramData\xx.jpg", FileMode.Open, FileAccess.Read, FileShare.Read)
                        Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\DataGridViewExport3.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
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
                My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox3.Text & "\DataGridViewExport1.pdf")
                Using inputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\DataGridViewExport3.pdf", FileMode.Open, FileAccess.Read, FileShare.Read)
                    Using inputImageStream As Stream = New FileStream(Konfigurime.TextBox2.Text, FileMode.Open, FileAccess.Read, FileShare.Read)
                        Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Shitje.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
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
                My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox3.Text & "\DataGridViewExport3.pdf")
                My.Computer.FileSystem.DeleteFile("C:\ProgramData\xx.jpg")
                System.Diagnostics.Process.Start(Konfigurime.TextBox3.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Shitje.pdf")
            End If
        End If
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim regDate As DateTime = Date.Now
        Dim strDate As String = regDate.ToString("dd/MM/yyyy")
        If Not DataGridView1.Rows.Count > 0 Then
            MsgBox("Fatura eshte bosh!Me pare gjeneroni nje fature!", MsgBoxStyle.Information)
        Else
            If My.Settings.logo = "" Or My.Settings.ruajraportet = "" Then
                MsgBox("Ju nuk keni zgjedhur nje logo ose vendodhjen se ku do te ruhen faturat.Shko tek Konfigurimet!", MsgBoxStyle.Information)
            Else
                Try
                    Konfigurime.TextBox1.Text = My.Settings.backup
                    Konfigurime.TextBox2.Text = My.Settings.logo
                    Konfigurime.TextBox3.Text = My.Settings.ruajraportet
                    Konfigurime.TextBox4.Text = My.Settings.faturatofert
                    lidhjeprint.Open()
                    setitedhenave2print.Clear
                    adaptori2print = New OleDbDataAdapter("SELECT Produkti,Njesia,Sasia,Cmimi,Vlera_Pa_TVSH,TVSH,Vlera_Me_TVSH,Zbritje_ne_perq FROM Shitjet WHERE(Kodi_Shitjes Like '" & TextBox8.Text & "%')", lidhjeprint)
                    adaptori2print.Fill(setitedhenave2print, "tedhena")
                    DataGridView8.Columns.Clear()
                    DataGridView8.Columns.Add("count1", "Nr.")
                    DataGridView8.DataSource = setitedhenave2print.Tables(0)
                    'DataGridView8.Columns("ID").Visible = False
                    lidhjeprint.Close()
                    Dim nr8 As Integer = 0
                    nr8 = DataGridView8.Rows.Count - 1
                    For i = 0 To nr8
                        DataGridView8.Rows(i).Cells(0).Value = i + 1
                    Next
                    DataGridView8.Refresh()
                    vlerapatvsh = (From row As DataGridViewRow In DataGridView8.Rows
                                   Where row.Cells(5).FormattedValue.ToString() <> String.Empty
                                   Select Convert.ToInt32(row.Cells(5).FormattedValue)).Sum().ToString()
                    vleraetvsh = (From row As DataGridViewRow In DataGridView8.Rows
                                  Where row.Cells(6).FormattedValue.ToString() <> String.Empty
                                  Select Convert.ToInt32(row.Cells(6).FormattedValue)).Sum().ToString()
                    vlerametvsh = (From row As DataGridViewRow In DataGridView8.Rows
                                   Where row.Cells(7).FormattedValue.ToString() <> String.Empty
                                   Select Convert.ToInt32(row.Cells(7).FormattedValue)).Sum().ToString()
                    DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(5).Value = vlerapatvsh
                    DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(6).Value = vleraetvsh
                    DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(7).Value = vlerametvsh
                    DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(4).Value = "TOTALI"
                    DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(0).Value = ""
                    Dim con_bleresit As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
                    Dim query_bleresit As String = "SELECT TOP 1 Data,Shitesi,Bleresi,Kodi_Shitjes FROM Shitjet WHERE(Kodi_Shitjes Like '" & TextBox8.Text & "%')"
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
                        Dim pdfTable As New PdfPTable(DataGridView8.ColumnCount)
                        pdfTable.DefaultCell.Padding = 3
                        pdfTable.WidthPercentage = 100
                        pdfTable.HorizontalAlignment = Element.ALIGN_LEFT
                        'Adding Header row
                        For Each column As DataGridViewColumn In DataGridView8.Columns
                            Dim cell As New PdfPCell(New Phrase(column.HeaderText.Replace("_", " ").Replace("ne perq", "%")))
                            cell.BorderWidthTop = 1
                            cell.BorderWidthBottom = 1
                            cell.BackgroundColor = New iTextSharp.text.BaseColor(250, 235, 215)
                            pdfTable.AddCell(cell)
                        Next
                        'Adding DataRow
                        Dim cellvalue As String = ""
                        Dim i As Integer = 0
                        For Each row As DataGridViewRow In DataGridView8.Rows
                            For Each cell As DataGridViewCell In row.Cells
                                cellvalue = cell.FormattedValue
                                pdfTable.AddCell(Convert.ToString(cellvalue))
                            Next
                        Next
                        'Exporting to PDF
                        Dim folderPath As String = Konfigurime.TextBox3.Text
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
                        Dim oldFile As String = Konfigurime.TextBox3.Text & "\DataGridViewExport.pdf"
                        Dim newFile As String = Konfigurime.TextBox3.Text & "\DataGridViewExport1.pdf"
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
                        cb.ShowTextAligned(1, emri_shitesit, 100, 597, 0)
                        Dim ad_shitesit As String = shita
                        cb.ShowTextAligned(1, ad_shitesit, 111, 578, 0)
                        Dim cel_shitesit As String = shitc
                        cb.ShowTextAligned(1, cel_shitesit, 97, 559, 0)
                        Dim nipt_shitesit As String = shitn
                        cb.ShowTextAligned(1, nipt_shitesit, 96, 541, 0)
                        Dim emri_bleresit As String = bleresi
                        cb.ShowTextAligned(1, emri_bleresit, 360, 597, 0)
                        Dim ad_bleres As String = bleresia
                        cb.ShowTextAligned(1, ad_bleres, 372, 577, 0)
                        Dim cel_ble As String = bleresic
                        cb.ShowTextAligned(1, cel_ble, 355, 558, 0)
                        Dim nipt_ble As String = bleresin
                        cb.ShowTextAligned(1, nipt_ble, 356, 540, 0)
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
                        My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox3.Text & "\DataGridViewExport.pdf")
                    End Using
                    Using inputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\DataGridViewExport1.pdf", FileMode.Open, FileAccess.Read, FileShare.Read)
                        My.Resources.xx.Save("C:\ProgramData\xx.jpg", Drawing.Imaging.ImageFormat.Jpeg)
                        Using inputImageStream As Stream = New FileStream("C:\ProgramData\xx.jpg", FileMode.Open, FileAccess.Read, FileShare.Read)
                            Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\DataGridViewExport3.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
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
                    My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox3.Text & "\DataGridViewExport1.pdf")
                    Using inputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\DataGridViewExport3.pdf", FileMode.Open, FileAccess.Read, FileShare.Read)
                        Using inputImageStream As Stream = New FileStream(Konfigurime.TextBox2.Text, FileMode.Open, FileAccess.Read, FileShare.Read)
                            Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Shitje.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
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
                    My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox3.Text & "\DataGridViewExport3.pdf")
                    My.Computer.FileSystem.DeleteFile("C:\ProgramData\xx.jpg")
                    ' System.Diagnostics.Process.Start(Konfigurime.TextBox3.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Shitje.pdf")
                    If bleresi Is Nothing Then
                        MsgBox("Dokumenti nuk ekziston!Duhet te ruani me pare faturen!", MsgBoxStyle.Information)
                    Else
                        Dim PrintPDF As New ProcessStartInfo
                        PrintPDF.UseShellExecute = True
                        PrintPDF.Verb = "print"
                        PrintPDF.WindowStyle = ProcessWindowStyle.Hidden
                        PrintPDF.FileName = Konfigurime.TextBox3.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & ".pdf"
                        Process.Start(PrintPDF)
                        Threading.Thread.Sleep(20000)
                        killProcess("Acrobat")
                        Threading.Thread.Sleep(10000)
                        My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox3.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Shitje.pdf")
                        MsgBox("Fatura u printua me sukses!", MsgBoxStyle.Information)
                    End If
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

    Sub ruajfaturen()
        Dim regDate As DateTime = Date.Now
        Dim strDate As String = regDate.ToString("dd/MM/yyyy")
        If Not DataGridView1.Rows.Count > 0 Then
            MsgBox("Fatura eshte bosh!Me pare gjeneroni nje fature!", MsgBoxStyle.Information)
        Else
            If My.Settings.logo = "" Or My.Settings.ruajraportet = "" Then
                MsgBox("Ju nuk keni zgjedhur nje logo ose vendodhjen se ku do te ruhen faturat.Shko tek Konfigurimet!", MsgBoxStyle.Information)
            Else
                Konfigurime.TextBox1.Text = My.Settings.backup
                Konfigurime.TextBox2.Text = My.Settings.logo
                Konfigurime.TextBox3.Text = My.Settings.ruajraportet
                Konfigurime.TextBox4.Text = My.Settings.faturatofert
                lidhjeprint.Open()
                setitedhenave2print.Clear
                adaptori2print = New OleDbDataAdapter("SELECT Produkti,Njesia,Sasia,Cmimi,Vlera_Pa_TVSH,TVSH,Vlera_Me_TVSH,Zbritje_ne_perq FROM Shitjet WHERE(Kodi_Shitjes Like '" & TextBox8.Text & "%')", lidhjeprint)
                adaptori2print.Fill(setitedhenave2print, "tedhena")
                DataGridView8.Columns.Clear()
                DataGridView8.Columns.Add("count1", "Nr.")
                DataGridView8.DataSource = setitedhenave2print.Tables(0)
                'DataGridView8.Columns("ID").Visible = False
                lidhjeprint.Close()
                Dim nr8 As Integer = 0
                nr8 = DataGridView8.Rows.Count - 1
                For i = 0 To nr8
                    DataGridView8.Rows(i).Cells(0).Value = i + 1
                Next
                DataGridView8.Refresh()
                vlerapatvsh = (From row As DataGridViewRow In DataGridView8.Rows
                               Where row.Cells(5).FormattedValue.ToString() <> String.Empty
                               Select Convert.ToInt32(row.Cells(5).FormattedValue)).Sum().ToString()
                vleraetvsh = (From row As DataGridViewRow In DataGridView8.Rows
                              Where row.Cells(6).FormattedValue.ToString() <> String.Empty
                              Select Convert.ToInt32(row.Cells(6).FormattedValue)).Sum().ToString()
                vlerametvsh = (From row As DataGridViewRow In DataGridView8.Rows
                               Where row.Cells(7).FormattedValue.ToString() <> String.Empty
                               Select Convert.ToInt32(row.Cells(7).FormattedValue)).Sum().ToString()
                DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(5).Value = vlerapatvsh
                DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(6).Value = vleraetvsh
                DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(7).Value = vlerametvsh
                DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(4).Value = "TOTALI"
                DataGridView8.Rows(DataGridView8.Rows.Count - 1).Cells(0).Value = ""
                Dim con_bleresit As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
                Dim query_bleresit As String = "SELECT TOP 1 Data,Shitesi,Bleresi,Kodi_Shitjes FROM Shitjet WHERE(Kodi_Shitjes Like '" & TextBox8.Text & "%')"
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
                    Dim pdfTable As New PdfPTable(DataGridView8.ColumnCount)
                    pdfTable.DefaultCell.Padding = 3
                    pdfTable.WidthPercentage = 100
                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT
                    'Adding Header row
                    For Each column As DataGridViewColumn In DataGridView8.Columns
                        Dim cell As New PdfPCell(New Phrase(column.HeaderText.Replace("_", " ").Replace("ne perq", "%")))
                        cell.BorderWidthTop = 1
                        cell.BorderWidthBottom = 1
                        cell.BackgroundColor = New iTextSharp.text.BaseColor(250, 235, 215)
                        pdfTable.AddCell(cell)
                    Next
                    'Adding DataRow
                    Dim cellvalue As String = ""
                    Dim i As Integer = 0
                    For Each row As DataGridViewRow In DataGridView8.Rows
                        For Each cell As DataGridViewCell In row.Cells
                            cellvalue = cell.FormattedValue
                            pdfTable.AddCell(Convert.ToString(cellvalue))
                        Next
                    Next
                    'Exporting to PDF
                    Dim folderPath As String = Konfigurime.TextBox3.Text
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
                    Dim oldFile As String = Konfigurime.TextBox3.Text & "\DataGridViewExport.pdf"
                    Dim newFile As String = Konfigurime.TextBox3.Text & "\DataGridViewExport1.pdf"
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
                    cb.ShowTextAligned(1, emri_shitesit, 100, 597, 0)
                    Dim ad_shitesit As String = shita
                    cb.ShowTextAligned(1, ad_shitesit, 111, 578, 0)
                    Dim cel_shitesit As String = shitc
                    cb.ShowTextAligned(1, cel_shitesit, 97, 559, 0)
                    Dim nipt_shitesit As String = shitn
                    cb.ShowTextAligned(1, nipt_shitesit, 96, 541, 0)
                    Dim emri_bleresit As String = bleresi
                    cb.ShowTextAligned(1, emri_bleresit, 360, 597, 0)
                    Dim ad_bleres As String = bleresia
                    cb.ShowTextAligned(1, ad_bleres, 372, 577, 0)
                    Dim cel_ble As String = bleresic
                    cb.ShowTextAligned(1, cel_ble, 355, 558, 0)
                    Dim nipt_ble As String = bleresin
                    cb.ShowTextAligned(1, nipt_ble, 356, 540, 0)
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
                    My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox3.Text & "\DataGridViewExport.pdf")
                End Using
                Using inputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\DataGridViewExport1.pdf", FileMode.Open, FileAccess.Read, FileShare.Read)
                    My.Resources.xx.Save("C:\ProgramData\xx.jpg", Drawing.Imaging.ImageFormat.Jpeg)
                    Using inputImageStream As Stream = New FileStream("C:\ProgramData\xx.jpg", FileMode.Open, FileAccess.Read, FileShare.Read)
                        Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\DataGridViewExport3.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
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
                My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox3.Text & "\DataGridViewExport1.pdf")
                Using inputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\DataGridViewExport3.pdf", FileMode.Open, FileAccess.Read, FileShare.Read)
                    Using inputImageStream As Stream = New FileStream(Konfigurime.TextBox2.Text, FileMode.Open, FileAccess.Read, FileShare.Read)
                        Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Shitje.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
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
                My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox3.Text & "\DataGridViewExport3.pdf")
                My.Computer.FileSystem.DeleteFile("C:\ProgramData\xx.jpg")
                System.Diagnostics.Process.Start(Konfigurime.TextBox3.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Shitje.pdf")
            End If
        End If
    End Sub
End Class