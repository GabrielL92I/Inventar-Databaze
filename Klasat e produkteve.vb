Imports System.Data.OleDb
Imports System.Linq
Public Class Klasat_e_produkteve
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
    End Sub
    Private Sub _Load(ByVal sender As Object, ByVal e As EventArgs)
        _form_resize._get_initial_size()
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
    Public Sub CreateReader_load_produktet_klasa(ByVal connectionString As String,
    ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                If ListBox1.Items.Contains(reader(1).ToString) Or ListBox2.Items.Contains(reader(1).ToString) Or ListBox3.Items.Contains(reader(1).ToString) Or ListBox4.Items.Contains(reader(1).ToString) Or ListBox5.Items.Contains(reader(1).ToString) Or ListBox6.Items.Contains(reader(1).ToString) Or ListBox7.Items.Contains(reader(1).ToString) Or ListBox8.Items.Contains(reader(1).ToString) Then

                Else
                    ListBox1.Items.Add(reader(1).ToString)
                End If
            End While
            reader.Close()
            connection.Close()
        End Using
    End Sub
    'Lexon Bleresit ne ComboBox2
    Public Sub Kap_Bleresit(ByVal connectionString As String,
   ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                ComboBox2.Items.Add(reader(1).ToString())
                If reader.HasRows = True Then
                    ComboBox2.SelectedIndex = 0
                Else
                End If
            End While
            reader.Close()
        End Using
    End Sub
    Public Sub Merrtedhenat_Klasat(ByVal rreshtiaktualdg71ofert121klasat)
        Try
            Label11.Text = setitedhenavedg71ofert121klasat.Tables("tedhena").Rows(rreshtiaktualdg71ofert121klasat)("ID")
            TextBox1.Text = setitedhenavedg71ofert121klasat.Tables("tedhena").Rows(rreshtiaktualdg71ofert121klasat)("Klasa_A")
            TextBox2.Text = setitedhenavedg71ofert121klasat.Tables("tedhena").Rows(rreshtiaktualdg71ofert121klasat)("Klasa_B")
            TextBox3.Text = setitedhenavedg71ofert121klasat.Tables("tedhena").Rows(rreshtiaktualdg71ofert121klasat)("Klasa_C")
            TextBox5.Text = setitedhenavedg71ofert121klasat.Tables("tedhena").Rows(rreshtiaktualdg71ofert121klasat)("Klasa_D")
            TextBox6.Text = setitedhenavedg71ofert121klasat.Tables("tedhena").Rows(rreshtiaktualdg71ofert121klasat)("Klasa_E")
            TextBox7.Text = setitedhenavedg71ofert121klasat.Tables("tedhena").Rows(rreshtiaktualdg71ofert121klasat)("Klasa_F")
            TextBox8.Text = setitedhenavedg71ofert121klasat.Tables("tedhena").Rows(rreshtiaktualdg71ofert121klasat)("Klasa_G")
        Catch ex As Exception
        End Try
    End Sub
    Public Function autofillbleresitklasat()
        setitedhenave2autofill.clear()

        lidhje2autofill.Open()
        adaptori2autofill = New OleDbDataAdapter("SELECT Bleresi FROM Klasat ORDER BY ID", lidhje2autofill)
        adaptori2autofill.Fill(setitedhenave2autofill, "tedhena")
        Dim col As New AutoCompleteStringCollection
        Dim i As Integer
        For i = 0 To setitedhenave2autofill.Tables(0).Rows.Count - 1
            col.Add(setitedhenave2autofill.Tables(0).Rows(i)("Bleresi").ToString())
        Next
        TextBox4.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox4.AutoCompleteCustomSource = col
        TextBox4.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        lidhje2autofill.Close()
        Return Nothing
    End Function
    Private Sub Klasat_e_produkteve_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        autofillbleresitklasat()
        With DataGridView1
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktualdg71ofert121klasat = 0
        lidhjedg71ofert121klasat.Open()
        setitedhenavedg71ofert121klasat.clear()
        adaptoridg71ofert121klasat = New OleDbDataAdapter("SELECT * FROM Klasat ORDER BY ID", lidhjedg71ofert121klasat)
        adaptoridg71ofert121klasat.Fill(setitedhenavedg71ofert121klasat, "tedhena")
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("count1", "Nr.")
        DataGridView1.DataSource = setitedhenavedg71ofert121klasat.Tables(0)
        DataGridView1.Columns("ID").Visible = False
        Dim nr1 As Integer = 0
        nr1 = DataGridView1.Rows.Count - 1
        For i = 0 To nr1
            DataGridView1.Rows(i).Cells(0).Value = i + 1
        Next
        If setitedhenavedg71ofert121klasat.Tables(0).Rows.Count > 0 Then
            DataGridView1.Columns("ID").Visible = False
            Merrtedhenat_Klasat(rreshtiaktualdg71ofert121klasat)
        Else

        End If
        lidhjedg71ofert121klasat.Close()
        If DataGridView1.Rows.Count = 0 Then
            Button3.Enabled = False
            Button4.Enabled = False
        Else
            Button3.Enabled = True
            Button4.Enabled = True
        End If
        Dim column As DataGridViewColumn = DataGridView1.Columns(0)
        column.Width = 30
        'Kap bleresit nga tabela Klientet dhe i eksporton ne ComboBox1
        Dim con_bleresit As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim query_bleresit As String = "SELECT * FROM Bleresit ORDER BY ID"
        Kap_Bleresit(con_bleresit, query_bleresit)
        Try
            If My.Settings.YourItems IsNot Nothing Then
                For Each S As String In My.Settings.YourItems
                    ListBox1.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems1 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems1
                    ListBox2.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems2 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems2
                    ListBox3.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems3 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems3
                    ListBox4.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems4 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems4
                    ListBox5.Items.Add(S)
                Next

            End If
            If My.Settings.YourItems5 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems5
                    ListBox6.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems6 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems6
                    ListBox7.Items.Add(S)
                Next
            End If
            If My.Settings.YourItems7 IsNot Nothing Then
                For Each S As String In My.Settings.YourItems7
                    ListBox8.Items.Add(S)
                Next
            End If
        Catch ex As Exception

        End Try
        ListBox1.AllowDrop = True
        ListBox2.AllowDrop = True
        ListBox3.AllowDrop = True
        ListBox4.AllowDrop = True
        ListBox5.AllowDrop = True
        ListBox6.AllowDrop = True
        ListBox7.AllowDrop = True
        ListBox8.AllowDrop = True
        Dim conyklasa As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim queyklasa As String = "SELECT * FROM Produktet ORDER BY ID"
        CreateReader_load_produktet_klasa(conyklasa, queyklasa)
    End Sub
    Private Sub Klasat_e_produkteve_Closing(sender As Object, e As EventArgs) Handles MyBase.Closing
        My.Settings.YourItems = New Specialized.StringCollection
        For Each S As String In ListBox1.Items
            My.Settings.YourItems.Add(S)
        Next
        My.Settings.Save()
        My.Settings.YourItems1 = New Specialized.StringCollection
        For Each S As String In ListBox2.Items
            My.Settings.YourItems1.Add(S)
        Next
        My.Settings.Save()
        My.Settings.YourItems2 = New Specialized.StringCollection
        For Each S As String In ListBox3.Items
            My.Settings.YourItems2.Add(S)
        Next
        My.Settings.Save()
        My.Settings.YourItems3 = New Specialized.StringCollection
        For Each S As String In ListBox4.Items
            My.Settings.YourItems3.Add(S)
        Next
        My.Settings.YourItems4 = New Specialized.StringCollection
        For Each S As String In ListBox5.Items
            My.Settings.YourItems4.Add(S)
        Next
        My.Settings.Save()
        My.Settings.YourItems5 = New Specialized.StringCollection
        For Each S As String In ListBox6.Items
            My.Settings.YourItems5.Add(S)
        Next
        My.Settings.Save()
        My.Settings.YourItems6 = New Specialized.StringCollection
        For Each S As String In ListBox7.Items
            My.Settings.YourItems6.Add(S)
        Next
        My.Settings.Save()
        My.Settings.YourItems7 = New Specialized.StringCollection
        For Each S As String In ListBox8.Items
            My.Settings.YourItems7.Add(S)
        Next
        My.Settings.Save()
        Farmacia.Show()
    End Sub
    Private Sub lst_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox8.MouseDown, ListBox7.MouseDown, ListBox6.MouseDown, ListBox5.MouseDown, ListBox4.MouseDown, ListBox3.MouseDown, ListBox2.MouseDown, ListBox1.MouseDown
        Dim lst As ListBox = DirectCast(sender, ListBox)
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim index As Integer = lst.IndexFromPoint(e.X, e.Y)
            If index <> ListBox.NoMatches Then
                Dim item As String = lst.Items(index)
                Dim drop_effect As DragDropEffects =
                    lst.DoDragDrop(
                        lst.Items(index),
                        DragDropEffects.Move Or DragDropEffects.Copy)
                ' If it was moved, remove the item from this list.
                If drop_effect = DragDropEffects.Move Then
                    ' See if the user dropped the item in this ListBox
                    ' at a higher position.
                    If lst.Items(index) = item Then
                        ' The item has not moved.
                        lst.Items.RemoveAt(index)
                    Else
                        ' The item has moved.
                        lst.Items.RemoveAt(index + 1)
                    End If
                End If
            End If
        End If
    End Sub
    ' Display the appropriate cursor.
    Private Sub lstFruits_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox8.DragOver, ListBox7.DragOver, ListBox6.DragOver, ListBox5.DragOver, ListBox4.DragOver, ListBox3.DragOver, ListBox2.DragOver, ListBox1.DragOver
        Const KEY_CTRL As Integer = 8

        If Not (e.Data.GetDataPresent(GetType(System.String))) Then
            e.Effect = DragDropEffects.None
        ElseIf (e.KeyState And KEY_CTRL) And
        (e.AllowedEffect And DragDropEffects.Copy) = DragDropEffects.Copy Then
            ' Copy.
            e.Effect = DragDropEffects.Copy
        ElseIf (e.AllowedEffect And DragDropEffects.Move) = DragDropEffects.Move Then
            ' Move.
            e.Effect = DragDropEffects.Move
        End If
    End Sub
    ' Drop the entry in the list.
    Private Sub lstFruits_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox8.DragDrop, ListBox7.DragDrop, ListBox6.DragDrop, ListBox5.DragDrop, ListBox4.DragDrop, ListBox3.DragDrop, ListBox2.DragDrop, ListBox1.DragDrop
        If e.Data.GetDataPresent(GetType(System.String)) Then
            If (e.Effect = DragDropEffects.Copy) Or
            (e.Effect = DragDropEffects.Move) Then
                Dim lst As ListBox = DirectCast(sender, ListBox)
                Dim item As Object = CType(e.Data.GetData(GetType(System.String)), System.Object)
                Dim pt As Point = lst.PointToClient(New Point(e.X, e.Y))
                Dim index As Integer = lst.IndexFromPoint(pt.X, pt.Y)
                If index = ListBox.NoMatches Then
                    lst.Items.Add(item)
                Else
                    lst.Items.Insert(index, item)
                End If
            End If
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim nrekzistues As New List(Of Integer)
        For Each kolone As DataGridViewRow In DataGridView1.Rows
            nrekzistues.Add(CInt(kolone.Cells(0).Value))
        Next
        Dim existingNumbers As New List(Of Integer)
        For Each r As DataGridViewRow In DataGridView1.Rows
            existingNumbers.Add(CInt(r.Cells(0).Value))
        Next
        If existingNumbers.Count = 0 And nrekzistues.Count = 0 Then
            Dim Str As String
            Try
                Str = "insert into Klasat values("
                Str += "1"
                Str += ","
                Str += """" & ComboBox2.Text & """"
                Str += ","
                Str += """" & TextBox1.Text.Trim & """"
                Str += ","
                Str += """" & TextBox2.Text.Trim() & """"
                Str += ","
                Str += """" & TextBox3.Text.Trim() & """"
                Str += ","
                Str += """" & TextBox5.Text.Trim & """"
                Str += ","
                Str += """" & TextBox6.Text.Trim() & """"
                Str += ","
                Str += """" & TextBox7.Text.Trim() & """"
                Str += ","
                Str += """" & TextBox8.Text.Trim() & """"
                Str += ")"
                lidhje211.Open()
                query211 = New OleDbCommand(Str, lidhje211)
                query211.ExecuteNonQuery()
                setitedhenave211.Clear()
                adaptori211 = New OleDbDataAdapter("SELECT * FROM Klasat ORDER BY ID", lidhje211)
                adaptori211.Fill(setitedhenave211, "tedhena")
                DataGridView1.DataSource = setitedhenave211.Tables(0)
                If setitedhenave211.Tables(0).Rows.Count > 0 Then
                    DataGridView1.Columns("ID").Visible = False
                Else

                End If
                Dim nr1 As Integer = 0
                nr1 = DataGridView1.Rows.Count - 1
                For i = 0 To nr1
                    DataGridView1.Rows(i).Cells(0).Value = i + 1
                Next
                MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                lidhje211.Close()
            Catch ex As Exception
                MsgBox("Nuk u shtua", MsgBoxStyle.Information)
                lidhje211.Close()
            End Try
        Else
            Dim missingNumbers = Enumerable.Range(existingNumbers.First, existingNumbers.Max() - existingNumbers.First + 1).Except(existingNumbers)
            Dim max = nrekzistues.Max() + 1
            If missingNumbers.Count = 0 Then
                Dim Str As String
                Try
                    Str = "insert into Klasat values("
                    Str += max.ToString
                    Str += ","
                    Str += """" & ComboBox2.Text & """"
                    Str += ","
                    Str += """" & TextBox1.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox2.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox3.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox5.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox6.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox7.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox8.Text.Trim() & """"
                    Str += ")"
                    lidhje211.Open()
                    query211 = New OleDbCommand(Str, lidhje211)
                    query211.ExecuteNonQuery()
                    setitedhenave211.Clear()
                    adaptori211 = New OleDbDataAdapter("SELECT * FROM Klasat ORDER BY ID", lidhje211)
                    adaptori211.Fill(setitedhenave211, "tedhena")
                    DataGridView1.DataSource = setitedhenave211.Tables(0)
                    If setitedhenave211.Tables(0).Rows.Count > 0 Then
                        DataGridView1.Columns("ID").Visible = False
                    Else

                    End If
                    Dim nr1 As Integer = 0
                    nr1 = DataGridView1.Rows.Count - 1
                    For i = 0 To nr1
                        DataGridView1.Rows(i).Cells(0).Value = i + 1
                    Next
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    lidhje211.Close()
                Catch ex As Exception
                    MsgBox("Nuk u shtua", MsgBoxStyle.Information)
                    lidhje211.Close()
                End Try
            Else
                Dim Str As String
                Try
                    Str = "insert into Klasat values("
                    Str += missingNumbers.First.ToString
                    Str += ","
                    Str += """" & ComboBox2.Text & """"
                    Str += ","
                    Str += """" & TextBox1.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox2.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox3.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox5.Text.Trim & """"
                    Str += ","
                    Str += """" & TextBox6.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox7.Text.Trim() & """"
                    Str += ","
                    Str += """" & TextBox8.Text.Trim() & """"
                    Str += ")"
                    lidhje211.Open()
                    query211 = New OleDbCommand(Str, lidhje211)
                    query211.ExecuteNonQuery()
                    setitedhenave211.Clear()
                    adaptori211 = New OleDbDataAdapter("SELECT * FROM Klasat ORDER BY ID", lidhje211)
                    adaptori211.Fill(setitedhenave211, "tedhena")
                    DataGridView1.DataSource = setitedhenave211.Tables(0)
                    If setitedhenave211.Tables(0).Rows.Count > 0 Then
                        DataGridView1.Columns("ID").Visible = False
                    Else

                    End If
                    Dim nr1 As Integer = 0
                    nr1 = DataGridView1.Rows.Count - 1
                    For i = 0 To nr1
                        DataGridView1.Rows(i).Cells(0).Value = i + 1
                    Next
                    MsgBox("U shtua me sukses", MsgBoxStyle.Information)
                    lidhje211.Close()
                Catch ex As Exception
                    MsgBox("Nuk u shtua", MsgBoxStyle.Information)
                    lidhje211.Close()
                End Try
            End If
        End If
        Button3.Enabled = True
        Button4.Enabled = True
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim Str As String
            Str = "update Klasat set Klasa_A="
            Str += """" & TextBox1.Text & """"
            Str += " where ID="
            Str += Label11.Text
            lidhje211.Open()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Klasat set Klasa_B="
            Str += """" & TextBox2.Text & """"
            Str += " where ID="
            Str += Label11.Text
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Klasat set Klasa_C="
            Str += """" & TextBox3.Text & """"
            Str += " where ID="
            Str += Label11.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Klasat set Klasa_D="
            Str += """" & TextBox5.Text & """"
            Str += " where ID="
            Str += Label11.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Klasat set Klasa_E="
            Str += """" & TextBox6.Text & """"
            Str += " where ID="
            Str += Label11.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Klasat set Klasa_F="
            Str += """" & TextBox7.Text & """"
            Str += " where ID="
            Str += Label11.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()

            lidhje211.Open()
            Str = "update Klasat set Klasa_G="
            Str += """" & TextBox8.Text & """"
            Str += " where ID="
            Str += Label11.Text.Trim()
            query211 = New OleDbCommand(Str, lidhje211)
            query211.ExecuteNonQuery()
            lidhje211.Close()
            DataGridView1.DataSource = Nothing
            With DataGridView1
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            rreshtiaktualdg71ofert121klasat = 0
            lidhjedg71ofert121klasat.Open()
            setitedhenavedg71ofert121klasat.clear
            adaptoridg71ofert121klasat = New OleDbDataAdapter("SELECT * FROM Klasat ORDER BY ID", lidhjedg71ofert121klasat)
            adaptoridg71ofert121klasat.Fill(setitedhenavedg71ofert121klasat, "tedhena")
            DataGridView1.DataSource = setitedhenavedg71ofert121klasat.Tables(0)
            DataGridView1.Columns("ID").Visible = False
            Dim nr1 As Integer = 0
            nr1 = DataGridView1.Rows.Count - 1
            For i = 0 To nr1
                DataGridView1.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView1.Refresh()
            Merrtedhenat_Klasat(rreshtiaktualdg71ofert121klasat)
            lidhjedg71ofert121klasat.Close()
            MsgBox("U rifreskua me sukses", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("Nuk u rifreskua me sukses", MsgBoxStyle.Information)
        End Try
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If (e.RowIndex >= 0) Then
            Dim row As DataGridViewRow = DataGridView1.Rows.Item(e.RowIndex)
            TextBox1.Text = row.Cells.Item("Klasa_A").Value.ToString
            TextBox2.Text = row.Cells.Item("Klasa_B").Value.ToString
            TextBox3.Text = row.Cells.Item("Klasa_C").Value.ToString
            TextBox5.Text = row.Cells.Item("Klasa_D").Value.ToString
            TextBox6.Text = row.Cells.Item("Klasa_E").Value.ToString
            TextBox7.Text = row.Cells.Item("Klasa_F").Value.ToString
            TextBox8.Text = row.Cells.Item("Klasa_G").Value.ToString
            Label11.Text = row.Cells.Item("ID").Value.ToString
            ComboBox2.Text = row.Cells.Item("Bleresi").Value.ToString
        End If
    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        lidhjedg71ofert121klasatkerk.Close()
        lidhjedg71ofert121klasatkerk.Open()
        adaptoridg71ofert121klasatkerk = New OleDbDataAdapter("SELECT * FROM Klasat WHERE(ID Like '" &
                                             TextBox4.Text & "%' OR Bleresi Like '" & TextBox4.Text & "%' OR Klasa_A Like '" &
                                             TextBox4.Text & "%' OR  Klasa_B Like '" & TextBox4.Text & "%' OR Klasa_C Like '" &
                                             TextBox4.Text & "%' )", Lidhje.lidhjedg71ofert121klasatkerk)
        Dim dataTablecc As New DataTable
        adaptoridg71ofert121klasatkerk.Fill(dataTablecc)
        DataGridView1.DataSource = dataTablecc
        Dim nr1 As Integer = 0
        nr1 = DataGridView1.Rows.Count - 1
        For i = 0 To nr1
            DataGridView1.Rows(i).Cells(0).Value = i + 1
        Next
        Merrtedhenat_Klasat(rreshtiaktualdg71ofert121klasat)
        DataGridView1.CurrentCell = Nothing
        If DataGridView1.CurrentCell Is Nothing Then

        Else
            Dim row As DataGridViewRow = Me.DataGridView1.Rows.Item(rreshtiaktual2)
            TextBox1.Text = row.Cells.Item("Klasa_A").Value.ToString
            TextBox2.Text = row.Cells.Item("Klasa_B").Value.ToString
            TextBox3.Text = row.Cells.Item("Klasa_C").Value.ToString
            ComboBox2.Text = row.Cells.Item("Bleresi").Value.ToString
            lidhjedg71ofert121klasatkerk.Close()
            DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            DataGridView1.RefreshEdit()
            lidhjedg71ofert121klasatkerk.Close()
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim Str As String
        Try
            Str = "delete from Klasat where ID="
            Str += DataGridView1.CurrentRow.Cells(1).Value.ToString
            lidhjefshijklasat.Open()
            queryfshijklasat = New OleDbCommand(Str, lidhjefshijklasat)
            queryfshijklasat.ExecuteNonQuery()
            setitedhenavefshijklasat.clear()
            adaptorifshijklasat = New OleDbDataAdapter("SELECT * FROM Klasat ORDER BY ID", lidhjefshijklasat)
            adaptorifshijklasat.Fill(setitedhenavefshijklasat, "tedhena")
            DataGridView1.DataSource = setitedhenavefshijklasat.Tables(0)
            Dim nr1 As Integer = 0
            nr1 = DataGridView1.Rows.Count - 1
            For i = 0 To nr1
                DataGridView1.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView1.Refresh()
            MsgBox("U fshi me sukses", MsgBoxStyle.Information)
            If rreshtiaktualfshijklasat > 0 Then
                rreshtiaktualfshijklasat -= 1
                Merrtedhenat_Klasat(rreshtiaktualfshijklasat)
            End If
            lidhjefshijklasat.Close()
        Catch ex As Exception
            MsgBox("Nuk u fshi", MsgBoxStyle.Information)
            lidhjefshijklasat.Close()
        End Try
    End Sub
End Class