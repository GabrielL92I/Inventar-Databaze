Imports System.Data.OleDb
Public Class Fjalkalimi_Databazes
    Dim path As String = My.Settings.ruajdtbpath & "\tedhena.accdb;"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            MsgBox("Me pare vendos fjalkalimin e ri!", MsgBoxStyle.Information)
        Else
            Try
                Dim cn As OleDbConnection = New OleDbConnection
                cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Mode=Share Deny Read|Share Deny Write;Persist Security Info=False;Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
                cn.Open()
                Dim cmd As OleDbCommand = New OleDbCommand
                cmd.Connection = cn
                cmd.CommandText = "ALTER DATABASE PASSWORD [" & TextBox2.Text.Trim() & "][" & TextBox1.Text.Trim() & "]"
                'cmd.CommandText = "ALTER DATABASE PASSWORD [newPassword][OldPassword]"
                cmd.ExecuteNonQuery()
                cn.Close()
                Hyrje.TextBox3.Text = TextBox2.Text
                My.Settings.dtb1 = Hyrje.TextBox3.Text
                My.Settings.Save()
                MsgBox("Fjalkalimi u ndryshua me sukses!", MsgBoxStyle.Information)
                MsgBox("Programi do te rihapet!", MsgBoxStyle.Information)
                Application.Restart()
            Catch ex As Exception
                MsgBox("Dicka shkoi keq!", MsgBoxStyle.Information)
            End Try
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.PasswordChar = ""
        Else
            TextBox1.PasswordChar = "*"
        End If
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox2.PasswordChar = ""
        Else
            TextBox2.PasswordChar = "*"
        End If
    End Sub
    Private Sub Fjalkalimi_Databazes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = Hyrje.TextBox3.Text
    End Sub
    Private Sub Fjalkalimi_Databazes_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Me.Hide()
        Hyrje.Show()
    End Sub
End Class