Public Class Konfigurime
    Private Sub Konfigurime_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = My.Settings.backup
        TextBox2.Text = My.Settings.logo
        TextBox3.Text = My.Settings.ruajraportet
        TextBox4.Text = My.Settings.faturatofert
        TextBox5.Text = My.Settings.ruajgjendjen
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dialog As New FolderBrowserDialog()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = "C:\"
        dialog.Description = "Zgjidh vendodhjen se ku doni te ruhet kopja e databazes!"
        If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = dialog.SelectedPath & "\tedhena_backup.accdb"
            My.Settings.backup = TextBox1.Text
            My.Settings.Save()
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'OpenFileDialog1.Filter = "Excel Files (*.*)|*.All"
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK _
       Then
            TextBox2.Text = OpenFileDialog1.FileName
            My.Settings.logo = TextBox2.Text
            My.Settings.Save()
        End If
    End Sub
    Private Sub Konfigurime_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        My.Settings.backup = TextBox1.Text
        My.Settings.logo = TextBox2.Text
        My.Settings.ruajraportet = TextBox3.Text
        My.Settings.faturatofert = TextBox4.Text
        My.Settings.ruajgjendjen = TextBox5.Text
        My.Settings.Save()
        Farmacia.Show()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dialog As New FolderBrowserDialog()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = "C:\"
        dialog.Description = "Zgjidh vendodhjen se ku doni te ruhen faturat e shitjeve!"
        If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TextBox3.Text = dialog.SelectedPath
            My.Settings.ruajraportet = TextBox3.Text
            My.Settings.Save()
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim dialog As New FolderBrowserDialog()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = "C:\"
        dialog.Description = "Zgjidh vendodhjen se ku doni te ruhen faturat e ofertave!"
        If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TextBox4.Text = dialog.SelectedPath
            My.Settings.faturatofert = TextBox4.Text
            My.Settings.Save()
        End If
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim dialog As New FolderBrowserDialog()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = "C:\"
        dialog.Description = "Zgjidh vendodhjen se ku doni te ruhen raportet e gjendes se magazines!"
        If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TextBox5.Text = dialog.SelectedPath
            My.Settings.ruajgjendjen = TextBox5.Text
            My.Settings.Save()
        End If
    End Sub
End Class