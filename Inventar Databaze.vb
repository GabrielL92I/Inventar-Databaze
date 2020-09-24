Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Public Class Farmacia
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
    Dim Lidhje1 As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Dim adaptor1 As OleDbDataAdapter
    Dim setidatave1 = New DataSet
    Dim query1 As OleDbCommand
    Dim Lidhje1b As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Dim adaptor1b As OleDbDataAdapter
    Dim setidatave1b = New DataSet
    Dim query1b As OleDbCommand
    Dim Lidhje1bc As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Dim adaptor1bc As OleDbDataAdapter
    Dim setidatave1bc = New DataSet
    Dim query1bc As OleDbCommand
    Dim Lidhje1bcx As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Dim adaptor1bcx As OleDbDataAdapter
    Dim setidatave1bcx = New DataSet
    Dim query1bcx As OleDbCommand
    Dim nr1 As Integer = 0
    Dim nr2 As Integer = 0
    Dim nr3 As Integer = 0
    Dim nr4 As Integer = 0
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        If Not DataGridView7.Rows.Count > 0 Then
            MsgBox("Fatura eshte bosh!Me pare gjeneroni nje fature!", MsgBoxStyle.Information)
        Else
            Dim regDate1 As DateTime = Date.Now
            Dim strDate1 As String = regDate1.ToString("dd/MM/yyyy")
            If My.Settings.logo = "" Or My.Settings.faturatofert = "" Then
                MsgBox("Ju nuk keni zgjedhur nje logo ose vendodhjen se ku do te ruhen faturat.Shko tek Konfigurimet!", MsgBoxStyle.Information)
            Else
                Try
                    DataGridView9.DataSource = Nothing
                    Konfigurime.TextBox1.Text = My.Settings.backup
                    Konfigurime.TextBox2.Text = My.Settings.logo
                    Konfigurime.TextBox3.Text = My.Settings.ruajraportet
                    Konfigurime.TextBox4.Text = My.Settings.faturatofert
                    lidhjeprint.Open()
                    setitedhenave2print.Clear
                    adaptori2print = New OleDbDataAdapter("SELECT Produkti,Njesia,Sasia,Cmimi,Vlera_Pa_TVSH,TVSH,Vlera_Me_TVSH,Zbritje_ne_perq FROM Ofertat WHERE(Kodi_Ofertes Like '" & TextBox38.Text & "%')", lidhjeprint)
                    adaptori2print.Fill(setitedhenave2print, "tedhena")
                    DataGridView9.Columns.Clear()
                    DataGridView9.Columns.Add("count1", "Nr.")
                    DataGridView9.DataSource = setitedhenave2print.Tables(0)
                    'DataGridView8.Columns("ID").Visible = False
                    lidhjeprint.Close()
                    Dim nr9 As Integer = 0
                    nr9 = DataGridView9.Rows.Count - 1
                    For i = 0 To nr9
                        DataGridView9.Rows(i).Cells(0).Value = i + 1
                    Next
                    DataGridView9.Refresh()
                    vlerapatvsh = (From row As DataGridViewRow In DataGridView9.Rows
                                   Where row.Cells(5).FormattedValue.ToString() <> String.Empty
                                   Select Convert.ToInt32(row.Cells(5).FormattedValue)).Sum().ToString()
                    vleraetvsh = (From row As DataGridViewRow In DataGridView9.Rows
                                  Where row.Cells(6).FormattedValue.ToString() <> String.Empty
                                  Select Convert.ToInt32(row.Cells(6).FormattedValue)).Sum().ToString()
                    vlerametvsh = (From row As DataGridViewRow In DataGridView9.Rows
                                   Where row.Cells(7).FormattedValue.ToString() <> String.Empty
                                   Select Convert.ToInt32(row.Cells(7).FormattedValue)).Sum().ToString()
                    DataGridView9.Rows(DataGridView9.Rows.Count - 1).Cells(5).Value = vlerapatvsh
                    DataGridView9.Rows(DataGridView9.Rows.Count - 1).Cells(6).Value = vleraetvsh
                    DataGridView9.Rows(DataGridView9.Rows.Count - 1).Cells(7).Value = vlerametvsh
                    DataGridView9.Rows(DataGridView9.Rows.Count - 1).Cells(4).Value = "TOTALI"
                    DataGridView9.Rows(DataGridView9.Rows.Count - 1).Cells(0).Value = ""
                    Dim con_bleresit As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
                    Dim query_bleresit As String = "SELECT TOP 1 Data,Shitesi,Bleresi,Kodi_Ofertes FROM Ofertat WHERE(Kodi_Ofertes Like '" & TextBox38.Text & "%')"
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
                        Dim pdfTable As New PdfPTable(DataGridView9.ColumnCount)
                        pdfTable.DefaultCell.Padding = 3
                        pdfTable.WidthPercentage = 100
                        pdfTable.HorizontalAlignment = Element.ALIGN_LEFT
                        'Adding Header row
                        For Each column As DataGridViewColumn In DataGridView9.Columns
                            Dim cell As New PdfPCell(New Phrase(column.HeaderText.Replace("_", " ").Replace("ne perq", "%")))
                            cell.BorderWidthTop = 1
                            cell.BorderWidthBottom = 1
                            cell.BackgroundColor = New iTextSharp.text.BaseColor(250, 235, 215)
                            pdfTable.AddCell(cell)
                        Next
                        'Adding DataRow
                        Dim cellvalue As String = ""
                        Dim i As Integer = 0
                        For Each row As DataGridViewRow In DataGridView9.Rows
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
                            Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox4.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate1.Replace("/", "-") & "_" & TextBox38.Text & "_Oferte.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
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
                    Dim regDate As DateTime = Date.Now
                    Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                    Dim PrintPDF As New ProcessStartInfo
                    PrintPDF.UseShellExecute = True
                    PrintPDF.Verb = "print"
                    PrintPDF.WindowStyle = ProcessWindowStyle.Hidden
                    PrintPDF.FileName = Konfigurime.TextBox4.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox38.Text & "_Oferte.pdf"
                    Process.Start(PrintPDF)
                    Threading.Thread.Sleep(20000)
                    killProcess("Acrobat")
                    Threading.Thread.Sleep(10000)
                    My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox4.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox38.Text & "_Oferte.pdf")
                    MsgBox("Fatura u printua me sukses!", MsgBoxStyle.Information)
                    'System.Diagnostics.Process.Start(Konfigurime.TextBox4.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate1.Replace("/", "-") & "_" & TextBox38.Text & "_Oferte.pdf")
                Catch ex As Exception
                    MsgBox("Dokumenti eshte i hapur.Mbyll dokumentin dhe provo perseri!", MsgBoxStyle.Information)
                End Try
            End If
        End If
    End Sub
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        If Not DataGridView7.Rows.Count > 0 Then
            MsgBox("Fatura eshte bosh!Me pare gjeneroni nje fature!", MsgBoxStyle.Information)
        Else
            Dim regDate As DateTime = Date.Now
            Dim strDate As String = regDate.ToString("dd/MM/yyyy")
            If My.Settings.logo = "" Or My.Settings.faturatofert = "" Then
                MsgBox("Ju nuk keni zgjedhur nje logo ose vendodhjen se ku do te ruhen faturat.Shko tek Konfigurimet!", MsgBoxStyle.Information)
            Else
                Try
                    DataGridView9.DataSource = Nothing
                    Konfigurime.TextBox1.Text = My.Settings.backup
                    Konfigurime.TextBox2.Text = My.Settings.logo
                    Konfigurime.TextBox3.Text = My.Settings.ruajraportet
                    Konfigurime.TextBox4.Text = My.Settings.faturatofert
                    lidhjeprint.Open()
                    setitedhenave2print.Clear
                    adaptori2print = New OleDbDataAdapter("SELECT Produkti,Njesia,Sasia,Cmimi,Vlera_Pa_TVSH,TVSH,Vlera_Me_TVSH,Zbritje_ne_perq FROM Ofertat WHERE(Kodi_Ofertes Like '" & TextBox38.Text & "%')", lidhjeprint)
                    adaptori2print.Fill(setitedhenave2print, "tedhena")
                    DataGridView9.Columns.Clear()
                    DataGridView9.Columns.Add("count1", "Nr.")
                    DataGridView9.DataSource = setitedhenave2print.Tables(0)
                    'DataGridView8.Columns("ID").Visible = False
                    lidhjeprint.Close()
                    Dim nr9 As Integer = 0
                    nr9 = DataGridView9.Rows.Count - 1
                    For i = 0 To nr9
                        DataGridView9.Rows(i).Cells(0).Value = i + 1
                    Next
                    DataGridView9.Refresh()
                    vlerapatvsh = (From row As DataGridViewRow In DataGridView9.Rows
                                   Where row.Cells(5).FormattedValue.ToString() <> String.Empty
                                   Select Convert.ToInt32(row.Cells(5).FormattedValue)).Sum().ToString()
                    vleraetvsh = (From row As DataGridViewRow In DataGridView9.Rows
                                  Where row.Cells(6).FormattedValue.ToString() <> String.Empty
                                  Select Convert.ToInt32(row.Cells(6).FormattedValue)).Sum().ToString()
                    vlerametvsh = (From row As DataGridViewRow In DataGridView9.Rows
                                   Where row.Cells(7).FormattedValue.ToString() <> String.Empty
                                   Select Convert.ToInt32(row.Cells(7).FormattedValue)).Sum().ToString()
                    DataGridView9.Rows(DataGridView9.Rows.Count - 1).Cells(5).Value = vlerapatvsh
                    DataGridView9.Rows(DataGridView9.Rows.Count - 1).Cells(6).Value = vleraetvsh
                    DataGridView9.Rows(DataGridView9.Rows.Count - 1).Cells(7).Value = vlerametvsh
                    DataGridView9.Rows(DataGridView9.Rows.Count - 1).Cells(4).Value = "TOTALI"
                    DataGridView9.Rows(DataGridView9.Rows.Count - 1).Cells(0).Value = ""
                    Dim con_bleresit As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
                    Dim query_bleresit As String = "SELECT TOP 1 Data,Shitesi,Bleresi,Kodi_Ofertes FROM Ofertat WHERE(Kodi_Ofertes Like '" & TextBox38.Text & "%')"
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
                        Dim pdfTable As New PdfPTable(DataGridView9.ColumnCount)
                        pdfTable.DefaultCell.Padding = 3
                        pdfTable.WidthPercentage = 100
                        pdfTable.HorizontalAlignment = Element.ALIGN_LEFT
                        'Adding Header row
                        For Each column As DataGridViewColumn In DataGridView9.Columns
                            Dim cell As New PdfPCell(New Phrase(column.HeaderText.Replace("_", " ").Replace("ne perq", "%")))
                            cell.BorderWidthTop = 1
                            cell.BorderWidthBottom = 1
                            cell.BackgroundColor = New iTextSharp.text.BaseColor(250, 235, 215)
                            pdfTable.AddCell(cell)
                        Next
                        'Adding DataRow
                        Dim cellvalue As String = ""
                        Dim i As Integer = 0
                        For Each row As DataGridViewRow In DataGridView9.Rows
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
                            Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox4.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox38.Text & "_Oferte.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
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
                    System.Diagnostics.Process.Start(Konfigurime.TextBox4.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox38.Text & "_Oferte.pdf")
                Catch ex As Exception
                    MsgBox("Dokumenti eshte i hapur.Mbyll dokumentin dhe provo perseri!", MsgBoxStyle.Information)
                End Try
            End If
        End If
    End Sub
    'Leximi i te dhenave nga database(tabelat "dhenat,Shitjet") ne datagridview1 & datagridview2
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.SetStyle(ControlStyles.UserPaint, True)
        Konfigurime.TextBox1.Text = My.Settings.backup
        Konfigurime.TextBox2.Text = My.Settings.logo
        Konfigurime.TextBox3.Text = My.Settings.ruajraportet
        Konfigurime.TextBox4.Text = My.Settings.faturatofert
        '   Me.DataGridView1.RowHeadersVisible = False
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd/MM/yyyy"
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        DateTimePicker2.CustomFormat = "dd/MM/yyyy"
        DateTimePicker3.Format = DateTimePickerFormat.Custom
        DateTimePicker3.CustomFormat = "dd/MM/yyyy"
        DateTimePicker4.Format = DateTimePickerFormat.Custom
        DateTimePicker4.CustomFormat = "dd/MM/yyyy"
        DateTimePicker5.Format = DateTimePickerFormat.Custom
        DateTimePicker5.CustomFormat = "dd/MM/yyyy"
        DateTimePicker6.Format = DateTimePickerFormat.Custom
        DateTimePicker6.CustomFormat = "dd/MM/yyyy"
        ' Me.TabControl1.Size = New Size(1122, 540)
        'Me.Width = 1125
        ' Me.Height = 578
        'Lidhje datagrid1
        With DataGridView1
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual2 = 0
        lidhje2.Open()
        adaptori2 = New OleDbDataAdapter("SELECT * FROM Dhenat ORDER BY ID", lidhje2)
        adaptori2.Fill(setitedhenave2, "tedhena")
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("count1", "Nr.")
        DataGridView1.DataSource = setitedhenave2.Tables(0)
        DataGridView1.Columns("ID").Visible = False
        nr1 = DataGridView1.Rows.Count - 1
        For i = 0 To nr1
            DataGridView1.Rows(i).Cells(0).Value = i + 1
        Next
        If setitedhenave2.Tables(0).Rows.Count > 0 Then
            Merrtedhenat(rreshtiaktual2)
        Else

        End If
        lidhje2.Close()
        Dim bg2 As System.ComponentModel.BackgroundWorker
        bg2 = New System.ComponentModel.BackgroundWorker
        AddHandler bg2.DoWork, AddressOf dtg2
        bg2.RunWorkerAsync()
        Dim bgw3 As System.ComponentModel.BackgroundWorker
        bgw3 = New System.ComponentModel.BackgroundWorker
        AddHandler bgw3.DoWork, AddressOf dtg3
        bgw3.RunWorkerAsync()
        Dim bgw4 As System.ComponentModel.BackgroundWorker
        bgw4 = New System.ComponentModel.BackgroundWorker
        AddHandler bgw4.DoWork, AddressOf dtg4
        bgw4.RunWorkerAsync()
        Dim bgw5 As System.ComponentModel.BackgroundWorker
        bgw5 = New System.ComponentModel.BackgroundWorker
        AddHandler bgw5.DoWork, AddressOf dtg5
        bgw5.RunWorkerAsync()
        Dim bgw6 As System.ComponentModel.BackgroundWorker
        bgw6 = New System.ComponentModel.BackgroundWorker
        AddHandler bgw6.DoWork, AddressOf dtg6
        bgw6.RunWorkerAsync()
        Dim bgw7 As System.ComponentModel.BackgroundWorker
        bgw7 = New System.ComponentModel.BackgroundWorker
        AddHandler bgw7.DoWork, AddressOf dtg7
        bgw7.RunWorkerAsync()
        If DataGridView1.Rows.Count = 0 Then
            Button2.Enabled = False
            Button3.Enabled = False
        Else
            Button2.Enabled = True
            Button3.Enabled = True
        End If

    End Sub
    Private Sub dtg2()
        With DataGridView2
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual2shitjet = 0
        lidhje2ypp.Open()
        setitedhenave2ypp.clear
        adaptori2ypp = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", lidhje2ypp)
        adaptori2ypp.Fill(setitedhenave2ypp, "tedhena")
        DataGridView2.Columns.Clear()
        DataGridView2.Columns.Add("count1", "Nr.")
        DataGridView2.DataSource = setitedhenave2ypp.Tables(0)
        DataGridView2.Columns("ID").Visible = False
        'Merrtedhenat_Shitjet(rreshtiaktual2shitjet)
        lidhje2ypp.Close()
        Dim nr2 As Integer = 0
        nr2 = DataGridView2.Rows.Count - 1
        For i = 0 To nr2
            DataGridView2.Rows(i).Cells(0).Value = i + 1
        Next
        lidhje2ypp.Close()
        If setitedhenave2ypp.Tables(0).Rows.Count > 0 Then
            Merrtedhenat_Shitjet(rreshtiaktual2shitjet)
            Button20.Text = "Ruaj Faturen(" & setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Kodi_Shitjes") & ")"
            Button21.Text = "Printo Faturen(" & setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Kodi_Shitjes") & ")"
        Else

        End If
        If Not DataGridView2.Rows.Count > 0 Then
            Button4.Enabled = False
            Button6.Enabled = False
        Else
            Button4.Enabled = True
            Button6.Enabled = True
        End If
    End Sub
    Private Sub dtg3()
        With DataGridView3
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktualdg3 = 0
        lidhjedg3.Open()
        setitedhenavedg3.clear
        adaptoridg3 = New OleDbDataAdapter("SELECT * FROM Dhenat ORDER BY ID", lidhjedg3)
        adaptoridg3.Fill(setitedhenavedg3, "tedhena")
        DataGridView3.Columns.Clear()
        DataGridView3.Columns.Add("count1", "Nr.")
        DataGridView3.DataSource = setitedhenavedg3.Tables(0)
        DataGridView3.Columns("ID").Visible = False
        lidhjedg3.Close()
    End Sub
    Private Sub dtg4()
        With DataGridView4
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktualdg4 = 0
        lidhjedg4.Open()
        setitedhenavedg4.clear
        adaptoridg4 = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", lidhjedg4)
        adaptoridg4.Fill(setitedhenavedg4, "tedhena")
        DataGridView4.Columns.Clear()
        DataGridView4.Columns.Add("count1", "Nr.")
        DataGridView4.DataSource = setitedhenavedg4.Tables(0)
        DataGridView4.Columns("ID").Visible = False
        lidhjedg4.Close()
    End Sub
    Private Sub dtg5()
        With DataGridView5
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual2 = 0
        Lidhje1bc.Open()
        adaptor1bc = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", Lidhje1bc)
        adaptor1bc.Fill(setidatave1bc, "tedhena")
        DataGridView5.Columns.Clear()
        DataGridView5.Columns.Add("count1", "Nr.")
        DataGridView5.DataSource = setidatave1bc.Tables(0)
        DataGridView5.Columns("ID").Visible = False
        Lidhje1bc.Close()
        If DataGridView5.Rows.Count = 0 Then
            Button7.Enabled = False
            Button8.Enabled = False
            Button9.Enabled = False
            CheckBox1.Enabled = False
        Else
            Button7.Enabled = True
            Button8.Enabled = True
            Button9.Enabled = True
            CheckBox1.Enabled = True
        End If
    End Sub
    Private Sub dtg6()
        With DataGridView6
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual2 = 0
        Lidhje1bcx.Open()
        adaptor1bcx = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", Lidhje1bcx)
        adaptor1bcx.Fill(setidatave1bcx, "tedhena")
        DataGridView6.Columns.Clear()
        DataGridView6.Columns.Add("count1", "Nr.")
        DataGridView6.DataSource = setidatave1bcx.Tables(0)
        DataGridView6.Columns("ID").Visible = False
        Lidhje1bcx.Close()
        If DataGridView6.Rows.Count = 0 Then
            Button10.Enabled = False
            Button12.Enabled = False
            Button14.Enabled = False
            CheckBox2.Enabled = False
        Else
            Button10.Enabled = True
            Button12.Enabled = True
            Button14.Enabled = True
            CheckBox2.Enabled = True
        End If
    End Sub
    Private Sub dtg7()
        With DataGridView7
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktualdg71 = 0
        lidhjedg71.Open()
        adaptoridg71 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhjedg71)
        adaptoridg71.Fill(setitedhenavedg71, "tedhena")
        DataGridView7.Columns.Clear()
        DataGridView7.Columns.Add("count1", "Nr.")
        DataGridView7.DataSource = setitedhenavedg71.Tables(0)
        DataGridView7.Columns("ID").Visible = False
        Dim nr7 As Integer
        nr7 = DataGridView7.Rows.Count - 1
        For i = 0 To nr7
            DataGridView7.Rows(i).Cells(0).Value = i + 1
        Next
        If setitedhenavedg71.Tables(0).Rows.Count > 0 Then
            Merrtedhenat_Ofertat(rreshtiaktualdg71)
            Button23.Text = "Ruaj Faturen(" & setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Kodi_Ofertes") & ")"
            Button22.Text = "Printo Faturen(" & setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Kodi_Ofertes") & ")"
        Else
        End If
        lidhjedg71.Close()
        If DataGridView7.Rows.Count = 0 Then
            Button18.Enabled = False
            Button19.Enabled = False
        Else
            Button18.Enabled = True
            Button19.Enabled = True
        End If
    End Sub
    Private Sub Farmacia_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        lidhje2.Close()
        Lidhje1.Close()
        Hyrje.Close()
        Try
            System.IO.File.Copy(AppDomain.CurrentDomain.BaseDirectory & "tedhena.accdb", Konfigurime.TextBox1.Text, True)
            File.SetAttributes(Konfigurime.TextBox1.Text, FileAttributes.Hidden)
        Catch ex As Exception
            MsgBox("Procesi i ruajtjes se kopjes se databazes nuk funksionoi!", MsgBoxStyle.Information)
        End Try
    End Sub
    'Form1
    'Meren te dhenat e tabales "dhenat" neper textbox-e me ane te rreshtit te selektuar(i cili si fillim eshte 0 pra rreshti pare)
    Public Sub Merrtedhenat_Shitjet(ByVal rreshtiaktual2shitjet)
        Try
            TextBox7.Text = setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("ID")
            TextBox8.Text = setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Kodi_Shitjes")
            TextBox3.Text = setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Shitesi")
            TextBox4.Text = setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Bleresi")
            TextBox9.Text = setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Produkti")
            TextBox10.Text = setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Njesia")
            TextBox12.Text = setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Sasia")
            TextBox14.Text = setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Cmimi")
            TextBox15.Text = setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Vlera_Pa_TVSH")
            TextBox16.Text = setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("TVSH")
            TextBox11.Text = setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Vlera_Me_TVSH")
            TextBox17.Text = setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Zbritje_ne_perq")
        Catch ex As Exception
        End Try
    End Sub
    Public Sub Merrtedhenat_Ofertat(ByVal rreshtiaktualdg71)
        Try
            TextBox39.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("ID")
            TextBox38.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Kodi_Ofertes")
            TextBox36.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Shitesi")
            TextBox35.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Bleresi")
            TextBox40.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Klient_Ekzistues")
            TextBox34.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Produkti")
            TextBox33.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Njesia")
            TextBox32.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Sasia")
            TextBox31.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Cmimi")
            TextBox30.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Vlera_Pa_TVSH")
            TextBox29.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("TVSH")
            TextBox37.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Vlera_Me_TVSH")
            TextBox28.Text = setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Zbritje_ne_perq")
        Catch ex As Exception
        End Try
    End Sub
    'Klik ne Datagridview dhe meren te dhenat neper textbox ne tabelen "Shitjet"
    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If (e.RowIndex >= 0) Then
            Dim row As DataGridViewRow = DataGridView2.Rows.Item(e.RowIndex)
            TextBox7.Text = row.Cells.Item("ID").Value.ToString
            TextBox8.Text = row.Cells.Item("Kodi_Shitjes").Value.ToString
            TextBox3.Text = row.Cells.Item("Shitesi").Value.ToString
            TextBox4.Text = row.Cells.Item("Bleresi").Value.ToString
            TextBox9.Text = row.Cells.Item("Produkti").Value.ToString
            TextBox10.Text = row.Cells.Item("Njesia").Value.ToString
            TextBox12.Text = row.Cells.Item("Sasia").Value.ToString
            TextBox14.Text = row.Cells.Item("Cmimi").Value.ToString
            TextBox15.Text = row.Cells.Item("Vlera_Pa_TVSH").Value.ToString
            TextBox16.Text = row.Cells.Item("TVSH").Value.ToString
            TextBox11.Text = row.Cells.Item("Vlera_Me_TVSH").Value.ToString
            TextBox17.Text = row.Cells.Item("Zbritje_ne_perq").Value.ToString
            Button20.Text = "Ruaj Faturen(" & row.Cells.Item("Kodi_Shitjes").Value.ToString & ")"
            Button21.Text = "Printo Faturen(" & row.Cells.Item("Kodi_Shitjes").Value.ToString & ")"
        End If
    End Sub
    'Rifreskimi mbas editimit te nje prej qelizave ne tabelen "dhenat"
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Str As String
        Str = "update dhenat set Fornitori="
        Str += """" & TextBox19.Text & """"
        Str += " where ID="
        Str += TextBox1.Text.Trim()
        lidhje2.Open()
        query2 = New OleDbCommand(Str, lidhje2)
        query2.ExecuteNonQuery()
        lidhje2.Close()

        lidhje2.Open()
        Str = "update dhenat set Bleresi="
        Str += """" & TextBox20.Text & """"
        Str += " where ID="
        Str += TextBox1.Text.Trim()
        query2 = New OleDbCommand(Str, lidhje2)
        query2.ExecuteNonQuery()
        lidhje2.Close()

        lidhje2.Open()
        Str = "update dhenat set Produkti="
        Str += """" & TextBox26.Text & """"
        Str += " where ID="
        Str += TextBox1.Text.Trim()
        query2 = New OleDbCommand(Str, lidhje2)
        query2.ExecuteNonQuery()
        lidhje2.Close()

        lidhje2.Open()
        Str = "update dhenat set Njesia="
        Str += """" & TextBox5.Text & """"
        Str += " where ID="
        Str += TextBox1.Text.Trim()
        query2 = New OleDbCommand(Str, lidhje2)
        query2.ExecuteNonQuery()
        lidhje2.Close()

        lidhje2.Open()
        Str = "update dhenat set Sasia="
        Str += """" & TextBox21.Text & """"
        Str += " where ID="
        Str += TextBox1.Text.Trim()
        query2 = New OleDbCommand(Str, lidhje2)
        query2.ExecuteNonQuery()
        lidhje2.Close()

        lidhje2.Open()
        Str = "update dhenat set Cmimi="
        Str += """" & TextBox22.Text & """"
        Str += " where ID="
        Str += TextBox1.Text.Trim()
        query2 = New OleDbCommand(Str, lidhje2)
        query2.ExecuteNonQuery()
        lidhje2.Close()

        lidhje2.Open()
        Str = "update dhenat set Vlera_Pa_TVSH="
        Str += """" & TextBox23.Text & """"
        Str += " where ID="
        Str += TextBox1.Text.Trim()
        query2 = New OleDbCommand(Str, lidhje2)
        query2.ExecuteNonQuery()
        lidhje2.Close()

        lidhje2.Open()
        Str = "update dhenat set TVSH="
        Str += """" & TextBox24.Text & """"
        Str += " where ID="
        Str += TextBox1.Text.Trim()
        query2 = New OleDbCommand(Str, lidhje2)
        query2.ExecuteNonQuery()
        lidhje2.Close()

        lidhje2.Open()
        Str = "update dhenat set Vlera_Me_TVSH="
        Str += """" & TextBox25.Text & """"
        Str += " where ID="
        Str += TextBox1.Text.Trim()
        query2 = New OleDbCommand(Str, lidhje2)
        query2.ExecuteNonQuery()
        lidhje2.Close()

        lidhje2114.Open()
        setitedhenave2114.Clear()
        adaptori2114 = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", lidhje2114)
        adaptori2114.Fill(setitedhenave2114, "tedhena")
        DataGridView1.DataSource = setitedhenave2114.Tables(0)
        Dim nr1 As Integer = 0
        nr1 = DataGridView1.Rows.Count - 1
        For i = 0 To nr1
            DataGridView1.Rows(i).Cells(0).Value = i + 1
        Next
        lidhje2114.Close()
        MsgBox("U rifreskua me sukses", MsgBoxStyle.Information)
    End Sub
    'Fshirja e nje rreshti nga tabale "dhenat"
    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If RadioButton1.Checked = True Then
            For Each row As DataGridViewRow In DataGridView1.Rows
                lidhje2.Close()
                Dim Str As String
                Try
                    Str = "delete from dhenat where ID="
                    Str += row.Cells(1).Value.ToString
                    lidhje2.Open()
                    query2 = New OleDbCommand(Str, lidhje2)
                    query2.ExecuteNonQuery()
                    setitedhenave2.clear()
                    adaptori2 = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", lidhje2)
                    adaptori2.Fill(setitedhenave2, "tedhena")
                    If rreshtiaktual2 > 0 Then
                        rreshtiaktual2 -= 1
                        Merrtedhenat(rreshtiaktual2)
                    End If
                    lidhje2.Close()
                    TextBox6.Text = ""
                    MsgBox("U fshi me sukses", MsgBoxStyle.Information)
                Catch ex As Exception
                    MsgBox("Nuk u fshi me sukses", MsgBoxStyle.Information)

                    lidhje2.Close()
                End Try
            Next
        ElseIf RadioButton2.Checked = True Then
            Dim Str As String
            Try
                Str = "delete from dhenat where ID="
                Str += DataGridView1.CurrentRow.Cells(1).Value.ToString
                lidhje2.Open()
                query2 = New OleDbCommand(Str, lidhje2)
                query2.ExecuteNonQuery()
                setitedhenave2.clear()
                adaptori2 = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", lidhje2)
                adaptori2.Fill(setitedhenave2, "tedhena")
                nr1 = DataGridView1.Rows.Count - 1
                For i = 0 To nr1
                    DataGridView1.Rows(i).Cells(0).Value = i + 1
                Next
                If rreshtiaktual2 > 0 Then
                    rreshtiaktual2 -= 1
                    Merrtedhenat(rreshtiaktual2)
                End If
                lidhje2.Close()
                TextBox6.Text = ""
                MsgBox("U fshi me sukses", MsgBoxStyle.Information)
            Catch ex As Exception
                MsgBox("Nuk u fshi me sukses", MsgBoxStyle.Information)
                lidhje2.Close()
            End Try
        End If
        If DataGridView1.Rows.Count = 0 Then
            Button2.Enabled = False
            Button3.Enabled = False
        Else
            Button2.Enabled = True
            Button3.Enabled = True
        End If
    End Sub
    'Kerkimi sipas te gjitha kolonave(ID,Date,Emer,etj ne tabelen "dhenat")
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        lidhje2.Close()
        lidhje2.Open()
        adaptori2 = New OleDbDataAdapter("SELECT ID,Kodi_Blerjes,Produkti,Fornitori,Bleresi,Njesia,Sasia,Cmimi,Vlera_Pa_TVSH,TVSH,Vlera_me_TVSH FROM dhenat WHERE(ID Like '" &
        TextBox6.Text & "%' OR Data Like '" & TextBox6.Text & "%' OR Ora Like '" &
                                             TextBox6.Text & "%' OR  Kodi_Blerjes Like '" & TextBox6.Text & "%' OR Produkti Like '" &
                                             TextBox6.Text & "%' OR Fornitori Like '" & TextBox6.Text & "%' OR Bleresi Like '" &
                                             TextBox6.Text & "%' OR Njesia Like '" & TextBox6.Text & "%' OR Sasia Like '" &
                                             TextBox6.Text & "%' OR Cmimi Like '" & TextBox6.Text & "%' OR Vlera_Pa_TVSH Like '" &
                                             TextBox6.Text & "%' OR TVSH Like '" & TextBox6.Text & "%' OR Vlera_me_TVSH Like '" &
                                              TextBox6.Text & "%')", Lidhje.lidhje2)
        Dim dataTable As New DataTable
        adaptori2.Fill(dataTable)
        DataGridView1.DataSource = dataTable
        nr1 = DataGridView1.Rows.Count - 1
        For i = 0 To nr1
            DataGridView1.Rows(i).Cells(0).Value = i + 1
        Next
        Merrtedhenat(rreshtiaktual2)
        DataGridView1.CurrentCell = Nothing
        If DataGridView1.CurrentCell Is Nothing Then

        Else
            Dim row As DataGridViewRow = Me.DataGridView1.Rows.Item(rreshtiaktual2)
            TextBox1.Text = row.Cells.Item("ID").Value.ToString
            TextBox18.Text = row.Cells.Item("Kodi_Blerjes").Value.ToString
            TextBox19.Text = row.Cells.Item("Produkti").Value.ToString
            TextBox20.Text = row.Cells.Item("Fornitori").Value.ToString
            TextBox26.Text = row.Cells.Item("Bleresi").Value.ToString
            TextBox5.Text = row.Cells.Item("Njesia").Value.ToString
            TextBox21.Text = row.Cells.Item("Sasia").Value.ToString
            TextBox22.Text = row.Cells.Item("Cmimi").Value.ToString
            TextBox23.Text = row.Cells.Item("Vlera_Pa_TVSH").Value.ToString
            TextBox24.Text = row.Cells.Item("TVSH").Value.ToString
            TextBox25.Text = row.Cells.Item("Vlera_me_TVSH").Value.ToString
            lidhje2.Close()
            DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            DataGridView1.RefreshEdit()
            lidhje2.Close()
        End If
        If TextBox6.Text = "" Then
            Label108.Text = 0
            Label109.Text = 0
            Label110.Text = 0
            Label108.ForeColor = DefaultForeColor
            Label109.ForeColor = DefaultForeColor
            Label110.ForeColor = DefaultForeColor
        Else
            Dim vl1, vl2, vl3 As String
            vl1 = (From row As DataGridViewRow In DataGridView1.Rows
                   Where row.Cells(9).FormattedValue.ToString() <> String.Empty
                   Select Convert.ToInt32(row.Cells(9).FormattedValue)).Sum().ToString()
            vl2 = (From row As DataGridViewRow In DataGridView1.Rows
                   Where row.Cells(10).FormattedValue.ToString() <> String.Empty
                   Select Convert.ToInt32(row.Cells(10).FormattedValue)).Sum().ToString()
            vl3 = (From row As DataGridViewRow In DataGridView1.Rows
                   Where row.Cells(11).FormattedValue.ToString() <> String.Empty
                   Select Convert.ToInt32(row.Cells(11).FormattedValue)).Sum().ToString()
            Label108.ForeColor = Color.Red
            Label109.ForeColor = Color.Orange
            Label110.ForeColor = Color.ForestGreen
            Label108.Text = vl1
            Label109.Text = vl2
            Label110.Text = vl3
        End If
    End Sub
    'Meren te dhenat e tabales "Shitjet" neper textbox-e me ane te rreshtit te selektuar(i cili si fillim eshte 0 pra rreshti pare)
    Private Sub Merrtedhenat(ByVal rreshtiaktual2)
        Try
            TextBox1.Text = setitedhenave2.Tables("tedhena").Rows(rreshtiaktual2)("ID")
            TextBox18.Text = setitedhenave2.Tables("tedhena").Rows(rreshtiaktual2)("Kodi_Blerjes")
            TextBox19.Text = setitedhenave2.Tables("tedhena").Rows(rreshtiaktual2)("Produkti")
            TextBox20.Text = setitedhenave2.Tables("tedhena").Rows(rreshtiaktual2)("Fornitori")
            TextBox26.Text = setitedhenave2.Tables("tedhena").Rows(rreshtiaktual2)("Bleresi")
            TextBox5.Text = setitedhenave2.Tables("tedhena").Rows(rreshtiaktual2)("Njesia")
            TextBox21.Text = setitedhenave2.Tables("tedhena").Rows(rreshtiaktual2)("Sasia")
            TextBox22.Text = setitedhenave2.Tables("tedhena").Rows(rreshtiaktual2)("Cmimi")
            TextBox23.Text = setitedhenave2.Tables("tedhena").Rows(rreshtiaktual2)("Vlera_Pa_TVSH")
            TextBox24.Text = setitedhenave2.Tables("tedhena").Rows(rreshtiaktual2)("TVSH")
            TextBox25.Text = setitedhenave2.Tables("tedhena").Rows(rreshtiaktual2)("Vlera_me_TVSH")
        Catch ex As Exception
        End Try
    End Sub
    'Klik ne Datagridview dhe meren te dhenat neper textbox ne tabelen "dhenat"
    Private Sub DataGridView1_CellLidhje2tentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If (e.RowIndex >= 0) Then
            Dim row As DataGridViewRow = DataGridView1.Rows.Item(e.RowIndex)
            TextBox1.Text = row.Cells.Item("ID").Value.ToString
            TextBox18.Text = row.Cells.Item("Kodi_Blerjes").Value.ToString
            TextBox19.Text = row.Cells.Item("Produkti").Value.ToString
            TextBox20.Text = row.Cells.Item("Fornitori").Value.ToString
            TextBox26.Text = row.Cells.Item("Bleresi").Value.ToString
            TextBox5.Text = row.Cells.Item("Njesia").Value.ToString
            TextBox21.Text = row.Cells.Item("Sasia").Value.ToString
            TextBox22.Text = row.Cells.Item("Cmimi").Value.ToString
            TextBox23.Text = row.Cells.Item("Vlera_Pa_TVSH").Value.ToString
            TextBox24.Text = row.Cells.Item("TVSH").Value.ToString
            TextBox25.Text = row.Cells.Item("Vlera_me_TVSH").Value.ToString
        End If
    End Sub
    'Shtimi i nje rreshti te ri ne tabelen "Shitjet"
    'Rifreskimi mbas editimit te nje prej qelizave ne tabelen "Shitjet"
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Not DataGridView2.Rows.Count > 0 Then
            MsgBox("Lista eshte bosh!", MsgBoxStyle.Information)
        Else
            Dim Str As String
            Str = "update Shitjet set Shitesi="
            Str += """" & TextBox3.Text & """"
            Str += " where ID="
            Str += TextBox7.Text.Trim()
            lidhje2115.Open()
            query2115 = New OleDbCommand(Str, lidhje2115)
            query2115.ExecuteNonQuery()
            lidhje2115.Close()
            lidhje2115.Open()
            Str = "update Shitjet set Bleresi="
            Str += """" & TextBox4.Text & """"
            Str += " where ID="
            Str += TextBox7.Text.Trim()
            query2115 = New OleDbCommand(Str, lidhje2115)
            query2115.ExecuteNonQuery()
            lidhje2115.Close()
            lidhje2115.Open()
            Str = "update Shitjet set Produkti="
            Str += """" & TextBox9.Text & """"
            Str += " where ID="
            Str += TextBox7.Text.Trim()
            query2115 = New OleDbCommand(Str, lidhje2115)
            query2115.ExecuteNonQuery()
            lidhje2115.Close()
            lidhje2115.Open()
            Str = "update Shitjet set Njesia="
            Str += """" & TextBox10.Text & """"
            Str += " where ID="
            Str += TextBox7.Text.Trim()
            query2115 = New OleDbCommand(Str, lidhje2115)
            query2115.ExecuteNonQuery()
            lidhje2115.Close()
            lidhje2115.Open()
            Str = "update Shitjet set Sasia="
            Str += """" & TextBox12.Text & """"
            Str += " where ID="
            Str += TextBox7.Text.Trim()
            query2115 = New OleDbCommand(Str, lidhje2115)
            query2115.ExecuteNonQuery()
            lidhje2115.Close()
            lidhje2115.Open()
            Str = "update Shitjet set Cmimi="
            Str += """" & TextBox14.Text & """"
            Str += " where ID="
            Str += TextBox7.Text.Trim()
            query2115 = New OleDbCommand(Str, lidhje2115)
            query2115.ExecuteNonQuery()
            lidhje2115.Close()
            lidhje2115.Open()
            Str = "update Shitjet set Vlera_Pa_TVSH="
            Str += """" & TextBox15.Text & """"
            Str += " where ID="
            Str += TextBox7.Text.Trim()
            query2115 = New OleDbCommand(Str, lidhje2115)
            query2115.ExecuteNonQuery()
            lidhje2115.Close()
            lidhje2115.Open()
            Str = "update Shitjet set TVSH="
            Str += """" & TextBox16.Text & """"
            Str += " where ID="
            Str += TextBox7.Text.Trim()
            query2115 = New OleDbCommand(Str, lidhje2115)
            query2115.ExecuteNonQuery()
            lidhje2115.Close()
            lidhje2115.Open()
            Str = "update Shitjet set Vlera_Me_TVSH="
            Str += """" & TextBox11.Text & """"
            Str += " where ID="
            Str += TextBox7.Text.Trim()
            query2115 = New OleDbCommand(Str, lidhje2115)
            query2115.ExecuteNonQuery()
            lidhje2115.Close()
            lidhje2115.Open()
            Str = "update Shitjet set Zbritje_ne_perq="
            Str += """" & TextBox17.Text & """"
            Str += " where ID="
            Str += TextBox7.Text.Trim()
            query2115 = New OleDbCommand(Str, lidhje2115)
            query2115.ExecuteNonQuery()
            lidhje2115.Close()
            setitedhenave2115.Clear()
            lidhje2115.Open()
            adaptori2115 = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", lidhje2115)
            adaptori2115.Fill(setitedhenave2115, "tedhena")
            DataGridView2.DataSource = setitedhenave2115.Tables(0)
            DataGridView2.Columns("ID").Visible = False
            nr2 = DataGridView1.Rows.Count - 1
            For i = 0 To nr2
                DataGridView2.Rows(i).Cells(0).Value = i + 1
            Next
            lidhje2115.Close()
            MsgBox("U rifreskua me sukses", MsgBoxStyle.Information)
        End If
    End Sub
    'Fshirja e nje rreshti nga database tek tabela "dhenat"
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim Str As String
        Try
            Str = "delete from dhenat where ID="
            Str += TextBox1.Text.Trim
            lidhje2.Open()
            query2 = New OleDbCommand(Str, lidhje2)
            query2.ExecuteNonQuery()
            setitedhenave2.clear()
            adaptori2 = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", lidhje2)
            adaptori2.Fill(setitedhenave2, "tedhena")
            MsgBox("U fshi me sukses", MsgBoxStyle.Information)
            If rreshtiaktual2 > 0 Then
                rreshtiaktual2 -= 1
                Merrtedhenat(rreshtiaktual2)
            End If
            lidhje2.Close()
        Catch ex As Exception
            MsgBox("Nuk u fshi me sukses", MsgBoxStyle.Information)
            lidhje2.Close()
        End Try
    End Sub
    'Kerkimi sipas te gjitha kolonave(ID,Date,Emer,etj ne tabelen "Shitjet")
    'Gjenerimi i xhiros ditore
    Private Sub Button7_Click(sender As Object, e As EventArgs)
        DataGridView1.DataSource = setitedhenave2.Tables(0).DefaultView
        DataGridView1.Refresh()
        DataGridView2.DataSource = setidatave1.Tables(0).DefaultView
        DataGridView2.Refresh()
        Dim total1 As Integer = 0
        Dim regDate2 As DateTime = Date.Now
        Dim strDate2 As String = regDate2.ToString("dd/MM/yyyy")
        Lidhje1.Open()
        adaptor1 = New OleDbDataAdapter("SELECT ID, Data,Ora,Shitesi, Emri_Ilacit,Sasia,Vlera,Shuma_Totale,Bleresi FROM Shitjet WHERE(Data Like '" &
                                        strDate2 & "%')", Lidhje1)
        Dim dataTable1 As New DataTable
        adaptor1.Fill(dataTable1)
        DataGridView2.DataSource = dataTable1
        For i As Integer = 0 To DataGridView2.RowCount - 1
            total1 = total1 + DataGridView2.Rows(i).Cells(7).Value
        Next
        Label11.Enabled = True
        Label11.Text = total1
        Lidhje1.Close()
    End Sub
    'Gjenerimi i xhiros mujore
    Private Sub Button8_Click(sender As Object, e As EventArgs)
        DataGridView1.DataSource = setitedhenave2.Tables(0).DefaultView
        DataGridView1.Refresh()
        DataGridView2.DataSource = setidatave1.Tables(0).DefaultView
        DataGridView2.Refresh()
        Dim regDate4 As DateTime = Date.Now
        Dim regDate5 As DateTime = Date.Now
        Dim total2 As Integer = 0
        Dim strDate2 As String = regDate4.ToString("/MM/")
        Dim strDate3 As String = regDate5.ToString("yyyy")
        Lidhje1.Open()
        Dim dataTable1 As New DataTable
        For i As Integer = 1 To 31
            adaptor1 = New OleDbDataAdapter("SELECT ID, Data,Ora,Shitesi, Emri_Ilacit,Sasia,Vlera,Shuma_Totale,Bleresi FROM Shitjet WHERE(Data Like '" &
                                            i.ToString("D2") & strDate2 & strDate3 & "%')", Lidhje1)
            adaptor1.Fill(dataTable1)
            DataGridView2.DataSource = dataTable1
        Next
        For ii As Integer = 0 To DataGridView2.RowCount - 1
            total2 = total2 + DataGridView2.Rows(ii).Cells(7).Value
        Next
        adaptor1.Update(dataTable1)
        Label12.Enabled = True
        Label12.Text = total2
        Lidhje1.Close()
    End Sub
    'Gjenerimi i xhiros vjetore
    Private Sub Button9_Click(sender As Object, e As EventArgs)
        DataGridView1.DataSource = setitedhenave2.Tables(0).DefaultView
        DataGridView1.Refresh()
        DataGridView2.DataSource = setidatave1.Tables(0).DefaultView
        DataGridView2.Refresh()
        Dim regDate2 As DateTime = Date.Now
        Dim regDate3 As DateTime = Date.Now
        Dim total As Integer = 0
        Dim strDate2 As String = regDate2.ToString("/MM/")
        Dim strDate3 As String = regDate2.ToString("yyyy")
        Lidhje1.Open()
        Dim dataTable1 As New DataTable
        For j As Int32 = 1 To 12
            For i As Int32 = 1 To 31
                adaptor1 = New OleDbDataAdapter("SELECT ID, Data,Ora,Shitesi, Emri_Ilacit,Sasia,Vlera,Shuma_Totale,Bleresi FROM Shitjet WHERE(Data Like '" &
                                                i.ToString("D2") & "/" & j.ToString("D2") & "/" & strDate3 & "%')", Lidhje1)
                adaptor1.Fill(dataTable1)
                DataGridView2.DataSource = dataTable1
            Next
        Next
        For ii As Integer = 0 To DataGridView2.RowCount - 1
            total = total + DataGridView2.Rows(ii).Cells(7).Value
        Next
        adaptor1.Update(dataTable1)
        Label13.Enabled = True
        Label13.Text = total
        Lidhje1.Close()
    End Sub
    'Gjenerimi i raportit per gjendjen e sasise se ilaceve
    Dim nr7 As Integer = 0
    'Refreskimi ne gjendje fillestare
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TextBox2.Text = ""
        Lidhje1.Close()
        lidhje2.Close()
        With DataGridView2
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual = 0
        Lidhje1.Open()
        adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", Lidhje1)
        adaptor1.Fill(setidatave1, "tedhena")
        DataGridView2.DataSource = setidatave1.Tables(0)
        nr7 = DataGridView7.Rows.Count - 1
        For i = 0 To nr7
            DataGridView7.Rows(i).Cells(0).Value = i + 1
        Next
        DataGridView2.Refresh()
        Merrtedhenat_Shitjet(rreshtiaktual2shitjet)
        Lidhje1.Close()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Not DataGridView2.Rows.Count > 0 Then
            MsgBox("Lista eshte bosh!", MsgBoxStyle.Information)
        Else
            If RadioButton3.Checked = True Then
                For Each row As DataGridViewRow In DataGridView2.Rows
                    Dim Str As String
                    Try
                        Str = "delete from Shitjet where ID="
                        Str += row.Cells(0).Value
                        lidhje2.Open()
                        query2 = New OleDbCommand(Str, lidhje2)
                        query2.ExecuteNonQuery()
                        setidatave1.clear()
                        adaptori2 = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", lidhje2)
                        adaptori2.Fill(setidatave1, "tedhena")

                        If rreshtiaktual > 0 Then
                            rreshtiaktual -= 1
                            Merrtedhenat(rreshtiaktual)
                        End If
                        lidhje2.Close()
                    Catch ex As Exception
                        MsgBox("Nuk u fshi me sukses", MsgBoxStyle.Information)
                        lidhje2.Close()
                    End Try
                Next
                TextBox2.Text = ""
                MsgBox("U fshi me sukses", MsgBoxStyle.Information)
            ElseIf RadioButton4.Checked = True Then
                Dim Str As String
                Try
                    Str = "delete from Shitjet where ID="
                    Str += DataGridView2.CurrentRow.Cells(1).Value.ToString
                    lidhje2.Open()
                    query2 = New OleDbCommand(Str, lidhje2)
                    query2.ExecuteNonQuery()
                    setidatave1.clear()
                    adaptori2 = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", lidhje2)
                    adaptori2.Fill(setidatave1, "tedhena")
                    DataGridView2.DataSource = setidatave1.tables(0)
                    nr2 = DataGridView2.Rows.Count - 1
                    For i = 0 To nr2
                        DataGridView2.Rows(i).Cells(0).Value = i + 1
                    Next
                    If rreshtiaktual > 0 Then
                        rreshtiaktual -= 1
                        Merrtedhenat(rreshtiaktual)
                    End If
                    lidhje2.Close()
                Catch ex As Exception
                    MsgBox("Nuk u fshi me sukses", MsgBoxStyle.Information)
                    lidhje2.Close()
                End Try
                TextBox2.Text = ""
                MsgBox("U fshi me sukses", MsgBoxStyle.Information)
            End If
        End If
    End Sub
    Private Sub ShtoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShtoToolStripMenuItem.Click
        Me.Hide()
        Administrim.Show()
    End Sub
    Private Sub ShitjeEReToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShitjeEReToolStripMenuItem.Click
        setitedhenave2ypp.clear
        lidhje2ypp.Open()
        adaptori2ypp = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", lidhje2ypp)
        adaptori2ypp.Fill(setitedhenave2ypp, "tedhena")
        lidhje2ypp.Close()
        If setitedhenave2ypp.Tables(0).Rows.Count = 0 Then
            MsgBox("Ju nuk keni asnje hyrje ne databaze!Shitja nuk eshte e mundur!", MsgBoxStyle.Information)
        Else



            Me.Hide()

            Shit.Show()


        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Try
            lidhje2v.Close()
            lidhje2v.Open()
            adaptori2v = New OleDbDataAdapter("SELECT ID	,Kodi_Shitjes	,Data	,Ora	,Shitesi,	Bleresi,	Produkti,	Njesia	,Sasia	,Cmimi	,Vlera_Pa_TVSH,	TVSH,	Vlera_Me_TVSH,	Zbritje_ne_perq FROM Shitjet WHERE(ID Like '" &
                                                TextBox2.Text & "%' OR Kodi_Shitjes Like '" & TextBox2.Text & "%' OR Data Like '" &
                                                TextBox2.Text & "%' OR Ora Like '" & TextBox2.Text & "%'  OR Shitesi Like '" &
                                                TextBox2.Text & "%' OR Bleresi Like '" & TextBox2.Text & "%' OR Produkti Like '" &
                                                TextBox2.Text & "%' OR Njesia Like '" & TextBox2.Text & "%' OR Sasia Like '" &
                                                TextBox2.Text & "%' OR Cmimi Like '" & TextBox2.Text & "%' OR Vlera_Pa_TVSH Like '" &
                                                 TextBox2.Text & "%' OR TVSH Like '" & TextBox2.Text & "%' OR Vlera_Me_TVSH Like '" & TextBox2.Text & "%' )", Lidhje.lidhje2v)
            Dim dataTablev As New DataTable
            adaptori2v.Fill(dataTablev)
            DataGridView2.DataSource = dataTablev
            nr2 = DataGridView2.Rows.Count - 1
            For i = 0 To nr2
                DataGridView2.Rows(i).Cells(0).Value = i + 1
            Next
            Merrtedhenat_Shitjet(rreshtiaktual2shitjet)
            DataGridView2.CurrentCell = Nothing
            If DataGridView2.CurrentCell Is Nothing Then

            Else
                Dim row As DataGridViewRow = Me.DataGridView2.Rows.Item(rreshtiaktual2shitjet)
                TextBox7.Text = row.Cells.Item("ID").Value.ToString
                TextBox8.Text = row.Cells.Item("Kodi_Shitjes").Value.ToString
                TextBox3.Text = row.Cells.Item("Shitesi").Value.ToString
                TextBox4.Text = row.Cells.Item("Bleresi").Value.ToString
                TextBox9.Text = row.Cells.Item("Produkti").Value.ToString
                TextBox10.Text = row.Cells.Item("Njesia").Value.ToString
                TextBox12.Text = row.Cells.Item("Sasia").Value.ToString
                TextBox14.Text = row.Cells.Item("Cmimi").Value.ToString
                TextBox15.Text = row.Cells.Item("Vlera_Pa_TVSH").Value.ToString
                TextBox16.Text = row.Cells.Item("TVSH").Value.ToString
                TextBox11.Text = row.Cells.Item("Vlera_me_TVSH").Value.ToString
                TextBox17.Text = row.Cells.Item("Zbritje_ne_perq").Value.ToString
                lidhje2v.Close()
                DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
                DataGridView2.RefreshEdit()
                lidhje2v.Close()
            End If
            If TextBox2.Text = "" Then
                Label103.Text = 0
                Label104.Text = 0
                Label105.Text = 0
                Label103.ForeColor = DefaultForeColor
                Label104.ForeColor = DefaultForeColor
                Label105.ForeColor = DefaultForeColor
            Else
                Dim vl1, vl2, vl3 As String
                vl1 = (From row As DataGridViewRow In DataGridView2.Rows
                       Where row.Cells(11).FormattedValue.ToString() <> String.Empty
                       Select Convert.ToInt32(row.Cells(11).FormattedValue)).Sum().ToString()
                vl2 = (From row As DataGridViewRow In DataGridView2.Rows
                       Where row.Cells(12).FormattedValue.ToString() <> String.Empty
                       Select Convert.ToInt32(row.Cells(12).FormattedValue)).Sum().ToString()
                vl3 = (From row As DataGridViewRow In DataGridView2.Rows
                       Where row.Cells(13).FormattedValue.ToString() <> String.Empty
                       Select Convert.ToInt32(row.Cells(13).FormattedValue)).Sum().ToString()
                Label103.ForeColor = Color.Red
                Label104.ForeColor = Color.Orange
                Label105.ForeColor = Color.ForestGreen
                Label103.Text = vl1
                Label104.Text = vl2
                Label105.Text = vl3
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub BlerjeEReHyrjeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BlerjeEReHyrjeToolStripMenuItem.Click
        Me.Hide()
        Bli.Show()
    End Sub
    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.TabPages.Count = 2 Then
            Select Case TabControl1.SelectedIndex
                Case 0
                    ' Me.TabControl1.Size = New Size(1122, 540)
                    '  Me.Width = 1125
                    '  Me.Height = 578
                    With DataGridView2
                        .RowsDefaultCellStyle.BackColor = Color.Bisque
                        .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                    End With
                    DataGridView2.DataSource = Nothing
                    rreshtiaktual2shitjet = 0
                    lidhje2ypp.Open()
                    setitedhenave2ypp.clear
                    adaptori2ypp = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", lidhje2ypp)
                    adaptori2ypp.Fill(setitedhenave2ypp, "tedhena")
                    DataGridView2.Columns.Clear()
                    DataGridView2.Columns.Add("count1", "Nr.")
                    DataGridView2.DataSource = setitedhenave2ypp.Tables(0)
                    DataGridView2.Columns("ID").Visible = False
                    'Merrtedhenat_Shitjet(rreshtiaktual2shitjet)
                    lidhje2ypp.Close()
                    Dim nr2 As Integer = 0
                    nr2 = DataGridView2.Rows.Count - 1
                    For i = 0 To nr2
                        DataGridView2.Rows(i).Cells(0).Value = i + 1
                    Next
                    lidhje2ypp.Close()
                    If setitedhenave2ypp.Tables(0).Rows.Count > 0 Then
                        Merrtedhenat_Shitjet(rreshtiaktual2shitjet)
                        Button20.Text = "Ruaj Faturen(" & setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Kodi_Shitjes") & ")"
                        Button21.Text = "Printo Faturen(" & setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Kodi_Shitjes") & ")"
                    Else

                    End If
                    If Not DataGridView2.Rows.Count > 0 Then
                        Button4.Enabled = False
                        Button6.Enabled = False
                    Else
                        Button4.Enabled = True
                        Button6.Enabled = True
                    End If
                Case 1
                    '  Me.TabControl1.Size = New Size(1245, 616)
                    '  Me.Width = 1260
                    '  Me.Height = 654
                    With DataGridView7
                        .RowsDefaultCellStyle.BackColor = Color.Bisque
                        .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                    End With
                    DataGridView7.DataSource = Nothing
                    rreshtiaktualdg71 = 0
                    lidhjedg71.Open()
                    adaptoridg71 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhjedg71)
                    adaptoridg71.Fill(setitedhenavedg71, "tedhena")
                    DataGridView7.Columns.Clear()
                    DataGridView7.Columns.Add("count1", "Nr.")
                    DataGridView7.DataSource = setitedhenavedg71.Tables(0)
                    DataGridView7.Columns("ID").Visible = False
                    Dim nr7 As Integer
                    nr7 = DataGridView7.Rows.Count - 1
                    For i = 0 To nr7
                        DataGridView7.Rows(i).Cells(0).Value = i + 1
                    Next
                    If setitedhenavedg71.Tables(0).Rows.Count > 0 Then
                        Merrtedhenat_Ofertat(rreshtiaktualdg71)
                        Button23.Text = "Ruaj Faturen(" & setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Kodi_Ofertes") & ")"
                        Button22.Text = "Printo Faturen(" & setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Kodi_Ofertes") & ")"
                    Else
                    End If
                    lidhjedg71.Close()
                    If DataGridView7.Rows.Count = 0 Then
                        Button18.Enabled = False
                        Button19.Enabled = False
                    Else
                        Button18.Enabled = True
                        Button19.Enabled = True
                    End If
            End Select
        Else
            Select Case TabControl1.SelectedIndex
                Case 0
                    ' Me.TabControl1.Size = New Size(1122, 540)
                    '  Me.Width = 1125
                    '  Me.Height = 578
                    With DataGridView1
                        .RowsDefaultCellStyle.BackColor = Color.Bisque
                        .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                    End With
                    DataGridView1.DataSource = Nothing
                    rreshtiaktual2 = 0
                    lidhje2.Open()
                    adaptori2 = New OleDbDataAdapter("SELECT * FROM Dhenat ORDER BY ID", lidhje2)
                    adaptori2.Fill(setitedhenave2, "tedhena")
                    DataGridView1.Columns.Clear()
                    DataGridView1.Columns.Add("count1", "Nr.")
                    DataGridView1.DataSource = setitedhenave2.Tables(0)
                    DataGridView1.Columns("ID").Visible = False
                    nr1 = DataGridView1.Rows.Count - 1
                    For i = 0 To nr1
                        DataGridView1.Rows(i).Cells(0).Value = i + 1
                    Next
                    If setitedhenave2.Tables(0).Rows.Count > 0 Then
                        Merrtedhenat(rreshtiaktual2)
                    Else

                    End If
                    lidhje2.Close()
                Case 1
                    With DataGridView2
                        .RowsDefaultCellStyle.BackColor = Color.Bisque
                        .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                    End With
                    DataGridView2.DataSource = Nothing
                    rreshtiaktual2shitjet = 0
                    lidhje2ypp.Open()
                    setitedhenave2ypp.clear
                    adaptori2ypp = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", lidhje2ypp)
                    adaptori2ypp.Fill(setitedhenave2ypp, "tedhena")
                    DataGridView2.Columns.Clear()
                    DataGridView2.Columns.Add("count1", "Nr.")
                    DataGridView2.DataSource = setitedhenave2ypp.Tables(0)
                    DataGridView2.Columns("ID").Visible = False
                    'Merrtedhenat_Shitjet(rreshtiaktual2shitjet)
                    lidhje2ypp.Close()
                    Dim nr2 As Integer = 0
                    nr2 = DataGridView2.Rows.Count - 1
                    For i = 0 To nr2
                        DataGridView2.Rows(i).Cells(0).Value = i + 1
                    Next
                    lidhje2ypp.Close()
                    If setitedhenave2ypp.Tables(0).Rows.Count > 0 Then
                        Merrtedhenat_Shitjet(rreshtiaktual2shitjet)
                        Button20.Text = "Ruaj Faturen(" & setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Kodi_Shitjes") & ")"
                        Button21.Text = "Printo Faturen(" & setitedhenave2ypp.Tables("tedhena").Rows(rreshtiaktual2shitjet)("Kodi_Shitjes") & ")"
                    Else

                    End If
                    If Not DataGridView2.Rows.Count > 0 Then
                        Button4.Enabled = False
                        Button6.Enabled = False
                    Else
                        Button4.Enabled = True
                        Button6.Enabled = True
                    End If
                Case 2
                    '   Me.TabControl1.Size = New Size(1123, 619)
                    '   Me.Width = 1130
                    '   Me.Height = 657
                    With DataGridView6
                        .RowsDefaultCellStyle.BackColor = Color.Bisque
                        .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                    End With
                    DataGridView6.DataSource = Nothing
                    rreshtiaktual2 = 0
                    Lidhje1bcx.Open()
                    adaptor1bcx = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", Lidhje1bcx)
                    adaptor1bcx.Fill(setidatave1bcx, "tedhena")
                    DataGridView6.Columns.Clear()
                    DataGridView6.Columns.Add("count1", "Nr.")
                    DataGridView6.DataSource = setidatave1bcx.Tables(0)
                    DataGridView6.Columns("ID").Visible = False
                    Lidhje1bcx.Close()
                    If DataGridView6.Rows.Count = 0 Then
                        Button10.Enabled = False
                        Button12.Enabled = False
                        Button14.Enabled = False
                        CheckBox2.Enabled = False
                    Else
                        Button10.Enabled = True
                        Button12.Enabled = True
                        Button14.Enabled = True
                        CheckBox2.Enabled = True
                    End If

                    Label85.Text = ""
                    Label79.Text = ""
                    Label77.Text = ""
                    Label86.Text = ""
                    Label73.Text = ""
                    Label72.Text = ""
                    Label87.Text = ""
                    Label68.Text = ""
                    Label67.Text = ""
                    Label82.Text = ""
                    Label63.Text = ""
                    Label62.Text = ""
                Case 3
                    '  Me.TabControl1.Size = New Size(1246, 619)
                    '  Me.Width = 1255
                    '   Me.Height = 657
                    With DataGridView5
                        .RowsDefaultCellStyle.BackColor = Color.Bisque
                        .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                    End With
                    DataGridView5.DataSource = Nothing
                    rreshtiaktual2 = 0
                    Lidhje1bc.Open()
                    adaptor1bc = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", Lidhje1bc)
                    adaptor1bc.Fill(setidatave1bc, "tedhena")
                    DataGridView5.Columns.Clear()
                    DataGridView5.Columns.Add("count1", "Nr.")
                    DataGridView5.DataSource = setidatave1bc.Tables(0)
                    DataGridView5.Columns("ID").Visible = False
                    Lidhje1bc.Close()
                    If DataGridView5.Rows.Count = 0 Then
                        Button7.Enabled = False
                        Button8.Enabled = False
                        Button9.Enabled = False
                        CheckBox1.Enabled = False
                    Else
                        Button7.Enabled = True
                        Button8.Enabled = True
                        Button9.Enabled = True
                        CheckBox1.Enabled = True
                    End If


                    Label11.Text = ""
                    Label42.Text = ""
                    Label43.Text = ""
                    Label12.Text = ""
                    Label50.Text = ""
                    Label51.Text = ""
                    Label13.Text = ""
                    Label53.Text = ""
                    Label52.Text = ""
                    Label26.Text = ""
                    Label58.Text = ""
                    Label57.Text = ""
                Case 4
                    '  Me.TabControl1.Size = New Size(1296, 615)
                    '  Me.Width = 1312
                    '   Me.Height = 653
                    With DataGridView3
                        .RowsDefaultCellStyle.BackColor = Color.Bisque
                        .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                    End With
                    DataGridView3.DataSource = Nothing
                    rreshtiaktualdg3 = 0
                    lidhjedg3.Open()
                    setitedhenavedg3.clear
                    adaptoridg3 = New OleDbDataAdapter("SELECT * FROM Dhenat ORDER BY ID", lidhjedg3)
                    adaptoridg3.Fill(setitedhenavedg3, "tedhena")
                    DataGridView3.Columns.Clear()
                    DataGridView3.Columns.Add("count1", "Nr.")
                    DataGridView3.DataSource = setitedhenavedg3.Tables(0)
                    DataGridView3.Columns("ID").Visible = False
                    lidhjedg3.Close()



                    With DataGridView4
                        .RowsDefaultCellStyle.BackColor = Color.Bisque
                        .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                    End With
                    DataGridView4.DataSource = Nothing
                    rreshtiaktualdg4 = 0
                    lidhjedg4.Open()
                    setitedhenavedg4.clear
                    adaptoridg4 = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", lidhjedg4)
                    adaptoridg4.Fill(setitedhenavedg4, "tedhena")
                    DataGridView4.Columns.Clear()
                    DataGridView4.Columns.Add("count1", "Nr.")
                    DataGridView4.DataSource = setitedhenavedg4.Tables(0)
                    DataGridView4.Columns("ID").Visible = False
                    lidhjedg4.Close()
                    Dim nr3 As Integer = 0
                    nr3 = DataGridView3.Rows.Count - 1
                    For i = 0 To nr3
                        DataGridView3.Rows(i).Cells(0).Value = i + 1
                    Next
                    Dim nr4 As Integer = 0
                    nr4 = DataGridView4.Rows.Count - 1
                    For i = 0 To nr4
                        DataGridView4.Rows(i).Cells(0).Value = i + 1
                    Next
                Case 5
                    '  Me.TabControl1.Size = New Size(1373, 617)
                    '  Me.Width = 1380
                    '  Me.Height = 655
                    With DataGridView7
                        .RowsDefaultCellStyle.BackColor = Color.Bisque
                        .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
                    End With
                    DataGridView7.DataSource = Nothing
                    rreshtiaktualdg71 = 0
                    lidhjedg71.Open()
                    adaptoridg71 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhjedg71)
                    adaptoridg71.Fill(setitedhenavedg71, "tedhena")
                    DataGridView7.Columns.Clear()
                    DataGridView7.Columns.Add("count1", "Nr.")
                    DataGridView7.DataSource = setitedhenavedg71.Tables(0)
                    DataGridView7.Columns("ID").Visible = False
                    Dim nr7 As Integer
                    nr7 = DataGridView7.Rows.Count - 1
                    For i = 0 To nr7
                        DataGridView7.Rows(i).Cells(0).Value = i + 1
                    Next
                    If setitedhenavedg71.Tables(0).Rows.Count > 0 Then
                        Merrtedhenat_Ofertat(rreshtiaktualdg71)
                        Button23.Text = "Ruaj Faturen(" & setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Kodi_Ofertes") & ")"
                        Button22.Text = "Printo Faturen(" & setitedhenavedg71.Tables("tedhena").Rows(rreshtiaktualdg71)("Kodi_Ofertes") & ")"
                    Else
                    End If
                    lidhjedg71.Close()
                    If DataGridView7.Rows.Count = 0 Then
                        Button18.Enabled = False
                        Button19.Enabled = False
                    Else
                        Button18.Enabled = True
                        Button19.Enabled = True
                    End If
            End Select
        End If
        If DataGridView1.Rows.Count = 0 Then
            Button2.Enabled = False
            Button3.Enabled = False
        Else
            Button2.Enabled = True
            Button3.Enabled = True
        End If
    End Sub
    Private Sub TextBox27_TextChanged(sender As Object, e As EventArgs) Handles TextBox27.TextChanged
        If TextBox27.Text = "" Then
            setitedhenavedg31.clear
            setitedhenavedg41.clear
            Label17.Text = 0
            Label18.Text = 0
            Label19.Text = 0
            Label17.Enabled = False
            Label18.Enabled = False
            Label19.Enabled = False
            Label17.ForeColor = DefaultForeColor
            Label18.ForeColor = DefaultForeColor
            Label19.ForeColor = DefaultForeColor
            Lidhje1.Close()
            lidhje2.Close()
            With DataGridView3
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            rreshtiaktualdg31 = 0
            lidhjedg31.Open()
            adaptoridg31 = New OleDbDataAdapter("SELECT * FROM dhenat WHERE(Produkti Like '" &
                                            TextBox27.Text & "%') ", lidhjedg31)
            adaptoridg31.Fill(setitedhenavedg31, "tedhena")
            DataGridView3.DataSource = setitedhenavedg31.Tables(0)
            nr3 = DataGridView3.Rows.Count - 1
            For i = 0 To nr3
                DataGridView3.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView3.Refresh()
            lidhjedg31.Close()
            With DataGridView4
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            rreshtiaktualdg41 = 0
            lidhjedg41.Open()
            adaptoridg41 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Produkti Like '" &
                                            TextBox27.Text & "%') ", lidhjedg41)
            adaptoridg41.Fill(setitedhenavedg41, "tedhena")
            DataGridView4.DataSource = setitedhenavedg41.Tables(0)
            nr4 = DataGridView4.Rows.Count - 1
            For i = 0 To nr4
                DataGridView4.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView4.Refresh()
            lidhjedg41.Close()
        Else
            setitedhenavedg31.clear
            setitedhenavedg41.clear
            lidhje2.Close()
            Lidhje1.Close()
            Label17.Enabled = True
            Label18.Enabled = True
            Label19.Enabled = True
            Label17.ForeColor = Color.Orange
            Label18.ForeColor = Color.Red
            Label19.ForeColor = Color.Green
            Dim total12 As Integer = 0
            lidhje2.Open()
            adaptori2 = New OleDbDataAdapter("SELECT * FROM dhenat WHERE(Produkti Like '" &
                                            TextBox27.Text & "%') ORDER BY Kodi_Blerjes", lidhje2)
            Dim dataTable As New DataTable
            adaptori2.Fill(dataTable)
            DataGridView3.DataSource = dataTable
            nr3 = DataGridView3.Rows.Count - 1
            For i = 0 To nr3
                DataGridView3.Rows(i).Cells(0).Value = i + 1
            Next
            For i As Integer = 0 To DataGridView3.RowCount - 1
                total12 = total12 + DataGridView3.Rows(i).Cells(9).Value
            Next
            Label17.Enabled = True
            Label17.Text = total12
            lidhje2.Close()
            Dim total13 As Integer = 0
            Lidhje1.Open()
            adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Produkti Like '" &
                                            TextBox27.Text & "%') ORDER BY Kodi_Shitjes", Lidhje1)
            Dim dataTable1 As New DataTable
            adaptor1.Fill(dataTable1)
            DataGridView4.DataSource = dataTable1
            lidhje2v.Close()
            DataGridView4.Sort(DataGridView4.Columns(1), System.ComponentModel.ListSortDirection.Ascending)
            DataGridView4.RefreshEdit()
            lidhje2v.Close()
            nr4 = DataGridView4.Rows.Count - 1
            For i = 0 To nr4
                DataGridView4.Rows(i).Cells(0).Value = i + 1
            Next
            For i As Integer = 0 To DataGridView4.RowCount - 1
                total13 = total13 + DataGridView4.Rows(i).Cells(9).Value
            Next
            Label18.Enabled = True
            Label18.Text = total13
            Lidhje1.Close()
            Label19.Enabled = True
            Label19.Text = Label17.Text - Label18.Text
        End If
    End Sub
    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        Label12.Enabled = False
        Label50.Enabled = False
        Label51.Enabled = False
        Label13.Enabled = False
        Label53.Enabled = False
        Label52.Enabled = False
        Label12.Text = 0
        Label50.Text = 0
        Label51.Text = 0
        Label13.Text = 0
        Label53.Text = 0
        Label52.Text = 0
        setitedhenave2117.clear
        With DataGridView5
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual2117 = 0
        lidhje2117.Open()
        adaptori2117 = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", lidhje2117)
        adaptori2117.Fill(setitedhenave2117, "tedhena")
        DataGridView5.DataSource = setitedhenave2117.Tables(0)
        lidhje2117.Close()
        Dim total1 As Integer = 0
        Dim total2 As Integer = 0
        Dim total3 As Integer = 0
        Dim regDate2 As DateTime = Date.Now
        Dim strDate2 As String = regDate2.ToString("dd/MM/yyyy")
        Lidhje1.Open()
        adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Data Like '" &
                                        strDate2 & "%')", Lidhje1)
        Dim dataTable1 As New DataTable
        adaptor1.Fill(dataTable1)
        DataGridView5.DataSource = dataTable1
        DataGridView5.Columns("ID").Visible = False
        Dim nr5 As Integer
        nr5 = DataGridView5.Rows.Count - 1
        For i = 0 To nr5
            DataGridView5.Rows(i).Cells(0).Value = i + 1
        Next
        For i As Integer = 0 To DataGridView5.RowCount - 1
            total1 = total1 + DataGridView5.Rows(i).Cells(11).Value
            total2 = total2 + DataGridView5.Rows(i).Cells(12).Value
            total3 = total3 + DataGridView5.Rows(i).Cells(13).Value
        Next
        Label11.Enabled = True
        Label42.Enabled = True
        Label43.Enabled = True
        Label11.ForeColor = Color.Orange
        Label42.ForeColor = Color.Red
        Label43.ForeColor = Color.Green
        Label11.Text = total1.ToString("#,#", CultureInfo.InvariantCulture)
        Label42.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
        Label43.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
        Lidhje1.Close()
        MsgBox("U gjenerua me sukses", MsgBoxStyle.Information)
    End Sub
    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        Label11.Enabled = False
        Label42.Enabled = False
        Label43.Enabled = False
        Label13.Enabled = False
        Label53.Enabled = False
        Label52.Enabled = False
        Label11.Text = 0
        Label42.Text = 0
        Label43.Text = 0
        Label13.Text = 0
        Label53.Text = 0
        Label52.Text = 0
        setitedhenave2117.clear
        With DataGridView5
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual2117 = 0
        lidhje2117.Open()
        adaptori2117 = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", lidhje2117)
        adaptori2117.Fill(setitedhenave2117, "tedhena")
        DataGridView5.DataSource = setitedhenave2117.Tables(0)
        lidhje2117.Close()
        Dim total11 As Integer = 0
        Dim total21 As Integer = 0
        Dim total31 As Integer = 0
        Dim regDate4 As DateTime = Date.Now
        Dim regDate5 As DateTime = Date.Now
        Dim strDate2 As String = regDate4.ToString("/MM/")
        Dim strDate3 As String = regDate5.ToString("yyyy")
        Lidhje1.Open()
        Dim dataTable1 As New DataTable
        For i As Integer = 1 To 31
            adaptor1 = New OleDbDataAdapter("SELECT* FROM Shitjet WHERE(Data Like '" &
                                            i.ToString("D2") & strDate2 & strDate3 & "%')", Lidhje1)
            adaptor1.Fill(dataTable1)
            DataGridView5.DataSource = dataTable1
        Next
        DataGridView5.Columns("ID").Visible = False
        Dim nr5 As Integer
        nr5 = DataGridView5.Rows.Count - 1
        For i = 0 To nr5
            DataGridView5.Rows(i).Cells(0).Value = i + 1
        Next
        For ii As Integer = 0 To DataGridView5.RowCount - 1
            total11 = total11 + DataGridView5.Rows(ii).Cells(11).Value
            total21 = total21 + DataGridView5.Rows(ii).Cells(12).Value
            total31 = total31 + DataGridView5.Rows(ii).Cells(13).Value
        Next
        adaptor1.Update(dataTable1)
        Label12.Enabled = True
        Label50.Enabled = True
        Label51.Enabled = True
        Label12.Text = total11.ToString("#,#", CultureInfo.InvariantCulture)
        Label50.Text = total21.ToString("#,#", CultureInfo.InvariantCulture)
        Label51.Text = total31.ToString("#,#", CultureInfo.InvariantCulture)
        Label12.ForeColor = Color.Orange
        Label50.ForeColor = Color.Red
        Label51.ForeColor = Color.Green
        Lidhje1.Close()
        MsgBox("U gjenerua me sukses", MsgBoxStyle.Information)
    End Sub
    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        Label11.Enabled = False
        Label42.Enabled = False
        Label43.Enabled = False
        Label12.Enabled = False
        Label50.Enabled = False
        Label51.Enabled = False
        Label11.Text = 0
        Label42.Text = 0
        Label43.Text = 0
        Label12.Text = 0
        Label50.Text = 0
        Label51.Text = 0
        setitedhenave2117.clear
        Dim regDate2 As DateTime = Date.Now
        Dim regDate3 As DateTime = Date.Now
        Dim strDate2 As String = regDate2.ToString("/MM/")
        Dim strDate3 As String = regDate2.ToString("yyyy")
        Dim total111 As Integer = 0
        Dim total121 As Integer = 0
        Dim total131 As Integer = 0
        Lidhje1.Open()
        Dim dataTable1 As New DataTable
        For j As Int32 = 1 To 12
            For i As Int32 = 1 To 31
                adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Data Like '" &
                                                i.ToString("D2") & "/" & j.ToString("D2") & "/" & strDate3 & "%')", Lidhje1)
                adaptor1.Fill(dataTable1)
                DataGridView5.DataSource = dataTable1
            Next
        Next
        DataGridView5.Columns("ID").Visible = False
        Dim nr5 As Integer
        nr5 = DataGridView5.Rows.Count - 1
        For i = 0 To nr5
            DataGridView5.Rows(i).Cells(0).Value = i + 1
        Next
        For ii As Integer = 0 To DataGridView5.RowCount - 1
            total111 = total111 + DataGridView5.Rows(ii).Cells(11).Value
            total121 = total121 + DataGridView5.Rows(ii).Cells(12).Value
            total131 = total131 + DataGridView5.Rows(ii).Cells(13).Value
        Next
        adaptor1.Update(dataTable1)
        Label13.Enabled = True
        Label53.Enabled = True
        Label52.Enabled = True
        Label13.ForeColor = Color.Orange
        Label53.ForeColor = Color.Red
        Label52.ForeColor = Color.Green
        Label13.Text = total111.ToString("#,#", CultureInfo.InvariantCulture)
        Label53.Text = total121.ToString("#,#", CultureInfo.InvariantCulture)
        Label52.Text = total131.ToString("#,#", CultureInfo.InvariantCulture)
        Lidhje1.Close()
        MsgBox("U gjenerua me sukses", MsgBoxStyle.Information)
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If DateTimePicker1.Value.ToString("ddMMyyyy") = DateTimePicker2.Value.ToString("ddMMyyyy") Then
            Dim totalall As Integer = 0
            Dim totalall1 As Integer = 0
            Dim totalall2 As Integer = 0
            DataGridView5.DataSource = setitedhenave2.Tables(0).DefaultView
            DataGridView5.Refresh()
            Dim dataTable11 As New DataTable
            Lidhje1.Open()
            Dim day As Int32 = CInt(DateTimePicker1.Value.ToString("dd"))
            Dim month As Int32 = (DateTimePicker1.Value.ToString("MM"))
            adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Data Like '" &
                                    day.ToString("d2") & "/" & month.ToString("d2") & "/" & DateTimePicker1.Value.ToString("yyyy") & "%')", Lidhje1)
            adaptor1.Fill(dataTable11)
            DataGridView5.DataSource = dataTable11
            DataGridView5.Columns("ID").Visible = False
            Dim nr5 As Integer
            nr5 = DataGridView5.Rows.Count - 1
            For i = 0 To nr5
                DataGridView5.Rows(i).Cells(0).Value = i + 1
            Next
            For ii As Integer = 0 To DataGridView5.RowCount - 1
                totalall = totalall + DataGridView5.Rows(ii).Cells(11).Value
                totalall1 = totalall1 + DataGridView5.Rows(ii).Cells(12).Value
                totalall2 = totalall2 + DataGridView5.Rows(ii).Cells(13).Value
            Next
            adaptor1.Update(dataTable11)
            Label26.Enabled = True
            Label58.Enabled = True
            Label57.Enabled = True
            Label26.ForeColor = Color.Orange
            Label58.ForeColor = Color.Red
            Label57.ForeColor = Color.Green
            Label26.Text = totalall.ToString("#,#", CultureInfo.InvariantCulture)
            Label58.Text = totalall1.ToString("#,#", CultureInfo.InvariantCulture)
            Label57.Text = totalall2.ToString("#,#", CultureInfo.InvariantCulture)
            Lidhje1.Close()
            MsgBox("U gjenerua me sukses", MsgBoxStyle.Information)
        Else
            Dim totalall As Integer = 0
            Dim totalall1 As Integer = 0
            Dim totalall2 As Integer = 0
            DataGridView5.DataSource = setitedhenave2.Tables(0).DefaultView
            DataGridView5.Refresh()
            Dim dataTable1 As New DataTable
            Lidhje1.Open()
            Dim day As Int32 = CInt(DateTimePicker1.Value.ToString("dd"))
            Dim month As Int32 = (DateTimePicker1.Value.ToString("MM"))
            adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Data BETWEEN " & "'" &
                                              DateTimePicker1.Value.ToString("dd/MM/yyyy") & "'" & " AND " & "'" & DateTimePicker2.Value.ToString("dd/MM/yyyy") & "')", Lidhje1)
            adaptor1.Fill(dataTable1)
            DataGridView5.DataSource = dataTable1
            DataGridView5.Columns("ID").Visible = False
            Dim nr5 As Integer
            nr5 = DataGridView5.Rows.Count - 1
            For i = 0 To nr5
                DataGridView5.Rows(i).Cells(0).Value = i + 1
            Next
            For ii As Integer = 0 To DataGridView5.RowCount - 1
                totalall = totalall + DataGridView5.Rows(ii).Cells(11).Value
                totalall1 = totalall1 + DataGridView5.Rows(ii).Cells(12).Value
                totalall2 = totalall2 + DataGridView5.Rows(ii).Cells(13).Value
            Next
            adaptor1.Update(dataTable1)
            Label26.Enabled = True
            Label58.Enabled = True
            Label57.Enabled = True
            Label26.ForeColor = Color.Orange
            Label58.ForeColor = Color.Red
            Label57.ForeColor = Color.Green
            Label26.Text = totalall.ToString("#,#", CultureInfo.InvariantCulture)
            Label58.Text = totalall1.ToString("#,#", CultureInfo.InvariantCulture)
            Label57.Text = totalall2.ToString("#,#", CultureInfo.InvariantCulture)
            Lidhje1.Close()
            MsgBox("U gjenerua me sukses", MsgBoxStyle.Information)
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Button7.Enabled = False
            Button8.Enabled = False
            Button9.Enabled = False
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = True
            Button13.Enabled = True
            Label13.Enabled = False
            Label53.Enabled = False
            Label52.Enabled = False
            Label12.Enabled = False
            Label50.Enabled = False
            Label51.Enabled = False
            Label11.Enabled = False
            Label42.Enabled = False
            Label43.Enabled = False
            Label13.ForeColor = DefaultForeColor
            Label53.ForeColor = DefaultForeColor
            Label52.ForeColor = DefaultForeColor
            Label12.ForeColor = DefaultForeColor
            Label50.ForeColor = DefaultForeColor
            Label51.ForeColor = DefaultForeColor
            Label11.ForeColor = DefaultForeColor
            Label42.ForeColor = DefaultForeColor
            Label43.ForeColor = DefaultForeColor
            Label13.Text = 0
            Label53.Text = 0
            Label52.Text = 0
            Label12.Text = 0
            Label50.Text = 0
            Label51.Text = 0
            Label11.Text = 0
            Label42.Text = 0
            Label43.Text = 0
            CheckBox1.ForeColor = Color.ForestGreen
        Else
            Label26.Text = 0
            Label58.Text = 0
            Label57.Text = 0
            Label26.ForeColor = DefaultForeColor
            Label58.ForeColor = DefaultForeColor
            Label57.ForeColor = DefaultForeColor
            Button7.Enabled = True
            Button8.Enabled = True
            Button9.Enabled = True
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            Button13.Enabled = False
            CheckBox1.ForeColor = Color.Red
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox6.Text = ""
        Lidhje1.Close()
        lidhje2.Close()
        With DataGridView1
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual = 0
        Lidhje1.Open()
        adaptor1 = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", Lidhje1)
        adaptor1.Fill(setitedhenave2, "tedhena")
        DataGridView1.DataSource = setitedhenave2.Tables(0)
        nr1 = DataGridView1.Rows.Count - 1
        For i = 0 To nr1
            DataGridView1.Rows(i).Cells(0).Value = i + 1
        Next
        DataGridView1.Refresh()
        Merrtedhenat(rreshtiaktual2)
        Lidhje1.Close()
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Label86.Enabled = False
        Label73.Enabled = False
        Label72.Enabled = False
        Label87.Enabled = False
        Label68.Enabled = False
        Label67.Enabled = False
        Label86.Text = 0
        Label73.Text = 0
        Label72.Text = 0
        Label87.Text = 0
        Label68.Text = 0
        Label67.Text = 0
        setitedhenave2116.clear
        With DataGridView6
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual2 = 0
        lidhje2116.Open()
        adaptori2116 = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", lidhje2116)
        adaptori2116.Fill(setitedhenave2116, "tedhena")
        DataGridView6.DataSource = setitedhenave2116.Tables(0)
        lidhje2116.Close()
        Dim total1 As Integer = 0
        Dim total2 As Integer = 0
        Dim total3 As Integer = 0
        Dim regDate2 As DateTime = Date.Now
        Dim strDate2 As String = regDate2.ToString("dd/MM/yyyy")
        Lidhje1.Open()
        adaptor1 = New OleDbDataAdapter("SELECT * FROM dhenat WHERE(Data Like '" &
                                        strDate2 & "%')", Lidhje1)
        Dim dataTable6 As New DataTable
        adaptor1.Fill(dataTable6)
        DataGridView6.DataSource = dataTable6
        DataGridView6.Columns("ID").Visible = False
        Dim nr6 As Integer
        nr6 = DataGridView6.Rows.Count - 1
        For i = 0 To nr6
            DataGridView6.Rows(i).Cells(0).Value = i + 1
        Next
        For i As Integer = 0 To DataGridView6.RowCount - 1
            total1 = total1 + DataGridView6.Rows(i).Cells(11).Value
            total2 = total2 + DataGridView6.Rows(i).Cells(12).Value
            total3 = total3 + DataGridView6.Rows(i).Cells(13).Value
        Next
        Label85.Enabled = True
        Label79.Enabled = True
        Label77.Enabled = True
        Label85.ForeColor = Color.Orange
        Label79.ForeColor = Color.Red
        Label77.ForeColor = Color.Green
        Label85.Text = total1.ToString("#,#", CultureInfo.InvariantCulture)
        Label79.Text = total2.ToString("#,#", CultureInfo.InvariantCulture)
        Label77.Text = total3.ToString("#,#", CultureInfo.InvariantCulture)
        Lidhje1.Close()
        MsgBox("U gjenerua me sukses", MsgBoxStyle.Information)
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Label85.Enabled = False
        Label79.Enabled = False
        Label77.Enabled = False
        Label87.Enabled = False
        Label68.Enabled = False
        Label67.Enabled = False
        Label85.Text = 0
        Label79.Text = 0
        Label77.Text = 0
        Label87.Text = 0
        Label68.Text = 0
        Label67.Text = 0
        setitedhenave2116.clear
        With DataGridView6
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual2116 = 0
        lidhje2116.Open()
        adaptori2116 = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", lidhje2116)
        adaptori2116.Fill(setitedhenave2116, "tedhena")
        DataGridView6.DataSource = setitedhenave2116.Tables(0)
        lidhje2116.Close()
        Dim total11 As Integer = 0
        Dim total21 As Integer = 0
        Dim total31 As Integer = 0
        Dim regDate4 As DateTime = Date.Now
        Dim regDate5 As DateTime = Date.Now
        Dim strDate2 As String = regDate4.ToString("/MM/")
        Dim strDate3 As String = regDate5.ToString("yyyy")
        Lidhje1.Open()
        Dim dataTable7 As New DataTable
        For i As Integer = 1 To 31
            adaptor1 = New OleDbDataAdapter("SELECT* FROM dhenat WHERE(Data Like '" &
                                            i.ToString("D2") & strDate2 & strDate3 & "%')", Lidhje1)
            adaptor1.Fill(dataTable7)
            DataGridView6.DataSource = dataTable7
        Next
        DataGridView6.Columns("ID").Visible = False
        Dim nr6 As Integer
        nr6 = DataGridView6.Rows.Count - 1
        For i = 0 To nr6
            DataGridView6.Rows(i).Cells(0).Value = i + 1
        Next
        For ii As Integer = 0 To DataGridView6.RowCount - 1
            total11 = total11 + DataGridView6.Rows(ii).Cells(11).Value
            total21 = total21 + DataGridView6.Rows(ii).Cells(12).Value
            total31 = total31 + DataGridView6.Rows(ii).Cells(13).Value
        Next
        adaptor1.Update(dataTable7)
        Label86.Enabled = True
        Label73.Enabled = True
        Label72.Enabled = True
        Label86.Text = total11.ToString("#,#", CultureInfo.InvariantCulture)
        Label73.Text = total21.ToString("#,#", CultureInfo.InvariantCulture)
        Label72.Text = total31.ToString("#,#", CultureInfo.InvariantCulture)
        Label86.ForeColor = Color.Orange
        Label73.ForeColor = Color.Red
        Label72.ForeColor = Color.Green
        Lidhje1.Close()
        MsgBox("U gjenerua me sukses", MsgBoxStyle.Information)
    End Sub
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Label85.Enabled = False
        Label79.Enabled = False
        Label77.Enabled = False
        Label86.Enabled = False
        Label73.Enabled = False
        Label72.Enabled = False
        Label85.Text = 0
        Label79.Text = 0
        Label77.Text = 0
        Label86.Text = 0
        Label73.Text = 0
        Label72.Text = 0
        setitedhenave2116.clear
        Dim regDate2 As DateTime = Date.Now
        Dim regDate3 As DateTime = Date.Now
        Dim strDate2 As String = regDate2.ToString("/MM/")
        Dim strDate3 As String = regDate2.ToString("yyyy")
        Dim total111 As Integer = 0
        Dim total121 As Integer = 0
        Dim total131 As Integer = 0
        Lidhje1.Open()
        Dim dataTable8 As New DataTable
        For j As Int32 = 1 To 12
            For i As Int32 = 1 To 31
                adaptor1 = New OleDbDataAdapter("SELECT * FROM dhenat WHERE(Data Like '" &
                                                i.ToString("D2") & "/" & j.ToString("D2") & "/" & strDate3 & "%')", Lidhje1)
                adaptor1.Fill(dataTable8)
                DataGridView6.DataSource = dataTable8
            Next
        Next
        DataGridView6.Columns("ID").Visible = False
        Dim nr6 As Integer
        nr6 = DataGridView6.Rows.Count - 1
        For i = 0 To nr6
            DataGridView6.Rows(i).Cells(0).Value = i + 1
        Next
        For ii As Integer = 0 To DataGridView6.RowCount - 1
            total111 = total111 + DataGridView6.Rows(ii).Cells(11).Value
            total121 = total121 + DataGridView6.Rows(ii).Cells(12).Value
            total131 = total131 + DataGridView6.Rows(ii).Cells(13).Value
        Next
        adaptor1.Update(dataTable8)
        Label87.Enabled = True
        Label68.Enabled = True
        Label67.Enabled = True
        Label87.ForeColor = Color.Orange
        Label68.ForeColor = Color.Red
        Label67.ForeColor = Color.Green
        Label87.Text = total111.ToString("#,#", CultureInfo.InvariantCulture)
        Label68.Text = total121.ToString("#,#", CultureInfo.InvariantCulture)
        Label67.Text = total131.ToString("#,#", CultureInfo.InvariantCulture)
        Lidhje1.Close()
        MsgBox("U gjenerua me sukses", MsgBoxStyle.Information)
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If DateTimePicker4.Value.ToString("ddMMyyyy") = DateTimePicker3.Value.ToString("ddMMyyyy") Then
            Dim totalall As Integer = 0
            Dim totalall1 As Integer = 0
            Dim totalall2 As Integer = 0
            setitedhenave2116.clear
            Dim dataTable11 As New DataTable
            Lidhje1.Open()
            Dim day As Int32 = CInt(DateTimePicker4.Value.ToString("dd"))
            Dim month As Int32 = (DateTimePicker4.Value.ToString("MM"))
            adaptor1 = New OleDbDataAdapter("SELECT * FROM dhenat WHERE(Data Like '" &
                                    day.ToString("d2") & "/" & month.ToString("d2") & "/" & DateTimePicker4.Value.ToString("yyyy") & "%')", Lidhje1)
            adaptor1.Fill(dataTable11)
            DataGridView6.DataSource = dataTable11
            DataGridView6.Columns("ID").Visible = False
            Dim nr6 As Integer
            nr6 = DataGridView6.Rows.Count - 1
            For i = 0 To nr6
                DataGridView6.Rows(i).Cells(0).Value = i + 1
            Next
            For ii As Integer = 0 To DataGridView6.RowCount - 1
                totalall = totalall + DataGridView6.Rows(ii).Cells(11).Value
                totalall1 = totalall1 + DataGridView6.Rows(ii).Cells(12).Value
                totalall2 = totalall2 + DataGridView6.Rows(ii).Cells(13).Value
            Next
            adaptor1.Update(dataTable11)
            Label82.Enabled = True
            Label63.Enabled = True
            Label62.Enabled = True
            Label82.ForeColor = Color.Orange
            Label63.ForeColor = Color.Red
            Label62.ForeColor = Color.Green
            Label82.Text = totalall1.ToString("#,#", CultureInfo.InvariantCulture)
            Label62.Text = totalall.ToString("#,#", CultureInfo.InvariantCulture)
            Label63.Text = totalall2.ToString("#,#", CultureInfo.InvariantCulture)
            Lidhje1.Close()
            MsgBox("U gjenerua me sukses", MsgBoxStyle.Information)
        Else
            Dim totalall As Integer = 0
            Dim totalall1 As Integer = 0
            Dim totalall2 As Integer = 0
            DataGridView6.DataSource = setitedhenave2.Tables(0).DefaultView
            DataGridView6.Refresh()
            Dim dataTable1 As New DataTable
            Lidhje1.Open()
            Dim day As Int32 = CInt(DateTimePicker4.Value.ToString("dd"))
            Dim month As Int32 = (DateTimePicker4.Value.ToString("MM"))
            adaptor1 = New OleDbDataAdapter("SELECT * FROM dhenat WHERE(Data BETWEEN " & "'" &
                                              DateTimePicker4.Value.ToString("dd/MM/yyyy") & "'" & " AND " & "'" & DateTimePicker3.Value.ToString("dd/MM/yyyy") & "')", Lidhje1)
            adaptor1.Fill(dataTable1)
            DataGridView6.DataSource = dataTable1
            DataGridView6.Columns("ID").Visible = False
            Dim nr6 As Integer
            nr6 = DataGridView6.Rows.Count - 1
            For i = 0 To nr6
                DataGridView6.Rows(i).Cells(0).Value = i + 1
            Next
            For ii As Integer = 0 To DataGridView6.RowCount - 1
                totalall = totalall + DataGridView6.Rows(ii).Cells(11).Value
                totalall1 = totalall1 + DataGridView6.Rows(ii).Cells(12).Value
                totalall2 = totalall2 + DataGridView6.Rows(ii).Cells(13).Value
            Next
            adaptor1.Update(dataTable1)
            Label82.Enabled = True
            Label63.Enabled = True
            Label62.Enabled = True
            Label82.ForeColor = Color.Orange
            Label63.ForeColor = Color.Red
            Label62.ForeColor = Color.Green
            Label82.Text = totalall.ToString("#,#", CultureInfo.InvariantCulture)
            Label63.Text = totalall1.ToString("#,#", CultureInfo.InvariantCulture)
            Label62.Text = totalall2.ToString("#,#", CultureInfo.InvariantCulture)
            Lidhje1.Close()
            MsgBox("U gjenerua me sukses", MsgBoxStyle.Information)
        End If
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Button10.Enabled = False
            Button12.Enabled = False
            Button14.Enabled = False
            DateTimePicker4.Enabled = True
            DateTimePicker3.Enabled = True
            Button5.Enabled = True
            Label85.Enabled = False
            Label79.Enabled = False
            Label77.Enabled = False
            Label86.Enabled = False
            Label73.Enabled = False
            Label72.Enabled = False
            Label87.Enabled = False
            Label68.Enabled = False
            Label67.Enabled = False
            Label85.ForeColor = DefaultForeColor
            Label79.ForeColor = DefaultForeColor
            Label77.ForeColor = DefaultForeColor
            Label86.ForeColor = DefaultForeColor
            Label73.ForeColor = DefaultForeColor
            Label72.ForeColor = DefaultForeColor
            Label87.ForeColor = DefaultForeColor
            Label68.ForeColor = DefaultForeColor
            Label67.ForeColor = DefaultForeColor
            Label85.Text = 0
            Label79.Text = 0
            Label77.Text = 0
            Label86.Text = 0
            Label73.Text = 0
            Label72.Text = 0
            Label87.Text = 0
            Label68.Text = 0
            Label67.Text = 0
            CheckBox2.ForeColor = Color.ForestGreen
        Else
            Label82.Text = 0
            Label63.Text = 0
            Label62.Text = 0
            Label82.ForeColor = DefaultForeColor
            Label63.ForeColor = DefaultForeColor
            Label62.ForeColor = DefaultForeColor
            Button10.Enabled = True
            Button12.Enabled = True
            Button14.Enabled = True
            DateTimePicker4.Enabled = False
            DateTimePicker3.Enabled = False
            Button5.Enabled = False
            CheckBox2.ForeColor = Color.Red
        End If
    End Sub
    Public Sub CreateReader_load_produktet_sasialimit(ByVal connectionString As String,
  ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                Dim total12 As Integer = 0
                lidhje2.Open()
                adaptori2 = New OleDbDataAdapter("SELECT * FROM dhenat WHERE(Produkti Like '" &
                                            reader(1).ToString() & "%') AND (Data BETWEEN " & "'" &
                                              DateTimePicker5.Value.ToString("dd/MM/yyyy") & "'" & " AND " & "'" & DateTimePicker6.Value.ToString("dd/MM/yyyy") & "')", lidhje2)
                Dim dataTable As New DataTable
                adaptori2.Fill(dataTable)
                DataGridView3.DataSource = dataTable
                For i As Integer = 0 To DataGridView3.RowCount - 1
                    total12 = total12 + DataGridView3.Rows(i).Cells(9).Value
                Next
                Label17.Enabled = True
                Label17.Text = total12
                lidhje2.Close()
                Dim total13 As Integer = 0
                Lidhje1.Open()
                adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Produkti Like '" &
                                            reader(1).ToString() & "%') AND (Data BETWEEN " & "'" &
                                              DateTimePicker5.Value.ToString("dd/MM/yyyy") & "'" & " AND " & "'" & DateTimePicker6.Value.ToString("dd/MM/yyyy") & "')", Lidhje1)
                Dim dataTable1 As New DataTable
                adaptor1.Fill(dataTable1)
                DataGridView4.DataSource = dataTable1
                For i As Integer = 0 To DataGridView4.RowCount - 1
                    total13 = total13 + DataGridView4.Rows(i).Cells(9).Value
                Next
                Label18.Enabled = True
                Label18.Text = total13
                Lidhje1.Close()
                Label19.Enabled = True
                Label19.Text = Label17.Text - Label18.Text
                If Label19.Text < reader(6).ToString() Or Label19.Text < reader(6).ToString() + CInt(TextBox41.Text) Then


                    Magazina.ListBox1.Items.Add(reader(1).ToString())
                    Magazina.ListBox2.Items.Add(Label19.Text)
                    Magazina.ListBox3.Items.Add(reader(6).ToString())

                Else


                End If
                For i As Integer = 0 To Magazina.ListBox1.Items.Count - 1
                    Dim dr As New DataGridViewRow
                    Magazina.DataGridView2.Rows.Add(dr)
                    Magazina.DataGridView2.Rows(i).Cells(1).Value = Magazina.ListBox1.Items(i)
                    Magazina.DataGridView2.Rows(i).Cells(0).Value = i + 1
                    Magazina.DataGridView2.AllowUserToAddRows = False
                Next
                For i As Integer = 0 To Magazina.ListBox2.Items.Count - 1
                    Dim dr As New DataGridViewRow
                    Magazina.DataGridView2.Rows.Add(dr)
                    Magazina.DataGridView2.Rows(i).Cells(2).Value = Magazina.ListBox2.Items(i)
                    Magazina.DataGridView2.AllowUserToAddRows = False
                Next
                For i As Integer = 0 To Magazina.ListBox3.Items.Count - 1
                    Dim dr As New DataGridViewRow
                    Magazina.DataGridView2.Rows.Add(dr)
                    Magazina.DataGridView2.Rows(i).Cells(3).Value = Magazina.ListBox3.Items(i)
                    Magazina.DataGridView2.AllowUserToAddRows = False
                Next

                For r As Integer = Magazina.DataGridView2.Rows.Count - 1 To 0 Step -1
                    Dim empty As Boolean = True
                    For Each cell As DataGridViewCell In Magazina.DataGridView2.Rows(r).Cells
                        If Not IsNothing(cell.Value) Then
                            empty = False
                            Exit For
                        End If
                    Next
                    If empty Then Magazina.DataGridView2.Rows.RemoveAt(r)
                Next
                Me.Hide()
                Magazina.Show()
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




                Dim total12 As Integer = 0
                lidhje2.Open()
                adaptori2 = New OleDbDataAdapter("SELECT * FROM dhenat WHERE(Produkti Like '" &
                                            reader(1).ToString() & "%') AND (Data BETWEEN " & "'" &
                                              DateTimePicker5.Value.ToString("dd/MM/yyyy") & "'" & " AND " & "'" & DateTimePicker6.Value.ToString("dd/MM/yyyy") & "')", lidhje2)
                Dim dataTable As New DataTable
                adaptori2.Fill(dataTable)
                DataGridView3.DataSource = dataTable

                For i As Integer = 0 To DataGridView3.RowCount - 1
                    total12 = total12 + DataGridView3.Rows(i).Cells(9).Value
                Next
                Label17.Enabled = True
                Label17.Text = total12
                lidhje2.Close()
                Dim total13 As Integer = 0
                Lidhje1.Open()
                adaptor1 = New OleDbDataAdapter("SELECT * FROM Shitjet WHERE(Produkti Like '" &
                                            reader(1).ToString() & "%') AND (Data BETWEEN " & "'" &
                                              DateTimePicker5.Value.ToString("dd/MM/yyyy") & "'" & " AND " & "'" & DateTimePicker6.Value.ToString("dd/MM/yyyy") & "')", Lidhje1)
                Dim dataTable1 As New DataTable
                adaptor1.Fill(dataTable1)
                DataGridView4.DataSource = dataTable1

                For i As Integer = 0 To DataGridView4.RowCount - 1
                    total13 = total13 + DataGridView4.Rows(i).Cells(9).Value
                Next
                Label18.Enabled = True
                Label18.Text = total13
                Lidhje1.Close()
                Label19.Enabled = True
                Label19.Text = Label17.Text - Label18.Text



                Magazina.ListBox1.Items.Add(reader(1).ToString())
                Magazina.ListBox2.Items.Add(Label19.Text)
                Magazina.ListBox3.Items.Add(reader(4).ToString())
                Magazina.ListBox4.Items.Add(reader(2).ToString())
                Magazina.ListBox5.Items.Add(Label19.Text * reader(4).ToString())



                For i As Integer = 0 To Magazina.ListBox1.Items.Count - 1
                    Dim dr As New DataGridViewRow
                    Magazina.DataGridView1.Rows.Add(dr)
                    Magazina.DataGridView1.Rows(i).Cells(1).Value = Magazina.ListBox1.Items(i)
                    Magazina.DataGridView1.Rows(i).Cells(0).Value = i + 1
                    Magazina.DataGridView1.AllowUserToAddRows = False
                Next




                For i As Integer = 0 To Magazina.ListBox2.Items.Count - 1
                    Dim dr As New DataGridViewRow
                    Magazina.DataGridView1.Rows.Add(dr)
                    Magazina.DataGridView1.Rows(i).Cells(3).Value = Magazina.ListBox2.Items(i)
                    Magazina.DataGridView1.AllowUserToAddRows = False
                Next
                For i As Integer = 0 To Magazina.ListBox3.Items.Count - 1
                    Dim dr As New DataGridViewRow
                    Magazina.DataGridView1.Rows.Add(dr)
                    Magazina.DataGridView1.Rows(i).Cells(4).Value = Magazina.ListBox3.Items(i)
                    Magazina.DataGridView1.AllowUserToAddRows = False
                Next
                For i As Integer = 0 To Magazina.ListBox4.Items.Count - 1
                    Dim dr As New DataGridViewRow
                    Magazina.DataGridView1.Rows.Add(dr)
                    Magazina.DataGridView1.Rows(i).Cells(2).Value = Magazina.ListBox4.Items(i)
                    Magazina.DataGridView1.AllowUserToAddRows = False
                Next

                Dim tota As Integer = 0
                Dim l As Integer = 0
                For l = 0 To Magazina.ListBox5.Items.Count - 1
                    Dim dr As New DataGridViewRow
                    Magazina.DataGridView1.Rows.Add(dr)
                    Magazina.DataGridView1.Rows(l).Cells(5).Value = Magazina.ListBox5.Items(l)
                    Magazina.DataGridView1.AllowUserToAddRows = False

                    tota = tota + Magazina.DataGridView1.Rows(l).Cells(5).Value



                Next




                Magazina.DataGridView1.Rows(Magazina.ListBox5.Items.Count - 1 + 1).Cells(5).Value = "Totali: " & tota & " ALL"
                Magazina.Label5.Text = tota & " ALL"




                Me.Hide()
                Magazina.Show()

            End While

            reader.Close()

            Magazina.DataGridView1.Rows(Magazina.ListBox1.Items.Count).Cells(5).Style.BackColor = Color.Green

        End Using
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If RadioButton7.Checked = True Then
            setitedhenavedg31.clear
            setitedhenavedg41.clear
            Magazina.DataGridView2.Visible = False
            Magazina.DataGridView1.Visible = True
            Magazina.Button2.Visible = False
            Label17.Text = 0
            Label18.Text = 0
            Label19.Text = 0
            Dim c1 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
            Dim q1 As String = "SELECT * FROM Produktet ORDER BY ID"
            CreateReader_load_produktet(c1, q1)
            Application.DoEvents()
            Label17.Text = 0
            Label18.Text = 0
            Label19.Text = 0
            Label17.Enabled = False
            Label18.Enabled = False
            Label19.Enabled = False
            TextBox27.Text = ""
            With DataGridView3
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            rreshtiaktualdg31 = 0
            lidhjedg31.Open()
            adaptoridg31 = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", lidhjedg31)
            adaptoridg31.Fill(setitedhenavedg31, "tedhena")
            DataGridView3.DataSource = setitedhenavedg31.Tables(0)
            nr3 = DataGridView3.Rows.Count - 1
            For i = 0 To nr3
                DataGridView3.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView3.Refresh()
            lidhjedg31.Close()
            With DataGridView4
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            rreshtiaktualdg41 = 0
            lidhjedg41.Open()
            adaptoridg41 = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", lidhjedg41)
            adaptoridg41.Fill(setitedhenavedg41, "tedhena")
            DataGridView4.DataSource = setitedhenavedg41.Tables(0)
            nr4 = DataGridView4.Rows.Count - 1
            For i = 0 To nr4
                DataGridView4.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView4.Refresh()
            lidhjedg41.Close()




        ElseIf RadioButton8.Checked = True Then
            setitedhenavedg31.clear
            setitedhenavedg41.clear
            Magazina.Button2.Visible = True
            Magazina.DataGridView2.Visible = True
            Magazina.DataGridView1.Visible = False
            Label17.Text = 0
            Label18.Text = 0
            Label19.Text = 0
            Dim c1 As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
            Dim q1 As String = "SELECT * FROM Produktet ORDER BY ID"
            CreateReader_load_produktet_sasialimit(c1, q1)
            Application.DoEvents()
            Label17.Text = 0
            Label18.Text = 0
            Label19.Text = 0
            Label17.Enabled = False
            Label18.Enabled = False
            Label19.Enabled = False
            TextBox27.Text = ""
            With DataGridView3
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            rreshtiaktualdg31 = 0
            lidhjedg31.Open()
            adaptoridg31 = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", lidhjedg31)
            adaptoridg31.Fill(setitedhenavedg31, "tedhena")
            DataGridView3.DataSource = setitedhenavedg31.Tables(0)
            nr3 = DataGridView3.Rows.Count - 1
            For i = 0 To nr3
                DataGridView3.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView3.Refresh()
            lidhjedg31.Close()
            With DataGridView4
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            rreshtiaktualdg41 = 0
            lidhjedg41.Open()
            adaptoridg41 = New OleDbDataAdapter("SELECT * FROM Shitjet ORDER BY ID", lidhjedg41)
            adaptoridg41.Fill(setitedhenavedg41, "tedhena")
            DataGridView4.DataSource = setitedhenavedg41.Tables(0)
            nr4 = DataGridView4.Rows.Count - 1
            For i = 0 To nr4
                DataGridView4.Rows(i).Cells(0).Value = i + 1
            Next
            DataGridView4.Refresh()
            lidhjedg41.Close()



        End If

    End Sub
    Public Sub CreateReader_load_produktetdate(ByVal connectionString As String,
  ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                If reader(0).ToString() = 1 Then
                    MsgBox("Celja e magazines u be me date: " & reader(2).ToString(), MsgBoxStyle.Information)
                End If
            End While
            reader.Close()
        End Using
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Dim cd As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";"
        Dim qd As String = "SELECT * FROM dhenat ORDER BY ID"
        CreateReader_load_produktetdate(cd, qd)
    End Sub
    Private Sub TabPage5_Click(sender As Object, e As EventArgs) Handles TabPage5.Click
        Application.DoEvents()
    End Sub
    Private Sub DataGridView7_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView7.CellContentClick
        Dim row As DataGridViewRow = DataGridView7.Rows.Item(e.RowIndex)
        TextBox39.Text = row.Cells.Item("ID").Value.ToString
        TextBox38.Text = row.Cells.Item("Kodi_Ofertes").Value.ToString
        TextBox36.Text = row.Cells.Item("Shitesi").Value.ToString
        TextBox35.Text = row.Cells.Item("Bleresi").Value.ToString
        TextBox40.Text = row.Cells.Item("Klient_Ekzistues").Value.ToString
        TextBox34.Text = row.Cells.Item("Produkti").Value.ToString
        TextBox33.Text = row.Cells.Item("Njesia").Value.ToString
        TextBox32.Text = row.Cells.Item("Sasia").Value.ToString
        TextBox31.Text = row.Cells.Item("Cmimi").Value.ToString
        TextBox30.Text = row.Cells.Item("Vlera_Pa_TVSH").Value.ToString
        TextBox29.Text = row.Cells.Item("TVSH").Value.ToString
        TextBox37.Text = row.Cells.Item("Vlera_Me_TVSH").Value.ToString
        TextBox28.Text = row.Cells.Item("Zbritje_ne_perq").Value.ToString
        Button23.Text = "Ruaj Faturen(" & row.Cells.Item("Kodi_Ofertes").Value.ToString & ")"
        Button22.Text = "Printo Faturen(" & row.Cells.Item("Kodi_Ofertes").Value.ToString & ")"
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim Str As String
        Str = "update Ofertat set Shitesi="
        Str += """" & TextBox36.Text & """"
        Str += " where ID="
        Str += TextBox39.Text.Trim()
        lidhje2118.Open()
        query2118 = New OleDbCommand(Str, lidhje2118)
        query2118.ExecuteNonQuery()
        lidhje2118.Close()
        lidhje2118.Open()
        Str = "update Ofertat set Bleresi="
        Str += """" & TextBox35.Text & """"
        Str += " where ID="
        Str += TextBox39.Text.Trim()
        query2118 = New OleDbCommand(Str, lidhje2118)
        query2118.ExecuteNonQuery()
        lidhje2118.Close()
        lidhje2118.Open()
        Str = "update Ofertat set Klient_Ekzistues="
        Str += """" & TextBox40.Text & """"
        Str += " where ID="
        Str += TextBox39.Text.Trim()
        query2118 = New OleDbCommand(Str, lidhje2118)
        query2118.ExecuteNonQuery()
        lidhje2118.Close()
        lidhje2118.Open()
        Str = "update Ofertat set Produkti="
        Str += """" & TextBox34.Text & """"
        Str += " where ID="
        Str += TextBox39.Text.Trim()
        query2118 = New OleDbCommand(Str, lidhje2118)
        query2118.ExecuteNonQuery()
        lidhje2118.Close()
        lidhje2118.Open()
        Str = "update Ofertat set Njesia="
        Str += """" & TextBox33.Text & """"
        Str += " where ID="
        Str += TextBox39.Text.Trim()
        query2118 = New OleDbCommand(Str, lidhje2118)
        query2118.ExecuteNonQuery()
        lidhje2118.Close()
        lidhje2118.Open()
        Str = "update Ofertat set Sasia="
        Str += """" & TextBox32.Text & """"
        Str += " where ID="
        Str += TextBox39.Text.Trim()
        query2118 = New OleDbCommand(Str, lidhje2118)
        query2118.ExecuteNonQuery()
        lidhje2118.Close()
        lidhje2118.Open()
        Str = "update Ofertat set Cmimi="
        Str += """" & TextBox31.Text & """"
        Str += " where ID="
        Str += TextBox39.Text.Trim()
        query2118 = New OleDbCommand(Str, lidhje2118)
        query2118.ExecuteNonQuery()
        lidhje2118.Close()
        lidhje2118.Open()
        Str = "update Ofertat set Vlera_Pa_TVSH="
        Str += """" & TextBox30.Text & """"
        Str += " where ID="
        Str += TextBox39.Text.Trim()
        query2118 = New OleDbCommand(Str, lidhje2118)
        query2118.ExecuteNonQuery()
        lidhje2118.Close()
        lidhje2118.Open()
        Str = "update Ofertat set TVSH="
        Str += """" & TextBox29.Text & """"
        Str += " where ID="
        Str += TextBox39.Text.Trim()
        query2118 = New OleDbCommand(Str, lidhje2118)
        query2118.ExecuteNonQuery()
        lidhje2118.Close()
        lidhje2118.Open()
        Str = "update Ofertat set Vlera_Me_TVSH="
        Str += """" & TextBox37.Text & """"
        Str += " where ID="
        Str += TextBox39.Text.Trim()
        query2118 = New OleDbCommand(Str, lidhje2118)
        query2118.ExecuteNonQuery()
        lidhje2118.Close()
        lidhje2118.Open()
        Str = "update Ofertat set Zbritje_ne_perq="
        Str += """" & TextBox28.Text & """"
        Str += " where ID="
        Str += TextBox39.Text.Trim()
        query2118 = New OleDbCommand(Str, lidhje2118)
        query2118.ExecuteNonQuery()
        lidhje2118.Close()
        setitedhenave2118.Clear()
        adaptori2118 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhje2118)
        adaptori2118.Fill(setitedhenave2118, "tedhena")
        DataGridView7.DataSource = setitedhenave2118.Tables(0)
        nr7 = DataGridView7.Rows.Count - 1
        For i = 0 To nr7
            DataGridView7.Rows(i).Cells(0).Value = i + 1
        Next
        MsgBox("U rifreskua me sukses", MsgBoxStyle.Information)
    End Sub
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If RadioButton6.Checked = True Then
            For Each row As DataGridViewRow In DataGridView7.Rows
                Dim Str As String
                Try
                    Str = "delete from Ofertat where ID="
                    Str += row.Cells(0).Value
                    lidhje2.Open()
                    query2 = New OleDbCommand(Str, lidhje2)
                    query2.ExecuteNonQuery()
                    setidatave1.clear()
                    adaptori2 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhje2)
                    adaptori2.Fill(setidatave1, "tedhena")

                    If rreshtiaktual > 0 Then
                        rreshtiaktual -= 1
                        Merrtedhenat(rreshtiaktual)
                    End If
                    lidhje2.Close()
                    TextBox2.Text = ""
                    MsgBox("U fshi me sukses", MsgBoxStyle.Information)
                Catch ex As Exception
                    MsgBox("Nuk u fshi me sukses", MsgBoxStyle.Information)
                    lidhje2.Close()
                End Try
            Next
        ElseIf RadioButton5.Checked = True Then
            Dim Str As String
            Try
                Str = "delete from Ofertat where ID="
                Str += DataGridView7.CurrentRow.Cells(1).Value.ToString
                lidhje2.Open()
                query2 = New OleDbCommand(Str, lidhje2)
                query2.ExecuteNonQuery()
                setidatave1.clear()
                adaptori2 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", lidhje2)
                adaptori2.Fill(setidatave1, "tedhena")
                DataGridView7.DataSource = setidatave1.tables(0)
                nr7 = DataGridView7.Rows.Count - 1
                For i = 0 To nr7
                    DataGridView7.Rows(i).Cells(0).Value = i + 1
                Next
                If rreshtiaktual > 0 Then
                    rreshtiaktual -= 1
                    Merrtedhenat(rreshtiaktual)
                End If
                lidhje2.Close()
                TextBox2.Text = ""
                MsgBox("U fshi me sukses", MsgBoxStyle.Information)
            Catch ex As Exception
                MsgBox("Nuk u fshi me sukses", MsgBoxStyle.Information)
                lidhje2.Close()
            End Try
        End If
    End Sub
    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        lidhje2v.Close()
        lidhje2v.Open()
        adaptori2v = New OleDbDataAdapter("SELECT * FROM Ofertat WHERE(ID Like '" &
                                        TextBox13.Text & "%' OR Kodi_Ofertes Like '" & TextBox13.Text & "%' OR Data Like '" &
                                        TextBox13.Text & "%' OR Ora Like '" & TextBox13.Text & "%'  OR Shitesi Like '" &
                                        TextBox13.Text & "%' OR Bleresi Like '" & TextBox13.Text & "%' OR Klient_Ekzistues Like '" & TextBox13.Text & "%' OR Produkti Like '" &
                                        TextBox13.Text & "%' OR Njesia Like '" & TextBox13.Text & "%' OR Sasia Like '" &
                                        TextBox13.Text & "%' OR Cmimi Like '" & TextBox13.Text & "%' OR Vlera_Pa_TVSH Like '" &
                                         TextBox13.Text & "%' OR TVSH Like '" & TextBox13.Text & "%' OR Vlera_Me_TVSH Like '" & TextBox13.Text & "%' )", Lidhje.lidhje2v)
        Dim dataTablev As New DataTable
        adaptori2v.Fill(dataTablev)
        DataGridView7.DataSource = dataTablev
        nr7 = DataGridView7.Rows.Count - 1
        For i = 0 To nr7
            DataGridView7.Rows(i).Cells(0).Value = i + 1
        Next
        Merrtedhenat_Ofertat(rreshtiaktualdg71)
        DataGridView7.CurrentCell = Nothing
        If DataGridView7.CurrentCell Is Nothing Then
        Else
            Dim row As DataGridViewRow = Me.DataGridView7.Rows.Item(rreshtiaktual2shitjet)
            TextBox39.Text = row.Cells.Item("ID").Value.ToString
            TextBox38.Text = row.Cells.Item("Kodi_Ofertes").Value.ToString
            TextBox36.Text = row.Cells.Item("Shitesi").Value.ToString
            TextBox35.Text = row.Cells.Item("Bleresi").Value.ToString
            TextBox40.Text = row.Cells.Item("Kient_Ekzistues").Value.ToString
            TextBox34.Text = row.Cells.Item("Produkti").Value.ToString
            TextBox33.Text = row.Cells.Item("Njesia").Value.ToString
            TextBox32.Text = row.Cells.Item("Sasia").Value.ToString
            TextBox31.Text = row.Cells.Item("Cmimi").Value.ToString
            TextBox30.Text = row.Cells.Item("Vlera_Pa_TVSH").Value.ToString
            TextBox29.Text = row.Cells.Item("TVSH").Value.ToString
            TextBox37.Text = row.Cells.Item("Vlera_me_TVSH").Value.ToString
            TextBox28.Text = row.Cells.Item("Zbritje_ne_perq").Value.ToString
            lidhje2v.Close()
            DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            DataGridView2.RefreshEdit()
            lidhje2v.Close()
        End If
        If TextBox13.Text = "" Then
            Label112.Text = 0
            Label113.Text = 0
            Label114.Text = 0
            Label112.ForeColor = DefaultForeColor
            Label113.ForeColor = DefaultForeColor
            Label114.ForeColor = DefaultForeColor
        Else
            Dim vl1, vl2, vl3 As String
            vl1 = (From row As DataGridViewRow In DataGridView7.Rows
                   Where row.Cells(12).FormattedValue.ToString() <> String.Empty
                   Select Convert.ToInt32(row.Cells(12).FormattedValue)).Sum().ToString()
            vl2 = (From row As DataGridViewRow In DataGridView7.Rows
                   Where row.Cells(13).FormattedValue.ToString() <> String.Empty
                   Select Convert.ToInt32(row.Cells(13).FormattedValue)).Sum().ToString()
            vl3 = (From row As DataGridViewRow In DataGridView7.Rows
                   Where row.Cells(14).FormattedValue.ToString() <> String.Empty
                   Select Convert.ToInt32(row.Cells(14).FormattedValue)).Sum().ToString()
            Label112.ForeColor = Color.Red
            Label113.ForeColor = Color.Orange
            Label114.ForeColor = Color.ForestGreen
            Label112.Text = vl1
            Label113.Text = vl2
            Label114.Text = vl3
        End If
    End Sub
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        TextBox13.Text = ""
        Lidhje1.Close()
        lidhje2.Close()
        With DataGridView7
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        rreshtiaktual = 0
        Lidhje1.Open()
        adaptor1 = New OleDbDataAdapter("SELECT * FROM Ofertat ORDER BY ID", Lidhje1)
        adaptor1.Fill(setidatave1, "tedhena")
        DataGridView7.DataSource = setidatave1.Tables(0)
        DataGridView7.Refresh()
        Merrtedhenat_Ofertat(rreshtiaktualdg71)
        Lidhje1.Close()
    End Sub
    Private Sub OferteEReOferteShitjeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OferteEReOferteShitjeToolStripMenuItem.Click
        setitedhenave2ypp.clear
        lidhje2ypp.Open()
        adaptori2ypp = New OleDbDataAdapter("SELECT * FROM dhenat ORDER BY ID", lidhje2ypp)
        adaptori2ypp.Fill(setitedhenave2ypp, "tedhena")
        lidhje2ypp.Close()
        If setitedhenave2ypp.Tables(0).Rows.Count = 0 Then
            MsgBox("Ju nuk keni asnje hyrje ne databaze!Oferta shitja nuk eshte e mundur!", MsgBoxStyle.Information)
        Else
            Me.Hide()
            Ofert.Show()
        End If
    End Sub
    Private Sub KlasatEProdukteveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KlasatEProdukteveToolStripMenuItem.Click
        Klasat_e_produkteve.Show()
        Me.Hide()
    End Sub
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Me.Hide()
        Konfigurime.Show()
    End Sub
    Public lidhjeprint As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Path & "Jet OLEDB:Database Password=" & Hyrje.TextBox3.Text & ";")
    Public adaptori2print As OleDbDataAdapter
    Public lexuesiprint As OleDbDataReader
    Public query2print As OleDbCommand
    Public setitedhenave2print = New DataSet
    Public rreshtiaktual2print As Integer
    Dim data, shitesi, bleresi, kodi_shit, shita, shitc, shitn, bleresia, bleresic, bleresin As String

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        If RadioButton7.Checked = True Then
            Label115.Visible = False
            TextBox41.Visible = False
        Else
            Label115.Visible = True
            TextBox41.Visible = True
        End If
    End Sub

    Private Sub StatistikaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatistikaToolStripMenuItem.Click
        Me.Hide()
        Statistika.Show()
    End Sub
    Dim vlerapatvsh, vleraetvsh, vlerametvsh As String
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        If Not DataGridView2.Rows.Count > 0 Then
            MsgBox("Fatura eshte bosh!Me pare gjeneroni nje fature!", MsgBoxStyle.Information)
        Else
            Try
                Dim regDate As DateTime = Date.Now
                Dim strDate As String = regDate.ToString("dd/MM/yyyy")
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
            Catch ex As Exception
                MsgBox("Dokumenti eshte i hapur.Mbyll dokumentin dhe provo perseri!", MsgBoxStyle.Information)
            End Try
        End If
    End Sub
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
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
                            Dim regDate1 As DateTime = Date.Now
                            Dim strDate1 As String = regDate1.ToString("dd/MM/yyyy")
                            Using outputPdfStream As Stream = New FileStream(Konfigurime.TextBox3.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate1.Replace("/", "-") & "_" & TextBox8.Text & "_Shitje.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
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
                    Dim regDate As DateTime = Date.Now
                    Dim strDate As String = regDate.ToString("dd/MM/yyyy")
                    Dim PrintPDF As New ProcessStartInfo
                    PrintPDF.UseShellExecute = True
                    PrintPDF.Verb = "print"
                    PrintPDF.WindowStyle = ProcessWindowStyle.Hidden
                    PrintPDF.FileName = Konfigurime.TextBox3.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Shitje.pdf"
                    Process.Start(PrintPDF)
                    Threading.Thread.Sleep(20000)
                    killProcess("Acrobat")
                    Threading.Thread.Sleep(10000)
                    My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox3.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Shitje.pdf")
                    MsgBox("Fatura u printua me sukses!", MsgBoxStyle.Information)
                    'System.Diagnostics.Process.Start(Konfigurime.TextBox3.Text & "\" & bleresi.Replace(" ", "_") & "_" & strDate.Replace("/", "-") & "_" & TextBox8.Text & "_Shitje.pdf")
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
End Class