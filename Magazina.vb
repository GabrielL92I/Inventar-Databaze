Imports System.Data.OleDb
Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports Microsoft.Office.Interop
Public Class Magazina
    Dim path As String = My.Settings.ruajdtbpath & "tedhena.accdb;"
    Dim xlApp As Microsoft.Office.Interop.Excel.Application
    Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim Th As Threading.Thread
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
    Public Sub CreateReader_magazina(ByVal connectionString As String,
   ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                ComboBox1.Items.Add(reader(1).ToString())
            End While
            reader.Close()
            ComboBox1.SelectedIndex = 0
        End Using
    End Sub
    Private Sub Loadthings()
        With DataGridView1
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        Dim conmag As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & My.Settings.dtb1 & ";"
        Dim quemag As String = "SELECT * FROM Shitesi ORDER BY ID"
        CreateReader_magazina(conmag, quemag)
        Application.DoEvents()
    End Sub
    Private Sub Magazina_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Konfigurime.TextBox1.Text = My.Settings.backup
        Konfigurime.TextBox2.Text = My.Settings.logo
        Konfigurime.TextBox3.Text = My.Settings.ruajraportet
        Konfigurime.TextBox4.Text = My.Settings.faturatofert
        Konfigurime.TextBox5.Text = My.Settings.ruajgjendjen
        Label4.Text = Farmacia.DateTimePicker6.Value.ToString("dd/MM/yyyy")
        CheckForIllegalCrossThreadCalls = False
        Th = New Threading.Thread(AddressOf Loadthings)
        Th.Start()
    End Sub
    Private Sub Magazina_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Farmacia.Show()
        Try
            Th.Abort()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SaveFileDialog1.Filter = "Excel Files (*.xlsx*)|*.xlsx"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK _
       Then




            Dim sheetIndex As Integer
            Dim Ex As Object
            Dim Wb As Object
            Dim Ws As Object
            Ex = CreateObject("Excel.Application")
            Wb = Ex.workbooks.add
            Dim col, row As Integer
            Dim rawData(DataGridView1.Rows.Count, DataGridView1.Columns.Count - 1) As Object
            For col = 0 To DataGridView1.Columns.Count - 1
                rawData(0, col) = DataGridView1.Columns(col).HeaderText.ToUpper
            Next
            For col = 0 To DataGridView1.Columns.Count - 1
                For row = 0 To DataGridView1.Rows.Count - 1
                    rawData(row + 1, col) = DataGridView1.Rows(row).Cells(col).Value
                Next
            Next
            Dim finalColLetter As String = String.Empty
            finalColLetter = ExcelColName(DataGridView1.Columns.Count)
            sheetIndex += 1
            Ws = Wb.Worksheets(sheetIndex)
            Dim excelRange As String = String.Format("A7:{0}{1}", finalColLetter, DataGridView1.Rows.Count + 1)
            Ws.Range(excelRange, Type.Missing).Value2 = rawData
            Ws = Nothing
            Wb.SaveAs(SaveFileDialog1.FileName, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing)
            Wb.Close(True, Type.Missing, Type.Missing)
            Wb = Nothing
            Ex.Quit()
            Ex = Nothing
            GC.Collect()
            xlApp = GetObject("", "Excel.Application")
            xlBook = xlApp.Workbooks.Open(SaveFileDialog1.FileName)
            xlSheet = xlBook.Worksheets("sheet1")
            xlApp.Visible = True
            xlBook.Sheets("sheet1").Cells(ListBox5.Items.Count - 1 + 9, 5).Value = "Totali:"
            xlBook.Sheets("sheet1").Cells(ListBox5.Items.Count - 1 + 9, 6).Value = Label5.Text
            xlBook.Sheets("sheet1").Cells(ListBox5.Items.Count - 1 + 13, 2).Value = ComboBox1.Text
            xlBook.Sheets("sheet1").Cells(ListBox5.Items.Count - 1 + 13, 2).EntireRow.Font.Bold = True
            xlBook.Sheets("sheet1").Cells(1, 2).ColumnWidth = 32
            xlBook.Sheets("sheet1").Cells(4, 1).ColumnWidth = 12
            xlBook.Sheets("sheet1").Cells(4, 5).ColumnWidth = 10
            xlBook.Sheets("sheet1").Cells(4, 6).ColumnWidth = 10
            xlBook.Sheets("sheet1").Cells(2, 2).Font.Size = 30
            xlBook.Sheets("sheet1").Cells(2, 2).Value = "Inventari"
            xlBook.Sheets("sheet1").Range("A6:F" & ListBox5.Items.Count - 1 + 9).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            xlBook.Sheets("sheet1").Cells(2, 2).EntireRow.Font.Bold = True
            xlBook.Sheets("sheet1").Range(xlBook.Sheets("sheet1").Cells(2, 2), xlBook.Sheets("sheet1").Cells(2, 3)).Merge
            xlBook.Sheets("sheet1").Cells(3, 1).Value = Label1.Text
            xlBook.Sheets("sheet1").Cells(3, 2).Value = ComboBox1.Text
            xlBook.Sheets("sheet1").Cells(4, 1).Value = Label2.Text
            xlBook.Sheets("sheet1").Cells(4, 2).Value = TextBox2.Text
            xlBook.Sheets("sheet1").Cells(3, 5).Value = Label3.Text
            xlBook.Sheets("sheet1").Cells(3, 6).Value = Label4.Text
            xlBook.Sheets("sheet1").Range("A1:H100").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
            xlBook.Sheets("sheet1").Cells(2, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight
            xlBook.Sheets("sheet1").Cells(2, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight
            xlBook.Sheets("sheet1").Cells(ListBox5.Items.Count - 1 + 13, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight
            xlBook.Save()

            xlBook.ActiveSheet.ExportAsFixedFormat(0, SaveFileDialog1.FileName.Replace(".xlsx", ".pdf"))
            xlBook.Close(True)
            xlApp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
            Dim proc As Process
            For Each proc In Process.GetProcessesByName("EXCEL")
                proc.Kill()
            Next






        End If
        MsgBox("U eksportua me sukses.", MsgBoxStyle.Information)
    End Sub
    Public Function ExcelColName(ByVal Col As Integer) As String
        If Col < 0 And Col > 256 Then
            MsgBox("Invalid Argument", MsgBoxStyle.Critical)
            Return Nothing
            Exit Function
        End If
        Dim i As Int16
        Dim r As Int16
        Dim S As String
        If Col <= 26 Then
            S = Chr(Col + 64)
        Else
            r = Col Mod 26
            i = System.Math.Floor(Col / 26)
            If r = 0 Then
                r = 26
                i = i - 1
            End If
            S = Chr(i + 64) & Chr(r + 64)
        End If
        ExcelColName = S
    End Function
    Public Sub CreateReadernipt(ByVal connectionString As String,
   ByVal queryString As String)
        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(queryString, connection)
            connection.Open()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                TextBox2.Text = (reader(4).ToString())
            End While
            reader.Close()
        End Using
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim connipt As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & path & "Jet OLEDB:Database Password=" & My.Settings.dtb1 & ";"
        Dim quenipt As String = "SELECT * FROM Shitesi WHERE(Emri LIKE '" & ComboBox1.Text & "%')"
        CreateReadernipt(connipt, quenipt)
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ComboBox1.DropDownStyle = ComboBoxStyle.DropDown
            TextBox2.ReadOnly = False
        Else
            ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
            TextBox2.ReadOnly = True
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If My.Settings.ruajgjendjen = "" Then
            MsgBox("Ju nuk keni zgjedhur vendodhjen se ku do te ruhen raportet.Shko tek Konfigurimet!", MsgBoxStyle.Information)
        Else
            Dim numri As String = ""
            numri = GenerateRandomString(6)
            Dim regDate As DateTime = Date.Now
            Dim strDate As String = regDate.ToString("dd/MM/yyyy")
            Dim pdfTable As New PdfPTable(DataGridView2.ColumnCount)
            pdfTable.DefaultCell.Padding = 3
            pdfTable.WidthPercentage = 100
            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT
            For Each column As DataGridViewColumn In DataGridView2.Columns
                Dim cell As New PdfPCell(New Phrase(column.HeaderText))
                cell.BorderWidthTop = 1
                cell.BorderWidthBottom = 1
                cell.BackgroundColor = New iTextSharp.text.BaseColor(250, 235, 215)
                pdfTable.AddCell(cell)
            Next
            Dim cellvalue As String = ""
            Dim i As Integer = 0
            For Each row As DataGridViewRow In DataGridView2.Rows
                For Each cell As DataGridViewCell In row.Cells
                    cellvalue = cell.FormattedValue
                    pdfTable.AddCell(Convert.ToString(cellvalue))
                Next
            Next
            Dim folderPath As String = Konfigurime.TextBox5.Text
            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
            End If
            Using stream As New FileStream(folderPath & "\DataGridViewExport.pdf", FileMode.Create)
                Dim pdfDoc As Document = New Document(PageSize.A4, 10.0F, 10.0F, 100, 0F)
                PdfWriter.GetInstance(pdfDoc, stream)
                Dim pdfDest As PdfDestination = New PdfDestination(PdfDestination.XYZ, 0, pdfDoc.PageSize.Height, 1.0F)
                pdfDoc.Open()
                pdfDoc.Add(pdfTable)
                pdfDoc.Close()
                stream.Close()
            End Using
            Dim oldFile As String = Konfigurime.TextBox5.Text & "\DataGridViewExport.pdf"
            Dim newFile As String = Konfigurime.TextBox5.Text & "\" & numri & "_Gjendja_produkteve_" & strDate.Replace("/", "-") & ".pdf"
            Dim reader As New PdfReader(oldFile)
            Dim size As Rectangle = reader.GetPageSizeWithRotation(1)
            Dim document As New Document(size)
            Dim fs As New FileStream(newFile, FileMode.Create, FileAccess.Write)
            Dim writer As PdfWriter = PdfWriter.GetInstance(document, fs)
            document.Open()
            Dim cb As PdfContentByte = writer.DirectContent
            Dim bf As BaseFont = BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)
            cb.SetColorFill(BaseColor.BLACK)
            cb.SetFontAndSize(bf, 20)
            cb.BeginText()
            Dim head As String = "Gjendja e magazines per produktet"
            cb.ShowTextAligned(0, head, 130, 800, 0)
            cb.EndText()
            Dim page As PdfImportedPage = writer.GetImportedPage(reader, 1)
            cb.AddTemplate(page, 0, 0)
            document.Close()
            fs.Close()
            writer.Close()
            reader.Close()
            My.Computer.FileSystem.DeleteFile(Konfigurime.TextBox5.Text & "\DataGridViewExport.pdf")
            System.Diagnostics.Process.Start(Konfigurime.TextBox5.Text & "\" & numri & "_Gjendja_produkteve_" & strDate.Replace("/", "-") & ".pdf")
        End If
    End Sub
    Public Function GenerateRandomString(ByRef iLength As Integer) As String
        Dim rdm As New Random()
        Dim allowChrs() As Char = "123456789".ToCharArray()
        Dim sResult As String = ""
        For i As Integer = 0 To iLength - 1
            sResult += allowChrs(rdm.Next(0, allowChrs.Length))
        Next
        Return sResult
    End Function
End Class