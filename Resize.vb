Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

Public Class clsResize
    Private _arr_control_storage As List(Of System.Drawing.Rectangle) = New List(Of System.Drawing.Rectangle)()
    Private showRowHeader As Boolean = False

    Public Sub New(ByVal _form_ As Form)
        form = _form_
        _formSize = _form_.ClientSize
        _fontsize = _form_.Font.Size
    End Sub

    Private Property _fontsize As Single
    Private Property _formSize As System.Drawing.SizeF
    Private Property form As Form

    Public Sub _get_initial_size()
        Dim _controls = _get_all_controls(form)

        For Each control As Control In _controls
            _arr_control_storage.Add(control.Bounds)
            '   If control.[GetType]() = GetType(DataGridView) Then _dgv_Column_Adjust((CType(control, DataGridView)), showRowHeader)
            'control.Refresh()
        Next
    End Sub
    Public Sub _get_initial_size1()
        Dim _controls = _get_all_controls(form)

        For Each control As Control In _controls
            _arr_control_storage.Add(control.Bounds)
            If control.[GetType]() = GetType(DataGridView) Then _dgv_Column_Adjust((CType(control, DataGridView)), showRowHeader)
            control.Refresh()
        Next
    End Sub
    Public Sub _resize()
        Dim _form_ratio_width As Double = CDbl(form.ClientSize.Width) / CDbl(_formSize.Width)
        Dim _form_ratio_height As Double = CDbl(form.ClientSize.Height) / CDbl(_formSize.Height)
        Dim _controls = _get_all_controls(form)
        Dim _pos As Integer = -1

        For Each control As Control In _controls
            _pos += 1
            Dim _controlSize As System.Drawing.Size = New System.Drawing.Size(CInt((_arr_control_storage(_pos).Width * _form_ratio_width)), CInt((_arr_control_storage(_pos).Height * _form_ratio_height)))
            Dim _controlposition As System.Drawing.Point = New System.Drawing.Point(CInt((_arr_control_storage(_pos).X * _form_ratio_width)), CInt((_arr_control_storage(_pos).Y * _form_ratio_height)))
            control.Bounds = New System.Drawing.Rectangle(_controlposition, _controlSize)
            If control.[GetType]() = GetType(DataGridView) Then _dgv_Column_Adjust((CType(control, DataGridView)), showRowHeader)
            control.Font = New System.Drawing.Font(form.Font.FontFamily, CSng((((Convert.ToDouble(_fontsize) * _form_ratio_width) / 2) + ((Convert.ToDouble(_fontsize) * _form_ratio_height) / 2))))
        Next
    End Sub


    Public Sub _resize1()
        Dim _form_ratio_width As Double = CDbl(form.ClientSize.Width) / CDbl(_formSize.Width)
        Dim _form_ratio_height As Double = CDbl(form.ClientSize.Height) / CDbl(_formSize.Height)
        Dim _controls = _get_all_controls(form)
        Dim _pos As Integer = -1

        For Each control As Control In _controls
            _pos += 1
            Dim _controlSize As System.Drawing.Size = New System.Drawing.Size(CInt((_arr_control_storage(_pos).Width * _form_ratio_width)), CInt((_arr_control_storage(_pos).Height * _form_ratio_height)))
            Dim _controlposition As System.Drawing.Point = New System.Drawing.Point(CInt((_arr_control_storage(_pos).X * _form_ratio_width)), CInt((_arr_control_storage(_pos).Y * _form_ratio_height)))
            control.Bounds = New System.Drawing.Rectangle(_controlposition, _controlSize)
            ' If control.[GetType]() = GetType(DataGridView) Then _dgv_Column_Adjust((CType(control, DataGridView)), showRowHeader)
            'control.Font = New System.Drawing.Font(form.Font.FontFamily, CSng((((Convert.ToDouble(_fontsize) * _form_ratio_width) / 2) + ((Convert.ToDouble(_fontsize) * _form_ratio_height) / 2))))
        Next
    End Sub

    Private Sub _dgv_Column_Adjust(ByVal dgv As DataGridView, ByVal _showRowHeader As Boolean)
        Dim intRowHeader As Integer = 0
        Const Hscrollbarwidth As Integer = 5

        If _showRowHeader Then
            intRowHeader = dgv.RowHeadersWidth
        Else
            dgv.RowHeadersVisible = False
        End If

        For i As Integer = 0 To dgv.ColumnCount - 1

            If dgv.Dock = DockStyle.Fill Then
                dgv.Columns(i).Width = ((dgv.Width - intRowHeader) / dgv.ColumnCount)
            Else
                dgv.Columns(i).Width = ((dgv.Width - intRowHeader - Hscrollbarwidth) / dgv.ColumnCount)
            End If
        Next
    End Sub

    Private Shared Function _get_all_controls(ByVal c As Control) As IEnumerable(Of Control)
        Return c.Controls.Cast(Of Control)().SelectMany(Function(item) _get_all_controls(item)).Concat(c.Controls.Cast(Of Control)()).Where(Function(control) control.Name <> String.Empty)
    End Function
End Class
