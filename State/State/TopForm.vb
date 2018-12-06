Imports System.Reflection
Imports System.Data.OleDb

Public Class TopForm

    'データベースのパス
    'Public dbFilePath As String = "\\PRIMERGYTX100S1\State\State.mdb"
    Public dbFilePath As String = "\\PRIMERGYTX100S1\Hakojun\事務\さかもと\State-動態-\State.mdb"
    Public DB_State As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFilePath
    Public DB_Patient As String = "\\PRIMERGYTX100S1\Patient\Patient.mdb"

    'エクセルのパス
    Public excelFilePass As String = "\\PRIMERGYTX100S1\State\State.xls"

    '.iniファイルのパス
    Public iniFilePath As String = "\\PRIMERGYTX100S1\State\State.ini"

    Private Const MAX_ROW_COUNT As Integer = 50

    Private disableCellStyle As DataGridViewCellStyle
    Private enableCellStyle As DataGridViewCellStyle

    Private ippanDt As DataTable
    Private ryouyouDt As DataTable

    Private ippanDisplayFlg As Boolean = False
    Private ryouyouDisplayFlg As Boolean = False

    '行ヘッダーのカレントセルを表す三角マークを非表示に設定する為のクラス。
    Public Class dgvRowHeaderCell

        'DataGridViewRowHeaderCell を継承
        Inherits DataGridViewRowHeaderCell

        'DataGridViewHeaderCell.Paint をオーバーライドして行ヘッダーを描画
        Protected Overrides Sub Paint(ByVal graphics As Graphics, ByVal clipBounds As Rectangle, _
           ByVal cellBounds As Rectangle, ByVal rowIndex As Integer, ByVal cellState As DataGridViewElementStates, _
           ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, _
           ByVal cellStyle As DataGridViewCellStyle, ByVal advancedBorderStyle As DataGridViewAdvancedBorderStyle, _
           ByVal paintParts As DataGridViewPaintParts)
            '標準セルの描画からセル内容の背景だけ除いた物を描画(-5)
            MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, _
                     formattedValue, errorText, cellStyle, advancedBorderStyle, _
                     Not DataGridViewPaintParts.ContentBackground)
        End Sub

    End Class

    Private Sub topForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'セルスタイル設定
        disableCellStyle = New DataGridViewCellStyle()
        disableCellStyle.BackColor = Color.FromKnownColor(KnownColor.Control)
        disableCellStyle.SelectionBackColor = Color.FromKnownColor(KnownColor.Control)
        disableCellStyle.SelectionForeColor = Color.Black

        'datagridview表示前設定
        ippanDataGridView.RowTemplate.HeaderCell = New dgvRowHeaderCell() '行ヘッダの三角マークを非表示に
        ippanDataGridView.RowTemplate.Height = 15 '行の高さ
        ryouyouDataGridView.RowTemplate.Height = 15 '行の高さ

        'dgv表示
        displayTable()
    End Sub

    Private Sub displayTable()
        Dim cnn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim sql As String = ""
        Dim da As OleDbDataAdapter = New OleDbDataAdapter()
        Dim ds As DataSet = New DataSet()
        cnn.Open(DB_State)

        '一般病棟部分表示
        ippanDt = New DataTable
        sql = "select P.Nam as 一般病棟, Int((Format(NOW(),'YYYYMMDD')-Format(P.Birth, 'YYYYMMDD'))/10000) as 年齢, S.State as 動態, P.Birth, P.Kana from [" & DB_Patient & "].UsrM as P left join StateD as S on (P.Nam = S.Nam and P.Nurse = 1 and P.Sanato = 0) order by P.Kana"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        da.Fill(ds, rs, "ippan")
        ippanDt = ds.Tables("ippan")
        fillBlankCell(ippanDt)
        ippanDataGridView.DataSource = ippanDt
        ippanDisplayFlg = True

        '療養病棟部分表示
        ryouyouDt = New DataTable
        sql = "select P.Nam as 療養病棟, Int((Format(NOW(),'YYYYMMDD')-Format(P.Birth, 'YYYYMMDD'))/10000) as 年齢, S.State as 動態, P.Birth, P.Kana from [" & DB_Patient & "].UsrM as P left join StateD as S on (P.Nam = S.Nam and P.Sanato = 1) order by P.Kana"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        da.Fill(ds, rs, "ryouyou")
        ryouyouDt = ds.Tables("ryouyou")
        fillBlankCell(ryouyouDt)
        ryouyouDataGridView.DataSource = ryouyouDt
        ryouyouDisplayFlg = True

        '表示設定
        settingDatagridview()

    End Sub

    Private Sub settingDatagridview()
        '一般
        '並び替えができないようにする
        For Each c As DataGridViewColumn In ippanDataGridView.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        With ippanDataGridView
            .ScrollBars = ScrollBars.None 'スクロールバーを非表示
            .AllowUserToAddRows = False '行追加禁止
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .EnableHeadersVisualStyles = False
            .MultiSelect = False
            .RowHeadersWidth = 27
            .SelectionMode = DataGridViewSelectionMode.CellSelect
            .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .ShowCellToolTips = False
            .DefaultCellStyle.SelectionForeColor = Color.Black
            With .Columns("一般病棟")
                .Width = 90
                .ReadOnly = True
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            With .Columns("年齢")
                .Width = 37
                .ReadOnly = True
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
            With .Columns("動態")
                .Width = 37
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
            .Columns("Birth").Visible = False
            .Columns("Kana").Visible = False
        End With
        setReadonlyCell(ippanDataGridView)

        '療養
        '並び替えができないようにする
        For Each c As DataGridViewColumn In ryouyouDataGridView.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        With ryouyouDataGridView
            .AllowUserToAddRows = False '行追加禁止
            .RowHeadersVisible = False '行ヘッダ非表示
            .AllowUserToResizeColumns = False '列の幅をユーザーが変更できないようにする
            .AllowUserToResizeRows = False '行の高さをユーザーが変更できないようにする
            .AllowUserToDeleteRows = False
            .EnableHeadersVisualStyles = False
            .MultiSelect = False
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .BackgroundColor = Color.FromKnownColor(KnownColor.Control)
            .ShowCellToolTips = False
            .DefaultCellStyle.SelectionForeColor = Color.Black
            With .Columns("療養病棟")
                .Width = 90
                .ReadOnly = True
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            With .Columns("年齢")
                .Width = 37
                .ReadOnly = True
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
            With .Columns("動態")
                .Width = 37
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
            .Columns("Birth").Visible = False
            .Columns("Kana").Visible = False
        End With
        setReadonlyCell(ryouyouDataGridView)
    End Sub

    Private Sub setReadonlyCell(dgv As DataGridView)
        Dim startIndex As Integer = MAX_ROW_COUNT
        For i As Integer = 0 To MAX_ROW_COUNT - 1
            If dgv(0, i).Value = "(平均年齢)" Then
                startIndex = i
                Exit For
            End If
        Next

        For i As Integer = startIndex To MAX_ROW_COUNT - 1
            With dgv("動態", i)
                .ReadOnly = True
                .Style = disableCellStyle
            End With
        Next

    End Sub

    Private Sub fillBlankCell(dt As DataTable)
        Dim rowCount As Integer = dt.Rows.Count
        '既にデータが50件あれば終了
        If rowCount = MAX_ROW_COUNT Then
            Return
        End If

        '平均年齢を計算し行追加
        Dim row As DataRow
        Dim sum As Integer = 0
        For i As Integer = 0 To rowCount - 1
            sum += dt.Rows(i).Item(1)
        Next
        sum = sum / rowCount
        row = dt.NewRow()
        row.Item(0) = "(平均年齢)"
        row.Item(1) = sum
        row.Item(2) = DBNull.Value
        dt.Rows.Add(row)
        rowCount += 1

        '50件になるまで空の行データを作成し追加
        For i As Integer = rowCount To MAX_ROW_COUNT - 1
            row = dt.NewRow()
            row.Item(0) = DBNull.Value '名前
            row.Item(1) = DBNull.Value '年齢
            row.Item(2) = DBNull.Value '動態
            dt.Rows.Add(row)
        Next

    End Sub

    Private Sub ippanDataGridView_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles ippanDataGridView.CellPainting
        '列ヘッダーかどうか調べる
        If e.ColumnIndex < 0 And e.RowIndex >= 0 Then
            'セルを描画する
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All)

            '行番号を描画する範囲を決定する
            'e.AdvancedBorderStyleやe.CellStyle.Paddingは無視しています
            Dim indexRect As Rectangle = e.CellBounds
            indexRect.Inflate(-2, -2)
            '行番号を描画する
            TextRenderer.DrawText(e.Graphics, _
                (e.RowIndex + 1).ToString(), _
                e.CellStyle.Font, _
                indexRect, _
                e.CellStyle.ForeColor, _
                TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            '描画が完了したことを知らせる
            e.Handled = True
        Else
            If e.ColumnIndex = 0 AndAlso e.RowIndex >= 0 Then
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None
                ippanDataGridView(e.ColumnIndex, e.RowIndex).Style = disableCellStyle
            ElseIf e.ColumnIndex = 1 AndAlso e.RowIndex >= 0 Then
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None
                ippanDataGridView(e.ColumnIndex, e.RowIndex).Style = disableCellStyle
            End If
        End If
    End Sub

    Private Sub ryouyouDataGridView_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles ryouyouDataGridView.CellPainting
        If e.ColumnIndex = 0 AndAlso e.RowIndex >= 0 Then
            e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None
            e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None
            ryouyouDataGridView(e.ColumnIndex, e.RowIndex).Style = disableCellStyle
        ElseIf e.ColumnIndex = 1 AndAlso e.RowIndex >= 0 Then
            e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None
            ryouyouDataGridView(e.ColumnIndex, e.RowIndex).Style = disableCellStyle
        End If
    End Sub

    Private Sub ippanDataGridView_GotFocus(sender As Object, e As System.EventArgs) Handles ippanDataGridView.GotFocus
        ryouyouDataGridView.CurrentCell.Selected = False
    End Sub

    Private Sub ryouyouDataGridView_GotFocus(sender As Object, e As System.EventArgs) Handles ryouyouDataGridView.GotFocus
        ippanDataGridView.CurrentCell.Selected = False
    End Sub

    Private Sub ippanDataGridView_MouseWheel(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles ippanDataGridView.MouseWheel
        Dim currentRowIndex As Integer = ippanDataGridView.FirstDisplayedScrollingRowIndex
        If e.Delta < 0 AndAlso (currentRowIndex <= MAX_ROW_COUNT - 3) Then
            ippanDataGridView.FirstDisplayedScrollingRowIndex += 3
        ElseIf e.Delta > 0 Then
            If currentRowIndex >= 3 Then
                ippanDataGridView.FirstDisplayedScrollingRowIndex -= 3
            ElseIf 0 < currentRowIndex AndAlso currentRowIndex <= 2 Then
                ippanDataGridView.FirstDisplayedScrollingRowIndex = 0
            End If
        End If
    End Sub

    Private Sub DataGridView_Scroll(sender As Object, e As System.Windows.Forms.ScrollEventArgs) Handles ippanDataGridView.Scroll, ryouyouDataGridView.Scroll
        If sender Is ippanDataGridView Then
            ryouyouDataGridView.FirstDisplayedScrollingRowIndex = ippanDataGridView.FirstDisplayedScrollingRowIndex
            ryouyouDataGridView.FirstDisplayedScrollingColumnIndex = ippanDataGridView.FirstDisplayedScrollingColumnIndex
        ElseIf sender Is ryouyouDataGridView Then
            ippanDataGridView.FirstDisplayedScrollingRowIndex = ryouyouDataGridView.FirstDisplayedScrollingRowIndex
            ippanDataGridView.FirstDisplayedScrollingColumnIndex = ryouyouDataGridView.FirstDisplayedScrollingColumnIndex
        End If
    End Sub

    Private Sub dataGridViewTextBox_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
        Dim tb As TextBox = CType(sender, TextBox)

        '数字(0～3)まで入力可(delete or backspace は0を入力)
        If e.KeyCode = Keys.Back OrElse e.KeyCode = Keys.Delete Then
            tb.Text = 0
            e.SuppressKeyPress = True
        ElseIf Keys.D0 <= e.KeyCode AndAlso e.KeyCode <= Keys.D3 Then
            tb.Text = Chr(e.KeyCode)
            e.SuppressKeyPress = True
        ElseIf Keys.NumPad0 <= e.KeyCode AndAlso e.KeyCode <= Keys.NumPad3 Then
            tb.Text = Chr(e.KeyCode - 48)
            e.SuppressKeyPress = True
        Else
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub dataGridView_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles ippanDataGridView.EditingControlShowing, ryouyouDataGridView.EditingControlShowing
        If TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
            Dim dgv As DataGridView = CType(sender, DataGridView)
            Dim tb As DataGridViewTextBoxEditingControl = CType(e.Control, DataGridViewTextBoxEditingControl)

            tb.ImeMode = Windows.Forms.ImeMode.Disable
            tb.MaxLength = 1

            'イベントハンドラを削除
            RemoveHandler tb.KeyDown, AddressOf dataGridViewTextBox_KeyDown

            '該当する列か調べる
            If dgv.CurrentCell.OwningColumn.Name = "動態" Then
                'イベントハンドラを追加
                AddHandler tb.KeyDown, AddressOf dataGridViewTextBox_KeyDown
            End If
        End If
    End Sub
End Class
