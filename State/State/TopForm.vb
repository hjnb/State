Imports System.Reflection
Imports System.Data.OleDb
Imports System.Runtime.InteropServices

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

    Public Sub New()
        InitializeComponent()

        Me.StartPosition = FormStartPosition.CenterScreen
    End Sub

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
        sql = "select P.Nam as 一般病棟, Int((Format(NOW(),'YYYYMMDD')-Format(P.Birth, 'YYYYMMDD'))/10000) as 年齢, S.State as 動態, P.Birth, P.Kana from [" & DB_Patient & "].UsrM as P left join StateD as S on (P.Nam = S.Nam and P.Nurse = 1) order by P.Kana"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        da.Fill(ds, rs, "ippan")
        ippanDt = ds.Tables("ippan")
        removeDuplicateDataRow(ippanDt)
        fillBlankCell(ippanDt)
        ippanDataGridView.DataSource = ippanDt
        ippanDisplayFlg = True

        '療養病棟部分表示
        ryouyouDt = New DataTable
        sql = "select P.Nam as 療養病棟, Int((Format(NOW(),'YYYYMMDD')-Format(P.Birth, 'YYYYMMDD'))/10000) as 年齢, S.State as 動態, P.Birth, P.Kana from [" & DB_Patient & "].UsrM as P left join StateD as S on (P.Nam = S.Nam  and P.Sanato = 1) order by P.Kana"
        rs.Open(sql, cnn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        da.Fill(ds, rs, "ryouyou")
        ryouyouDt = ds.Tables("ryouyou")
        removeDuplicateDataRow(ryouyouDt)
        fillBlankCell(ryouyouDt)
        ryouyouDataGridView.DataSource = ryouyouDt
        ryouyouDisplayFlg = True

        '表示設定
        settingDatagridview()

        '
        ippanDataGridView.CurrentCell = Nothing
        ryouyouDataGridView.CurrentCell = Nothing

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
            .GridColor = Color.FromKnownColor(KnownColor.Control)
            .BorderStyle = BorderStyle.None
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
            .GridColor = Color.FromKnownColor(KnownColor.Control)
            .BorderStyle = BorderStyle.None
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

    Private Sub removeDuplicateDataRow(dt As DataTable)
        Dim namList As New List(Of String)
        For i As Integer = dt.Rows.Count - 1 To 0 Step -1
            If existsListNam(dt.Rows(i).Item(0), namList) Then
                dt.Rows.RemoveAt(i)
                Continue For
            End If
            namList.Add(dt.Rows(i).Item(0))
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
        sum = Math.Floor(sum / rowCount)
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

    Private Sub ippanDataGridView_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles ippanDataGridView.CellEndEdit
        ippanDataGridView(e.ColumnIndex, e.RowIndex).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        If ippanDisplayFlg = True AndAlso ryouyouDisplayFlg = True Then
            Dim targetName As String = ippanDt.Rows(e.RowIndex).Item(0)
            For Each row As DataRow In ryouyouDt.Rows
                If IsDBNull(row.Item(0)) Then
                    Exit For
                ElseIf row.Item(0) = targetName Then
                    row.Item(2) = ippanDt.Rows(e.RowIndex).Item(2)
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub ryouyouDataGridView_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles ryouyouDataGridView.CellEndEdit
        ryouyouDataGridView(e.ColumnIndex, e.RowIndex).Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        If ippanDisplayFlg = True AndAlso ryouyouDisplayFlg = True Then
            Dim targetName As String = ryouyouDt.Rows(e.RowIndex).Item(0)
            For Each row As DataRow In ippanDt.Rows
                If IsDBNull(row.Item(0)) Then
                    Exit For
                ElseIf row.Item(0) = targetName Then
                    row.Item(2) = ryouyouDt.Rows(e.RowIndex).Item(2)
                    Exit For
                End If
            Next
        End If
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
        If Not IsNothing(ryouyouDataGridView.CurrentCell) Then
            ryouyouDataGridView.CurrentCell.Selected = False
        End If
    End Sub

    Private Sub ryouyouDataGridView_GotFocus(sender As Object, e As System.EventArgs) Handles ryouyouDataGridView.GotFocus
        If Not IsNothing(ippanDataGridView.CurrentCell) Then
            ippanDataGridView.CurrentCell.Selected = False
        End If
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
            tb.Text = "0"
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
                dgv.CurrentCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                'イベントハンドラを追加
                AddHandler tb.KeyDown, AddressOf dataGridViewTextBox_KeyDown
            End If
        End If
    End Sub

    Private Function getInputRowCount(dt As DataTable) As Integer
        Dim count As Integer = 0
        For Each row As DataRow In dt.Rows
            If row(0) = "(平均年齢)" Then
                Exit For
            End If
            count += 1
        Next
        Return count
    End Function

    Private Function existsListNam(nam As String, namList As List(Of String)) As Boolean
        For Each n As String In namList
            If n = nam Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub regist_Click(sender As System.Object, e As System.EventArgs) Handles regist.Click
        Dim cn As New ADODB.Connection()
        cn.Open(DB_State)
        Dim cmd As New ADODB.Command()
        cmd.ActiveConnection = cn

        '削除
        cmd.CommandText = "Delete from StateD"
        cmd.Execute()

        '行数取得
        Dim ippanRowCount As Integer = getInputRowCount(ippanDt)
        Dim ryouyouRowCount As Integer = getInputRowCount(ryouyouDt)

        '登録
        Dim registNamList As New List(Of String)
        Dim rs As New ADODB.Recordset
        rs.Open("StateD", cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)
        '一般病棟
        For i As Integer = 0 To ippanRowCount
            If Not IsDBNull(ippanDt.Rows(i).Item("動態")) AndAlso ippanDt.Rows(i).Item("動態") <> 0 AndAlso Not existsListNam(ippanDt.Rows(i).Item("一般病棟"), registNamList) Then
                registNamList.Add(ippanDt.Rows(i).Item("一般病棟"))
                With rs
                    .AddNew()
                    .Fields("Nam").Value = ippanDt.Rows(i).Item("一般病棟")
                    .Fields("Birth").Value = ippanDt.Rows(i).Item("Birth")
                    .Fields("State").Value = ippanDt.Rows(i).Item("動態")
                End With
            End If
        Next
        '療養病棟
        For i As Integer = 0 To ryouyouRowCount
            If Not IsDBNull(ryouyouDt.Rows(i).Item("動態")) AndAlso ryouyouDt.Rows(i).Item("動態") <> 0 AndAlso Not existsListNam(ryouyouDt.Rows(i).Item("療養病棟"), registNamList) Then
                registNamList.Add(ryouyouDt.Rows(i).Item("療養病棟"))
                With rs
                    .AddNew()
                    .Fields("Nam").Value = ryouyouDt.Rows(i).Item("療養病棟")
                    .Fields("Birth").Value = ryouyouDt.Rows(i).Item("Birth")
                    .Fields("State").Value = ryouyouDt.Rows(i).Item("動態")
                End With
            End If
        Next
        rs.Update()
        cn.Close()

        '再表示
        ippanDataGridView.DataSource = Nothing
        ryouyouDataGridView.DataSource = Nothing
        displayTable()

        MsgBox("登録しました。")
    End Sub

    Private Sub print_Click(sender As System.Object, e As System.EventArgs) Handles print.Click
        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheet As Object

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        objWorkBook = objWorkBooks.Open(excelFilePass)
        oSheet = objWorkBook.Worksheets("動態報告")

        '文字削除処理
        oSheet.Range("F2").value = ""
        oSheet.Range("B5").value = ""
        oSheet.Range("F5").value = ""

        '現在日付
        Dim nowDate As Date = DateTime.Now
        oSheet.Range("F2").value = nowDate.ToString("yyyy/MM/dd")

        '一般病棟部分
        Dim ippanCharArray As String() = {"C", "D", "E"}
        Dim ippanRowCount As Integer = 0
        For Each row As DataRow In ippanDt.Rows
            If row.Item("一般病棟") = "(平均年齢)" Then
                Exit For
            End If
            oSheet.Range("B" & (5 + ippanRowCount)).value = row.Item("一般病棟")
            If Not IsDBNull(row.Item("動態")) AndAlso row.Item("動態") <> 0 Then
                oSheet.Range(ippanCharArray((row.Item("動態") - 1)) & (5 + ippanRowCount)).value = "※"
            End If
            ippanRowCount += 1
        Next

        '療養病棟部分
        Dim ryouyouCharArray As String() = {"G", "H", "I"}
        Dim ryouyouRowCount As Integer = 0
        For Each row As DataRow In ryouyouDt.Rows
            If row.Item("療養病棟") = "(平均年齢)" Then
                Exit For
            End If
            oSheet.Range("F" & (5 + ryouyouRowCount)).value = row.Item("療養病棟")
            If Not IsDBNull(row.Item("動態")) AndAlso row.Item("動態") <> 0 Then
                oSheet.Range(ryouyouCharArray((row.Item("動態") - 1)) & (5 + ryouyouRowCount)).value = "※"
            End If
            ryouyouRowCount += 1
        Next


        '変更保存確認ダイアログ非表示
        objExcel.DisplayAlerts = False

        '印刷 or 印刷プレビュー
        If rbtnPreview.Checked Then
            objExcel.Visible = True
            oSheet.PrintPreview(1)
        ElseIf rbtnPrint.Checked Then
            oSheet.printOut()
        End If

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing
        objWorkBook = Nothing
        objExcel = Nothing
    End Sub
End Class
