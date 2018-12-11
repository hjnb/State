Public Class ExDataGridView
    Inherits DataGridView

    Protected Overrides Function ProcessKeyEventArgs(ByRef m As System.Windows.Forms.Message) As Boolean
        Dim code As Integer = CInt(m.WParam)
        If code = Keys.Left OrElse code = Keys.Right OrElse code = Keys.Up OrElse code = Keys.Down OrElse code = Keys.Enter OrElse (Keys.NumPad0 <= code AndAlso code <= Keys.NumPad3) OrElse (Keys.D0 <= code AndAlso code <= Keys.D3) Then
            Return MyBase.ProcessKeyEventArgs(m)
        Else
            m.WParam = Keys.F2
            Return MyBase.ProcessKeyEventArgs(m)
        End If
    End Function

    Private Sub ExDataGridView_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles Me.CellPainting
        '選択したセルに枠を付ける
        If e.ColumnIndex >= 0 AndAlso e.RowIndex >= 0 AndAlso (e.PaintParts And DataGridViewPaintParts.Background) = DataGridViewPaintParts.Background Then
            e.Graphics.FillRectangle(New SolidBrush(e.CellStyle.BackColor), e.CellBounds)

            If (e.PaintParts And DataGridViewPaintParts.SelectionBackground) = DataGridViewPaintParts.SelectionBackground AndAlso (e.State And DataGridViewElementStates.Selected) = DataGridViewElementStates.Selected Then
                e.Graphics.DrawRectangle(New Pen(Color.Black, 2I), e.CellBounds.X + 1I, e.CellBounds.Y + 1I, e.CellBounds.Width - 3I, e.CellBounds.Height - 3I)
            End If

            Dim pParts As DataGridViewPaintParts = e.PaintParts And Not DataGridViewPaintParts.Background
            e.Paint(e.ClipBounds, pParts)
            e.Handled = True
        End If
    End Sub


End Class
