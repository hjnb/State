<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TopForm
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.print = New System.Windows.Forms.Button()
        Me.regist = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.rbtnPreview = New System.Windows.Forms.RadioButton()
        Me.rbtnPrint = New System.Windows.Forms.RadioButton()
        Me.blueLine = New System.Windows.Forms.Label()
        Me.ryouyouDataGridView = New State.ExDataGridView(Me.components)
        Me.ippanDataGridView = New State.ExDataGridView(Me.components)
        CType(Me.ryouyouDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ippanDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'print
        '
        Me.print.Location = New System.Drawing.Point(317, 678)
        Me.print.Name = "print"
        Me.print.Size = New System.Drawing.Size(75, 31)
        Me.print.TabIndex = 11
        Me.print.Text = "印刷"
        Me.print.UseVisualStyleBackColor = True
        '
        'regist
        '
        Me.regist.Location = New System.Drawing.Point(225, 678)
        Me.regist.Name = "regist"
        Me.regist.Size = New System.Drawing.Size(75, 31)
        Me.regist.TabIndex = 10
        Me.regist.Text = "登録"
        Me.regist.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(16, 645)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(221, 12)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "動態　　0:クリア　1:危篤　2:厳重観察　3:観察"
        '
        'rbtnPreview
        '
        Me.rbtnPreview.AutoSize = True
        Me.rbtnPreview.Location = New System.Drawing.Point(323, 643)
        Me.rbtnPreview.Name = "rbtnPreview"
        Me.rbtnPreview.Size = New System.Drawing.Size(67, 16)
        Me.rbtnPreview.TabIndex = 14
        Me.rbtnPreview.Text = "プレビュー"
        Me.rbtnPreview.UseVisualStyleBackColor = True
        '
        'rbtnPrint
        '
        Me.rbtnPrint.AutoSize = True
        Me.rbtnPrint.Checked = True
        Me.rbtnPrint.Location = New System.Drawing.Point(259, 643)
        Me.rbtnPrint.Name = "rbtnPrint"
        Me.rbtnPrint.Size = New System.Drawing.Size(47, 16)
        Me.rbtnPrint.TabIndex = 0
        Me.rbtnPrint.TabStop = True
        Me.rbtnPrint.Text = "印刷"
        Me.rbtnPrint.UseVisualStyleBackColor = True
        '
        'blueLine
        '
        Me.blueLine.BackColor = System.Drawing.Color.Blue
        Me.blueLine.Location = New System.Drawing.Point(206, 6)
        Me.blueLine.Name = "blueLine"
        Me.blueLine.Size = New System.Drawing.Size(1, 623)
        Me.blueLine.TabIndex = 16
        Me.blueLine.Text = "Label2"
        '
        'ryouyouDataGridView
        '
        Me.ryouyouDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ryouyouDataGridView.Location = New System.Drawing.Point(206, 6)
        Me.ryouyouDataGridView.Name = "ryouyouDataGridView"
        Me.ryouyouDataGridView.RowTemplate.Height = 21
        Me.ryouyouDataGridView.Size = New System.Drawing.Size(184, 623)
        Me.ryouyouDataGridView.TabIndex = 13
        '
        'ippanDataGridView
        '
        Me.ippanDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ippanDataGridView.Location = New System.Drawing.Point(15, 6)
        Me.ippanDataGridView.Name = "ippanDataGridView"
        Me.ippanDataGridView.RowTemplate.Height = 21
        Me.ippanDataGridView.Size = New System.Drawing.Size(193, 623)
        Me.ippanDataGridView.TabIndex = 12
        '
        'TopForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(409, 720)
        Me.Controls.Add(Me.blueLine)
        Me.Controls.Add(Me.rbtnPrint)
        Me.Controls.Add(Me.rbtnPreview)
        Me.Controls.Add(Me.ryouyouDataGridView)
        Me.Controls.Add(Me.ippanDataGridView)
        Me.Controls.Add(Me.print)
        Me.Controls.Add(Me.regist)
        Me.Controls.Add(Me.Label1)
        Me.Name = "TopForm"
        Me.Text = "State"
        CType(Me.ryouyouDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ippanDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents print As System.Windows.Forms.Button
    Friend WithEvents regist As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ippanDataGridView As State.ExDataGridView
    Friend WithEvents ryouyouDataGridView As State.ExDataGridView
    Friend WithEvents rbtnPreview As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnPrint As System.Windows.Forms.RadioButton
    Friend WithEvents blueLine As System.Windows.Forms.Label

End Class
