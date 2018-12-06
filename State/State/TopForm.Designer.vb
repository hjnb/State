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
        Me.ippanDataGridView = New State.ExDataGridView(Me.components)
        Me.ryouyouDataGridView = New State.ExDataGridView(Me.components)
        CType(Me.ippanDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ryouyouDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'print
        '
        Me.print.Location = New System.Drawing.Point(357, 640)
        Me.print.Name = "print"
        Me.print.Size = New System.Drawing.Size(75, 31)
        Me.print.TabIndex = 11
        Me.print.Text = "印刷"
        Me.print.UseVisualStyleBackColor = True
        '
        'regist
        '
        Me.regist.Location = New System.Drawing.Point(265, 640)
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
        Me.Label1.Location = New System.Drawing.Point(17, 645)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(221, 12)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "動態　　0:クリア　1:危篤　2:厳重観察　3:観察"
        '
        'ippanDataGridView
        '
        Me.ippanDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ippanDataGridView.Location = New System.Drawing.Point(30, 6)
        Me.ippanDataGridView.Name = "ippanDataGridView"
        Me.ippanDataGridView.RowTemplate.Height = 21
        Me.ippanDataGridView.Size = New System.Drawing.Size(192, 623)
        Me.ippanDataGridView.TabIndex = 12
        '
        'ryouyouDataGridView
        '
        Me.ryouyouDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ryouyouDataGridView.Location = New System.Drawing.Point(221, 6)
        Me.ryouyouDataGridView.Name = "ryouyouDataGridView"
        Me.ryouyouDataGridView.RowTemplate.Height = 21
        Me.ryouyouDataGridView.Size = New System.Drawing.Size(184, 623)
        Me.ryouyouDataGridView.TabIndex = 13
        '
        'TopForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(449, 683)
        Me.Controls.Add(Me.ryouyouDataGridView)
        Me.Controls.Add(Me.ippanDataGridView)
        Me.Controls.Add(Me.print)
        Me.Controls.Add(Me.regist)
        Me.Controls.Add(Me.Label1)
        Me.Name = "TopForm"
        Me.Text = "State"
        CType(Me.ippanDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ryouyouDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents print As System.Windows.Forms.Button
    Friend WithEvents regist As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ippanDataGridView As State.ExDataGridView
    Friend WithEvents ryouyouDataGridView As State.ExDataGridView

End Class
