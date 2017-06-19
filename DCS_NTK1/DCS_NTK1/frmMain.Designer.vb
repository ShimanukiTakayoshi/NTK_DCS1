<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.SysmacCJ = New OMRON.Compolet.SYSMAC.SysmacCJ(Me.components)
        Me.timScan = New System.Windows.Forms.Timer(Me.components)
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.dgvEq = New System.Windows.Forms.DataGridView()
        Me.dgvQu = New System.Windows.Forms.DataGridView()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        CType(Me.dgvEq, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvQu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SysmacCJ
        '
        Me.SysmacCJ.Active = True
        Me.SysmacCJ.DeviceName = ""
        Me.SysmacCJ.HeartBeatTimer = CType(0, Long)
        Me.SysmacCJ.NetworkAddress = CType(0, Short)
        Me.SysmacCJ.NodeAddress = CType(1, Short)
        Me.SysmacCJ.ReceiveTimeLimit = CType(750, Long)
        Me.SysmacCJ.RetryCount = 0
        Me.SysmacCJ.UnitAddress = CType(0, Short)
        Me.SysmacCJ.UseNameServer = False
        Me.SysmacCJ.UseSpecificPort = False
        '
        'timScan
        '
        Me.timScan.Enabled = True
        Me.timScan.Interval = 1000
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(12, 642)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(148, 19)
        Me.TextBox1.TabIndex = 0
        '
        'dgvEq
        '
        Me.dgvEq.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvEq.Location = New System.Drawing.Point(12, 12)
        Me.dgvEq.Name = "dgvEq"
        Me.dgvEq.RowTemplate.Height = 21
        Me.dgvEq.Size = New System.Drawing.Size(240, 150)
        Me.dgvEq.TabIndex = 1
        '
        'dgvQu
        '
        Me.dgvQu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvQu.Location = New System.Drawing.Point(12, 238)
        Me.dgvQu.Name = "dgvQu"
        Me.dgvQu.RowTemplate.Height = 21
        Me.dgvQu.Size = New System.Drawing.Size(240, 150)
        Me.dgvQu.TabIndex = 2
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(960, 318)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(43, 19)
        Me.TextBox2.TabIndex = 3
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(960, 283)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(43, 29)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "PLC"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(969, 354)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(33, 24)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(947, 426)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(76, 19)
        Me.TextBox3.TabIndex = 6
        Me.TextBox3.Visible = False
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(947, 384)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(76, 19)
        Me.TextBox4.TabIndex = 7
        Me.TextBox4.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1026, 673)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.dgvQu)
        Me.Controls.Add(Me.dgvEq)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "frmMain"
        Me.Text = "抵抗測定機データ収集"
        CType(Me.dgvEq, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvQu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SysmacCJ As OMRON.Compolet.SYSMAC.SysmacCJ
    Friend WithEvents timScan As System.Windows.Forms.Timer
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents dgvEq As DataGridView
    Friend WithEvents dgvQu As DataGridView
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
End Class
