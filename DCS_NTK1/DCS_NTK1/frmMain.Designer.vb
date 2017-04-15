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
        Me.SysmacCJ.NodeAddress = CType(2, Short)
        Me.SysmacCJ.ReceiveTimeLimit = CType(750, Long)
        Me.SysmacCJ.RetryCount = 0
        Me.SysmacCJ.UnitAddress = CType(0, Short)
        Me.SysmacCJ.UseNameServer = False
        Me.SysmacCJ.UseSpecificPort = False
        '
        'timScan
        '
        Me.timScan.Enabled = True
        Me.timScan.Interval = 250
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
        Me.dgvQu.Location = New System.Drawing.Point(12, 298)
        Me.dgvQu.Name = "dgvQu"
        Me.dgvQu.RowTemplate.Height = 21
        Me.dgvQu.Size = New System.Drawing.Size(240, 150)
        Me.dgvQu.TabIndex = 2
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(893, 673)
        Me.Controls.Add(Me.dgvQu)
        Me.Controls.Add(Me.dgvEq)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "frmMain"
        Me.Text = "DCS_NTK"
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
End Class
