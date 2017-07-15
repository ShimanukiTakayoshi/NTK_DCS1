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
        Me.dgvEq = New System.Windows.Forms.DataGridView()
        Me.dgvQu = New System.Windows.Forms.DataGridView()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnEnd = New System.Windows.Forms.Button()
        Me.btnData = New System.Windows.Forms.Button()
        Me.timHeartBeat = New System.Windows.Forms.Timer(Me.components)
        Me.btnDebugMode = New System.Windows.Forms.Button()
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
        'dgvEq
        '
        Me.dgvEq.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvEq.Location = New System.Drawing.Point(3, 12)
        Me.dgvEq.Name = "dgvEq"
        Me.dgvEq.RowTemplate.Height = 21
        Me.dgvEq.Size = New System.Drawing.Size(240, 150)
        Me.dgvEq.TabIndex = 1
        '
        'dgvQu
        '
        Me.dgvQu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvQu.Location = New System.Drawing.Point(3, 238)
        Me.dgvQu.Name = "dgvQu"
        Me.dgvQu.RowTemplate.Height = 21
        Me.dgvQu.Size = New System.Drawing.Size(240, 150)
        Me.dgvQu.TabIndex = 2
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(52, 705)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(43, 19)
        Me.TextBox2.TabIndex = 3
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(3, 702)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(43, 22)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "PLC"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(101, 702)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(33, 24)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(222, 705)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(76, 19)
        Me.TextBox3.TabIndex = 6
        Me.TextBox3.Visible = False
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(140, 705)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(76, 19)
        Me.TextBox4.TabIndex = 7
        Me.TextBox4.Visible = False
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(304, 706)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(56, 18)
        Me.btnStart.TabIndex = 8
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnEnd
        '
        Me.btnEnd.Location = New System.Drawing.Point(366, 706)
        Me.btnEnd.Name = "btnEnd"
        Me.btnEnd.Size = New System.Drawing.Size(55, 20)
        Me.btnEnd.TabIndex = 9
        Me.btnEnd.Text = "End"
        Me.btnEnd.UseVisualStyleBackColor = True
        '
        'btnData
        '
        Me.btnData.Location = New System.Drawing.Point(427, 706)
        Me.btnData.Name = "btnData"
        Me.btnData.Size = New System.Drawing.Size(55, 20)
        Me.btnData.TabIndex = 10
        Me.btnData.Text = "DataGet"
        Me.btnData.UseVisualStyleBackColor = True
        '
        'timHeartBeat
        '
        Me.timHeartBeat.Enabled = True
        Me.timHeartBeat.Interval = 500
        '
        'btnDebugMode
        '
        Me.btnDebugMode.Location = New System.Drawing.Point(152, 1)
        Me.btnDebugMode.Name = "btnDebugMode"
        Me.btnDebugMode.Size = New System.Drawing.Size(53, 11)
        Me.btnDebugMode.TabIndex = 11
        Me.btnDebugMode.UseVisualStyleBackColor = True
        Me.btnDebugMode.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1008, 729)
        Me.Controls.Add(Me.btnDebugMode)
        Me.Controls.Add(Me.btnData)
        Me.Controls.Add(Me.btnEnd)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.dgvQu)
        Me.Controls.Add(Me.dgvEq)
        Me.Name = "frmMain"
        Me.Text = "抵抗測定機データ収集"
        CType(Me.dgvEq, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvQu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SysmacCJ As OMRON.Compolet.SYSMAC.SysmacCJ
    Friend WithEvents timScan As System.Windows.Forms.Timer
    Friend WithEvents dgvEq As DataGridView
    Friend WithEvents dgvQu As DataGridView
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents btnStart As Button
    Friend WithEvents btnEnd As Button
    Friend WithEvents btnData As Button
    Friend WithEvents timHeartBeat As System.Windows.Forms.Timer
    Friend WithEvents btnDebugMode As System.Windows.Forms.Button
End Class
