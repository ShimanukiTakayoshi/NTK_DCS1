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
        Me.SysmacCJ1 = New OMRON.Compolet.SYSMAC.SysmacCJ(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'SysmacCJ1
        '
        Me.SysmacCJ1.Active = True
        Me.SysmacCJ1.DeviceName = ""
        Me.SysmacCJ1.HeartBeatTimer = CType(0, Long)
        Me.SysmacCJ1.NetworkAddress = CType(0, Short)
        Me.SysmacCJ1.NodeAddress = CType(2, Short)
        Me.SysmacCJ1.ReceiveTimeLimit = CType(750, Long)
        Me.SysmacCJ1.RetryCount = 0
        Me.SysmacCJ1.UnitAddress = CType(0, Short)
        Me.SysmacCJ1.UseNameServer = False
        Me.SysmacCJ1.UseSpecificPort = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(44, 49)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(148, 19)
        Me.TextBox1.TabIndex = 0
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "frmMain"
        Me.Text = "DCS_NTK"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SysmacCJ1 As OMRON.Compolet.SYSMAC.SysmacCJ
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox

End Class
