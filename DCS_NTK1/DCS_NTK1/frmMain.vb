Public Class frmMain

    'PLC通信アドレス設定
    Public AckAddress As Long = 0           'PLCへ受信OK返答
    Public StartTriggerAdress As Long = 0   'ｽﾀｰﾄﾄﾘｶﾞ
    Public EndTriggerAdress As Long = 0     'ｴﾝﾄﾞﾄﾘｶﾞ
    Public ElementNoAddress As Long = 0     '素子品番
    Public LotNoAddress As Long = 0         'ﾒｯｷﾛｯﾄ
    Public OperatorAddress As Long = 0      '作業者
    Public StartTimeAddress As Long = 0     '仕掛時間
    Public EndTimeAddress As Long = 0       '完了時間
    Public ProbeAddress As Long = 0         'ﾌﾟﾛｰﾌﾞ使用回数先頭ｱﾄﾞﾚｽ


    Public ElementNo As String = ""
    Public LotNo As String = ""
    Public OperatorNo As String = ""
    Public StartTime As String = ""
    Public EndTime As String = ""
    Public ProbeData(9） As Long

    Public PlcReadingFlag As Boolean = False

    Public DebugFlag As Boolean = True
    Public tmp0 As Long = 0

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        initialize
    End Sub

    Private Sub initialize()
        DGVClear(dgvEq)
        Me.Width = 1024
        Me.Height = 768
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        dgvEq.Width = 900
        dgvEq.Height = 200
        Dim cstyle1 As New DataGridViewCellStyle
        cstyle1.Alignment = DataGridViewContentAlignment.MiddleRight
        Dim columnHeaderStyle As DataGridViewCellStyle = dgvEq.ColumnHeadersDefaultCellStyle
        dgvEq.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        columnHeaderStyle.Font = New Font("ＭＳ ゴシック", 8)
        dgvEq.Columns.Add("0", "素子品番")
        dgvEq.Columns.Add("1", "ﾒｯｷﾛｯﾄ" & vbCrLf & "No.  ")
        dgvEq.Columns.Add("2", "作業者")
        dgvEq.Columns.Add("3", "仕掛時間")
        dgvEq.Columns.Add("4", "完了時間")
        dgvEq.Columns.Add("5", "検知抵抗" & vbCrLf & "上ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        dgvEq.Columns.Add("6", "検知抵抗" & vbCrLf & "横ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        dgvEq.Columns.Add("7", "検知抵抗" & vbCrLf & "下ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        dgvEq.Columns.Add("8", "全長抵抗①" & vbCrLf & "上ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        dgvEq.Columns.Add("9", "全長抵抗①" & vbCrLf & "下ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        dgvEq.Columns.Add("10", "全長抵抗②" & vbCrLf & "上ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        dgvEq.Columns.Add("11", "全長抵抗②" & vbCrLf & "横ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        dgvEq.Columns.Add("12", "全長抵抗②" & vbCrLf & "斜めﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        For i As Integer = 0 To 12
            dgvEq.Columns(i).DefaultCellStyle = cstyle1
            dgvEq.Columns(i).Width = 90
        Next i
        For i As Integer = 0 To 9
            dgvEq.Rows.Add("")
        Next
        dgvEq.RowHeadersVisible = False
        dgvEq.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvEq.CurrentCell = Nothing         '選択されているセルをなくす
    End Sub


    Private Sub timScan_Tick(sender As Object, e As EventArgs) Handles timScan.Tick
        If Not PlcReadingFlag Then
            If PlcRead(StartTriggerAdress) <> 0 Then
                PlcWrite(StartTimeAddress, 0)
                If Not DebugFlag Then
                    StartProcess()
                Else
                    StartProcessDebug()
                End If
                DrawChartSetubi
            End If
        End If
    End Sub

    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles timScan.Tick
    '    tmp0 = PlcRead(10301)
    '    TextBox1.Text = Str(tmp0)
    'End Sub

    Public Sub StartProcess()
        PlcReadingFlag = True
        ElementNo = PlcReadStrings(ElementNoAddress, 8)
        LotNo = PlcReadStrings(LotNoAddress, 8)
        OperatorNo = PlcReadStrings(LotNoAddress, 8)
        StartTime = Trim(CStr(Now))
        PlcReadingFlag = False
    End Sub

    Public Sub StartProcessDebug()
        PlcReadingFlag = True
        ElementNo = "Element0"
        LotNo = "Lot00000"
        OperatorNo = "Operator"
        StartTime = Trim(CStr(Now))
        PlcReadingFlag = False
    End Sub

    Public Sub DrawChartSetubi()

    End Sub

    Public Function PlcRead(address As Long) As Long
        '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
        'シーケンサーのDMメモリーより「address」にて指定したアドレスの内容を読み込む
        '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
        If Not DebugFlag Then
            Try
                PlcRead = SysmacCJ.DM(address)
            Catch ex As Exception
                If MsgBox("PLC－PC通信ｴﾗｰ" & vbCr & "DCSを終了してよいですか？", CType(vbOKCancel + vbExclamation, MsgBoxStyle)) = vbOK Then
                    Application.Exit()
                End If
                Application.Exit()
                PlcRead = 0
            End Try
        Else
            PlcRead = 99
        End If
    End Function

    Public Sub PlcWrite(address As Long, value As Long)
        '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
        'シーケンサーのDMメモリーへ「address」にて指定したアドレスに｢value｣の値を書き込む
        '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
        If Not DebugFlag Then
            Try
                SysmacCJ.DM(address) = value
            Catch ex As Exception
                If MsgBox("PLC－PC通信ｴﾗｰ" & vbCr & "DCSを終了してよいですか？", CType(vbOKCancel + vbExclamation, MsgBoxStyle)) = vbOK Then
                    Application.Exit()
                End If
            End Try
        End If
    End Sub

    Public Function PlcReadStrings(address As Long, length As Integer) As String
        '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
        'シーケンサーのDMメモリーより「address」にて指定したアドレスの内容を「length」文字数分のアスキーデータを読み込む
        '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
        Dim tmp1(99) As Integer
        Dim tmp2(99) As String
        PlcReadStrings = ""
        If Not DebugFlag Then
            Try
                tmp1 = SysmacCJ.ReadMemoryWordInteger(OMRON.Compolet.SYSMAC.SysmacCJ.MemoryTypes.DM, address, length, OMRON.Compolet.SYSMAC.SysmacCJ.DataTypes.BIN)
                For i As Integer = 0 To length
                    tmp2(i) = Hex(tmp1(i))
                    PlcReadStrings += HexAsc(tmp2(i))
                Next
            Catch ex As Exception
                If MsgBox("PLC－PC通信ｴﾗｰ" & vbCr & "DCSを終了してよいですか？", CType(vbOKCancel + vbExclamation, MsgBoxStyle)) = vbOK Then
                    Application.Exit()
                End If
                Application.Exit()
                PlcReadStrings = ""
            End Try
        Else
            PlcReadStrings = ""
        End If
    End Function

    Public Function HexAsc(x As String) As String
        If Len(x) > 4 Then x = Strings.Right(x, 4)
        If Len(x) < 4 Then x = Strings.Right("0000", (4 - Len(x))) + x
        Dim s1 As String = Strings.Left(x, 2)
        Dim s2 As String = Strings.Right(x, 2)
        If Val("&h" + s1) >= &H20 And Val("&h" + s1) <= &H7E Then
            HexAsc = Chr(CInt("&h" + s1))
        Else
            HexAsc = "?"
        End If
        If Val("&h" + s2) >= &H20 And Val("&h" + s2) <= &H7E Then
            HexAsc += Chr(CInt("&h" + s2))
        Else
            HexAsc += "?"
        End If
    End Function

    Public Sub DGVClear(ByVal dgv As DataGridView)
        With dgv
            '列数が>0なら表示されていると判断し、一旦消去(表示速度には影響なし)
            If .Rows.Count > 0 Then
                .Columns.Clear()                          'コレクションを空にします(行・列削除)
                .DataSource = Nothing                     'DataSource に既定値を設定
                .DefaultCellStyle = Nothing               'セルスタイルを初期値に設定
                .RowHeadersDefaultCellStyle = Nothing     '行ヘッダーを初期値に設定
                .RowHeadersVisible = True                 '行ヘッダーを表示
                'フォントの高さ＋9 (フォントサイズ 9 の場合、12+9= 21 となる
                .RowTemplate.Height = 21                  'デフォルトの行の高さを設定(表示後では有効にならない)
                .ColumnHeadersDefaultCellStyle = Nothing  '列ヘッダーを初期値に設定
                .ColumnHeadersVisible = True              '列ヘッダーを表示
                .ColumnHeadersHeight = 23                 '列ヘッダーの高さを既定値に設定
                .TopLeftHeaderCell = Nothing              '左端上端のヘッダーを初期値に設定
                '奇数行に適用される既定のセルスタイルを初期値に設定   
                .AlternatingRowsDefaultCellStyle = Nothing
                'セルの境界線スタイルを初期値(一重線の境界線)に設定
                .AdvancedCellBorderStyle.All = DataGridViewAdvancedCellBorderStyle.Single
                .GridColor = SystemColors.ControlDark     'セルを区切るグリッド線の色を初期値に設定
                .Refresh()                                '再描画
            End If
        End With
    End Sub

    Private Sub dgvEq_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvEq.CellClick
        dgvEq.CurrentCell = Nothing         '選択されているセルをなくす
    End Sub

End Class

