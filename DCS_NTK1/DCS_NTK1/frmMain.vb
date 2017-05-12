Public Class frmMain
    'データ保存関連
    Public SaveFolder As String = "c:\NTK" 'CSVファイル保存先メインフォルダ
    Public SaveSubFolder As String = ""    'CSVファイル保存先サブホルダ
    Public SaveFileName As String = ""     'CSVファイル名
    Public Gouki As Integer = 1            '号機番号
    Public SaveTimeH As String = "7"       'データ保存ファイル切替時間(H)
    Public SaveTimeM As String = "26"      'データ保存ファイル切替時間(M)
    'PLC通信アドレス設定
    Public AckAddress As Long = 0           'PLCへ受信OK返答
    Public StartTriggerAdress As Long = 0   'ｽﾀｰﾄﾄﾘｶﾞ
    Public EndTriggerAdress As Long = 0     'ｴﾝﾄﾞﾄﾘｶﾞ
    Public QuTriggerAdress As Long = 0      '品質ﾃﾞｰﾀﾄﾘｶﾞ
    Public ElementNoAddress As Long = 0     '素子品番
    Public LotNoAddress As Long = 0         'ﾒｯｷﾛｯﾄ
    Public OperatorAddress As Long = 0      '作業者
    Public StartTimeAddress As Long = 0     '仕掛時間
    Public EndTimeAddress As Long = 0       '完了時間
    Public ProbeAddress As Long = 0         'ﾌﾟﾛｰﾌﾞ使用回数先頭ｱﾄﾞﾚｽ
    Public QuAddress As Long = 0            '品質ﾃﾞｰﾀ先頭ｱﾄﾞﾚｽ

    Public ElementNo As String = ""         '素子品番
    Public LotNo As String = ""             'ﾒｯｷﾛｯﾄNo.
    Public OperatorNo As String = ""        '作業者
    Public StartTime As String = ""         '仕掛時間
    Public EndTime As String = ""           '完了時間
    Public ProbeData(9） As Long            '各ﾌﾟﾛｰﾌﾞ使用回数

    Public QuHizuke(4) As String
    Public QuType(4) As String
    Public QuLot(4) As String
    Public QuWorkNo(4) As String
    Public QuIcnikime(4) As String
    Public QuKenchiResister(4) As String
    Public QuKenchiKekka(4) As String
    Public QuKenchiRetry(4) As String
    Public QuZenchoResister(4) As String
    Public QuZenchoKekka(4) As String
    Public QuZenchoRetry(4) As String
    Public QuPoshiton(4) As String
    Public QuIndexNo(4) As String
    Public QuZaika(4) As String

    Public QuData(70) As String             '品質ﾃﾞｰﾀ

    Public StackData(13, 110) As String     '装置 直近n=100個分ﾃﾞｰﾀ
    Public StackCounter As Integer = 0      '装置 ｽﾀｯｸｶｳﾝﾀｰ
    Public QuStackData(13, 110) As String '測定 直近n=100個分ﾃﾞｰﾀ
    Public QuStackCounter As Integer = 0  '測定 ｽﾀｯｸｶｳﾝﾀｰ

    Public PlcReadingFlag As Boolean = False    'PLC通信中ﾌﾗｸﾞ
    Public SaveDataFirstFlag As Boolean = True  '初回ﾃﾞｰﾀ保存ﾌﾗｸﾞ

    Public DebugFlag As Boolean = True
    Public tmp0 As Long = 0
    Public TmpLong(20) As Long
    Public TmpInt(20) As Long

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        initialize()
    End Sub

    Private Sub initialize()
        '設備結果ﾃﾞｰﾀｼｰﾄ
        DGVClear(dgvEq)
        Me.Width = 1024
        Me.Height = 768
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        dgvEq.Width = 900
        dgvEq.Height = 260
        Dim cstyle1 As New DataGridViewCellStyle
        cstyle1.Alignment = DataGridViewContentAlignment.MiddleRight
        Dim columnHeaderStyle As DataGridViewCellStyle = dgvEq.ColumnHeadersDefaultCellStyle
        dgvEq.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        columnHeaderStyle.Font = New Font("ＭＳ ゴシック", 6)
        dgvEq.Columns.Add("0", "素子" & vbCrLf & "品番")
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
            dgvEq.Columns(i).Width = 60
        Next i
        dgvEq.Columns(3).Width = 110
        dgvEq.Columns(4).Width = 110
        For i As Integer = 0 To 99
            dgvEq.Rows.Add("")
        Next
        dgvEq.RowHeadersVisible = False
        dgvEq.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvEq.CurrentCell = Nothing         '選択されているセルをなくす
        '品質結果ﾃﾞｰﾀｼｰﾄ
        DGVClear(dgvQu)
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        dgvQu.Width = 1000
        dgvQu.Height = 400
        Dim cstyle2 As New DataGridViewCellStyle
        cstyle1.Alignment = DataGridViewContentAlignment.MiddleRight
        Dim columnHeaderStyle2 As DataGridViewCellStyle = dgvQu.ColumnHeadersDefaultCellStyle
        dgvQu.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        columnHeaderStyle2.Font = New Font("ＭＳ ゴシック", 6)
        dgvQu.Columns.Add("0", "日付")
        dgvQu.Columns.Add("1", "品番")
        dgvQu.Columns.Add("2", "ﾛｯﾄNo.")
        dgvQu.Columns.Add("3", "ﾜｰｸNo")
        dgvQu.Columns.Add("4", "位置決め")
        dgvQu.Columns.Add("5", "検知部抵抗")
        dgvQu.Columns.Add("6", "結果")
        dgvQu.Columns.Add("7", "ﾘﾄﾗｲ")
        dgvQu.Columns.Add("8", "全長抵抗")
        dgvQu.Columns.Add("9", "結果")
        dgvQu.Columns.Add("10", "ﾘﾄﾗｲ")
        dgvQu.Columns.Add("11", "測定ﾎﾟｼﾞｼｮﾝ")
        dgvQu.Columns.Add("12", "ｲﾝﾃﾞｯｸｽ治具No.")
        For i As Integer = 0 To 12
            dgvQu.Columns(i).DefaultCellStyle = cstyle1
            dgvQu.Columns(i).Width = 60
        Next i
        For i As Integer = 0 To 99
            dgvQu.Rows.Add("")
        Next
        dgvQu.RowHeadersVisible = False
        dgvQu.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvQu.CurrentCell = Nothing         '選択されているセルをなくす

    End Sub

    Private Sub timScan_Tick(sender As Object, e As EventArgs) Handles timScan.Tick
        If Not PlcReadingFlag Then
            Main()
        End If
    End Sub

    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles timScan.Tick
    '    tmp0 = PlcRead(10301)
    '    TextBox1.Text = Str(tmp0)
    'End Sub

    Public Sub Main()
        'スタートトリガ監視
        If PlcRead(StartTriggerAdress) <> 0 Then
            PlcWrite(StartTriggerAdress, 0)
            StackCounter += 1
            If Not DebugFlag Then
                StartProcess()
                GetProbeData()
                StackSet()
            Else
                StartProcessDebug()
                GetProbeDataDebug()
                StackSet()
            End If
            DrawChartSetubi()
        End If
        'エンドトリガ監視
        If PlcRead(EndTriggerAdress) <> 0 Then
            PlcWrite(EndTriggerAdress, 0)
            EndProcess()
            If Not DebugFlag Then
                GetProbeData()
                StackSet()
            Else
                GetProbeDataDebug()
                StackSet()
            End If
            DrawChartSetubi()
            SaveData()
        End If
        '測定データトリガ監視
        If PlcRead(QuTriggerAdress) <> 0 Then
            PlcWrite(QuTriggerAdress, 0)
            QuStackCounter += 1
            If Not DebugFlag Then
                GetQuData()
                QuStackSet()
            Else
                StartProcessDebug()
                GetProbeDataDebug()
                StackSet()
            End If
            DrawChartSetubi()
        End If
    End Sub

    Public Sub StartProcess()
        PlcReadingFlag = True
        ElementNo = PlcReadStrings(ElementNoAddress, 8)
        LotNo = PlcReadStrings(LotNoAddress, 8)
        OperatorNo = PlcReadStrings(LotNoAddress, 8)
        StartTime = Trim(CStr(Now))
        PlcReadingFlag = False
    End Sub

    Public Sub EndProcess()
        EndTime = Trim(CStr(Now))
    End Sub

    Public Sub StartProcessDebug()
        PlcReadingFlag = True
        ElementNo = "Element0"
        LotNo = "Lot00000"
        OperatorNo = "Operator"
        StartTime = Trim(CStr(Now))
        PlcReadingFlag = False
    End Sub

    Private Sub GetProbeData()
        PlcReadingFlag = True
        PlcReadDWord(ProbeAddress, 8)
        PlcReadingFlag = False
        For i As Integer = 0 To 7
            ProbeData(i) = TmpLong(i)
        Next
    End Sub

    Private Sub GetProbeDataDebug()
        PlcReadingFlag = True
        For i As Integer = 0 To 7
            ProbeData(i) = CLng(Int(Rnd(1) * 99999999))
        Next
        PlcReadingFlag = False
    End Sub

    Private Sub GetQuData()
        PlcReadingFlag = True
        PlcReadDWord(QuAddress, 12)
        PlcReadWord(QuAddress + 13, 31)
        PlcReadingFlag = False
        For i As Integer = 0 To 11
            QuData(i) = Trim(CStr(TmpLong(i)))
        Next
        For i As Integer = 0 To 30
            QuData(12 + i) = Trim(CStr(TmpInt(i)))
        Next
    End Sub

    Private Sub GetQuDataDebug()
        PlcReadingFlag = True
        For i As Integer = 0 To 11
            QuData(i) = Trim(CStr(Int(Rnd(1) * 99999999)))
        Next
        For i As Integer = 0 To 30
            QuData(12 + i) = Trim(CStr(Int(Rnd(1) * 99999999)))
        Next
        PlcReadingFlag = False
        QuDataDecode()
    End Sub

    Private Sub QuDataDecode()
        'Public QuHizuke(4) As String
        'Public QuType(4) As String
        'Public QuLot(4) As String
        'Public QuWorkNo(4) As String
        'Public QuIcnikime(4) As String
        'Public QuKenchiResister(4) As String
        'Public QuKenchiKekka(4) As String
        'Public QuKenchiRetry(4) As String
        'Public QuZenchoResister(4) As String
        'Public QuZenchoKekka(4) As String
        'Public QuZenchoRetry(4) As String
        'Public QuPoshiton(4) As String
        'Public QuIndexNo As String
        For i As Integer = 0 To 3
            '日付
            QuHizuke(i) = Trim(Str(Now))
            '品番
            QuType(i) = ElementNo
            'ロット
            QuLot(i) = LotNo
            'ワークNo.
            Dim tmp1 As String = "X"
            If QuData(0) = "1" Then tmp1 = "A"
            If QuData(0) = "2" Then tmp1 = "B"
            If QuData(0) = "3" Then tmp1 = "C"
            If QuData(0) = "4" Then tmp1 = "D"
            If QuData(0) = "5" Then tmp1 = "E"
            If QuData(0) = "6" Then tmp1 = "F"
            QuWorkNo(i) = tmp1 & "-" & Trim(Str(Val(QuData(1)) + i))
            '位置決め判定
            If QuData(3 + i) = "0" Then
                QuIcnikime(i) = "NG"
            ElseIf QuData(3 + i) = "1" Then
                QuIcnikime(i) = "OK"
            Else
                QuIcnikime(i) = "-"
            End If
            '検知抵抗 測定値
            QuKenchiResister(i) = QuData(40 + i)
            '検知抵抗 結果
            If QuData(7 + i) = "0" Then
                QuKenchiKekka(i) = "NG"
            ElseIf QuData(7 + i) = "1" Then
                QuKenchiKekka(i) = "OK"
            Else
                QuKenchiKekka(i) = "-"
            End If
            '検知抵抗 ﾘﾄﾗｲ
            If QuData(19 + i) = "1" Then
                QuKenchiRetry(i) = "R"
            Else
                QuKenchiRetry(i) = "-"
            End If
            '全長抵抗_測定値
            Dim ZenchouMeas1 As String = QuData(44 + i)
            Dim ZenchouMeas2 As String = QuData(48 + i)
            Dim ZenchouJudge1 As String = QuData(11 + i)
            Dim ZenchouJudge2 As String = QuData(15 + i)
            Dim ZenchouRetry1 As String = QuData(23 + i)
            Dim ZenchouRetry2 As String = QuData(27 + i)
            If ZenchouJudge1 <> "0" And ZenchouJudge1 <> "1" And ZenchouJudge2 <> "0" And ZenchouJudge2 <> "1" Then
                QuZenchoResister(i) = "-"
            ElseIf (ZenchouJudge1 = "0" Or ZenchouJudge1 = "1") And ZenchouJudge2 <> "0" And ZenchouJudge2 <> "1" Then
                QuZenchoResister(i) = ZenchouMeas1
            Else
                QuZenchoResister(i) = ZenchouMeas2
            End If


        Next i
    End Sub


    Private Sub StackSet()
        StackData(0, StackCounter) = ElementNo
        StackData(1, StackCounter) = LotNo
        StackData(2, StackCounter) = OperatorNo
        StackData(3, StackCounter) = StartTime
        StackData(4, StackCounter) = EndTime
        For i As Integer = 5 To 12
            StackData(i, StackCounter) = CStr(ProbeData(i - 5))
        Next
        If StackCounter > 100 Then
            For i As Integer = 1 To 100
                For j As Integer = 0 To 12
                    StackData(j, i) = StackData(j, i + 1)
                Next
            Next
            StackCounter = 100
        End If
    End Sub


    Private Sub QuStackSet()
        For i As Integer = 0 To 3
            StackData(0, QuStackCounter + i) = Trim(Str(Now))
            StackData(1, QuStackCounter + i) = ElementNo
            StackData(2, QuStackCounter + i) = LotNo
            StackData(3, QuStackCounter + i) = WorkNo(i)
            StackData(4, QuStackCounter + i) = Icnikime(i)
            StackData(5, QuStackCounter + i) = KenchiResister(i)
            StackData(6, QuStackCounter + i) = KenchiKekka(i)
            StackData(7, QuStackCounter + i) = KenchiRetry(i)
            StackData(8, QuStackCounter + i) = ZenchoResister(i)
            StackData(9, QuStackCounter + i) = ZenchoKekka(i)
            StackData(10, QuStackCounter + i) = ZenchoRetry(i)
            StackData(11, QuStackCounter + i) = Trim(Str(i + 1))
            StackData(12, QuStackCounter + i) = IndexNo
        Next
        For i As Integer = 5 To 12
            StackData(i, StackCounter) = CStr(ProbeData(i - 5))
        Next
        If StackCounter > 100 Then
            For i As Integer = 1 To 100
                For j As Integer = 0 To 12
                    StackData(j, i) = StackData(j, i + 1)
                Next
            Next
            StackCounter = 100
        End If
    End Sub


    Public Sub DrawChartSetubi()
        For i As Integer = 0 To StackCounter
            For j As Integer = 0 To 12
                dgvEq.Item(j, i).Value = StackData(j, i + 1)
            Next
        Next
    End Sub

    'PLC通信

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

    Public Sub PlcReadDWord(address As Long, length As Integer)
        '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
        'シーケンサーのDMメモリーより「address」にて指定したアドレスの内容を「length」ダブルワード分の整数を読み込む
        '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
        Dim tmp1(99) As Long
        If Not DebugFlag Then
            Try
                tmp1 = SysmacCJ.ReadMemoryDwordLong(OMRON.Compolet.SYSMAC.SysmacCJ.MemoryTypes.DM, address, length, OMRON.Compolet.SYSMAC.SysmacCJ.DataTypes.BIN)
                For i As Integer = 0 To length
                    TmpLong(i) = tmp1(i)
                Next
            Catch ex As Exception
                If MsgBox("PLC－PC通信ｴﾗｰ" & vbCr & "DCSを終了してよいですか？", CType(vbOKCancel + vbExclamation, MsgBoxStyle)) = vbOK Then
                    Application.Exit()
                End If
                Application.Exit()
            End Try
        Else
            For i As Integer = 0 To length
                TmpLong(i) = 99999999
            Next i
        End If
    End Sub

    Public Sub PlcReadWord(address As Long, length As Integer)
        '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
        'シーケンサーのDMメモリーより「address」にて指定したアドレスの内容を「length」ワード分の整数を読み込む
        '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
        Dim tmp1(99) As Integer
        If Not DebugFlag Then
            Try
                tmp1 = SysmacCJ.ReadMemoryWordInteger(OMRON.Compolet.SYSMAC.SysmacCJ.MemoryTypes.DM, address, length, OMRON.Compolet.SYSMAC.SysmacCJ.DataTypes.BIN)
                For i As Integer = 0 To length
                    TmpInt(i) = tmp1(i)
                Next
            Catch ex As Exception
                If MsgBox("PLC－PC通信ｴﾗｰ" & vbCr & "DCSを終了してよいですか？", CType(vbOKCancel + vbExclamation, MsgBoxStyle)) = vbOK Then
                    Application.Exit()
                End If
                Application.Exit()
            End Try
        Else
            For i As Integer = 0 To length
                TmpInt(i) = 999999
            Next i
        End If
    End Sub

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

    '各ファイル保存

    Public Sub CreateSaveFolder()
        '保存先フォルダ生成
        Dim dt As DateTime = DateTime.Now
        Dim b As String = dt.ToString
        SaveSubFolder = Strings.Left(Trim(b), 4) + Strings.Mid(Trim(b), 6, 2)
        Dim di As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(SaveFolder + "\" + SaveSubFolder)
        di.Create()
    End Sub

    Public Sub CreateSaveFileName()
        '保存ファイル生成
        Dim dt As DateTime = DateTime.Now
        Dim b As String = dt.ToString
        SaveFileName = Strings.Left(Trim(b), 4) + Strings.Mid(Trim(b), 6, 2) + Strings.Mid(Trim(b), 9, 2)
        Dim Title As String = ""
        Title = "素子品番,ﾒｯｷﾛｯﾄ,作業者,仕掛時間,完了時間,検知上,検知横,検知下,全1上,全1下,全2上,全2横,全2斜" + vbCrLf
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "setubi" + Trim(Str(Gouki)) + "_" + SaveFileName + ".CSV", Title, True)
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "setubi" + Trim(Str(Gouki)) + "_" + SaveFileName + ".BKF", Title, True)
    End Sub

    Public Sub SaveData()
        '起動初回確認
        If SaveDataFirstFlag Then
            CreateSaveFolder()
            CreateSaveFileName()
            SaveDataFirstFlag = False
        End If
        '現在時刻確認
        Dim NowYearMonth As String = Replace(Strings.Left(CStr(Now), 7), "/", "")
        Dim NowDate As String = Replace(Strings.Left(CStr(Now), 10), "/", "")
        Dim NowTime As String = Replace(Strings.Mid(CStr(Now), 12, 5), ":", "")
        If NowDate <> SaveFileName And Val(NowTime) >= Val(SaveTimeH + SaveTimeM) Then
            If NowYearMonth <> SaveSubFolder Then CreateSaveFolder()
            CreateSaveFileName()
        End If
        'データ保存
        Dim InputString As String = ""
        For i As Integer = 0 To 12
            InputString = InputString + StackData(i, StackCounter) + ","
        Next
        InputString = InputString & vbCrLf
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "setubi" + Trim(Str(Gouki)) + "_" + SaveFileName + ".CSV", InputString, True)
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "setubi" + Trim(Str(Gouki)) + "_" + SaveFileName + ".BKF", InputString, True)
    End Sub

End Class

