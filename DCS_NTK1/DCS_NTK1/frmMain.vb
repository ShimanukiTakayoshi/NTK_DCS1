Public Class frmMain
    'データ保存関連
    Public SaveFolder As String = "c:\NTK"  'CSVファイル保存先メインフォルダ
    Public SaveSubFolder As String = ""     'CSVファイル保存先サブホルダ
    Public SaveFileName As String = ""      'CSVファイル名
    Public SaveFileNameQu As String = ""    'CSVファイル名
    Public Gouki As Integer = 1             '号機番号
    Public SaveTimeH As String = "00"       'データ保存ファイル切替時間(H)
    Public SaveTimeM As String = "00"       'データ保存ファイル切替時間(M)
    'PLC通信アドレス設定
    Public StartTriggerAdress As Long = 12301   'ｽﾀｰﾄﾄﾘｶﾞ
	Public EndTriggerAdress As Long = 12302     'ｴﾝﾄﾞﾄﾘｶﾞ
    Public QuTriggerAdress As Long = 12300      '品質ﾃﾞｰﾀﾄﾘｶﾞ

    Public ElementNo As String = ""     '素子品番
    Public LotNo As String = ""         'ﾒｯｷﾛｯﾄNo.
    Public OperatorNo As String = ""    '作業者
    Public StartTime As String = ""     '仕掛時間
    Public EndTime As String = ""       '完了時間
    Public ProcessTime As String = ""   '処理時間
    Public ProbeData(9) As Long         '各ﾌﾟﾛｰﾌﾞ使用回数
    Public StartTimeValue As Long = 0   '開始時間(秒数)
    Public EndTimeValue As Long = 0     '終了時間(秒数)
    Public dtNow As DateTime

    Public SayaNo As Integer = 0        'サヤ番号
    Public SayaPosi(4) As Integer       'サヤ上ﾜｰｸ位置番号
    Public IndexNo As Integer = 0       'ｲﾝﾃﾞｯｸｽ ｽﾃｰｼｮﾝ番号
    Public JudgeLead(4) As Integer      'ﾘｰﾄﾞ位置決めOK/NG
    Public JudgeDet(4) As Integer       '検知抵抗OK/NG
    Public JudgeLen1(4) As Integer      '全長抵抗1 OK/NG
    Public JudgeLen2(4) As Integer      '全長抵抗2 OK/NG
    'Public RetryLead(4) As Integer      'ﾘｰﾄﾞ位置決めﾘﾄﾗｲ有無
    Public RetryDet(4) As Integer       '検知抵抗ﾘﾄﾗｲ有無
    Public RetryLen1(4) As Integer      '全長抵抗1ﾘﾄﾗｲ有無
    Public RetryLen2(4) As Integer      '全長抵抗2ﾘﾄﾗｲ有無
    Public Zaika(4) As Integer          '在荷有無
    Public ReDet(4) As String           '検知抵抗測定値
    Public ReLen1(4) As String          '全長抵抗1測定値
    Public ReLen2(4) As String          '全長抵抗2測定値
    Public ReShift(24) As String        'ﾃﾞﾊﾞｯｸﾞﾓﾆﾀ用ｼﾌﾄﾃﾞｰﾀ
    'Public FormatData(100) As String
    Public EqChartRow As Integer = 8      '設備ﾁｬｰﾄ_ｽｸﾛｰﾙ制御用
    Public QuChartRow As Integer = 4      '品質ﾁｬｰﾄ_ｽｸﾛｰﾙ制御用
    Public QuData(4, 70) As String        '取得_品質ﾃﾞｰﾀ
    Public StackData(13, 110) As String   '設備ﾁｬｰﾄ 直近n=100個分ﾃﾞｰﾀ
    Public StackCounter As Integer = 0    '設備ﾁｬｰﾄ ｽﾀｯｸｶｳﾝﾀｰ
    Public QuStackData(500, 12) As String '品質ﾁｬｰﾄ 直近n=100個分ﾃﾞｰﾀ
    Public QuStackCounter As Integer = 0  '品質ﾁｬｰﾄ ｽﾀｯｸｶｳﾝﾀｰ

    Public EqStartedFlag As Boolean = False         '設備ﾃﾞｰﾀ集計開始ﾌﾗｸﾞ
    Public PlcReadingFlag As Boolean = False        'PLC通信中ﾌﾗｸﾞ
    Public SaveDataFirstFlag As Boolean = True      '設備ﾃﾞｰﾀ 初回ﾃﾞｰﾀ保存ﾌﾗｸﾞ
    Public SaveDataFirstFlagQu As Boolean = True    '品質ﾃﾞｰﾀ 初回ﾃﾞｰﾀ保存ﾌﾗｸﾞ

    Public DebugFlag As Boolean = False             'ﾃﾞﾊﾞｯｸﾞﾌﾗｸﾞ
    Public DebugSatrtFlag As Boolean = False        'ﾃﾞﾊﾞｯｸﾞ用設備ﾃﾞｰﾀ取得開始ﾌﾗｸﾞ
    Public DebugEndFlag As Boolean = False          'ﾃﾞﾊﾞｯｸﾞ用設備ﾃﾞｰﾀ取得完了ﾌﾗｸﾞ
    Public DebugDataFlag As Boolean = False         'ﾃﾞﾊﾞｯｸﾞ用品質ﾃﾞｰﾀ取得ﾌﾗｸﾞ
    Public SayaPosiDebugFlag As Boolean = True      'ｻﾔﾎﾟｼﾞｼｮﾝ生成ﾌﾗｸﾞ(もともとﾃﾞﾊﾞｯｸﾞ用だったが、通常使用となった。)
    Public TmpLong(20) As Long                      '汎用
    Public TmpInt(299) As Long                      '汎用

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '二重起動防止
        If PrevInstance() Then
            Application.Exit()
        End If
        '初期設定
        initialize()
    End Sub

    Private Sub initialize()
        'ﾃﾞﾊﾞｯｸﾞ用ｺﾝﾄﾛｰﾙ表示/非表示
        If DebugFlag Then
            TextBox2.Visible = True
            TextBox3.Visible = True
            TextBox4.Visible = True
            btnStart.Visible = True
            btnEnd.Visible = True
            btnData.Visible = True
            Button1.Visible = True
            Button2.Visible = True
        Else
            TextBox2.Visible = False
            TextBox3.Visible = False
            TextBox4.Visible = False
            btnStart.Visible = False
            btnEnd.Visible = False
            btnData.Visible = False
            Button1.Visible = False
            Button2.Visible = False
        End If
        '設備結果ﾃﾞｰﾀｼｰﾄ
        DGVClear(dgvEq)
        Me.Width = 1024
        Me.Height = 768
        Me.Left = 0
        Me.Top = 0
        Me.StartPosition = FormStartPosition.Manual
        dgvEq.Width = 990
        dgvEq.Height = 263 - 21 * 2
        Dim cstyle1 As New DataGridViewCellStyle
        cstyle1.Alignment = DataGridViewContentAlignment.MiddleCenter
        Dim columnHeaderStyle As DataGridViewCellStyle = dgvEq.ColumnHeadersDefaultCellStyle
        dgvEq.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        columnHeaderStyle.Font = New Font("ＭＳ ゴシック", 8)
        dgvEq.Columns.Add("0", "品番")
        dgvEq.Columns.Add("1", "Lot")
        dgvEq.Columns.Add("2", "ｵﾍﾟﾚｰﾀ")
        dgvEq.Columns.Add("3", "仕掛時間")
        dgvEq.Columns.Add("4", "完了時間")
        dgvEq.Columns.Add("5", "処理時間")
        dgvEq.Columns.Add("6", "検知上")
        dgvEq.Columns.Add("7", "検知横")
        dgvEq.Columns.Add("8", "検知下")
        dgvEq.Columns.Add("9", "全長①上")
        dgvEq.Columns.Add("10", "全長①下")
        dgvEq.Columns.Add("11", "全長②上")
        dgvEq.Columns.Add("12", "全長②横")
        dgvEq.Columns.Add("13", "全長②斜")
        For i As Integer = 0 To 4
            dgvEq.Columns(i).DefaultCellStyle = cstyle1
            dgvEq.Columns(i).Width = 59
        Next i
        For i As Integer = 5 To 13
            dgvEq.Columns(i).DefaultCellStyle = cstyle1
            dgvEq.Columns(i).Width = 63
        Next i
        dgvEq.Columns(2).Width = 65
        dgvEq.Columns(3).Width = 110
        dgvEq.Columns(4).Width = 110
        For i As Integer = 0 To 8
            dgvEq.Rows.Add("")
        Next
        dgvEq.RowHeadersVisible = False
        dgvEq.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvEq.CurrentCell = Nothing         '選択されているセルをなくす
        '品質結果ﾃﾞｰﾀｼｰﾄ
        DGVClear(dgvQu)
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        dgvQu.Width = 990 '945
        dgvQu.Height = 410 + 21 * 2
        Dim cstyle2 As New DataGridViewCellStyle
        cstyle1.Alignment = DataGridViewContentAlignment.MiddleCenter
        Dim columnHeaderStyle2 As DataGridViewCellStyle = dgvQu.ColumnHeadersDefaultCellStyle
        dgvQu.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        columnHeaderStyle2.Font = New Font("ＭＳ ゴシック", 8)
        dgvQu.Columns.Add("0", "日付")
        dgvQu.Columns.Add("1", "品番")
        dgvQu.Columns.Add("2", "ﾛｯﾄNo.")
        dgvQu.Columns.Add("3", "ﾜｰｸNo")
        dgvQu.Columns.Add("4", "位置決め")
        dgvQu.Columns.Add("5", "検知抵抗")
        dgvQu.Columns.Add("6", "結果")
        dgvQu.Columns.Add("7", "ﾘﾄﾗｲ")
        dgvQu.Columns.Add("8", "全長抵抗")
        dgvQu.Columns.Add("9", "結果")
        dgvQu.Columns.Add("10", "ﾘﾄﾗｲ")
        dgvQu.Columns.Add("11", "測定ﾎﾟｼﾞｼｮﾝ")
        dgvQu.Columns.Add("12", "ｲﾝﾃﾞｯｸｽ治具No.")
        For i As Integer = 0 To 12
            dgvQu.Columns(i).DefaultCellStyle = cstyle1
            dgvQu.Columns(i).Width = 62
        Next i
        For i As Integer = 0 To 19
            dgvQu.Rows.Add("")
        Next
        dgvQu.Columns(0).Width = 110
        dgvQu.Columns(1).Width = 85
        dgvQu.Columns(2).Width = 85
        dgvQu.Columns(3).Width = 71
        dgvQu.Columns(5).Width = 92
        dgvQu.Columns(8).Width = 92
        dgvQu.RowHeadersVisible = False
        dgvQu.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvQu.CurrentCell = Nothing         '選択されているセルをなくす
        dgvQu.FirstDisplayedScrollingRowIndex = 0
    End Sub

    Private Sub timScan_Tick(sender As Object, e As EventArgs) Handles timScan.Tick
        TextBox2.Text = Str(PlcRead(10410))
        If Not PlcReadingFlag Then
            Main()
        End If
    End Sub

	Public Sub Main()
        'スタートトリガ監視
        If PlcRead(StartTriggerAdress) <> 0 Or DebugSatrtFlag Then
            DebugSatrtFlag = False
            PlcWrite(StartTriggerAdress, 0)
            If Not EqStartedFlag Then
                EqStartedFlag = True
                StartTime = Trim(CStr(Now))
                dtNow = DateTime.Now
                StartTimeValue = TimeValue(dtNow)
                StackCounter += 1
                GetPlcData()
                'ElementNo = "Element123"
                'LotNo = "Lot123"
                EndTime = ""
                ProcessTime = ""
                For i As Integer = 0 To 7
                    ProbeData(i) = 0
                Next
                StackSet()
                DrawChartSetubi()
                MakeElementFolder()
                'MakeLotFile()
            End If
        End If
        'エンドトリガ監視
        If PlcRead(EndTriggerAdress) <> 0 Or DebugEndFlag Then
            DebugEndFlag = False
            PlcWrite(EndTriggerAdress, 0)
            If EqStartedFlag Then
                EndTime = Trim(CStr(Now))
                dtNow = DateTime.Now
                EndTimeValue = TimeValue(dtNow)
                If StartTimeValue > EndTimeValue Then
                    EndTimeValue += 86400
                End If
                Dim a0 As Long = EndTimeValue - StartTimeValue
                Dim s0 As String = Trim(Str(Int(a0 / 60)))
                Dim s1 As String = Trim(Str(a0 Mod 60))
                If Len(s1) = 1 Then s1 = "0" + s1
                ProcessTime = s0 & ":" & s1
                GetPlcData()
                StackSet()
                DrawChartSetubi()
                SaveData()
                EqStartedFlag = False
            End If
        End If
        '測定データトリガ監視
        If PlcRead(QuTriggerAdress) <> 0 Or DebugDataFlag Then
            DebugDataFlag = False
            PlcWrite(QuTriggerAdress, 0)
            QuStackCounter += 1
            GetPlcData()
            If (SayaNo >= 1 And SayaNo <= 12) And (SayaPosi(0) > 0 And SayaPosi(1) > 0 And SayaPosi(2) > 0 And SayaPosi(3) > 0) Then
                ChFormat()
                QuStackSet()
                DrawChartQu()
                SaveDataQu()
            Else
                QuStackCounter -= 1
            End If
        End If
    End Sub

    Public Sub GetPlcData()
        If Not DebugFlag Then
            PlcReadingFlag = True
            PlcReadWord(12000, 240)
            PlcReadingFlag = False
            '設備ﾃﾞｰﾀ読込
            ElementNo = HexAsc(Hex(TmpInt(0))) & HexAsc(Hex(TmpInt(1))) & HexAsc(Hex(TmpInt(2))) & HexAsc(Hex(TmpInt(3)))
            LotNo = HexAsc(Hex(TmpInt(4))) & HexAsc(Hex(TmpInt(5))) & HexAsc(Hex(TmpInt(6))) & HexAsc(Hex(TmpInt(7)))
            OperatorNo = HexAsc(Hex(TmpInt(8))) & HexAsc(Hex(TmpInt(9))) & HexAsc(Hex(TmpInt(10))) & HexAsc(Hex(TmpInt(11)))
            For i As Short = 0 To 7
                ProbeData(i) = CLng(Val(Hex(TmpInt(i * 2 + 13)) & Hex(TmpInt(i * 2 + 12))))
            Next i
            '品質ﾃﾞｰﾀ読込
            SayaNo = CInt(TmpInt(100))
            If Not SayaPosiDebugFlag Then
                For i As Short = 0 To 3
                    SayaPosi(i) = CInt(TmpInt(101 + i))
                Next
            Else
                For i As Short = 0 To 3
                    SayaPosi(i) = CInt(TmpInt(220 + i))
                Next
            End If
            IndexNo = CInt(TmpInt(105))
            For i As Short = 0 To 3
                JudgeLead(i) = CInt(TmpInt(106 + i))
            Next
            For i As Short = 0 To 3
                JudgeDet(i) = CInt(TmpInt(110 + i))
            Next
            For i As Short = 0 To 3
                JudgeLen1(i) = CInt(TmpInt(114 + i))
            Next
            For i As Short = 0 To 3
                JudgeLen2(i) = CInt(TmpInt(118 + i))
            Next
            For i As Short = 0 To 3
                RetryDet(i) = CInt(TmpInt(122 + i))
            Next
            For i As Short = 0 To 3
                RetryLen1(i) = CInt(TmpInt(126 + i))
            Next
            For i As Short = 0 To 3
                RetryLen2(i) = CInt(TmpInt(130 + i))
            Next
            For i As Short = 0 To 3
                Zaika(i) = CInt(TmpInt(134 + i))
            Next
            For i As Short = 0 To 3
                ReDet(i) = CStr(DBcdLong(TmpInt(140 + i * 2), TmpInt(141 + i * 2)))
            Next
            For i As Short = 0 To 3
                ReLen1(i) = CStr(DBcdLong(TmpInt(148 + i * 2), TmpInt(149 + i * 2)))
            Next
            For i As Short = 0 To 3
                ReLen2(i) = CStr(DBcdLong(TmpInt(156 + i * 2), TmpInt(157 + i * 2)))
            Next
            For i As Short = 0 To 23
                ReShift(i) = CStr(DBcdLong(TmpInt(170 + i * 2), TmpInt(171 + i * 2)))
            Next
        Else
            ElementNo = "Ele" & Trim(CStr(Int(Rnd(1) * 99999)))
            LotNo = "Lot" & Trim(CStr(Int(Rnd(1) * 99999)))
            OperatorNo = "Ope" & Trim(CStr(Int(Rnd(1) * 99999)))
            For i As Short = 0 To 7
                ProbeData(i) = CLng((Int(Rnd(1) * 9999999)))
            Next i
            '品質ﾃﾞｰﾀ読込
            SayaNo = CInt(Int(Rnd(1) * 6) + 1)
            For i As Short = 0 To 3
                SayaPosi(i) = CInt(Int(Rnd(1) * 200) + 1)
            Next
            IndexNo = CInt(Int(Rnd(1) * 8) + 1)
            For i As Short = 0 To 3
                JudgeLead(i) = CInt(Int(Rnd(1) * 3))
            Next
            For i As Short = 0 To 3
                JudgeDet(i) = CInt(Int(Rnd(1) * 3))
            Next
            For i As Short = 0 To 3
                JudgeLen1(i) = CInt(Int(Rnd(1) * 3))
            Next
            For i As Short = 0 To 3
                JudgeLen2(i) = CInt(Int(Rnd(1) * 3))
            Next
            For i As Short = 0 To 3
                RetryDet(i) = CInt(Int(Rnd(1) * 2))
            Next
            For i As Short = 0 To 3
                RetryLen1(i) = CInt(Int(Rnd(1) * 2))
            Next
            For i As Short = 0 To 3
                RetryLen2(i) = CInt(Int(Rnd(1) * 2))
            Next
            For i As Short = 0 To 3
                Zaika(i) = CInt(Int(Rnd(1) * 2))
            Next
            For i As Short = 0 To 3
                ReDet(i) = CType(CLng((Int(Rnd(1) * 100000))), String)
            Next
            For i As Short = 0 To 3
                ReLen1(i) = CType(CLng((Int(Rnd(1) * 100000))), String)
            Next
            For i As Short = 0 To 3
                ReLen2(i) = CType(CLng((Int(Rnd(1) * 100000))), String)
            Next
        End If
        If DebugFlag Then
            ElementNo = "Ele123y8"
            LotNo = "Lot12399"
        End If
    End Sub

	Public Sub ChFormat()
		'さやNoアルファベット変換
		Dim s0 As String = ""
		Dim s1 As String = ""
		Select Case SayaNo
			Case 1
				s0 = "A"
			Case 2
				s0 = "B"
			Case 3
				s0 = "C"
			Case 4
				s0 = "D"
			Case 5
				s0 = "E"
			Case 6
				s0 = "F"
			Case 7
				s0 = "R1"
			Case 8
				s0 = "R2"
			Case 9
				s0 = "R3"
			Case 10
				s0 = "R4"
			Case 11
				s0 = "R5"
			Case 12
				s0 = "R6"
			Case Else
				s0 = "X"
		End Select
		'全長抵抗１or２選択
		Dim ReLen(4) As String
		Dim JudgeLen(4) As Integer
		Dim RetryLen(4) As Integer
		Dim a0 As Long = CLng(Val(ReLen1(0)) + Val(ReLen1(1)) + Val(ReLen1(2)) + Val(ReLen1(3)))
		If a0 <> 0 Then
			ReLen(0) = ReLen1(0) : ReLen(1) = ReLen1(1) : ReLen(2) = ReLen1(2) : ReLen(3) = ReLen1(3)
			JudgeLen(0) = JudgeLen1(0) : JudgeLen(1) = JudgeLen1(1) : JudgeLen(2) = JudgeLen1(2) : JudgeLen(3) = JudgeLen1(3)
			RetryLen(0) = RetryLen1(0) : RetryLen(1) = RetryLen1(1) : RetryLen(2) = RetryLen1(2) : RetryLen(3) = RetryLen1(3)
		Else
			ReLen(0) = ReLen2(0) : ReLen(1) = ReLen2(1) : ReLen(2) = ReLen2(2) : ReLen(3) = ReLen2(3)
			JudgeLen(0) = JudgeLen2(0) : JudgeLen(1) = JudgeLen2(1) : JudgeLen(2) = JudgeLen2(2) : JudgeLen(3) = JudgeLen2(3)
			RetryLen(0) = RetryLen2(0) : RetryLen(1) = RetryLen2(1) : RetryLen(2) = RetryLen2(2) : RetryLen(3) = RetryLen2(3)
		End If
		For i As Integer = 0 To 3
			QuData(i, 0) = CType(Now, String)
			QuData(i, 1) = ElementNo
			QuData(i, 2) = LotNo
			QuData(i, 3) = s0 & "-" & CType(SayaPosi(i), String)
			Select Case JudgeLead(i)
				Case 0
					QuData(i, 4) = "NG"
				Case 1
					QuData(i, 4) = "OK"
				Case Else
					QuData(i, 4) = "--"
			End Select
            QuData(i, 5) = ChangeData(ReDet(i))
			Select Case JudgeDet(i)
				Case 0
					QuData(i, 6) = "NG"
				Case 1
					QuData(i, 6) = "OK"
				Case Else
					QuData(i, 6) = "--"
			End Select
			Select Case RetryDet(i)
				Case 1
					QuData(i, 7) = "R"
				Case Else
					QuData(i, 7) = " "
			End Select
            QuData(i, 8) = ChangeData(ReLen(i))
			Select Case JudgeLen(i)
				Case 0
					QuData(i, 9) = "NG"
				Case 1
					QuData(i, 9) = "OK"
				Case Else
					QuData(i, 9) = "--"
			End Select
            Select Case RetryLen(i)
                Case 1
                    QuData(i, 10) = "R"
                Case Else
                    QuData(i, 10) = " "
            End Select
			QuData(i, 11) = CType(i + 1, String)
			QuData(i, 12) = CType(IndexNo, String)
		Next i
	End Sub

	Private Sub StackSet()
		StackData(0, StackCounter) = ElementNo
		StackData(1, StackCounter) = LotNo
		StackData(2, StackCounter) = OperatorNo
		StackData(3, StackCounter) = StartTime
		StackData(4, StackCounter) = EndTime
        StackData(5, StackCounter) = ProcessTime
        For i As Integer = 6 To 13
            StackData(i, StackCounter) = CStr(ProbeData(i - 6))
        Next
		If StackCounter > 100 Then
			For i As Integer = 1 To 100
                For j As Integer = 0 To 13
                    StackData(j, i) = StackData(j, i + 1)
                Next
			Next
			StackCounter = 100
		End If
	End Sub

	Private Sub QuStackSet()
		For i As Integer = 0 To 3
			For j As Integer = 0 To 12
				QuStackData(QuStackCounter * 4 + i, j) = QuData(i, j)
			Next j
		Next i
		If QuStackCounter > 100 Then
			For i As Integer = 1 To 100
				For j As Integer = 0 To 12
                    QuStackData(i * 4 + 0, j) = QuStackData((i + 1) * 4 + 0, j)
                    QuStackData(i * 4 + 1, j) = QuStackData((i + 1) * 4 + 1, j)
                    QuStackData(i * 4 + 2, j) = QuStackData((i + 1) * 4 + 2, j)
                    QuStackData(i * 4 + 3, j) = QuStackData((i + 1) * 4 + 3, j)
                Next
			Next
			QuStackCounter = 100
		End If
	End Sub

    Public Sub DrawChartSetubi()
        If StackCounter >= EqChartRow + 1 Then
            dgvEq.Rows.Add("")
            EqChartRow = StackCounter
        End If
        For i As Integer = 0 To StackCounter
            For j As Integer = 0 To 13
                dgvEq.Item(j, i).Value = StackData(j, i + 1)
            Next
        Next
        If dgvEq.RowCount > 11 Then
            dgvEq.FirstDisplayedScrollingRowIndex = (StackCounter - 9) * 1 + 0
        End If
    End Sub

	Public Sub DrawChartQu()
        If QuStackCounter >= QuChartRow + 1 Then
            dgvQu.Rows.Add("")
            dgvQu.Rows.Add("")
            dgvQu.Rows.Add("")
            dgvQu.Rows.Add("")
            QuChartRow = QuStackCounter
        End If
        For i As Integer = 0 To QuStackCounter
            For j As Integer = 0 To 12
                dgvQu.Item(j, i * 4 + 0).Value = QuStackData((i + 1) * 4 + 0, j)
                dgvQu.Item(j, i * 4 + 1).Value = QuStackData((i + 1) * 4 + 1, j)
                dgvQu.Item(j, i * 4 + 2).Value = QuStackData((i + 1) * 4 + 2, j)
                dgvQu.Item(j, i * 4 + 3).Value = QuStackData((i + 1) * 4 + 3, j)
            Next
        Next
        If dgvQu.RowCount > 28 Then
            dgvQu.FirstDisplayedScrollingRowIndex = (QuStackCounter - 5) * 4 + 0
        End If
    End Sub

    Public Function TimeValue(x As DateTime) As Long
        TimeValue = x.Second
        TimeValue += (x.Minute) * 60
        TimeValue += (x.Hour) * 60 * 60
    End Function

    'PLC通信

    Public Function PlcRead(address As Long) As Long
        '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
        'シーケンサーのDMメモリーより「address」にて指定したアドレスの内容を読み込む
        '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
        If Not DebugFlag Then
            Try
                PlcRead = SysmacCJ.DM(address)
            Catch ex As Exception
                timScan.Enabled = False
                If MsgBox("PLC－PC通信ｴﾗｰ" & vbCr & "DCSを終了してよいですか？", CType(vbOKCancel + vbExclamation, MsgBoxStyle)) = vbOK Then
                    Application.Exit()
                End If
                timScan.Enabled = True
                PlcRead = 0
            End Try
        Else
            PlcRead = 0
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
                timScan.Enabled = False
                If MsgBox("PLC－PC通信ｴﾗｰ" & vbCr & "DCSを終了してよいですか？", CType(vbOKCancel + vbExclamation, MsgBoxStyle)) = vbOK Then
                    Application.Exit()
                End If
                timScan.Enabled = True
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
                timScan.Enabled = False
                If MsgBox("PLC－PC通信ｴﾗｰ" & vbCr & "DCSを終了してよいですか？", CType(vbOKCancel + vbExclamation, MsgBoxStyle)) = vbOK Then
                    Application.Exit()
                End If
                timScan.Enabled = True
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
                timScan.Enabled = False
                If MsgBox("PLC－PC通信ｴﾗｰ" & vbCr & "DCSを終了してよいですか？", CType(vbOKCancel + vbExclamation, MsgBoxStyle)) = vbOK Then
                    Application.Exit()
                End If
                timScan.Enabled = True
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
        Dim tmp1(299) As Integer
        If Not DebugFlag Then
            Try
                tmp1 = SysmacCJ.ReadMemoryWordInteger(OMRON.Compolet.SYSMAC.SysmacCJ.MemoryTypes.DM, address, length + 1, OMRON.Compolet.SYSMAC.SysmacCJ.DataTypes.BIN)
                For i As Integer = 0 To length
                    TmpInt(i) = tmp1(i)
                Next
            Catch ex As Exception
                timScan.Enabled = False
                If MsgBox("PLC－PC通信ｴﾗｰ" & vbCr & "DCSを終了してよいですか？", CType(vbOKCancel + vbExclamation, MsgBoxStyle)) = vbOK Then
                    Application.Exit()
                End If
                timScan.Enabled = True
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

    Public Function Hex4(x As String) As String
        Dim a0 As Integer = Len(Trim(x))
        If a0 < 4 Then
            Hex4 = Strings.Left("0000", 4 - a0) + Trim(x)
        Else
            Hex4 = Strings.Left(Trim(x), 4)
        End If
    End Function

    Public Function DBcdLong(bl As Long, bh As Long) As Long
        Dim s1 As String = Hex4(Trim(Hex(bh)))
        Dim s2 As String = Hex4(Trim(Hex(bl)))
        Dim s3 As String = s1 + s2
        DBcdLong = CLng(Val(s3))
    End Function

    Public Function ChangeData(s0 As String) As String
        Dim a0 As Double = 0
        Dim s1 As String = ""
        Dim a1 As Integer = 0
        Dim a2 As Integer = 0
        If IsNumeric(s0) Then
            a0 = Val(s0) / 1000
            s1 = Trim(Str(a0))
            a1 = Len(s1)
            If Strings.Left(s1, 1) = "." Then
                If a1 > 4 Then s1 = Strings.Left(s1, 4)
                If a1 < 4 Then s1 = s1 & (Strings.Left("0000", 4 - a1))
                s1 = "0" + s1
            Else
                a2 = InStr(s1, ".")
                If a2 = 0 Then
                    s1 = s1 + ".000"
                Else
                    Select Case a1 - a2
                        Case 1
                            s1 = s1 + "00"
                        Case 2
                            s1 = s1 + "0"
                        Case Else
                            s1 = s1
                    End Select
                End If
            End If
        End If
        ChangeData = s1
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

    Public Shared Function PrevInstance() As Boolean
        ' ---------------------------------------------------------------------------
        '    同名のプロセスが起動しているかどうかを示す値を返します。
        '    同名のプロセスが起動中の場合は True。それ以外は False。
        ' ---------------------------------------------------------------------------
        ' このアプリケーションのプロセス名を取得
        Dim stThisProcess As String = System.Diagnostics.Process.GetCurrentProcess().ProcessName
        ' 同名のプロセスが他に存在する場合は、既に起動していると判断する
        If System.Diagnostics.Process.GetProcessesByName(stThisProcess).Length > 1 Then
            Return True
        End If
        ' 存在しない場合は False を返す
        Return False
    End Function

    '各ファイル保存

    Public Sub MakeElementFolder()
        If ElementNo <> "" And InStr(ElementNo, "?") = 0 Then
            Dim di As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(SaveFolder + "\" + ElementNo)
            di.Create()
        Else
            '
        End If
    End Sub

    Public Sub MakeLotFile()
        '保存ファイル生成
        SaveFileNameQu = LotNo
        Dim Title As String = ""
        Title = "日付,素子品番,ﾛｯﾄNo,ﾜｰｸNo,位置決め,検知抵抗,結果,ﾘﾄﾗｲ,全長抵抗,結果,ﾘﾄﾗｲ,測定ﾎﾟｼﾞｼｮﾝ,ｲﾝﾃﾞｯｸｽ治具No" + vbCrLf
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + ElementNo + "\" + SaveFileNameQu + ".CSV", Title, True)
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + ElementNo + "\" + SaveFileNameQu + ".BKF", Title, True)
        SaveDataFirstFlagQu = False
    End Sub

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
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "setubi_" + SaveFileName + ".CSV", Title, True)
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "setubi_" + SaveFileName + ".BKF", Title, True)
    End Sub

    Public Sub CreateSaveFileNameQu()
        '保存ファイル生成
        Dim dt As DateTime = DateTime.Now
        Dim b As String = dt.ToString
        SaveFileNameQu = Strings.Left(Trim(b), 4) + Strings.Mid(Trim(b), 6, 2) + Strings.Mid(Trim(b), 9, 2)
        Dim Title As String = ""
        Title = "日付,素子品番,ﾛｯﾄNo,ﾜｰｸNo,位置決め,検知抵抗,結果,ﾘﾄﾗｲ,全長抵抗,結果,ﾘﾄﾗｲ,測定ﾎﾟｼﾞｼｮﾝ,ｲﾝﾃﾞｯｸｽ治具No" + vbCrLf
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "hinshitu" + Trim(Str(Gouki)) + "_" + SaveFileNameQu + ".CSV", Title, True)
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "Hinshitu" + Trim(Str(Gouki)) + "_" + SaveFileNameQu + ".BKF", Title, True)
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
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "setubi_" + SaveFileName + ".CSV", InputString, True)
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "setubi_" + SaveFileName + ".BKF", InputString, True)
    End Sub

    Public Sub SaveDataQu()
        '起動初回確認
        If SaveDataFirstFlagQu Then
            MakeElementFolder()
            MakeLotFile()
            SaveDataFirstFlagQu = False
        End If
        '保存先フォルダ確認＆生成
        Dim x1 As String = SaveFolder + "\" + ElementNo + "\"
        If Not System.IO.Directory.Exists(x1) Then
            MakeElementFolder()
        End If
        '保存ファイルの確認
        Dim x2 As String = SaveFolder + "\" + ElementNo + "\" + SaveFileNameQu + ".CSV"
        If Not System.IO.File.Exists(x2) Then
            MakeLotFile()
        End If
        'データ保存
        Dim InputString As String = ""
        For i As Integer = 0 To 3
            InputString = ""
            For j As Integer = 0 To 12
                InputString = InputString + QuStackData(QuStackCounter * 4 + i, j) + ","
            Next
            InputString = InputString & vbCrLf
            My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + ElementNo + "\" + SaveFileNameQu + ".CSV", InputString, True)
            My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + ElementNo + "\" + SaveFileNameQu + ".BKF", InputString, True)
        Next
    End Sub

    'ﾃﾞﾊﾞｯｸﾞ用

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmDebug.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox3.Text = ChangeData(TextBox4.Text)
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        DebugSatrtFlag = True
    End Sub

    Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click
        DebugEndFlag = True
    End Sub

    Private Sub btnData_Click(sender As Object, e As EventArgs) Handles btnData.Click
        DebugDataFlag = True
    End Sub
End Class

