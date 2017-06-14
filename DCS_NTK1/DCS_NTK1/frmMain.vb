Public Class frmMain
    'データ保存関連
    Public SaveFolder As String = "c:\NTK" 'CSVファイル保存先メインフォルダ
    Public SaveSubFolder As String = ""    'CSVファイル保存先サブホルダ
    Public SaveFileName As String = ""     'CSVファイル名
    Public SaveFileNameQu As String = ""     'CSVファイル名
    Public Gouki As Integer = 1            '号機番号
    Public SaveTimeH As String = "7"       'データ保存ファイル切替時間(H)
    Public SaveTimeM As String = "26"      'データ保存ファイル切替時間(M)
    'PLC通信アドレス設定
    Public AckAddress As Long = 0           'PLCへ受信OK返答
    Public StartTriggerAdress As Long = 12301   'ｽﾀｰﾄﾄﾘｶﾞ
    Public EndTriggerAdress As Long = 12302     'ｴﾝﾄﾞﾄﾘｶﾞ
    Public QuTriggerAdress As Long = 12300      '品質ﾃﾞｰﾀﾄﾘｶﾞ
	Public StartTriggerOkAdress As Long = 12210   'ｽﾀｰﾄﾄﾘｶﾞ受信OK
	Public EndTriggerOkAdress As Long = 12211     'ｴﾝﾄﾞﾄﾘｶﾞ受信OK
	Public QuTriggerOkAdress As Long = 12212      '品質ﾃﾞｰﾀﾄﾘｶﾞ受信OK
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

	Public SayaNo As Integer = 0
	Public SayaPosi(4) As Integer
	Public IndexNo As Integer = 0
	Public JudgeLead(4) As Integer
	Public JudgeDet(4) As Integer
	Public JudgeLen1(4) As Integer
	Public JudgeLen2(4) As Integer
	Public RetryLead(4) As Integer
	Public RetryDet(4) As Integer
	Public RetryLen1(4) As Integer
	Public RetryLen2(4) As Integer
	Public Zaika(4) As Integer
	Public ReDet(4) As String
	Public ReLen1(4) As String
    Public ReLen2(4) As String
    Public ReShift(24) As String
    Public FormatData(100) As String

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

	Public QuData(4, 70) As String             '品質ﾃﾞｰﾀ

	Public StackData(13, 110) As String     '装置 直近n=100個分ﾃﾞｰﾀ
    Public StackCounter As Integer = 0      '装置 ｽﾀｯｸｶｳﾝﾀｰ
	Public QuStackData(500, 12) As String '測定 直近n=100個分ﾃﾞｰﾀ
    Public QuStackCounter As Integer = 0  '測定 ｽﾀｯｸｶｳﾝﾀｰ

    Public EqStartedFlag As Boolean = False '設備ﾃﾞｰﾀ集計中ﾌﾗｸﾞ

    Public PlcReadingFlag As Boolean = False    'PLC通信中ﾌﾗｸﾞ
    Public SaveDataFirstFlag As Boolean = True  '初回ﾃﾞｰﾀ保存ﾌﾗｸﾞ
    Public SaveDataFirstFlagQu As Boolean = True  '初回ﾃﾞｰﾀ保存ﾌﾗｸﾞ

    Public DebugFlag As Boolean = False
    Public tmp0 As Long = 0
    Public TmpLong(20) As Long
    Public TmpInt(299) As Long

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        initialize()
    End Sub

    Private Sub initialize()
        '設備結果ﾃﾞｰﾀｼｰﾄ
        DGVClear(dgvEq)
        Me.Width = 1024
        Me.Height = 768
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        dgvEq.Width = 990
        dgvEq.Height = 263
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
        dgvEq.Columns.Add("5", "検知上")
        dgvEq.Columns.Add("6", "検知横")
        dgvEq.Columns.Add("7", "検知下")
        dgvEq.Columns.Add("8", "全長①上")
        dgvEq.Columns.Add("9", "全長①下")
        dgvEq.Columns.Add("10", "全長②上")
        dgvEq.Columns.Add("11", "全長②横")
        dgvEq.Columns.Add("12", "全長②斜")
        'dgvEq.Columns.Add("0", "素子" & vbCrLf & "品番")
        'dgvEq.Columns.Add("1", "ﾒｯｷﾛｯﾄ" & vbCrLf & "No.  ")
        'dgvEq.Columns.Add("2", "作業者")
        'dgvEq.Columns.Add("3", "仕掛時間")
        'dgvEq.Columns.Add("4", "完了時間")
        'dgvEq.Columns.Add("5", "検知抵抗" & vbCrLf & "上ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        'dgvEq.Columns.Add("6", "検知抵抗" & vbCrLf & "横ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        'dgvEq.Columns.Add("7", "検知抵抗" & vbCrLf & "下ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        'dgvEq.Columns.Add("8", "全長抵抗①" & vbCrLf & "上ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        'dgvEq.Columns.Add("9", "全長抵抗①" & vbCrLf & "下ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        'dgvEq.Columns.Add("10", "全長抵抗②" & vbCrLf & "上ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        'dgvEq.Columns.Add("11", "全長抵抗②" & vbCrLf & "横ﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        'dgvEq.Columns.Add("12", "全長抵抗②" & vbCrLf & "斜めﾌﾟﾛｰﾌﾞ" & vbCrLf & "使用回数")
        For i As Integer = 0 To 4
            dgvEq.Columns(i).DefaultCellStyle = cstyle1
            dgvEq.Columns(i).Width = 62
        Next i
        For i As Integer = 5 To 12
            dgvEq.Columns(i).DefaultCellStyle = cstyle1
            dgvEq.Columns(i).Width = 70
        Next i
        dgvEq.Columns(2).Width = 65
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
        dgvQu.Width = 945
        dgvQu.Height = 410
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
		For i As Integer = 0 To 500
			dgvQu.Rows.Add("")
		Next
        dgvQu.Columns(0).Width = 110
        dgvQu.Columns(1).Width = 70
        dgvQu.Columns(2).Width = 70
        dgvQu.Columns(3).Width = 70
        dgvQu.Columns(5).Width = 85
        dgvQu.Columns(8).Width = 85
        dgvQu.RowHeadersVisible = False
        dgvQu.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvQu.CurrentCell = Nothing         '選択されているセルをなくす

    End Sub

    Private Sub timScan_Tick(sender As Object, e As EventArgs) Handles timScan.Tick
        tmp0 = PlcRead(10410)
        TextBox2.Text = Str(tmp0)
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
            If Not EqStartedFlag Then
                EqStartedFlag = True
                StartTime = Trim(CStr(Now))
                StackCounter += 1
                GetPlcData()
                EndTime = ""
                For i As Integer = 0 To 7
                    ProbeData(i) = 0
                Next
                StackSet()
                DrawChartSetubi()
            End If
        End If
        'エンドトリガ監視
        If PlcRead(EndTriggerAdress) <> 0 Then
            PlcWrite(EndTriggerAdress, 0)
            If EqStartedFlag Then
                EndTime = Trim(CStr(Now))
                GetPlcData()
                StackSet()
                DrawChartSetubi()
                SaveData()
                EqStartedFlag = False
            End If
        End If
        '測定データトリガ監視
        If PlcRead(QuTriggerAdress) <> 0 Then
            PlcWrite(QuTriggerAdress, 0)
            QuStackCounter += 1
            GetPlcData()
            ChFormat()
            QuStackSet()
            DrawChartQu()
            SaveDataQu()
        End If
    End Sub

    Public Sub StartProcess()
        PlcReadingFlag = True
        ElementNo = PlcReadStrings(ElementNoAddress, 8)
        LotNo = PlcReadStrings(LotNoAddress, 8)
        OperatorNo = PlcReadStrings(LotNoAddress, 8)
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
            For i As Short = 0 To 3
                SayaPosi(i) = CInt(TmpInt(101 + i))
            Next
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
                ReDet(i) = HexAsc(Hex(TmpInt(140 + i * 2))) & HexAsc(Hex(TmpInt(141 + i * 2)))
            Next
            For i As Short = 0 To 3
                ReLen1(i) = HexAsc(Hex(TmpInt(148 + i * 2))) & HexAsc(Hex(TmpInt(149 + i * 2)))
            Next
            For i As Short = 0 To 3
                ReLen2(i) = HexAsc(Hex(TmpInt(156 + i * 2))) & HexAsc(Hex(TmpInt(157 + i * 2)))
            Next
            For i As Short = 0 To 23
                ReShift(i) = HexAsc(Hex(TmpInt(170 + i * 2))) & HexAsc(Hex(TmpInt(171 + i * 2)))
            Next
        Else
            ElementNo = "Ele" & Trim(CStr(Int(Rnd(1) * 1000)))
            LotNo = "Lot" & Trim(CStr(Int(Rnd(1) * 1000)))
            OperatorNo = "Ope" & Trim(CStr(Int(Rnd(1) * 1000)))
            For i As Short = 0 To 7
                ProbeData(i) = CLng((Int(Rnd(1) * 100000)))
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
			QuData(i, 5) = ReDet(i)
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
			QuData(i, 8) = ReLen(i)
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
		For i As Integer = 0 To StackCounter
			For j As Integer = 0 To 12
				dgvEq.Item(j, i).Value = StackData(j, i + 1)
			Next
		Next
	End Sub

	Public Sub DrawChartQu()
        For i As Integer = 0 To QuStackCounter
            For j As Integer = 0 To 12
                dgvQu.Item(j, i * 4 + 0).Value = QuStackData((i + 1) * 4 + 0, j)
                dgvQu.Item(j, i * 4 + 1).Value = QuStackData((i + 1) * 4 + 1, j)
                dgvQu.Item(j, i * 4 + 2).Value = QuStackData((i + 1) * 4 + 2, j)
                dgvQu.Item(j, i * 4 + 3).Value = QuStackData((i + 1) * 4 + 3, j)
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
        Dim tmp1(299) As Integer
        If Not DebugFlag Then
            Try
                tmp1 = SysmacCJ.ReadMemoryWordInteger(OMRON.Compolet.SYSMAC.SysmacCJ.MemoryTypes.DM, address, length + 1, OMRON.Compolet.SYSMAC.SysmacCJ.DataTypes.BIN)
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
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "setubi" + Trim(Str(Gouki)) + "_" + SaveFileName + ".CSV", InputString, True)
        My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "setubi" + Trim(Str(Gouki)) + "_" + SaveFileName + ".BKF", InputString, True)
    End Sub

    Public Sub SaveDataQu()
        '起動初回確認
        If SaveDataFirstFlagQu Then
            CreateSaveFolder()
            CreateSaveFileNameQu()
            SaveDataFirstFlagQu = False
        End If
        '現在時刻確認
        Dim NowYearMonth As String = Replace(Strings.Left(CStr(Now), 7), "/", "")
        Dim NowDate As String = Replace(Strings.Left(CStr(Now), 10), "/", "")
        Dim NowTime As String = Replace(Strings.Mid(CStr(Now), 12, 5), ":", "")
        If NowDate <> SaveFileNameQu And Val(NowTime) >= Val(SaveTimeH + SaveTimeM) Then
            If NowYearMonth <> SaveSubFolder Then CreateSaveFolder()
            CreateSaveFileNameQu()
        End If
        'データ保存
        Dim InputString As String = ""
        For i As Integer = 0 To 3
            InputString = ""
            For j As Integer = 0 To 12
                InputString = InputString + QuStackData(QuStackCounter * 4 + i, j) + ","
            Next
            InputString = InputString & vbCrLf
            My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "hinshitu" + Trim(Str(Gouki)) + "_" + SaveFileNameQu + ".CSV", InputString, True)
            My.Computer.FileSystem.WriteAllText(SaveFolder + "\" + SaveSubFolder + "\" + "hinshitu" + Trim(Str(Gouki)) + "_" + SaveFileNameQu + ".BKF", InputString, True)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmDebug.Show()
    End Sub
End Class

