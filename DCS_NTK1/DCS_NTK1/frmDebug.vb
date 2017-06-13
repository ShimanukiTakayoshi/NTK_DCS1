Public Class frmDebug

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Not frmMain.PlcReadingFlag Then
            frmMain.PlcReadingFlag = True
            frmMain.GetPlcData()
            frmMain.PlcReadingFlag = False
            TextBox1.Text = frmMain.ElementNo
            TextBox2.Text = frmMain.LotNo
            TextBox3.Text = frmMain.OperatorNo
            TextBox4.Text = CStr(frmMain.ProbeData(0))
            TextBox5.Text = CStr(frmMain.ProbeData(1))
            TextBox6.Text = CStr(frmMain.ProbeData(2))
            TextBox7.Text = CStr(frmMain.ProbeData(3))
            TextBox8.Text = CStr(frmMain.ProbeData(4))
            TextBox9.Text = CStr(frmMain.ProbeData(5))
            TextBox10.Text = CStr(frmMain.ProbeData(6))
            TextBox11.Text = CStr(frmMain.ProbeData(7))
            TextBox12.Text = CStr(frmMain.SayaNo)
            TextBox13.Text = CStr(frmMain.SayaPosi(0))
            TextBox14.Text = CStr(frmMain.SayaPosi(1))
            TextBox15.Text = CStr(frmMain.SayaPosi(2))
            TextBox16.Text = CStr(frmMain.SayaPosi(3))
            TextBox17.Text = CStr(frmMain.IndexNo)
            TextBox18.Text = CStr(frmMain.JudgeLead(0)) & CStr(frmMain.JudgeLead(1)) & CStr(frmMain.JudgeLead(2)) & CStr(frmMain.JudgeLead(3))
            TextBox19.Text = CStr(frmMain.JudgeDet(0)) & CStr(frmMain.JudgeDet(1)) & CStr(frmMain.JudgeDet(2)) & CStr(frmMain.JudgeDet(3))
            TextBox20.Text = CStr(frmMain.JudgeLen1(0)) & CStr(frmMain.JudgeLen1(1)) & CStr(frmMain.JudgeLen1(2)) & CStr(frmMain.JudgeLen1(3))
            TextBox21.Text = CStr(frmMain.JudgeLen2(0)) & CStr(frmMain.JudgeLen2(1)) & CStr(frmMain.JudgeLen2(2)) & CStr(frmMain.JudgeLen2(3))
            TextBox22.Text = CStr(frmMain.RetryDet(0)) & CStr(frmMain.RetryDet(1)) & CStr(frmMain.RetryDet(2)) & CStr(frmMain.RetryDet(3))
            TextBox23.Text = CStr(frmMain.RetryLen1(0)) & CStr(frmMain.RetryLen1(1)) & CStr(frmMain.RetryLen1(2)) & CStr(frmMain.RetryLen1(3))
            TextBox24.Text = CStr(frmMain.RetryLen2(0)) & CStr(frmMain.RetryLen2(1)) & CStr(frmMain.RetryLen2(2)) & CStr(frmMain.RetryLen2(3))
            TextBox25.Text = CStr(frmMain.Zaika(0)) & CStr(frmMain.Zaika(1)) & CStr(frmMain.Zaika(2)) & CStr(frmMain.Zaika(3))
            TextBox26.Text = CStr(frmMain.ReDet(0))
            TextBox27.Text = CStr(frmMain.ReDet(1))
            TextBox28.Text = CStr(frmMain.ReDet(2))
            TextBox29.Text = CStr(frmMain.ReDet(3))
            TextBox30.Text = CStr(frmMain.ReLen1(0))
            TextBox31.Text = CStr(frmMain.ReLen1(1))
            TextBox32.Text = CStr(frmMain.ReLen1(2))
            TextBox33.Text = CStr(frmMain.ReLen1(3))
            TextBox34.Text = CStr(frmMain.ReLen2(0))
            TextBox35.Text = CStr(frmMain.ReLen2(1))
            TextBox36.Text = CStr(frmMain.ReLen2(2))
            TextBox37.Text = CStr(frmMain.ReLen2(3))



            'For i As Short = 0 To 3
            '    ReDet(i) = HexAsc(Hex(TmpInt(140 + i * 2))) & HexAsc(Hex(TmpInt(141 + i * 2)))
            'Next
            'For i As Short = 0 To 3
            '    ReLen1(i) = HexAsc(Hex(TmpInt(148 + i * 2))) & HexAsc(Hex(TmpInt(149 + i * 2)))
            'Next
            'For i As Short = 0 To 3
            '    ReLen2(i) = HexAsc(Hex(TmpInt(156 + i * 2))) & HexAsc(Hex(TmpInt(157 + i * 2)))
            'Next
        End If
    End Sub

 
End Class