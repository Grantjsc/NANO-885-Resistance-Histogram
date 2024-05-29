Imports System.Net.Mime.MediaTypeNames
Imports System.Text.RegularExpressions
Imports System.Threading

Public Class Form1

    Sub Form1test_result()
        Invoke(Sub()
                   Reading = CDec(txtTegamValue.Text)
                   'Reading = CDec(Tegam)
                   If Reading <= nominal * 2 Then
                       store_res(CDec(txtTested.Text)) = CDec(txtTegamValue.Text)
                       'runcount += 1
                   End If
                   runcount += 1

                   If Reading <= CDec(txtHiLimit.Text) And Reading >= CDec(txtLoLimit.Text) Then
                       txtResult.Text = "Good"
                       txtResult.FillColor = Color.Lime
                       txtTegamValue.FillColor = Color.Lime
                       txtGood.Text += 1
                   End If
                   If Reading > CDec(txtHiLimit.Text) Then
                       If Reading > nominal * 2 Then
                           txtResult.Text = "Open"
                           txtResult.FillColor = Color.Red
                           txtTegamValue.FillColor = Color.Red
                           txtOpen.Text += 1
                       Else
                           txtResult.Text = "High"
                           txtResult.FillColor = Color.Red
                           txtTegamValue.FillColor = Color.Red
                           txtHigh.Text += 1
                       End If
                   End If

                   If Reading < CDec(txtLoLimit.Text) Then
                       txtResult.Text = "Low"
                       txtResult.FillColor = Color.Red
                       txtTegamValue.FillColor = Color.Red
                       txtLow.Text += 1
                   End If
                   txtTested.Text += 1
                   txtYield.Text = Math.Round((CDec(txtGood.Text) / CDec(txtTested.Text)) * 100, 2) & " %"
                   txtRej.Text = CDec(txtHigh.Text) + CDec(txtOpen.Text) + CDec(txtLow.Text)
               End Sub)
    End Sub

    Sub Form1Prog_count()
        Invoke(Sub()
                   Reading = CDec(txtTegamValue.Text)
                   'Reading = CDec(Tegam)
                   If Reading > CDec(Label4.Text) Then
                       lblValue1.Text += 1
                       Progress1.Value += 1
                       If CDec(lblValue1.Text) >= CDec(Progress1.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading > CDec(Label5.Text) Then
                       lblValue2.Text += 1
                       Progress2.Value += 1
                       If CDec(lblValue2.Text) >= CDec(Progress2.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading > CDec(Label6.Text) Then
                       lblValue3.Text += 1
                       Progress3.Value += 1
                       If CDec(lblValue3.Text) >= CDec(Progress3.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading > CDec(Label7.Text) Then
                       lblValue4.Text += 1
                       Progress4.Value += 1
                       If CDec(lblValue4.Text) >= CDec(Progress4.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading > CDec(Label8.Text) Then
                       lblValue5.Text += 1
                       Progress5.Value += 1
                       If CDec(lblValue5.Text) >= CDec(Progress5.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading > CDec(Label9.Text) Then
                       lblValue6.Text += 1
                       Progress6.Value += 1
                       If CDec(lblValue6.Text) >= CDec(Progress6.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading > CDec(Label10.Text) Then
                       lblValue7.Text += 1
                       Progress7.Value += 1
                       If CDec(lblValue7.Text) >= CDec(Progress7.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading > CDec(Label11.Text) Then
                       lblValue8.Text += 1
                       Progress8.Value += 1
                       If CDec(lblValue8.Text) >= CDec(Progress8.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading > CDec(Label12.Text) Then
                       lblValue9.Text += 1
                       Progress9.Value += 1
                       If CDec(lblValue9.Text) >= CDec(Progress9.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading > CDec(Label13.Text) Then
                       lblValue10.Text += 1
                       Progress10.Value += 1
                       If CDec(lblValue10.Text) >= CDec(Progress10.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading > (CDec(Label14.Text) + splitval) Then
                       lblValue11.Text += 1
                       Progress11.Value += 1
                       If CDec(lblValue11.Text) >= CDec(Progress11.Maximum) Then
                           Prog_increase()
                       End If

                       '**************************************NOMINAL**************************************
                       'ElseIf Reading < CDec(Label13.Text) Then
                       '    lblValue13.Text += 1
                       '    Progress13.Value += 1
                       '    If CDec(lblValue13.Text) >= CDec(Progress13.Maximum) Then
                       '        Prog_increase()
                       '    End If
                   ElseIf Reading >= (CDec(Label14.Text) - splitval) And Reading <= (CDec(Label14.Text) + splitval) Then
                       lblValue12.Text += 1
                       Progress12.Value += 1
                       If CDec(lblValue12.Text) >= CDec(Progress12.Maximum) Then
                           Prog_increase()
                       End If
                       '**************************************NOMINAL**************************************\

                   ElseIf Reading < CDec(Label24.Text) Then
                       lblValue23.Text += 1
                       Progress23.Value += 1
                       If CDec(lblValue23.Text) >= CDec(Progress23.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading < CDec(Label23.Text) Then
                       lblValue22.Text += 1
                       Progress22.Value += 1
                       If CDec(lblValue22.Text) >= CDec(Progress22.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading < CDec(Label22.Text) Then
                       lblValue21.Text += 1
                       Progress21.Value += 1
                       If CDec(lblValue21.Text) >= CDec(Progress21.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading < CDec(Label21.Text) Then
                       lblValue20.Text += 1
                       Progress20.Value += 1
                       If CDec(lblValue20.Text) >= CDec(Progress20.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading < CDec(Label20.Text) Then
                       lblValue19.Text += 1
                       Progress19.Value += 1
                       If CDec(lblValue19.Text) >= CDec(Progress19.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading < CDec(Label19.Text) Then
                       lblValue18.Text += 1
                       Progress18.Value += 1
                       If CDec(lblValue18.Text) >= CDec(Progress18.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading < CDec(Label18.Text) Then
                       lblValue17.Text += 1
                       Progress17.Value += 1
                       If CDec(lblValue17.Text) >= CDec(Progress17.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading < CDec(Label17.Text) Then
                       lblValue16.Text += 1
                       Progress16.Value += 1
                       If CDec(lblValue16.Text) >= CDec(Progress16.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading < CDec(Label16.Text) Then
                       lblValue15.Text += 1
                       Progress15.Value += 1
                       If CDec(lblValue15.Text) >= CDec(Progress15.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading < CDec(Label15.Text) Then
                       lblValue14.Text += 1
                       Progress14.Value += 1
                       If CDec(lblValue14.Text) >= CDec(Progress14.Maximum) Then
                           Prog_increase()
                       End If
                   ElseIf Reading < (CDec(Label14.Text) - splitval) Then
                       lblValue13.Text += 1
                       Progress13.Value += 1
                       If CDec(lblValue13.Text) >= CDec(Progress13.Maximum) Then
                           Prog_increase()
                       End If
                   End If
               End Sub)
    End Sub

    Sub Form1statistics()
        average = 0
        For s As Integer = 0 To store_res.Length - 1
            If store_res(s) = "" Then
                If average > 0 Then
                    average = Math.Round(average / s, 2)
                    stdev = 0
                    For t As Integer = 0 To store_res.Length - 1
                        If store_res(t) = "" Then
                            If stdev > 0 Then
                                stdev = Math.Sqrt(CDec(stdev) / CDec(t))
                                cpk_max = (maximum - average) / (stdev * 3)
                                cpk_min = (average - minimum) / (stdev * 3)
                                txtCpk.Text = CStr(Math.Round(Math.Min(cpk_max, cpk_min), 2))
                                txtStdev.Text = CStr(Math.Round(stdev, 2))
                            End If
                            Exit For
                        Else
                            stdev = ((CDec(store_res(t)) - average) ^ 2) + stdev
                        End If
                    Next
                End If
                Exit For
            Else
                average = average + store_res(s)
            End If
        Next
    End Sub

    Private Sub txtPartNo_KeyUp(sender As Object, e As KeyEventArgs) Handles txtPartNo.KeyUp
        If e.KeyCode = Keys.Enter Then
            'EnterPartNumber()
            'txtLotNumber.Focus()

            CheckDatabase()
        End If
    End Sub

    Private Sub txtLotNumber_KeyUp(sender As Object, e As KeyEventArgs) Handles txtLotNumber.KeyUp
        If e.KeyCode = Keys.Enter Then
            LogStartTime()
            LogLotforCCD1()
            LogLotforCCD2()
            LogLotforCCD3()
            LogLotforCCD4()
            LogLotforCCD5()
            LogLotforCCD6()
            txtQty.Focus()
        End If
    End Sub

    Private Sub SerialPort1_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        'SerialPort1Receieved()

        TegamRead = SerialPort1.ReadLine
        TegamRead = TegamRead.Replace(" ", "")

        Dim regex As New Regex("-?\d+(\.\d+)?")
        Dim matches As MatchCollection = regex.Matches(TegamRead)
        For Each match As Match In matches
            Tegam = match.Value
        Next

        Invoke(Sub()
                   txtTegamValue.Text = Tegam
               End Sub)
        'Console.WriteLine(Tegam)
        Thread.Sleep(50)

        Form1test_result() 'GOOD/HIGH/LOW or OPEN 
        Form1Prog_count() 'histogram counter
        Form1statistics() 'comp of CPK and Standard Dev
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1Function()
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        btnStartFunction()
    End Sub

    Private Sub btnLot_Click(sender As Object, e As EventArgs) Handles btnLot.Click
        clear()
    End Sub

    Private Sub btnChange_Click(sender As Object, e As EventArgs) Handles btnChange.Click
        change_device()
    End Sub

    Private Sub btnClear2_Click(sender As Object, e As EventArgs) Handles btnClear2.Click
        change_device()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormLoad()
        'Call CenterToScreen()
        'Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub ViewLotDetailsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewLotDetailsToolStripMenuItem.Click
        ViewLotDetails()
    End Sub

    Private Sub TimerQtyChecking_Tick(sender As Object, e As EventArgs) Handles TimerQtyChecking.Tick
        If CInt(txtQty.Text) = CInt(txtGood.Text) Then
            TimerQtyChecking.Enabled = False
            ChangeToStart()
            btnStart.Enabled = False
            LogEndTime() 'Log the end Time
            CheckTable() 'Check table in db then save all results to db and export data to .csv file
        End If
    End Sub

    Private Sub txtQty_KeyUp(sender As Object, e As KeyEventArgs) Handles txtQty.KeyUp
        If e.KeyCode = Keys.Enter Then
            txtQty.ReadOnly = True
            btnStart.Enabled = True
            btnChange.Enabled = True
            btnLot.Enabled = True
            btnClear2.Enabled = False
        End If
    End Sub

    Private Sub txtQty_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQty.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If (Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57) Then
                e.Handled = True
                'MessageBox.Show("Please enter numeric value!", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If SerialPort1.IsOpen Then
            setTegam()
            SetHighLowAndNominal()
            SerialPort1.Write("E" & vbCr)
        Else
            MsgBox("Please input part number", MsgBoxStyle.Critical)
        End If

        'test_result() 'GOOD/HIGH/LOW or OPEN 
        'Prog_count() 'histogram counter
        'statistics() 'comp of CPK and Standard Dev
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim dialog As DialogResult
        dialog = MessageBox.Show("Do you really want to exit?", "NANO 885 Resistance Histogram", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If dialog = DialogResult.No Then
            e.Cancel = True
        Else
            'Application.ExitThread()
            If SerialPort1.IsOpen Then
                SerialPort1.Close()
            End If
            End
        End If
    End Sub

    Private Sub btnEMsave_Click(sender As Object, e As EventArgs) Handles btnEMsave.Click
        EmergencyExport()
    End Sub

    Private Sub ChangeSerialNameToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ChangeSerialNameToolStripMenuItem1.Click
        ChangeSerialCom()
    End Sub

    Private Sub ChangeDbTableReferenceToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ChangeDbTableReferenceToolStripMenuItem1.Click
        ChangeTableReferenceOfDb()
    End Sub

    Private Sub ChangePathOfPartNumberToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangePathOfPartNumberToolStripMenuItem.Click
        ChangePathMenuStrip()
    End Sub

    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
        TimerQtyChecking.Enabled = False
        ChangeToStart()
        btnStart.Enabled = False
        LogEndTime() 'Log the end Time
        CheckTable() 'Check table in db then save all results to db and export data to .csv file
    End Sub

    Private Sub CHnageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CHnageToolStripMenuItem.Click
        ChangeTiming_Form.ShowDialog()
    End Sub
End Class
