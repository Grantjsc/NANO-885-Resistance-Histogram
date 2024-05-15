Imports System.Configuration
Imports System.Data.Entity.Core.Common.EntitySql
Imports System.Data.OleDb
Imports System.Data.SQLite
Imports System.IO
Imports System.IO.Ports
Imports System.Threading
Imports LFPHWIADBLib

'********* NANO 885 Part Number List ********

'0885001.DR
'08851.25DR
'088501.6DR
'0885002.DR
'088502.5DR
'08853.15DR
'0885004.DR
'0885005.DR

'********* TEGAM Connection ********
' Femaleale to Male dsub9
' J 4-6        not sure J 7-8
'5             9
'2             2
'3             3
'8             4
'7             5


Module Form1Histogram_Module
    'Private NANO_db As New WIADb("Data Source=BTMESSQLPROD;Initial Catalog=LFPHNANO;User ID=mesaccount;Password=superfuse;Persist Security Info=True;")
    Public Tegam As String
    Public TegamRead As String
    Public Reading As Decimal
    Public store_res(50000) As String
    Public runcount As Integer

    Sub FormLoad()
        'Form1.WindowState = FormWindowState.Maximized
        Form1.btnStart.Enabled = False
        Form1.btnChange.Enabled = False
        Form1.btnLot.Enabled = False
        'System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False


        GetSerial1()
        GetSerial2()
        Form1.SerialPort1.PortName = TegamCOMname
        Form1.SerialPort2.PortName = TriggerCOMname

        'Form1.SerialPort1.Open()
        'Form1.SerialPort2.Open()
        labels()
    End Sub
    Sub labels()
        Form1.Label4.Text = "0"
        Form1.Label5.Text = "0"
        Form1.Label6.Text = "0"
        Form1.Label7.Text = "0"
        Form1.Label8.Text = "0"
        Form1.Label9.Text = "0"
        Form1.Label10.Text = "0"
        Form1.Label11.Text = "0"
        Form1.Label12.Text = "0"
        Form1.Label13.Text = "0"
        Form1.Label14.Text = "0"
        Form1.Label15.Text = "0"
        Form1.Label16.Text = "0"
        Form1.Label17.Text = "0"
        Form1.Label18.Text = "0"
        Form1.Label19.Text = "0"
        Form1.Label20.Text = "0"
        Form1.Label21.Text = "0"
        Form1.Label22.Text = "0"
        Form1.Label23.Text = "0"
        Form1.Label24.Text = "0"
        Form1.lblRange.Text = ""
    End Sub

    Sub labelResult()
        Form1.lblValue1.Text = "0"
        Form1.lblValue2.Text = "0"
        Form1.lblValue3.Text = "0"
        Form1.lblValue4.Text = "0"
        Form1.lblValue5.Text = "0"
        Form1.lblValue6.Text = "0"
        Form1.lblValue7.Text = "0"
        Form1.lblValue8.Text = "0"
        Form1.lblValue9.Text = "0"
        Form1.lblValue10.Text = "0"
        Form1.lblValue11.Text = "0"
        Form1.lblValue12.Text = "0"
        Form1.lblValue13.Text = "0"
        Form1.lblValue14.Text = "0"
        Form1.lblValue15.Text = "0"
        Form1.lblValue16.Text = "0"
        Form1.lblValue17.Text = "0"
        Form1.lblValue18.Text = "0"
        Form1.lblValue19.Text = "0"
        Form1.lblValue20.Text = "0"
        Form1.lblValue21.Text = "0"
        Form1.lblValue22.Text = "0"
        Form1.lblValue23.Text = "0"
    End Sub

    Sub ProgBar()
        Form1.Progress1.Value = 0
        Form1.Progress2.Value = 0
        Form1.Progress3.Value = 0
        Form1.Progress4.Value = 0
        Form1.Progress5.Value = 0
        Form1.Progress6.Value = 0
        Form1.Progress7.Value = 0
        Form1.Progress8.Value = 0
        Form1.Progress9.Value = 0
        Form1.Progress10.Value = 0
        Form1.Progress11.Value = 0
        Form1.Progress12.Value = 0
        Form1.Progress13.Value = 0
        Form1.Progress14.Value = 0
        Form1.Progress15.Value = 0
        Form1.Progress16.Value = 0
        Form1.Progress17.Value = 0
        Form1.Progress18.Value = 0
        Form1.Progress19.Value = 0
        Form1.Progress20.Value = 0
        Form1.Progress21.Value = 0
        Form1.Progress22.Value = 0
        Form1.Progress23.Value = 0

        'Maximum value
        Form1.Progress1.Maximum = 100
        Form1.Progress2.Maximum = 100
        Form1.Progress3.Maximum = 100
        Form1.Progress4.Maximum = 100
        Form1.Progress5.Maximum = 100
        Form1.Progress6.Maximum = 100
        Form1.Progress7.Maximum = 100
        Form1.Progress8.Maximum = 100
        Form1.Progress9.Maximum = 100
        Form1.Progress10.Maximum = 100
        Form1.Progress11.Maximum = 100
        Form1.Progress12.Maximum = 100
        Form1.Progress13.Maximum = 100
        Form1.Progress14.Maximum = 100
        Form1.Progress15.Maximum = 100
        Form1.Progress16.Maximum = 100
        Form1.Progress17.Maximum = 100
        Form1.Progress18.Maximum = 100
        Form1.Progress19.Maximum = 100
        Form1.Progress20.Maximum = 100
        Form1.Progress21.Maximum = 100
        Form1.Progress22.Maximum = 100
        Form1.Progress23.Maximum = 100
    End Sub

    Sub Prog_increase()
        Form1.Progress1.Maximum += 50
        Form1.Progress2.Maximum += 50
        Form1.Progress3.Maximum += 50
        Form1.Progress4.Maximum += 50
        Form1.Progress5.Maximum += 50
        Form1.Progress6.Maximum += 50
        Form1.Progress7.Maximum += 50
        Form1.Progress8.Maximum += 50
        Form1.Progress9.Maximum += 50
        Form1.Progress10.Maximum += 50
        Form1.Progress11.Maximum += 50
        Form1.Progress12.Maximum += 50
        Form1.Progress13.Maximum += 50
        Form1.Progress14.Maximum += 50
        Form1.Progress15.Maximum += 50
        Form1.Progress16.Maximum += 50
        Form1.Progress17.Maximum += 50
        Form1.Progress18.Maximum += 50
        Form1.Progress19.Maximum += 50
        Form1.Progress20.Maximum += 50
        Form1.Progress21.Maximum += 50
        Form1.Progress22.Maximum += 50
        Form1.Progress23.Maximum += 50
    End Sub

    Sub Prog_count()
        Reading = CDec(Form1.txtTegamValue.Text)
        'Reading = CDec(Tegam)
        If Reading > CDec(Form1.Label4.Text) Then
            Form1.lblValue1.Text += 1
            Form1.Progress1.Value += 1
            If CDec(Form1.lblValue1.Text) >= CDec(Form1.Progress1.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Form1.Label5.Text) Then
            Form1.lblValue2.Text += 1
            Form1.Progress2.Value += 1
            If CDec(Form1.lblValue2.Text) >= CDec(Form1.Progress2.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Form1.Label6.Text) Then
            Form1.lblValue3.Text += 1
            Form1.Progress3.Value += 1
            If CDec(Form1.lblValue3.Text) >= CDec(Form1.Progress3.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Form1.Label7.Text) Then
            Form1.lblValue4.Text += 1
            Form1.Progress4.Value += 1
            If CDec(Form1.lblValue4.Text) >= CDec(Form1.Progress4.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Form1.Label8.Text) Then
            Form1.lblValue5.Text += 1
            Form1.Progress5.Value += 1
            If CDec(Form1.lblValue5.Text) >= CDec(Form1.Progress5.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Form1.Label9.Text) Then
            Form1.lblValue6.Text += 1
            Form1.Progress6.Value += 1
            If CDec(Form1.lblValue6.Text) >= CDec(Form1.Progress6.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Form1.Label10.Text) Then
            Form1.lblValue7.Text += 1
            Form1.Progress7.Value += 1
            If CDec(Form1.lblValue7.Text) >= CDec(Form1.Progress7.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Form1.Label11.Text) Then
            Form1.lblValue8.Text += 1
            Form1.Progress8.Value += 1
            If CDec(Form1.lblValue8.Text) >= CDec(Form1.Progress8.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Form1.Label12.Text) Then
            Form1.lblValue9.Text += 1
            Form1.Progress9.Value += 1
            If CDec(Form1.lblValue9.Text) >= CDec(Form1.Progress9.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > CDec(Form1.Label13.Text) Then
            Form1.lblValue10.Text += 1
            Form1.Progress10.Value += 1
            If CDec(Form1.lblValue10.Text) >= CDec(Form1.Progress10.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading > (CDec(Form1.Label14.Text) + splitval) Then
            Form1.lblValue11.Text += 1
            Form1.Progress11.Value += 1
            If CDec(Form1.lblValue11.Text) >= CDec(Form1.Progress11.Maximum) Then
                Prog_increase()
            End If

            '**************************************NOMINAL**************************************
            'ElseIf Reading < CDec(Label13.Text) Then
            '    lblValue13.Text += 1
            '    Progress13.Value += 1
            '    If CDec(lblValue13.Text) >= CDec(Progress13.Maximum) Then
            '        Prog_increase()
            '    End If
        ElseIf Reading >= (CDec(Form1.Label14.Text) - splitval) And Reading <= (CDec(Form1.Label14.Text) + splitval) Then
            Form1.lblValue12.Text += 1
            Form1.Progress12.Value += 1
            If CDec(Form1.lblValue12.Text) >= CDec(Form1.Progress12.Maximum) Then
                Prog_increase()
            End If
            '**************************************NOMINAL**************************************\

        ElseIf Reading < CDec(Form1.Label24.Text) Then
            Form1.lblValue23.Text += 1
            Form1.Progress23.Value += 1
            If CDec(Form1.lblValue23.Text) >= CDec(Form1.Progress23.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Form1.Label23.Text) Then
            Form1.lblValue22.Text += 1
            Form1.Progress22.Value += 1
            If CDec(Form1.lblValue22.Text) >= CDec(Form1.Progress22.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Form1.Label22.Text) Then
            Form1.lblValue21.Text += 1
            Form1.Progress21.Value += 1
            If CDec(Form1.lblValue21.Text) >= CDec(Form1.Progress21.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Form1.Label21.Text) Then
            Form1.lblValue20.Text += 1
            Form1.Progress20.Value += 1
            If CDec(Form1.lblValue20.Text) >= CDec(Form1.Progress20.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Form1.Label20.Text) Then
            Form1.lblValue19.Text += 1
            Form1.Progress19.Value += 1
            If CDec(Form1.lblValue19.Text) >= CDec(Form1.Progress19.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Form1.Label19.Text) Then
            Form1.lblValue18.Text += 1
            Form1.Progress18.Value += 1
            If CDec(Form1.lblValue18.Text) >= CDec(Form1.Progress18.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Form1.Label18.Text) Then
            Form1.lblValue17.Text += 1
            Form1.Progress17.Value += 1
            If CDec(Form1.lblValue17.Text) >= CDec(Form1.Progress17.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Form1.Label17.Text) Then
            Form1.lblValue16.Text += 1
            Form1.Progress16.Value += 1
            If CDec(Form1.lblValue16.Text) >= CDec(Form1.Progress16.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Form1.Label16.Text) Then
            Form1.lblValue15.Text += 1
            Form1.Progress15.Value += 1
            If CDec(Form1.lblValue15.Text) >= CDec(Form1.Progress15.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < CDec(Form1.Label15.Text) Then
            Form1.lblValue14.Text += 1
            Form1.Progress14.Value += 1
            If CDec(Form1.lblValue14.Text) >= CDec(Form1.Progress14.Maximum) Then
                Prog_increase()
            End If
        ElseIf Reading < (CDec(Form1.Label14.Text) - splitval) Then
            Form1.lblValue13.Text += 1
            Form1.Progress13.Value += 1
            If CDec(Form1.lblValue13.Text) >= CDec(Form1.Progress13.Maximum) Then
                Prog_increase()
            End If
        End If
    End Sub

    Sub test_result()
        Reading = CDec(Form1.txtTegamValue.Text)
        'Reading = CDec(Tegam)
        If Reading <= nominal * 2 Then
            store_res(CDec(Form1.txtTested.Text)) = CDec(Form1.txtTegamValue.Text)
            'runcount += 1
        End If
        runcount += 1

        If Reading <= CDec(Form1.txtHiLimit.Text) And Reading >= CDec(Form1.txtLoLimit.Text) Then
            Form1.txtResult.Text = "Good"
            Form1.txtResult.FillColor = Color.Lime
            Form1.txtTegamValue.FillColor = Color.Lime
            Form1.txtGood.Text += 1
        End If
        If Reading > CDec(Form1.txtHiLimit.Text) Then
            If Reading > nominal * 2 Then
                Form1.txtResult.Text = "Open"
                Form1.txtResult.FillColor = Color.Red
                Form1.txtTegamValue.FillColor = Color.Red
                Form1.txtOpen.Text += 1
            Else
                Form1.txtResult.Text = "High"
                Form1.txtResult.FillColor = Color.Red
                Form1.txtTegamValue.FillColor = Color.Red
                Form1.txtHigh.Text += 1
            End If
        End If

        If Reading < CDec(Form1.txtLoLimit.Text) Then
            Form1.txtResult.Text = "Low"
            Form1.txtResult.FillColor = Color.Red
            Form1.txtTegamValue.FillColor = Color.Red
            Form1.txtLow.Text += 1
        End If
        Form1.txtTested.Text += 1
        Form1.txtYield.Text = Math.Round((CDec(Form1.txtGood.Text) / CDec(Form1.txtTested.Text)) * 100, 2) & " %"
        Form1.txtRej.Text = CDec(Form1.txtHigh.Text) + CDec(Form1.txtOpen.Text) + CDec(Form1.txtLow.Text)
    End Sub

    Public maximum As Decimal
    Const limit200ohm = 18000
    Const limit20ohm = 1800
    Const limit2ohm = 180
    Const limit200mohm = 18
    Const limit20mohm = 1.8
    Const limit2mohm = 0.18
    Public TegamRange As String

    Sub range()
        If maximum > limit20ohm Then
            If NANO_dbRating >= 0.1 Then
                Form1.lblRange.Text = "R8 = 20Ω range @ 10mA Test Current"
                'TegamRange = "R8"
                TegamRange = DbSetRange
            Else
                Form1.lblRange.Text = "R9 = 20Ω range @ 1mA Test Current"
                'TegamRange = "R9"
                TegamRange = DbSetRange
            End If
            multiplier = 1000
        ElseIf maximum > limit2ohm Then
            If NANO_dbRating >= 1 Then
                Form1.lblRange.Text = "R6 = 2Ω range @ 100mA Test Current"
                'TegamRange = "R6"
                TegamRange = DbSetRange
                multiplier = 1000
            ElseIf NANO_dbRating >= 0.1 Then
                Form1.lblRange.Text = "R7 = 2Ω range @ 10mA Test Current"
                'TegamRange = "R7"
                TegamRange = DbSetRange
                multiplier = 1000
            Else
                Form1.lblRange.Text = "R9 = 20Ω range @ 1mA Test Current"
                'TegamRange = "R9"
                TegamRange = DbSetRange
                multiplier = 1000
            End If
        ElseIf maximum > limit200mohm Then      ' Go to 200 mohm range
            If NANO_dbRating >= 1 Then
                Form1.lblRange.Text = "R5 = 200mΩ range @ 100mA Test Current"
                'TegamRange = "R5"
                TegamRange = DbSetRange
                multiplier = 1
            ElseIf NANO_dbRating >= 10 Then
                Form1.lblRange.Text = "R4 = 200mΩ range @ 1mA Test Current"
                'TegamRange = "R4"
                TegamRange = DbSetRange
                multiplier = 1
            Else
                Form1.lblRange.Text = "R7 = 2Ω range @ 10mA Test Current"
                'TegamRange = "R7"
                TegamRange = DbSetRange
                multiplier = 1000
            End If
        ElseIf maximum > limit20mohm Then     ' Go to 20 mohm range
            If NANO_dbRating >= 10 Then
                Form1.lblRange.Text = "R2 = 20mΩ range @ 1A Test Current"
                'TegamRange = "R2"
                TegamRange = DbSetRange
                multiplier = 1
            ElseIf NANO_dbRating >= 1 Then
                Form1.lblRange.Text = "R3 = 20mΩ range @ 100mA Test Current"
                'TegamRange = "R3"
                TegamRange = DbSetRange
                multiplier = 1
            Else
                Form1.lblRange.Text = "R8 = 2Ω range @ 10mA Test Current"
                'TegamRange = "R8"
                TegamRange = DbSetRange
                multiplier = 1000
            End If
        Else                          ' Go to 2 mohm range
            If NANO_dbRating >= 10 Then
                Form1.lblRange.Text = "R1 = 2Ω range @ 1A Test Current"
                'TegamRange = "R1"
                TegamRange = DbSetRange
                multiplier = 1
            Else
                Form1.lblRange.Text = "R3 = 20mΩ range @ 100mA Test Current"
                'TegamRange = "R3"
                TegamRange = DbSetRange
                multiplier = 1
            End If
        End If

        setTegam()
    End Sub

    Sub setTegam()
        Try
            Form1.SerialPort1.Write("P2" & vbCr) 'display mode
            Form1.SerialPort1.Write("X" & vbCr)
            Form1.SerialPort1.Write(TegamRange & vbCr) 'set range
            Form1.SerialPort1.Write("X" & vbCr)
            Form1.SerialPort1.Write("T3" & vbCr) 'delay one shot on talk
            Form1.SerialPort1.Write("X" & vbCr)
            Form1.SerialPort1.Write("D001" & vbCr) 'delay set to 1ms
            Form1.SerialPort1.Write("X" & vbCr)
            Thread.Sleep(10)
        Catch ex As Exception

            MsgBox(ex.Message, vbCritical)

        End Try
    End Sub

    Sub SerialPort1Receieved()
        TegamRead = Form1.SerialPort1.ReadLine
        Tegam = TegamRead.Substring(0, 6)
        'Form1.txtTegamValue.Text = TegamRead
        Console.WriteLine(TegamRead)

        ''Reading = Tegam
        ''txtTegamValue.Text = Reading
        'test_result() 'GOOD/HIGH/LOW or OPEN 
        'Prog_count() 'histogram counter
        'statistics() 'comp of CPK and Standard Dev
    End Sub

    Sub set_val()
        'maximum = (PICO_db.Maximum / multiplier)
        'minimum = (PICO_db.Minimum / multiplier)
        add = (CDec(Form1.txtHiLimit.Text) - nominal) / 9
        reduce = (nominal - CDec(Form1.txtLoLimit.Text)) / 9

        Form1.Label4.Text = Math.Round(nominal + (add * 10), 3)
        Form1.Label5.Text = Math.Round(nominal + (add * 9), 3)
        Form1.Label6.Text = Math.Round(nominal + (add * 8), 3)
        Form1.Label7.Text = Math.Round(nominal + (add * 7), 3)
        Form1.Label8.Text = Math.Round(nominal + (add * 6), 3)
        Form1.Label9.Text = Math.Round(nominal + (add * 5), 3)
        Form1.Label10.Text = Math.Round(nominal + (add * 4), 3)
        Form1.Label11.Text = Math.Round(nominal + (add * 3), 3)
        Form1.Label12.Text = Math.Round(nominal + (add * 2), 3)
        Form1.Label13.Text = Math.Round(nominal + (add * 1), 3)
        Form1.Label14.Text = nominal
        Form1.Label15.Text = Math.Round(nominal - (reduce * 1), 3)
        Form1.Label16.Text = Math.Round(nominal - (reduce * 2), 3)
        Form1.Label17.Text = Math.Round(nominal - (reduce * 3), 3)
        Form1.Label18.Text = Math.Round(nominal - (reduce * 4), 3)
        Form1.Label19.Text = Math.Round(nominal - (reduce * 5), 3)
        Form1.Label20.Text = Math.Round(nominal - (reduce * 6), 3)
        Form1.Label21.Text = Math.Round(nominal - (reduce * 7), 3)
        Form1.Label22.Text = Math.Round(nominal - (reduce * 8), 3)
        Form1.Label23.Text = Math.Round(nominal - (reduce * 9), 3)
        Form1.Label24.Text = Math.Round(nominal - (reduce * 10), 3)

        splitval = Math.Round((CDec(Form1.Label13.Text) - CDec(Form1.Label15.Text)) / 3, 3)
        Console.WriteLine(splitval)
    End Sub

    Public dbPath As String
    Sub GetPathforDB()
        Select Case Form1.txtPartNo.Text

            Case "0885001.DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("0885001.DR")
                Console.WriteLine(Path)
                dbPath = Path

            Case "08851.25DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("08851.25DR")
                Console.WriteLine(Path)
                dbPath = Path

            Case "088501.6DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("088501.6DR")
                Console.WriteLine(Path)
                dbPath = Path

            Case "0885002.DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("0885002.DR")
                Console.WriteLine(Path)
                dbPath = Path

            Case "088502.5DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("088502.5DR")
                Console.WriteLine(Path)
                dbPath = Path

            Case "08853.15DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("08853.15DR")
                Console.WriteLine(Path)
                dbPath = Path

            Case "0885004.DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("0885004.DR")
                Console.WriteLine(Path)
                dbPath = Path

            Case "0885005.DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("0885005.DR")
                Console.WriteLine(Path)
                dbPath = Path
        End Select
    End Sub

    Public nominal As Decimal
    Public minimum As Decimal
    Public HiLimit As Decimal
    Public LoLimit As Decimal
    Public getLoLim As Decimal
    Public getLolim2 As Decimal
    Public multiplier As Decimal '= 1000
    Public add As Decimal
    Public reduce As Decimal
    Public splitval As Decimal
    Public truemax As Decimal
    Public truemin As Decimal
    Public Eq1 As Decimal


    Sub SetHighLowAndNominal()
        ''************* For Comparator Mode ******************
        'Dim HL As String = Math.Round(Form1.txtHiLimit.Text * 10, 0)
        'Dim LL As String = Math.Round(Form1.txtLoLimit.Text * 10, 0)

        '************* For % Compare Mode ******************
        Dim HL As String = CInt(Form1.txtPositive.Text) * 100
        Dim LL As String = CInt(Form1.txtNegative.Text) * 100
        Dim Nom As String = Form1.txtNominal.Text * 100

        Form1.SerialPort1.Write("L2," & Nom & vbCr) 'Set nominal value in % comparator where n=00000-22999
        Console.WriteLine(Nom)
        Form1.SerialPort1.Write("X" & vbCr)
        Form1.SerialPort1.Write("L3," & HL & vbCr) 'Set high comparator limit value where n=00000-22999
        Form1.SerialPort1.Write("X" & vbCr)
        Form1.SerialPort1.Write("L4," & LL & vbCr) 'Set low comparator limit value where n=00000-22998
        Form1.SerialPort1.Write("X" & vbCr)

        'Form1.SerialPort1.Write("L0," & vbCr) 'Set high comparator limit value where n=00000-22999
        'Form1.SerialPort1.Write("L1," & vbCr) 'Set Low comparator limit value where n=00000-22999
        'Form1.SerialPort1.Write("L2," & Form1.txtNominal.Text & vbCr) 'Set nominal value in % comparator where n=00000-22999
        'Form1.SerialPort1.Write("L3," & vbCr) 'Set high % limit in % comparator where nnnn=00.00–99.99
        'Form1.SerialPort1.Write("L4," & vbCr) 'Set low % limit in % comparator where nnnn=00.00–99.99

    End Sub

    Public NANO_dbRating As Decimal
    Public NANO_dbNominal As Decimal
    Public NANO_dbMaximum As Integer
    Public NANO_dbMinimum As Integer
    Public DbSetRange As String


    Sub EnterPartNumber()
        'If NANO_db.IsValidPartNumber(Form1.txtPartNo.Text) Then

        '    If Not Form1.SerialPort1.IsOpen Then
        '        Form1.SerialPort1.Open()
        '    End If

        '    Console.WriteLine("Rating: " & NANO_db.Rating)
        '    Console.WriteLine("Nominal: " & NANO_db.Nominal)
        '    Console.WriteLine("Maximum: " & NANO_db.Maximum)
        '    Console.WriteLine("Minimum: " & NANO_db.Minimum)


        '    truemax = Math.Round(CDec(NANO_db.Nominal) * ((CDec(NANO_db.Maximum) / 100) + 1), 3)

        '    'Eq1 = (CDec(PICO_db.Nominal) * (CDec(PICO_db.Minimum) / 100))
        '    'truemin = Math.Round(CDec(PICO_db.Nominal) - Eq1, 3)

        '    maximum = truemax
        '    'minimum = truemin
        '    range() 'mutiplier set up

        '    nominal = NANO_db.Nominal / multiplier
        '    HiLimit = nominal * ((CDec(NANO_db.Maximum) / 100) + 1) 'High Limit value
        '    maximum = Math.Round(HiLimit, 3)

        '    getLoLim = 100 - NANO_db.Minimum 'get the low limit less 100
        '    getLolim2 = getLoLim / 100 'convert to %
        '    LoLimit = nominal * getLolim2 ' Low Limit Value
        '    minimum = Math.Round(LoLimit, 3)

        '    Form1.txtRating.Text = NANO_db.Rating
        '    Form1.txtPositive.Text = NANO_db.Maximum
        '    Form1.txtNegative.Text = NANO_db.Minimum
        '    Form1.txtNominal.Text = nominal
        '    Form1.txtHiLimit.Text = Math.Round(HiLimit, 3)
        '    Form1.txtLoLimit.Text = Math.Round(LoLimit, 3)

        '    set_val() 'ser the progress bar parameter
        '    ProgBar() 'progress bar value and maximum value

        '    Form1.txtPartNo.ReadOnly = True

        '    SetHighLowAndNominal()
        'Else

        'End If

        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Part Numbers.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885partnum")


        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From 885Details_tb WHERE PartNumber = '" + Form1.txtPartNo.Text + "'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then

                NANO_dbRating = Data.Rows(0).Item("Rating").ToString
                NANO_dbNominal = Data.Rows(0).Item("Nominal").ToString
                NANO_dbMaximum = Data.Rows(0).Item("Maximum").ToString
                NANO_dbMinimum = Data.Rows(0).Item("Minimum").ToString
                DbSetRange = Data.Rows(0).Item("Range").ToString
                Console.WriteLine("Rating: " & NANO_dbRating)
                Console.WriteLine("Nominal: " & NANO_dbNominal)
                Console.WriteLine("Maximum: " & NANO_dbMaximum)
                Console.WriteLine("Minimum: " & NANO_dbMinimum)
                Console.WriteLine("Range: " & DbSetRange)

                OFFLINEdb()
                Form1.txtLotNumber.Focus()
            Else
                MsgBox("Part number does not exist in the database.", MessageBoxIcon.Error)
                Form1.txtPartNo.Text = ""
                Form1.txtPartNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()
        End Try
    End Sub

    Sub OFFLINEdb()
        If Not Form1.SerialPort1.IsOpen Then
            Form1.SerialPort1.Open()
        End If

        truemax = Math.Round(CDec(NANO_dbNominal) * ((CDec(NANO_dbMaximum) / 100) + 1), 3)

        'Eq1 = (CDec(PICO_db.Nominal) * (CDec(PICO_db.Minimum) / 100))
        'truemin = Math.Round(CDec(PICO_db.Nominal) - Eq1, 3)

        maximum = truemax
        'minimum = truemin
        range() 'mutiplier set up

        nominal = NANO_dbNominal / multiplier
        HiLimit = nominal * ((CDec(NANO_dbMaximum) / 100) + 1) 'High Limit value
        maximum = Math.Round(HiLimit, 3)

        getLoLim = 100 - NANO_dbMinimum 'get the low limit less 100
        getLolim2 = getLoLim / 100 'convert to %
        LoLimit = nominal * getLolim2 ' Low Limit Value
        minimum = Math.Round(LoLimit, 3)

        Form1.txtRating.Text = NANO_dbRating
        Form1.txtPositive.Text = NANO_dbMaximum
        Form1.txtNegative.Text = NANO_dbMinimum
        Form1.txtNominal.Text = nominal
        Form1.txtHiLimit.Text = Math.Round(HiLimit, 3)
        Form1.txtLoLimit.Text = Math.Round(LoLimit, 3)

        set_val() 'ser the progress bar parameter
        ProgBar() 'progress bar value and maximum value

        Form1.txtPartNo.ReadOnly = True

        SetHighLowAndNominal()
    End Sub

    Sub change_device()
        Form1.btnStart.Enabled = False

        If Form1.TimerQtyChecking.Enabled = True Then
            Form1.TimerQtyChecking.Enabled = False
        End If

        labels()
        labelResult()
        ProgBar()
        Form1.txtTegamValue.Text = ""
        Form1.txtResult.Text = ""
        Form1.txtResult.FillColor = Color.Black
        Form1.txtTegamValue.FillColor = Color.Black

        Form1.txtPartNo.Text = ""
        Form1.txtLotNumber.Text = ""
        Form1.txtQty.Text = ""
        Form1.txtRating.Text = ""
        Form1.txtPositive.Text = ""
        Form1.txtNegative.Text = ""
        Form1.txtNominal.Text = ""
        Form1.txtHiLimit.Text = ""
        Form1.txtLoLimit.Text = ""
        Form1.txtTested.Text = "0"
        Form1.txtGood.Text = "0"
        Form1.txtHigh.Text = "0"
        Form1.txtOpen.Text = "0"
        Form1.txtLow.Text = "0"
        Form1.txtRej.Text = "0"
        Form1.txtStdev.Text = "0"
        Form1.txtCpk.Text = "0"
        Form1.txtYield.Text = ""

        Form1.txtPartNo.ReadOnly = False
        Form1.txtLotNumber.ReadOnly = False
        Form1.txtQty.ReadOnly = False

        runcount = 0
        Array.Clear(store_res, 0, store_res.Length)

        Form1.Timer1.Enabled = False
        Form1.btnClear2.Enabled = True
        Form1.btnChange.Enabled = False
        Form1.btnLot.Enabled = False
        Form1.btnStart.Text = "Start"
        Form1.btnStart.FillColor = Color.LightGreen
        Form1.btnStart.FillColor2 = Color.Green
        Form1.btnStart.ForeColor = Color.Black

        If Form1.SerialPort1.IsOpen Then
            Form1.SerialPort1.Close()
        End If

    End Sub

    Sub clear()

        If Form1.TimerQtyChecking.Enabled = True Then
            Form1.TimerQtyChecking.Enabled = False
        End If

        ProgBar()
        labelResult()
        Form1.txtTegamValue.Text = ""
        Form1.txtResult.Text = ""
        Form1.txtResult.FillColor = Color.Black
        Form1.txtTegamValue.FillColor = Color.Black

        Form1.txtLotNumber.Text = ""
        Form1.txtQty.Text = ""

        Form1.txtTested.Text = "0"
        Form1.txtGood.Text = "0"
        Form1.txtHigh.Text = "0"
        Form1.txtOpen.Text = "0"
        Form1.txtLow.Text = "0"
        Form1.txtRej.Text = "0"
        Form1.txtStdev.Text = "0"
        Form1.txtCpk.Text = "0"
        Form1.txtYield.Text = ""

        Form1.txtLotNumber.ReadOnly = False
        Form1.txtQty.ReadOnly = False

        runcount = 0
        Array.Clear(store_res, 0, store_res.Length)

        Form1.Timer1.Enabled = False
        Form1.btnStart.Enabled = False
        Form1.btnClear2.Enabled = True
        Form1.btnChange.Enabled = False
        Form1.btnLot.Enabled = False
        Form1.btnStart.Text = "Start"
        Form1.btnStart.FillColor = Color.LightGreen
        Form1.btnStart.FillColor2 = Color.Green
        Form1.btnStart.ForeColor = Color.Black
    End Sub

    Sub btnStartFunction()

        If Form1.btnStart.Text = "Start" Then
            If Not Form1.SerialPort2.IsOpen Then
                Form1.SerialPort2.Open()
            End If
            Form1.Timer1.Enabled = True
            Form1.TimerQtyChecking.Enabled = True
            Form1.btnStart.Text = "Stop"
            Form1.btnStart.FillColor = Color.Maroon
            Form1.btnStart.FillColor2 = Color.Red
            Form1.btnStart.ForeColor = Color.White

            Form1.btnChange.Enabled = False
            Form1.btnLot.Enabled = False
        Else
            Form1.Timer1.Enabled = False
            Form1.btnStart.Text = "Start"
            Form1.btnStart.FillColor = Color.LightGreen
            Form1.btnStart.FillColor2 = Color.Green
            Form1.btnStart.ForeColor = Color.Black

            Form1.btnChange.Enabled = True
            Form1.btnLot.Enabled = True
        End If
    End Sub

    Sub ChangeToStart()
        Form1.Timer1.Enabled = False
        Form1.btnStart.Text = "Start"
        Form1.btnStart.FillColor = Color.LightGreen
        Form1.btnStart.FillColor2 = Color.Green
        Form1.btnStart.ForeColor = Color.Black

        Form1.btnChange.Enabled = True
        Form1.btnLot.Enabled = True
    End Sub

    Public holding As Boolean
    Sub Timer1Function()
        If Form1.SerialPort2.CDHolding = True Then
            If holding = False Then
                holding = True
                Try

                    setTegam()
                    SetHighLowAndNominal()
                    Thread.Sleep(100)
                    Form1.SerialPort1.WriteLine("E" & vbCr)

                    'test_result() 'GOOD/HIGH/LOW or OPEN 
                    'Prog_count() 'histogram counter
                    'statistics() 'comp of CPK and Standard Dev
                Catch ex As Exception
                    MsgBox(ex.Message, vbCritical)
                End Try
            End If
        Else
            holding = False
        End If
    End Sub

    Public average As Decimal
    Public stdev As Decimal
    Public cpk_max As Decimal
    Public cpk_min As Decimal
    Sub statistics()
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
                                Form1.txtCpk.Text = CStr(Math.Round(Math.Min(cpk_max, cpk_min), 2))
                                Form1.txtStdev.Text = CStr(Math.Round(stdev, 2))
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

End Module


Module Form1MenuStrip_Module
    Sub ViewLotDetails()
        With LotDetails_Form
            .TopLevel = False
            Form1.PanelMain.Controls.Add(LotDetails_Form)
            .WindowState = FormWindowState.Maximized
            .BringToFront()
            .Show()
        End With
    End Sub

    Sub ChangeSerialCom()
        ChangeSerialPass_Form.ShowDialog()
    End Sub

    Sub ChangeTableReferenceOfDb()
        Password_Form.ShowDialog()
        'ChangeTableRef_Form.ShowDialog()
    End Sub

    Sub ChangePathMenuStrip()
        ChangePathPass_Form.ShowDialog()
    End Sub
End Module

Module LogLotnumber_Module

    Sub LogStartTime()
        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log"
        Dim connection As New OleDbConnection(connString)

        Dim mycommand As String
        Dim _dateNtime As String = Date.Now.ToString("yy-MM-dd HH:mm:ss:fff")
        Dim S_Ref As String = Date.Now.ToString("yyyy-MM-dd")
        Dim Lotnum As String = Form1.txtLotNumber.Text

        If Form1.txtLotNumber.Text = "" Then
            MsgBox("Please enter Lot number!", MessageBoxIcon.Error)
        Else

            Try
                connection.Open()
                mycommand = "INSERT INTO [LotHistory_tb] ([Lot_Num],[Start_time],[Start_Reference]) VALUES (@Lot, @_dateNtime, @Sref)"
                Using command As New OleDbCommand(mycommand, connection)
                    command.Parameters.AddWithValue("@Lot", Lotnum)
                    command.Parameters.AddWithValue("@_dateNtime", _dateNtime)
                    command.Parameters.AddWithValue("@Sref", S_Ref)
                    command.ExecuteNonQuery()
                End Using
                connection.Close()
            Catch ex As Exception
                MsgBox(ex.Message, vbCritical)
            End Try
        End If
    End Sub

    Sub LogEndTime()
        Dim Lot As String
        Dim E_Ref As String = Date.Now.ToString("yyyy-MM-dd")
        Dim connectionString As String = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")
        'Dim query As String = "UPDATE LotHistory_tb SET End_time = @NewValue, dbTableName = @dbname WHERE Lot_Num = @Lot"
        'Dim query As String = "UPDATE LotHistory_tb SET End_time = @NewValue WHERE Lot_Num = @Lot"
        Dim query As String = "UPDATE LotHistory_tb SET End_time = @NewValue, End_Reference = @ERef WHERE Lot_Num = @Lot"
        Dim Lotendtime As String = Date.Now.ToString("yy-MM-dd HH:mm:ss:fff")
        Lot = Form1.txtLotNumber.Text

        Using connection As New OleDbConnection(connectionString)
            Using command As New OleDbCommand(query, connection)
                command.Parameters.AddWithValue("@NewValue", Lotendtime)
                command.Parameters.AddWithValue("@ERef", E_Ref)
                command.Parameters.AddWithValue("@Lot", Lot)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Sub LogLotforCCD1()
        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log"
        Dim connection As New OleDbConnection(connString)

        Dim mycommand As String
        Dim Lotnum As String = Form1.txtLotNumber.Text

        If Form1.txtLotNumber.Text = "" Then
            MsgBox("Please enter Lot number!", MessageBoxIcon.Error)
        Else

            Try
                connection.Open()
                mycommand = "INSERT INTO [Camera1_tb] ([Lot_Num]) VALUES (@Lot)"
                Using command As New OleDbCommand(mycommand, connection)
                    command.Parameters.AddWithValue("@Lot", Lotnum)
                    command.ExecuteNonQuery()
                End Using
                connection.Close()
            Catch ex As Exception
                MsgBox(ex.Message, vbCritical)
            End Try
        End If
    End Sub

    Sub LogLotforCCD2()
        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log"
        Dim connection As New OleDbConnection(connString)

        Dim mycommand As String
        Dim Lotnum As String = Form1.txtLotNumber.Text

        If Form1.txtLotNumber.Text = "" Then
            MsgBox("Please enter Lot number!", MessageBoxIcon.Error)
        Else

            Try
                connection.Open()
                mycommand = "INSERT INTO [Camera2_tb] ([Lot_Num]) VALUES (@Lot)"
                Using command As New OleDbCommand(mycommand, connection)
                    command.Parameters.AddWithValue("@Lot", Lotnum)
                    command.ExecuteNonQuery()
                End Using
                connection.Close()
            Catch ex As Exception
                MsgBox(ex.Message, vbCritical)
            End Try
        End If
    End Sub

    Sub LogLotforCCD3()
        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log"
        Dim connection As New OleDbConnection(connString)

        Dim mycommand As String
        Dim Lotnum As String = Form1.txtLotNumber.Text

        If Form1.txtLotNumber.Text = "" Then
            MsgBox("Please enter Lot number!", MessageBoxIcon.Error)
        Else

            Try
                connection.Open()
                mycommand = "INSERT INTO [Camera3_tb] ([Lot_Num]) VALUES (@Lot)"
                Using command As New OleDbCommand(mycommand, connection)
                    command.Parameters.AddWithValue("@Lot", Lotnum)
                    command.ExecuteNonQuery()
                End Using
                connection.Close()
            Catch ex As Exception
                MsgBox(ex.Message, vbCritical)
            End Try
        End If
    End Sub

    Sub LogLotforCCD4()
        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log"
        Dim connection As New OleDbConnection(connString)

        Dim mycommand As String
        Dim Lotnum As String = Form1.txtLotNumber.Text

        If Form1.txtLotNumber.Text = "" Then
            MsgBox("Please enter Lot number!", MessageBoxIcon.Error)
        Else

            Try
                connection.Open()
                mycommand = "INSERT INTO [Camera4_tb] ([Lot_Num]) VALUES (@Lot)"
                Using command As New OleDbCommand(mycommand, connection)
                    command.Parameters.AddWithValue("@Lot", Lotnum)
                    command.ExecuteNonQuery()
                End Using
                connection.Close()
            Catch ex As Exception
                MsgBox(ex.Message, vbCritical)
            End Try
        End If
    End Sub

    Sub LogLotforCCD5()
        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log"
        Dim connection As New OleDbConnection(connString)

        Dim mycommand As String
        Dim Lotnum As String = Form1.txtLotNumber.Text

        If Form1.txtLotNumber.Text = "" Then
            MsgBox("Please enter Lot number!", MessageBoxIcon.Error)
        Else

            Try
                connection.Open()
                mycommand = "INSERT INTO [Camera5_tb] ([Lot_Num]) VALUES (@Lot)"
                Using command As New OleDbCommand(mycommand, connection)
                    command.Parameters.AddWithValue("@Lot", Lotnum)
                    command.ExecuteNonQuery()
                End Using
                connection.Close()
            Catch ex As Exception
                MsgBox(ex.Message, vbCritical)
            End Try
        End If
    End Sub

    Sub LogLotforCCD6()
        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log"
        Dim connection As New OleDbConnection(connString)

        Dim mycommand As String
        Dim Lotnum As String = Form1.txtLotNumber.Text

        If Form1.txtLotNumber.Text = "" Then
            MsgBox("Please enter Lot number!", MessageBoxIcon.Error)
        Else

            Try
                connection.Open()
                mycommand = "INSERT INTO [Camera6_tb] ([Lot_Num]) VALUES (@Lot)"
                Using command As New OleDbCommand(mycommand, connection)
                    command.Parameters.AddWithValue("@Lot", Lotnum)
                    command.ExecuteNonQuery()
                End Using
                connection.Close()
            Catch ex As Exception
                MsgBox(ex.Message, vbCritical)
            End Try
        End If
    End Sub
End Module

Module CreateCSV_Module
    Public StartOled As String
    Public EndOled As String
    Public dbNameoftb As String
    Public StartProRef As String
    Public EndProRef As String
    Sub GetLotNumberHistory()
        Dim Start As String = "Sample"

        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")


        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From LotHistory_tb WHERE Lot_Num = '" + Form1.txtLotNumber.Text + "'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then

                StartOled = Data.Rows(0).Item("Start_time").ToString
                EndOled = Data.Rows(0).Item("End_time").ToString
                dbNameoftb = Data.Rows(0).Item("dbTableName").ToString
                StartProRef = Data.Rows(0).Item("Start_Reference").ToString
                EndProRef = Data.Rows(0).Item("End_Reference").ToString
                Console.WriteLine(StartOled)
                Console.WriteLine(EndOled)
                Console.WriteLine(dbNameoftb)
            Else
                MsgBox("Lot number does not exist in the database.", MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()
        End Try
    End Sub

    Sub ExportToCSV()

        'GetPathforDB()
        GetLotNumberHistory()

        If StartProRef = EndProRef Then

            '"SELECT * FROM ProduceLog19748 WHERE Time BETWEEN '" & First & "'  AND '" & Last & "'"
            '"SELECT * FROM ProduceLog19783 WHERE Time > '" & StartOled & "' AND Time < '" & EndOled & "'"

            Dim connString As String = "Data Source=" & dbPath & ";Version=3"

            Dim query As String = "SELECT * FROM '" & dbNameoftb & "' WHERE Time > '" & StartOled & "' AND Time < '" & EndOled & "'"
            Dim dateNtime As String = Date.Now.ToString("MM-dd-yy hh_mmtt")
            Dim Year As String = Date.Now.ToString("yyyy")
            Dim Month As String = Date.Now.ToString("MMMM")
            Dim FolderPath As String = "C:\Nano 885 data\" & Year & "\" & Month & "\" & Form1.txtLotNumber.Text & " " & dateNtime & ".csv"

            Try

                ' Ensure the directory exists or create it
                Dim directoryPath As String = Path.GetDirectoryName(FolderPath)
                If Not Directory.Exists(directoryPath) Then
                    Directory.CreateDirectory(directoryPath)
                End If

                Using connection As New SQLite.SQLiteConnection(connString)
                    connection.Open()

                    Using command As New SQLiteCommand(query, connection)
                        Using reader As SQLiteDataReader = command.ExecuteReader()
                            Using writer As New StreamWriter(FolderPath)
                                ' Write the column headers
                                For i As Integer = 0 To reader.FieldCount - 1
                                    writer.Write($"{reader.GetName(i)},")
                                Next
                                writer.WriteLine()

                                ' Write the data
                                While reader.Read()
                                    For i As Integer = 0 To reader.FieldCount - 1
                                        writer.Write($"{reader(i)},")
                                    Next
                                    writer.WriteLine()
                                End While

                                ' CCD1
                                writer.Write(vbCrLf & "CCD 1 Total, OKAY, Plug Defect, Wire Defect")
                                writer.WriteLine()

                                ' Add values for CCD 1
                                writer.Write(CCD1Total & "," & CCD1OK & "," & CCD1PlugDefect & "," & CCD1WireDefect)
                                writer.WriteLine()

                                ' CCD2
                                writer.Write(vbCrLf & "CCD 2 Total, OKAY, Plug Defect, Wire Defect")
                                writer.WriteLine()

                                ' Add values for CCD2
                                writer.Write(CCD2Total & "," & CCD2OK & "," & CCD2PlugDefect & "," & CCD2WireDefect)
                                writer.WriteLine()

                                ' CCD3
                                writer.Write(vbCrLf & "CCD 3 Total, OKAY, Plug Defect, Wire Defect")
                                writer.WriteLine()

                                ' Add values for CCD3
                                writer.Write(CCD3Total & "," & CCD3OK & "," & CCD3PlugDefect & "," & CCD3WireDefect)
                                writer.WriteLine()

                                ' CCD4
                                writer.Write(vbCrLf & "CCD 4 Total, OKAY, Plug Defect, Wire Defect")
                                writer.WriteLine()

                                ' Add values for CCD4
                                writer.Write(CCD4Total & "," & CCD4OK & "," & CCD4PlugDefect & "," & CCD4WireDefect)
                                writer.WriteLine()

                                ' CCD5
                                writer.Write(vbCrLf & "CCD 5 Total, OKAY, Plug Defect, Sheet Defect")
                                writer.WriteLine()

                                ' Add values for CCD5
                                writer.Write(CCD5Total & "," & CCD5OK & "," & CCD5PlugDefect & "," & CCD5WireDefect)
                                writer.WriteLine()

                                ' CCD6
                                writer.Write(vbCrLf & "CCD 6 Total, OKAY, Plug Defect, Sheet Defect")
                                writer.WriteLine()

                                ' Add values for CCD6
                                writer.Write(CCD6Total & "," & CCD6OK & "," & CCD6PlugDefect & "," & CCD6WireDefect)
                                writer.WriteLine()
                            End Using
                        End Using
                    End Using
                End Using

                'MessageBox.Show("The data is saved in " & FolderPath)

            Catch ex As Exception
                MessageBox.Show("Error exporting data: " & ex.Message)
            Finally

            End Try

        Else
            FirstSheet()
            Thread.Sleep(500)
            SecondSheet()
        End If
    End Sub

    Sub FirstSheet()
        Dim PreviousTable = "ProduceLog" & PrevtbRef
        Dim connString As String = "Data Source=" & dbPath & ";Version=3"

        Dim query As String = "SELECT * FROM '" & PreviousTable & "' WHERE Time > '" & StartOled & "' AND Time < '" & EndOled & "'"
        Dim dateNtime As String = Date.Now.ToString("MM-dd-yy hh_mmtt")
        Dim Year As String = Date.Now.ToString("yyyy")
        Dim Month As String = Date.Now.ToString("MMMM")
        Dim FolderPath As String = "C:\Nano 885 data\" & Year & "\" & Month & "\First Sheet " & Form1.txtLotNumber.Text & " " & dateNtime & ".csv"

        Try

            ' Ensure the directory exists or create it
            Dim directoryPath As String = Path.GetDirectoryName(FolderPath)
            If Not Directory.Exists(directoryPath) Then
                Directory.CreateDirectory(directoryPath)
            End If

            Using connection As New SQLite.SQLiteConnection(connString)
                connection.Open()

                Using command As New SQLiteCommand(query, connection)
                    Using reader As SQLiteDataReader = command.ExecuteReader()
                        Using writer As New StreamWriter(FolderPath)
                            ' Write the column headers
                            For i As Integer = 0 To reader.FieldCount - 1
                                writer.Write($"{reader.GetName(i)},")
                            Next
                            writer.WriteLine()

                            ' Write the data
                            While reader.Read()
                                For i As Integer = 0 To reader.FieldCount - 1
                                    writer.Write($"{reader(i)},")
                                Next
                                writer.WriteLine()
                            End While
                        End Using
                    End Using
                End Using
            End Using

            'MessageBox.Show("The data is saved in " & FolderPath)

        Catch ex As Exception
            MessageBox.Show("Error exporting data: " & ex.Message)
        Finally

        End Try
    End Sub

    Sub SecondSheet()
        Dim connString As String = "Data Source=" & dbPath & ";Version=3"

        Dim query As String = "SELECT * FROM '" & dbNameoftb & "' WHERE Time > '" & StartOled & "' AND Time < '" & EndOled & "'"
        Dim dateNtime As String = Date.Now.ToString("MM-dd-yy hh_mmtt")
        Dim Year As String = Date.Now.ToString("yyyy")
        Dim Month As String = Date.Now.ToString("MMMM")
        Dim FolderPath As String = "C:\Nano 885 data\" & Year & "\" & Month & "\Second Sheet " & Form1.txtLotNumber.Text & " " & dateNtime & ".csv"

        Try

            ' Ensure the directory exists or create it
            Dim directoryPath As String = Path.GetDirectoryName(FolderPath)
            If Not Directory.Exists(directoryPath) Then
                Directory.CreateDirectory(directoryPath)
            End If

            Using connection As New SQLite.SQLiteConnection(connString)
                connection.Open()

                Using command As New SQLiteCommand(query, connection)
                    Using reader As SQLiteDataReader = command.ExecuteReader()
                        Using writer As New StreamWriter(FolderPath)
                            ' Write the column headers
                            For i As Integer = 0 To reader.FieldCount - 1
                                writer.Write($"{reader.GetName(i)},")
                            Next
                            writer.WriteLine()

                            ' Write the data
                            While reader.Read()
                                For i As Integer = 0 To reader.FieldCount - 1
                                    writer.Write($"{reader(i)},")
                                Next
                                writer.WriteLine()
                            End While

                            ' CCD1
                            writer.Write(vbCrLf & "CCD 1 Total, OKAY, Plug Defect, Wire Defect")
                            writer.WriteLine()

                            ' Add values for CCD 1
                            writer.Write(CCD1Total & "," & CCD1OK & "," & CCD1PlugDefect & "," & CCD1WireDefect)
                            writer.WriteLine()

                            ' CCD2
                            writer.Write(vbCrLf & "CCD 2 Total, OKAY, Plug Defect, Wire Defect")
                            writer.WriteLine()

                            ' Add values for CCD2
                            writer.Write(CCD2Total & "," & CCD2OK & "," & CCD2PlugDefect & "," & CCD2WireDefect)
                            writer.WriteLine()

                            ' CCD3
                            writer.Write(vbCrLf & "CCD 3 Total, OKAY, Plug Defect, Wire Defect")
                            writer.WriteLine()

                            ' Add values for CCD3
                            writer.Write(CCD3Total & "," & CCD3OK & "," & CCD3PlugDefect & "," & CCD3WireDefect)
                            writer.WriteLine()

                            ' CCD4
                            writer.Write(vbCrLf & "CCD 4 Total, OKAY, Plug Defect, Wire Defect")
                            writer.WriteLine()

                            ' Add values for CCD4
                            writer.Write(CCD4Total & "," & CCD4OK & "," & CCD4PlugDefect & "," & CCD4WireDefect)
                            writer.WriteLine()

                            ' CCD5
                            writer.Write(vbCrLf & "CCD 5 Total, OKAY, Plug Defect, Sheet Defect")
                            writer.WriteLine()

                            ' Add values for CCD5
                            writer.Write(CCD5Total & "," & CCD5OK & "," & CCD5PlugDefect & "," & CCD5WireDefect)
                            writer.WriteLine()

                            ' CCD6
                            writer.Write(vbCrLf & "CCD 6 Total, OKAY, Plug Defect, Sheet Defect")
                            writer.WriteLine()

                            ' Add values for CCD6
                            writer.Write(CCD6Total & "," & CCD6OK & "," & CCD6PlugDefect & "," & CCD6WireDefect)
                            writer.WriteLine()
                        End Using
                    End Using
                End Using
            End Using

            'MessageBox.Show("The data is saved in " & FolderPath)

        Catch ex As Exception
            MessageBox.Show("Error exporting data: " & ex.Message)
        Finally

        End Try
    End Sub
End Module

Module Appconfig_Module

    Public PrevtbRef As Integer
    Public PrevDateRef As String
    Public NewDateRef As String
    Public NewtbRef As Integer
    Sub GetPrevTable()
        Dim Prevtb As String = System.Configuration.ConfigurationManager.AppSettings("dbTableReference")
        Console.WriteLine(Prevtb)

        PrevtbRef = Prevtb
    End Sub

    Sub GetPrevDate()
        Dim Prevdate As String = System.Configuration.ConfigurationManager.AppSettings("dateReference")
        Console.WriteLine(Prevdate)

        PrevDateRef = Prevdate
    End Sub

    Public config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
    Sub UpdatedbTableReference()
        config.AppSettings.Settings("dbTableReference").Value = NewtbRef ' Rewrite
        config.Save(ConfigurationSaveMode.Modified) ' save the new value

        ConfigurationManager.RefreshSection("appSettings") 'refresh
    End Sub

    Sub UpdateDateReference()
        config.AppSettings.Settings("dateReference").Value = NewDateRef ' Rewrite 
        config.Save(ConfigurationSaveMode.Modified) ' save the new value

        ConfigurationManager.RefreshSection("appSettings") 'refresh
    End Sub


End Module

Module CameraResults_Module

    Sub CheckDatabase()

        GetPathforDB()
        If Not System.IO.File.Exists(dbPath) Then
            MessageBox.Show("Database file not found!.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            change_device()
            Exit Sub ' or handle the error in any other way as needed
        End If
        EnterPartNumber()
        'Form1.txtLotNumber.Focus()
    End Sub

    Public Supplier_tbName As String
    Sub CheckTable()
        Dim DateUpdate As String = Date.Now.ToString("yyyy-MM-dd")

        Dim connString As String = "Data Source=" & dbPath & ";Version=3"
        'Dim TableName As String = Date.Now.ToString("yyyy-MM-dd")

        GetPrevTable()
        GetPrevDate()

        If PrevDateRef = Date.Now.ToString("yyyy-MM-dd") Then
            Supplier_tbName = "ProduceLog" & PrevtbRef

            Dim tableExistsQuery As String = $"SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='" & Supplier_tbName & "';"

            Using connectionCheck As New SQLite.SQLiteConnection(connString)
                Using commandCheck As New SQLiteCommand(tableExistsQuery, connectionCheck)
                    connectionCheck.Open()

                    Dim tableCount As Integer = CInt(commandCheck.ExecuteScalar())

                    If tableCount = 0 Then
                        MessageBox.Show("The table " & Supplier_tbName & " does not exist in the database.", "Prev Db Table Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Form1.btnEMsave.Visible = True
                        Form1.btnStart.Enabled = False
                        Form1.btnChange.Enabled = False
                        Form1.btnLot.Enabled = False
                        Exit Sub ' or handle the error in any other way as needed
                    End If
                End Using
            End Using

            'Save Tablename in MS access
            Dim Lot As String
            Dim dbTableName As String = Supplier_tbName
            Dim connectionString As String = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")
            Dim query As String = "UPDATE LotHistory_tb SET dbTableName = @dbname WHERE Lot_Num = @Lot"
            Lot = Form1.txtLotNumber.Text

            Using connection As New OleDbConnection(connectionString)
                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@dbname", dbTableName)
                    command.Parameters.AddWithValue("@Lot", Lot)
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using

            Thread.Sleep(100)

            GetCCD1Result()
            LogResultofCCD1()

            GetCCD2Result()
            LogResultofCCD2()

            GetCCD3Result()
            LogResultofCCD3()

            GetCCD4Result()
            LogResultofCCD4()

            GetCCD5Result()
            LogResultofCCD5()

            GetCCD6Result()
            LogResultofCCD6()
            Thread.Sleep(100)

            ExportToCSV() 'Export to CSV File

            NewDateRef = DateUpdate
            UpdateDateReference()

            MsgBox(Form1.txtLotNumber.Text & " Done!")
        Else


            Dim LastDate As Date = PrevDateRef
            Dim now As Date = Date.Now
            Dim diff As TimeSpan = now - LastDate
            Dim days As Integer = diff.Days + 0
            Dim TableName As Integer = PrevtbRef + days
            Console.WriteLine("PrduceLog" & TableName)

            Supplier_tbName = "ProduceLog" & TableName

            Dim tableExistsQuery As String = $"SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='" & Supplier_tbName & "';"

            Using connectionCheck As New SQLite.SQLiteConnection(connString)
                Using commandCheck As New SQLiteCommand(tableExistsQuery, connectionCheck)
                    connectionCheck.Open()

                    Dim tableCount As Integer = CInt(commandCheck.ExecuteScalar())

                    If tableCount = 0 Then
                        MessageBox.Show("The table " & Supplier_tbName & " does not exist in the database.", "New Db Table Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Form1.btnEMsave.Visible = True
                        Form1.btnStart.Enabled = False
                        Form1.btnChange.Enabled = False
                        Form1.btnLot.Enabled = False
                        Exit Sub ' or handle the error in any other way as needed
                    End If
                End Using
            End Using

            'Save Tablename in MS access
            Dim Lot As String
            Dim dbTableName As String = Supplier_tbName
            Dim connectionString As String = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")
            Dim query As String = "UPDATE LotHistory_tb SET dbTableName = @dbname WHERE Lot_Num = @Lot"
            Lot = Form1.txtLotNumber.Text

            Using connection As New OleDbConnection(connectionString)
                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@dbname", dbTableName)
                    command.Parameters.AddWithValue("@Lot", Lot)
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using

            Thread.Sleep(100)

            GetCCD1Result()
            LogResultofCCD1()

            GetCCD2Result()
            LogResultofCCD2()

            GetCCD3Result()
            LogResultofCCD3()

            GetCCD4Result()
            LogResultofCCD4()

            GetCCD5Result()
            LogResultofCCD5()

            GetCCD6Result()
            LogResultofCCD6()
            Thread.Sleep(100)

            ExportToCSV() 'Export to CSV File

            NewtbRef = TableName
            UpdatedbTableReference()

            NewDateRef = DateUpdate
            UpdateDateReference()

            MsgBox(Form1.txtLotNumber.Text & " Done!")

        End If
    End Sub

    Sub EmergencyExport()
        Dim DateUpdate As String = Date.Now.ToString("yyyy-MM-dd")
        Dim connString As String = "Data Source=" & dbPath & ";Version=3"
        'Dim TableName As String = Date.Now.ToString("yyyy-MM-dd")

        GetPrevTable()
        GetPrevDate()

        If PrevDateRef = Date.Now.ToString("yyyy-MM-dd") Then
            Supplier_tbName = "ProduceLog" & PrevtbRef

            Dim tableExistsQuery As String = $"SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='" & Supplier_tbName & "';"

            Using connectionCheck As New SQLite.SQLiteConnection(connString)
                Using commandCheck As New SQLiteCommand(tableExistsQuery, connectionCheck)
                    connectionCheck.Open()

                    Dim tableCount As Integer = CInt(commandCheck.ExecuteScalar())

                    If tableCount = 0 Then
                        MessageBox.Show("The table " & Supplier_tbName & " does not exist in the database.", "Table Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub ' or handle the error in any other way as needed
                    End If
                End Using
            End Using

            'Save Tablename in MS access
            Dim Lot As String
            Dim dbTableName As String = Supplier_tbName
            Dim connectionString As String = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")
            Dim query As String = "UPDATE LotHistory_tb SET dbTableName = @dbname WHERE Lot_Num = @Lot"
            Lot = Form1.txtLotNumber.Text

            Using connection As New OleDbConnection(connectionString)
                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@dbname", dbTableName)
                    command.Parameters.AddWithValue("@Lot", Lot)
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using

            GetCCD1Result()
            LogResultofCCD1()

            GetCCD2Result()
            LogResultofCCD2()

            GetCCD3Result()
            LogResultofCCD3()

            GetCCD4Result()
            LogResultofCCD4()

            GetCCD5Result()
            LogResultofCCD5()

            GetCCD6Result()
            LogResultofCCD6()
            Thread.Sleep(100)

            ExportToCSV() 'Export to CSV File

            NewDateRef = DateUpdate
            UpdateDateReference()

            Form1.btnEMsave.Visible = False
            Form1.btnStart.Enabled = True
            Form1.btnChange.Enabled = True
            Form1.btnLot.Enabled = True

            MsgBox(Form1.txtLotNumber.Text & " Done!")

        Else
            'Dim DateUpdate As String = Date.Now.ToString("yyyy-MM-dd")

            Dim LastDate As Date = PrevDateRef
            Dim now As Date = Date.Now
            Dim diff As TimeSpan = now - LastDate
            Dim days As Integer = diff.Days + 0
            Dim TableName As Integer = PrevtbRef + days
            Console.WriteLine("PrduceLog" & TableName)

            Supplier_tbName = "ProduceLog" & TableName

            Dim tableExistsQuery As String = $"SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='" & Supplier_tbName & "';"

            Using connectionCheck As New SQLite.SQLiteConnection(connString)
                Using commandCheck As New SQLiteCommand(tableExistsQuery, connectionCheck)
                    connectionCheck.Open()

                    Dim tableCount As Integer = CInt(commandCheck.ExecuteScalar())

                    If tableCount = 0 Then
                        MessageBox.Show("The table " & Supplier_tbName & " does not exist in the database.", "Table Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub ' or handle the error in any other way as needed
                    End If
                End Using
            End Using

            'Save Tablename in MS access
            Dim Lot As String
            Dim dbTableName As String = Supplier_tbName
            Dim connectionString As String = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")
            Dim query As String = "UPDATE LotHistory_tb SET dbTableName = @dbname WHERE Lot_Num = @Lot"
            Lot = Form1.txtLotNumber.Text

            Using connection As New OleDbConnection(connectionString)
                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@dbname", dbTableName)
                    command.Parameters.AddWithValue("@Lot", Lot)
                    connection.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using

            GetCCD1Result()
            LogResultofCCD1()

            GetCCD2Result()
            LogResultofCCD2()

            GetCCD3Result()
            LogResultofCCD3()

            GetCCD4Result()
            LogResultofCCD4()

            GetCCD5Result()
            LogResultofCCD5()

            GetCCD6Result()
            LogResultofCCD6()
            Thread.Sleep(100)

            ExportToCSV() 'Export to CSV File

            NewtbRef = TableName
            UpdatedbTableReference()

            NewDateRef = DateUpdate
            UpdateDateReference()

            Form1.btnEMsave.Visible = False
            Form1.btnStart.Enabled = True
            Form1.btnChange.Enabled = True
            Form1.btnLot.Enabled = True

            MsgBox(Form1.txtLotNumber.Text & " Done!")

        End If
    End Sub

    Public CCD1Total As String
    Public CCD1OK As String
    Public CCD1PlugDefect As String
    Public CCD1WireDefect As String

    Sub GetCCD1Result()

        'Dim connString As String = "Data Source=C:\LF Database\ProduceLog.db;Version=3"

        'Dim query As String = "SELECT t1.* FROM ProduceLog19783 t1 WHERE t1.CCD = 1 AND NOT EXISTS 
        '                        (SELECT 1 FROM ProduceLog19783 t2 WHERE t2.CCD = t1.CCD AND t2.Time > t1.Time);"
        Try
            Dim connString As String = "Data Source=" & dbPath & ";Version=3"

            Dim TableName As String = Date.Now.ToString("yyyy-MM-dd")

            Dim query As String = "SELECT t1.* FROM '" & Supplier_tbName & "' t1 WHERE t1.CCD = 1 AND NOT EXISTS 
            (SELECT 1 FROM '" & Supplier_tbName & "' t2 WHERE t2.CCD = t1.CCD AND t2.Time > t1.Time);"

            Using connection As New SQLite.SQLiteConnection(connString)
                Using command As New SQLiteCommand(query, connection)
                    connection.Open()

                    Dim reader As SQLiteDataReader = command.ExecuteReader()
                    If reader.Read() Then

                        CCD1Total = reader("NUM").ToString
                        CCD1OK = reader("OK").ToString
                        CCD1PlugDefect = reader("PlugDefect").ToString
                        CCD1WireDefect = reader("WireDefect").ToString

                        'MessageBox.Show("Total: " & CCD1Total & ControlChars.NewLine & "Good: " & CCD1OK & ControlChars.NewLine & "Plug Defect: " & CCD1PlugDefect & ControlChars.NewLine & "Wire Defect: " & CCD1WireDefect & "  ", "885 Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("No data found.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

    Sub LogResultofCCD1()
        Dim Lot As String
        Dim connectionString As String = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")
        Dim query As String = "UPDATE Camera1_tb SET T_Tested = @totalVal, Okay = @OkayVal, PlugDefect = @PlugDefVal, WireDefect = @WireDeftVal WHERE Lot_Num = @Lot"
        Lot = Form1.txtLotNumber.Text
        Using connection As New OleDbConnection(connectionString)
            Using command As New OleDbCommand(query, connection)
                command.Parameters.AddWithValue("@totalVal", CCD1Total)
                command.Parameters.AddWithValue("@OkayVal", CCD1OK)
                command.Parameters.AddWithValue("@PlugDefVal", CCD1PlugDefect)
                command.Parameters.AddWithValue("@WireDeftVal", CCD1WireDefect)
                command.Parameters.AddWithValue("@Lot", Lot)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public CCD2Total As String
    Public CCD2OK As String
    Public CCD2PlugDefect As String
    Public CCD2WireDefect As String

    Sub GetCCD2Result()
        'Dim connString As String = "Data Source=C:\LF Database\ProduceLog.db;Version=3"

        'Dim query As String = "SELECT t1.* FROM ProduceLog19783 t1 WHERE t1.CCD = 2 AND NOT EXISTS 
        '                        (SELECT 1 FROM ProduceLog19783 t2 WHERE t2.CCD = t1.CCD AND t2.Time > t1.Time);"

        Dim connString As String = "Data Source=" & dbPath & ";Version=3"

        Dim TableName As String = Date.Now.ToString("yyyy-MM-dd")

        Dim query As String = "SELECT t1.* FROM '" & Supplier_tbName & "' t1 WHERE t1.CCD = 2 AND NOT EXISTS 
                                (SELECT 1 FROM '" & Supplier_tbName & "' t2 WHERE t2.CCD = t1.CCD AND t2.Time > t1.Time);"

        Using connection As New SQLite.SQLiteConnection(connString)
            Using command As New SQLiteCommand(query, connection)
                connection.Open()

                Dim reader As SQLiteDataReader = command.ExecuteReader()
                If reader.Read() Then

                    CCD2Total = reader("NUM").ToString
                    CCD2OK = reader("OK").ToString
                    CCD2PlugDefect = reader("PlugDefect").ToString
                    CCD2WireDefect = reader("WireDefect").ToString

                    ' MessageBox.Show("Total: " & CCD2Total & ControlChars.NewLine & "Good: " & CCD2OK & ControlChars.NewLine & "Plug Defect: " & CCD2PlugDefect & ControlChars.NewLine & "Wire Defect: " & CCD2WireDefect & "  ", "885 Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No data found.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Using

    End Sub

    Sub LogResultofCCD2()
        Dim Lot As String
        Dim connectionString As String = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")
        Dim query As String = "UPDATE Camera2_tb SET T_Tested = @totalVal, Okay = @OkayVal, PlugDefect = @PlugDefVal, WireDefect = @WireDeftVal WHERE Lot_Num = @Lot"
        Lot = Form1.txtLotNumber.Text
        Using connection As New OleDbConnection(connectionString)
            Using command As New OleDbCommand(query, connection)
                command.Parameters.AddWithValue("@totalVal", CCD2Total)
                command.Parameters.AddWithValue("@OkayVal", CCD2OK)
                command.Parameters.AddWithValue("@PlugDefVal", CCD2PlugDefect)
                command.Parameters.AddWithValue("@WireDeftVal", CCD2WireDefect)
                command.Parameters.AddWithValue("@Lot", Lot)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public CCD3Total As String
    Public CCD3OK As String
    Public CCD3PlugDefect As String
    Public CCD3WireDefect As String

    Sub GetCCD3Result()
        'Dim connString As String = "Data Source=C:\LF Database\ProduceLog.db;Version=3"

        'Dim query As String = "SELECT t1.* FROM ProduceLog19783 t1 WHERE t1.CCD = 3 AND NOT EXISTS 
        '                        (SELECT 1 FROM ProduceLog19783 t2 WHERE t2.CCD = t1.CCD AND t2.Time > t1.Time);"

        Dim connString As String = "Data Source=" & dbPath & ";Version=3"

        Dim TableName As String = Date.Now.ToString("yyyy-MM-dd")

        Dim query As String = "SELECT t1.* FROM '" & Supplier_tbName & "' t1 WHERE t1.CCD = 3 AND NOT EXISTS 
                                (SELECT 1 FROM '" & Supplier_tbName & "' t2 WHERE t2.CCD = t1.CCD AND t2.Time > t1.Time);"

        Using connection As New SQLite.SQLiteConnection(connString)
            Using command As New SQLiteCommand(query, connection)
                connection.Open()

                Dim reader As SQLiteDataReader = command.ExecuteReader()
                If reader.Read() Then

                    CCD3Total = reader("NUM").ToString
                    CCD3OK = reader("OK").ToString
                    CCD3PlugDefect = reader("PlugDefect").ToString
                    CCD3WireDefect = reader("WireDefect").ToString

                    ' MessageBox.Show("Total: " & CCD3Total & ControlChars.NewLine & "Good: " & CCD3OK & ControlChars.NewLine & "Plug Defect: " & CCD3PlugDefect & ControlChars.NewLine & "Wire Defect: " & CCD3WireDefect & "  ", "885 Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No data found.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Using

    End Sub

    Sub LogResultofCCD3()
        Dim Lot As String
        Dim connectionString As String = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")
        Dim query As String = "UPDATE Camera3_tb SET T_Tested = @totalVal, Okay = @OkayVal, PlugDefect = @PlugDefVal, WireDefect = @WireDeftVal WHERE Lot_Num = @Lot"
        Lot = Form1.txtLotNumber.Text
        Using connection As New OleDbConnection(connectionString)
            Using command As New OleDbCommand(query, connection)
                command.Parameters.AddWithValue("@totalVal", CCD3Total)
                command.Parameters.AddWithValue("@OkayVal", CCD3OK)
                command.Parameters.AddWithValue("@PlugDefVal", CCD3PlugDefect)
                command.Parameters.AddWithValue("@WireDeftVal", CCD3WireDefect)
                command.Parameters.AddWithValue("@Lot", Lot)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public CCD4Total As String
    Public CCD4OK As String
    Public CCD4PlugDefect As String
    Public CCD4WireDefect As String

    Sub GetCCD4Result()
        'Dim connString As String = "Data Source=C:\LF Database\ProduceLog.db;Version=3"

        'Dim query As String = "SELECT t1.* FROM ProduceLog19783 t1 WHERE t1.CCD = 4 AND NOT EXISTS 
        '                        (SELECT 1 FROM ProduceLog19783 t2 WHERE t2.CCD = t1.CCD AND t2.Time > t1.Time);"

        Dim connString As String = "Data Source=" & dbPath & ";Version=3"

        Dim TableName As String = Date.Now.ToString("yyyy-MM-dd")

        Dim query As String = "SELECT t1.* FROM '" & Supplier_tbName & "' t1 WHERE t1.CCD = 4 AND NOT EXISTS 
                                (SELECT 1 FROM '" & Supplier_tbName & "' t2 WHERE t2.CCD = t1.CCD AND t2.Time > t1.Time);"

        Using connection As New SQLite.SQLiteConnection(connString)
            Using command As New SQLiteCommand(query, connection)
                connection.Open()

                Dim reader As SQLiteDataReader = command.ExecuteReader()
                If reader.Read() Then

                    CCD4Total = reader("NUM").ToString
                    CCD4OK = reader("OK").ToString
                    CCD4PlugDefect = reader("PlugDefect").ToString
                    CCD4WireDefect = reader("WireDefect").ToString

                    ' MessageBox.Show("Total: " & CCD4Total & ControlChars.NewLine & "Good: " & CCD4OK & ControlChars.NewLine & "Plug Defect: " & CCD4PlugDefect & ControlChars.NewLine & "Wire Defect: " & CCD4WireDefect & "  ", "885 Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No data found.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Using

    End Sub

    Sub LogResultofCCD4()
        Dim Lot As String
        Dim connectionString As String = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")
        Dim query As String = "UPDATE Camera4_tb SET T_Tested = @totalVal, Okay = @OkayVal, PlugDefect = @PlugDefVal, WireDefect = @WireDeftVal WHERE Lot_Num = @Lot"
        Lot = Form1.txtLotNumber.Text
        Using connection As New OleDbConnection(connectionString)
            Using command As New OleDbCommand(query, connection)
                command.Parameters.AddWithValue("@totalVal", CCD4Total)
                command.Parameters.AddWithValue("@OkayVal", CCD4OK)
                command.Parameters.AddWithValue("@PlugDefVal", CCD4PlugDefect)
                command.Parameters.AddWithValue("@WireDeftVal", CCD4WireDefect)
                command.Parameters.AddWithValue("@Lot", Lot)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public CCD5Total As String
    Public CCD5OK As String
    Public CCD5PlugDefect As String
    Public CCD5WireDefect As String

    Sub GetCCD5Result()
        'Dim connString As String = "Data Source=C:\LF Database\ProduceLog.db;Version=3"

        'Dim query As String = "SELECT t1.* FROM ProduceLog19783 t1 WHERE t1.CCD = 5 AND NOT EXISTS 
        '                        (SELECT 1 FROM ProduceLog19783 t2 WHERE t2.CCD = t1.CCD AND t2.Time > t1.Time);"

        Dim connString As String = "Data Source=" & dbPath & ";Version=3"

        Dim TableName As String = Date.Now.ToString("yyyy-MM-dd")

        Dim query As String = "SELECT t1.* FROM '" & Supplier_tbName & "' t1 WHERE t1.CCD = 5 AND NOT EXISTS 
                                (SELECT 1 FROM '" & Supplier_tbName & "' t2 WHERE t2.CCD = t1.CCD AND t2.Time > t1.Time);"

        Using connection As New SQLite.SQLiteConnection(connString)
            Using command As New SQLiteCommand(query, connection)
                connection.Open()

                Dim reader As SQLiteDataReader = command.ExecuteReader()
                If reader.Read() Then

                    CCD5Total = reader("NUM").ToString
                    CCD5OK = reader("OK").ToString
                    CCD5PlugDefect = reader("PlugDefect").ToString
                    CCD5WireDefect = reader("SheetDefect").ToString

                    'MessageBox.Show("Total: " & CCD5Total & ControlChars.NewLine & "Good: " & CCD5OK & ControlChars.NewLine & "Plug Defect: " & CCD5PlugDefect & ControlChars.NewLine & "Wire Defect: " & CCD5WireDefect & "  ", "885 Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No data found.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Using

    End Sub

    Sub LogResultofCCD5()
        Dim Lot As String
        Dim connectionString As String = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")
        Dim query As String = "UPDATE Camera5_tb SET T_Tested = @totalVal, Okay = @OkayVal, PlugDefect = @PlugDefVal, WireDefect = @WireDeftVal WHERE Lot_Num = @Lot"
        Lot = Form1.txtLotNumber.Text
        Using connection As New OleDbConnection(connectionString)
            Using command As New OleDbCommand(query, connection)
                command.Parameters.AddWithValue("@totalVal", CCD5Total)
                command.Parameters.AddWithValue("@OkayVal", CCD5OK)
                command.Parameters.AddWithValue("@PlugDefVal", CCD5PlugDefect)
                command.Parameters.AddWithValue("@WireDeftVal", CCD5WireDefect)
                command.Parameters.AddWithValue("@Lot", Lot)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public CCD6Total As String
    Public CCD6OK As String
    Public CCD6PlugDefect As String
    Public CCD6WireDefect As String

    Sub GetCCD6Result()
        'Dim connString As String = "Data Source=C:\LF Database\ProduceLog.db;Version=3"

        'Dim query As String = "SELECT t1.* FROM ProduceLog19783 t1 WHERE t1.CCD = 6 AND NOT EXISTS 
        '                        (SELECT 1 FROM ProduceLog19783 t2 WHERE t2.CCD = t1.CCD AND t2.Time > t1.Time);"

        Dim connString As String = "Data Source=" & dbPath & ";Version=3"

        Dim TableName As String = Date.Now.ToString("yyyy-MM-dd")

        Dim query As String = "SELECT t1.* FROM '" & Supplier_tbName & "' t1 WHERE t1.CCD = 6 AND NOT EXISTS 
                                (SELECT 1 FROM '" & Supplier_tbName & "' t2 WHERE t2.CCD = t1.CCD AND t2.Time > t1.Time);"

        Using connection As New SQLite.SQLiteConnection(connString)
            Using command As New SQLiteCommand(query, connection)
                connection.Open()

                Dim reader As SQLiteDataReader = command.ExecuteReader()
                If reader.Read() Then

                    CCD6Total = reader("NUM").ToString
                    CCD6OK = reader("OK").ToString
                    CCD6PlugDefect = reader("PlugDefect").ToString
                    CCD6WireDefect = reader("SheetDefect").ToString

                    'MessageBox.Show("Total: " & CCD6Total & ControlChars.NewLine & "Good: " & CCD6OK & ControlChars.NewLine & "Plug Defect: " & CCD6PlugDefect & ControlChars.NewLine & "Wire Defect: " & CCD6WireDefect & "  ", "885 Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No data found.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Using

    End Sub

    Sub LogResultofCCD6()
        Dim Lot As String
        Dim connectionString As String = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")
        Dim query As String = "UPDATE Camera6_tb SET T_Tested = @totalVal, Okay = @OkayVal, PlugDefect = @PlugDefVal, WireDefect = @WireDeftVal WHERE Lot_Num = @Lot"
        Lot = Form1.txtLotNumber.Text
        Using connection As New OleDbConnection(connectionString)
            Using command As New OleDbCommand(query, connection)
                command.Parameters.AddWithValue("@totalVal", CCD6Total)
                command.Parameters.AddWithValue("@OkayVal", CCD6OK)
                command.Parameters.AddWithValue("@PlugDefVal", CCD6PlugDefect)
                command.Parameters.AddWithValue("@WireDeftVal", CCD6WireDefect)
                command.Parameters.AddWithValue("@Lot", Lot)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

End Module