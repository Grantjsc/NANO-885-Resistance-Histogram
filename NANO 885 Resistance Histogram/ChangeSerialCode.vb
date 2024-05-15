Imports System.Configuration
Imports System.IO.Ports

Module ChangeSerial_Module

    Public TegamCOMname As String
    Public TriggerCOMname As String


    Sub GetSerial1()
        Dim serial1 As String = System.Configuration.ConfigurationManager.AppSettings("Serial1")
        Console.WriteLine(serial1)

        TegamCOMname = serial1
    End Sub
    Sub GetSerial2()
        Dim serial2 As String = System.Configuration.ConfigurationManager.AppSettings("Serial2")
        Console.WriteLine(serial2)

        TriggerCOMname = serial2
    End Sub
    Sub LoadComPort1()
        Dim TegamportName As String() = SerialPort.GetPortNames()
        For Each COMname As String In TegamportName
            ChangeSerial.cboTegam.Items.Add(COMname)
        Next
    End Sub

    Sub LoadComPort2()
        Dim TrigportName As String() = SerialPort.GetPortNames()
        For Each ComPortName As String In TrigportName
            ChangeSerial.cboTrigger.Items.Add(ComPortName)
        Next
    End Sub

    Public config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

    Public NewSerial1 As String
    Public NewSerial2 As String

    Sub GetCOMport1()
        NewSerial1 = ChangeSerial.cboTegam.Text
    End Sub
    Sub ChangeTegamCOMname()
        config.AppSettings.Settings("Serial1").Value = NewSerial1 ' Rewrite 
        config.Save(ConfigurationSaveMode.Modified) ' save the new value

        ConfigurationManager.RefreshSection("appSettings") 'refresh
    End Sub
    Sub GetCOMport2()
        NewSerial2 = ChangeSerial.cboTrigger.Text
    End Sub
    Sub ChangeTriggerCOMname()
        config.AppSettings.Settings("Serial2").Value = NewSerial2 ' Rewrite 
        config.Save(ConfigurationSaveMode.Modified) ' save the new value

        ConfigurationManager.RefreshSection("appSettings") 'refresh
    End Sub

    Sub ChangeNames()
        If ChangeSerial.cboTegam.Text = "" And ChangeSerial.cboTrigger.Text = "" Then
            MsgBox("Please select SerialPort name!", MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(ChangeSerial.cboTegam.Text) Then
            GetCOMport2()
            ChangeTriggerCOMname()

            ChangeSerial.cboTegam.Items.Clear()
            ChangeSerial.cboTrigger.Items.Clear()
            ChangeSerial.Close()
        ElseIf String.IsNullOrEmpty(ChangeSerial.cboTrigger.Text) Then
            GetCOMport1()
            ChangeTegamCOMname()

            ChangeSerial.cboTegam.Items.Clear()
            ChangeSerial.cboTrigger.Items.Clear()
            ChangeSerial.Close()
        Else

            GetCOMport1()
            GetCOMport2()
            ChangeTegamCOMname()
            ChangeTriggerCOMname()

            ChangeSerial.cboTegam.Items.Clear()
            ChangeSerial.cboTrigger.Items.Clear()
            ChangeSerial.Close()

        End If

        GetSerial1()
        GetSerial2()
        Form1.SerialPort1.PortName = TegamCOMname
        Form1.SerialPort2.PortName = TriggerCOMname
    End Sub
End Module