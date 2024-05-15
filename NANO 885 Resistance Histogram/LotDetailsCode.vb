Imports System.Data.OleDb
Imports System.Data.SQLite
Imports System.Web.UI.WebControls

Module LotDetails_Module

    Public dbPathLot As String
    Sub GetPathforDBLotDetails()
        Select Case LotDetails_Form.txtPartNo.Text

            Case "0885001.DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("0885001.DR")
                Console.WriteLine(Path)
                dbPathLot = Path

            Case "08851.25DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("08851.25DR")
                Console.WriteLine(Path)
                dbPathLot = Path

            Case "088501.6DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("088501.6DR")
                Console.WriteLine(Path)
                dbPathLot = Path

            Case "0885002.DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("0885002.DR")
                Console.WriteLine(Path)
                dbPathLot = Path

            Case "088502.5DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("088502.5DR")
                Console.WriteLine(Path)
                dbPathLot = Path

            Case "08853.15DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("08853.15DR")
                Console.WriteLine(Path)
                dbPathLot = Path

            Case "0885004.DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("0885004.DR")
                Console.WriteLine(Path)
                dbPathLot = Path

            Case "0885005.DR"
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("0885005.DR")
                Console.WriteLine(Path)
                dbPathLot = Path
            Case Else
                MsgBox("Part number does not exist!", MsgBoxStyle.Critical)
                LotDetails_Form.txtPartNo.Text = ""
        End Select
    End Sub

    Sub CheckDatabaseLotDetails()

        GetPathforDBLotDetails()
        If Not System.IO.File.Exists(dbPathLot) Then
            MessageBox.Show("Database file not found!.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            LotDetails_Form.txtPartNo.Text = ""
            Exit Sub ' or handle the error in any other way as needed
        End If
        LotDetails_Form.txtLotNumber.Focus()
    End Sub


    Public table As New DataTable

    Public StartOledLotDetails As String
    Public EndOledLotDetails As String
    Public dbNameoftbLotDetails As String
    Sub GetLotNumberHistoryforLotDetials()
        Dim Start As String = "Sample"

        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")


        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From LotHistory_tb WHERE Lot_Num = '" + LotDetails_Form.txtLotNumber.Text + "'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then

                StartOledLotDetails = Data.Rows(0).Item("Start_time").ToString
                EndOledLotDetails = Data.Rows(0).Item("End_time").ToString
                dbNameoftbLotDetails = Data.Rows(0).Item("dbTableName").ToString
                Console.WriteLine(StartOledLotDetails)
                Console.WriteLine(EndOledLotDetails)
                Console.WriteLine(dbNameoftbLotDetails)

                LoadData()
            Else
                'MsgBox("Lot number does not exist in the database.", MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()
        End Try
    End Sub
    Sub LoadData()
        '"Data Source=C:\LF Database\ProduceLog.db;Version=3"

        '"Data Source=D:\IntelligentVisionPlatforms-x64-20230511\MicroMatch\server\127.0.0.1_6000\Project\0885004.DR\ProduceLog.db;Version=3"
        'Dim connString As String = "Data Source=C:\LF Database\ProduceLog.db;Version=3"

        'GetLotNumberHistoryforLotDetials()

        Dim connString As String = "Data Source=" & dbPathLot & ";Version=3"

        Dim connection As New SQLite.SQLiteConnection(connString)
        Dim command As New SQLiteCommand("", connection)

        connection.Open()

        If connection.State = ConnectionState.Open Then
            ' "SELECT * FROM '" & dbNameoftb & "' WHERE Time BETWEEN '" & StartOled & "'  AND '" & EndOled & "'"
            '"Select * From ProduceLog19783"

            command.Connection = connection
            command.CommandText = "SELECT * FROM '" & dbNameoftbLotDetails & "' WHERE Time BETWEEN '" & StartOledLotDetails & "'  AND '" & EndOledLotDetails & "'"
            'command.CommandText = "Select * From ProduceLog19783"

            Dim rdr As SQLiteDataReader = command.ExecuteReader

            'Dim table As New DataTable
            table.Load(rdr)

            LotDetails_Form.DataGridView1.DataSource = table

            ' Bold the header cells
            For Each column As DataGridViewColumn In LotDetails_Form.DataGridView1.Columns
                column.HeaderCell.Style.Font = New Font(LotDetails_Form.DataGridView1.Font, FontStyle.Bold)
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Next

        End If
        connection.Close()
    End Sub

    Sub ReturnHome()
        LotDetails_Form.Close()
    End Sub

End Module


'*********************** Get the Camera Result History ***************************
Module Cam1toCam6History_Module
    Public Cam1totalResult As String
    Public Cam1GoodResult As String
    Public Cam1PlugDefResult As String
    Public Cam1WireDefResult As String
    Sub GetResultHistoryCCD1()
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")

        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From Camera1_tb WHERE Lot_Num = '" + LotDetails_Form.txtLotNumber.Text + "'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then

                Cam1totalResult = Data.Rows(0).Item("T_Tested").ToString
                Cam1GoodResult = Data.Rows(0).Item("Okay").ToString
                Cam1PlugDefResult = Data.Rows(0).Item("PlugDefect").ToString
                Cam1WireDefResult = Data.Rows(0).Item("WireDefect").ToString

                LotDetails_Form.lblCam1Total.Text = Cam1totalResult
                LotDetails_Form.lblCam1Good.Text = Cam1GoodResult
                LotDetails_Form.lblCam1PDefect.Text = Cam1PlugDefResult
                LotDetails_Form.lblCam1WDefect.Text = Cam1WireDefResult
            Else
                MsgBox("Lot number does not exist in the database.", MessageBoxIcon.Error)
                LotDetails_Form.txtLotNumber.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()
        End Try
    End Sub

    Public Cam2totalResult As String
    Public Cam2GoodResult As String
    Public Cam2PlugDefResult As String
    Public Cam2WireDefResult As String
    Sub GetResultHistoryCCD2()
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")

        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From Camera2_tb WHERE Lot_Num = '" + LotDetails_Form.txtLotNumber.Text + "'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then

                Cam2totalResult = Data.Rows(0).Item("T_Tested").ToString
                Cam2GoodResult = Data.Rows(0).Item("Okay").ToString
                Cam2PlugDefResult = Data.Rows(0).Item("PlugDefect").ToString
                Cam2WireDefResult = Data.Rows(0).Item("WireDefect").ToString

                LotDetails_Form.lblCam2Total.Text = Cam2totalResult
                LotDetails_Form.lblCam2Good.Text = Cam2GoodResult
                LotDetails_Form.lblCam2PDefect.Text = Cam2PlugDefResult
                LotDetails_Form.lblCam2WDefect.Text = Cam2WireDefResult
            Else
                'MsgBox("Lot number does not exist in the database.", MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()
        End Try
    End Sub

    Public Cam3totalResult As String
    Public Cam3GoodResult As String
    Public Cam3PlugDefResult As String
    Public Cam3WireDefResult As String
    Sub GetResultHistoryCCD3()
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")

        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From Camera3_tb WHERE Lot_Num = '" + LotDetails_Form.txtLotNumber.Text + "'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then

                Cam3totalResult = Data.Rows(0).Item("T_Tested").ToString
                Cam3GoodResult = Data.Rows(0).Item("Okay").ToString
                Cam3PlugDefResult = Data.Rows(0).Item("PlugDefect").ToString
                Cam3WireDefResult = Data.Rows(0).Item("WireDefect").ToString

                LotDetails_Form.lblCam3Total.Text = Cam3totalResult
                LotDetails_Form.lblCam3Good.Text = Cam3GoodResult
                LotDetails_Form.lblCam3PDefect.Text = Cam3PlugDefResult
                LotDetails_Form.lblCam3WDefect.Text = Cam3WireDefResult
            Else
                'MsgBox("Lot number does not exist in the database.", MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()
        End Try
    End Sub

    Public Cam4totalResult As String
    Public Cam4GoodResult As String
    Public Cam4PlugDefResult As String
    Public Cam4WireDefResult As String
    Sub GetResultHistoryCCD4()
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")

        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From Camera4_tb WHERE Lot_Num = '" + LotDetails_Form.txtLotNumber.Text + "'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then

                Cam4totalResult = Data.Rows(0).Item("T_Tested").ToString
                Cam4GoodResult = Data.Rows(0).Item("Okay").ToString
                Cam4PlugDefResult = Data.Rows(0).Item("PlugDefect").ToString
                Cam4WireDefResult = Data.Rows(0).Item("WireDefect").ToString

                LotDetails_Form.lblCam4Total.Text = Cam4totalResult
                LotDetails_Form.lblCam4Good.Text = Cam4GoodResult
                LotDetails_Form.lblCam4PDefect.Text = Cam4PlugDefResult
                LotDetails_Form.lblCam4WDefect.Text = Cam4WireDefResult
            Else
                'MsgBox("Lot number does not exist in the database.", MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()
        End Try
    End Sub

    Public Cam5totalResult As String
    Public Cam5GoodResult As String
    Public Cam5PlugDefResult As String
    Public Cam5WireDefResult As String
    Sub GetResultHistoryCCD5()
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")

        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From Camera5_tb WHERE Lot_Num = '" + LotDetails_Form.txtLotNumber.Text + "'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then

                Cam5totalResult = Data.Rows(0).Item("T_Tested").ToString
                Cam5GoodResult = Data.Rows(0).Item("Okay").ToString
                Cam5PlugDefResult = Data.Rows(0).Item("PlugDefect").ToString
                Cam5WireDefResult = Data.Rows(0).Item("WireDefect").ToString

                LotDetails_Form.lblCam5Total.Text = Cam5totalResult
                LotDetails_Form.lblCam5Good.Text = Cam5GoodResult
                LotDetails_Form.lblCam5PDefect.Text = Cam5PlugDefResult
                LotDetails_Form.lblCam5WDefect.Text = Cam5WireDefResult
            Else
                'MsgBox("Lot number does not exist in the database.", MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()
        End Try
    End Sub

    Public Cam6totalResult As String
    Public Cam6GoodResult As String
    Public Cam6PlugDefResult As String
    Public Cam6WireDefResult As String
    Sub GetResultHistoryCCD6()
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\NANO 885 Data Log.accdb;Persist Security Info=True;Jet OLEDB:Database Password=lfnano885log")

        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From Camera6_tb WHERE Lot_Num = '" + LotDetails_Form.txtLotNumber.Text + "'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then

                Cam6totalResult = Data.Rows(0).Item("T_Tested").ToString
                Cam6GoodResult = Data.Rows(0).Item("Okay").ToString
                Cam6PlugDefResult = Data.Rows(0).Item("PlugDefect").ToString
                Cam6WireDefResult = Data.Rows(0).Item("WireDefect").ToString

                LotDetails_Form.lblCam6Total.Text = Cam6totalResult
                LotDetails_Form.lblCam6Good.Text = Cam6GoodResult
                LotDetails_Form.lblCam6PDefect.Text = Cam6PlugDefResult
                LotDetails_Form.lblCam6WDefect.Text = Cam6WireDefResult
            Else
                'MsgBox("Lot number does not exist in the database.", MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()
        End Try
    End Sub
End Module

