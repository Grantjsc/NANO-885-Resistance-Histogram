Imports System.Configuration

Module ChangeTableReference_Module

    Public tbRef As String

    Sub GetTableReference()
        Dim Tablenum As String = System.Configuration.ConfigurationManager.AppSettings("dbTableReference")
        Console.WriteLine(Tablenum)

        tbRef = Tablenum
    End Sub

    Public config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

    Public NewTableRefNum As String
    Sub GetNewTableRefNum()
        NewTableRefNum = ChangeTableRef_Form.txtNewRef.Text
    End Sub
    Sub ChangeTbRef()
        config.AppSettings.Settings("dbTableReference").Value = NewTableRefNum ' Rewrite Ohaus Weighing Scale COM name
        config.Save(ConfigurationSaveMode.Modified) ' save the new value

        ConfigurationManager.RefreshSection("appSettings") 'refresh
    End Sub

    Sub UpdateRef()
        If ChangeTableRef_Form.txtNewRef.Text = "" Then
            MsgBox("Please input new reference.", MsgBoxStyle.Critical)
        Else
            GetNewTableRefNum()
            ChangeTbRef()
            MsgBox("The database table reference number is now updated.", MsgBoxStyle.Information)
            ChangeTableRef_Form.Close()
        End If
    End Sub
End Module