Imports System.Configuration

Module ChangePath_Module

    Public NewDBPath As String
    Public config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

    Sub GetNewPathofPN()
        NewDBPath = ChangePath_Form.txtNewPath.Text & "\ProduceLog.db"
    End Sub

    Sub ChangePathofPN()
        Select Case ChangePath_Form.cboPartNumber.Text

            Case "0885001.DR"

                GetNewPathofPN()
                config.AppSettings.Settings("0885001.DR").Value = NewDBPath ' Rewrite 
                config.Save(ConfigurationSaveMode.Modified) ' save the new value

                ConfigurationManager.RefreshSection("appSettings") 'refresh

                MsgBox("The changes in the path are complete.", MsgBoxStyle.Information)
                ChangePath_Form.cboPartNumber.Text = Nothing
                ChangePath_Form.txtNewPath.Text = ""

            Case "08851.25DR"
                GetNewPathofPN()
                config.AppSettings.Settings("08851.25DR").Value = NewDBPath ' Rewrite 
                config.Save(ConfigurationSaveMode.Modified) ' save the new value

                ConfigurationManager.RefreshSection("appSettings") 'refresh

                MsgBox("The changes in the path are complete.", MsgBoxStyle.Information)
                ChangePath_Form.cboPartNumber.Text = Nothing
                ChangePath_Form.txtNewPath.Text = ""

            Case "088501.6DR"
                GetNewPathofPN()
                config.AppSettings.Settings("088501.6DR").Value = NewDBPath ' Rewrite 
                config.Save(ConfigurationSaveMode.Modified) ' save the new value

                ConfigurationManager.RefreshSection("appSettings") 'refresh

                MsgBox("The changes in the path are complete.", MsgBoxStyle.Information)
                ChangePath_Form.cboPartNumber.Text = Nothing
                ChangePath_Form.txtNewPath.Text = ""

            Case "0885002.DR"
                GetNewPathofPN()
                config.AppSettings.Settings("0885002.DR").Value = NewDBPath ' Rewrite 
                config.Save(ConfigurationSaveMode.Modified) ' save the new value

                ConfigurationManager.RefreshSection("appSettings") 'refresh

                MsgBox("The changes in the path are complete.", MsgBoxStyle.Information)
                ChangePath_Form.cboPartNumber.Text = Nothing
                ChangePath_Form.txtNewPath.Text = ""

            Case "088502.5DR"
                GetNewPathofPN()
                config.AppSettings.Settings("088502.5DR").Value = NewDBPath ' Rewrite 
                config.Save(ConfigurationSaveMode.Modified) ' save the new value

                ConfigurationManager.RefreshSection("appSettings") 'refresh

                MsgBox("The changes in the path are complete.", MsgBoxStyle.Information)
                ChangePath_Form.cboPartNumber.Text = Nothing
                ChangePath_Form.txtNewPath.Text = ""

            Case "08853.15DR"
                GetNewPathofPN()
                config.AppSettings.Settings("08853.15DR").Value = NewDBPath ' Rewrite 
                config.Save(ConfigurationSaveMode.Modified) ' save the new value

                ConfigurationManager.RefreshSection("appSettings") 'refresh

                MsgBox("The changes in the path are complete.", MsgBoxStyle.Information)
                ChangePath_Form.cboPartNumber.Text = Nothing
                ChangePath_Form.txtNewPath.Text = ""

            Case "0885004.DR"
                GetNewPathofPN()
                config.AppSettings.Settings("0885004.DR").Value = NewDBPath ' Rewrite 
                config.Save(ConfigurationSaveMode.Modified) ' save the new value

                ConfigurationManager.RefreshSection("appSettings") 'refresh

                MsgBox("The changes in the path are complete.", MsgBoxStyle.Information)
                ChangePath_Form.cboPartNumber.Text = Nothing
                ChangePath_Form.txtNewPath.Text = ""

            Case "0885005.DR"
                GetNewPathofPN()
                config.AppSettings.Settings("0885005.DR").Value = NewDBPath ' Rewrite 
                config.Save(ConfigurationSaveMode.Modified) ' save the new value

                ConfigurationManager.RefreshSection("appSettings") 'refresh

                MsgBox("The changes in the path are complete.", MsgBoxStyle.Information)
                ChangePath_Form.cboPartNumber.Text = Nothing
                ChangePath_Form.txtNewPath.Text = ""
        End Select
    End Sub
End Module