Public Class ChangeSerial
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        cboTegam.Items.Clear()
        cboTrigger.Items.Clear()
        Me.Close()
    End Sub

    Private Sub ChangeSerial_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadComPort1()
        LoadComPort2()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ChangeNames()
    End Sub
End Class