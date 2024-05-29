Public Class ChangeTiming_Form
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub ChangeTiming_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetTimer()
        txtNewTiming.Text = TimerTiming
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If txtNewTiming.Text = "" Then
            MsgBox("Please input new timer", MsgBoxStyle.Critical)
        Else
            GetNewTimer()
            ChangeOldTiming()
            GetTimer()
            Me.Close()
        End If
    End Sub
End Class