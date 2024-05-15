Public Class ChangePath_Form
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If cboPartNumber.Text = "" Then
            MsgBox("Plese input part number!", MsgBoxStyle.Critical)
        ElseIf txtNewPath.Text = "" Then
            MsgBox("Plese input new path!", MsgBoxStyle.Critical)
        Else
            ChangePathofPN()
        End If
    End Sub

    Private Sub ChangePath_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtNewPath.Text = ""
        cboPartNumber.Text = Nothing
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class