Public Class ChangeTableRef_Form
    Private Sub ChangeTableRef_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetTableReference()
        lblLastRef.Text = tbRef
        txtNewRef.Text = ""
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        UpdateRef()
    End Sub
End Class