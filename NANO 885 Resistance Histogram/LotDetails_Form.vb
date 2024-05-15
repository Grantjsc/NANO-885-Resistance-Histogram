Public Class LotDetails_Form
    Private Sub btnReturn_Click(sender As Object, e As EventArgs) Handles btnReturn.Click
        ReturnHome()
    End Sub

    Private Sub LotDetails_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'LoadData()
    End Sub

    Private Sub txtPartNo_KeyUp(sender As Object, e As KeyEventArgs) Handles txtPartNo.KeyUp
        If e.KeyCode = Keys.Enter Then
            CheckDatabaseLotDetails()
        End If
    End Sub

    Private Sub txtLotNumber_KeyUp(sender As Object, e As KeyEventArgs) Handles txtLotNumber.KeyUp
        If e.KeyCode = Keys.Enter Then
            If txtLotNumber.Text = "" Then
                MsgBox("Please enter lot number!", MsgBoxStyle.Critical)
            Else
                'LoadData()
                GetLotNumberHistoryforLotDetials()
                GetResultHistoryCCD1()
                GetResultHistoryCCD2()
                GetResultHistoryCCD3()
                GetResultHistoryCCD4()
                GetResultHistoryCCD5()
                GetResultHistoryCCD6()

                'GetLotNumberHistoryforLotDetials()
            End If
        End If
    End Sub
End Class