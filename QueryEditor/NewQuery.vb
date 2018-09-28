Public Class NewQuery
    Public ClearOutput As Boolean = False
    Public ClearParams As Boolean = False

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        ClearOutput = chk_ClearOutput.Checked
        ClearParams = chk_ClearParams.Checked

        Close()
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Close()
    End Sub
End Class

