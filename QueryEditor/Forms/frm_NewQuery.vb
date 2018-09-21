Imports System.Windows.Forms

Public Class frm_NewQuery

    Public ClearOutput As Boolean = False
    Public ClearParams As Boolean = False

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        Me.ClearOutput = Me.chk_ClearOutput.Checked
        Me.ClearParams = Me.chk_ClearParams.Checked

        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.Close()
    End Sub

    Private Sub frm_NewQuery_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class

