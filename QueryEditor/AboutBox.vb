Imports System.IO

Public Class AboutBox
    Private Sub AboutBox1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim appInf = My.Application.Info
        Dim applicationTitle = If(appInf.Title <> "",
                                  appInf.Title,
                                  Path.GetFileNameWithoutExtension(appInf.AssemblyName))

        Text = String.Format("About {0}", applicationTitle)

        txt_Version.Text = appInf.Version.ToString
        txt_Description.Text = appInf.Description
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles btn_OK.Click
        Close()
    End Sub
End Class
