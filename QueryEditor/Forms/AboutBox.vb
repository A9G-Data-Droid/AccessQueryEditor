Imports System.IO
Public Class AboutBox

    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim AppInf = My.Application.Info
        Dim ApplicationTitle = If(AppInf.Title <> "", _
                                  AppInf.Title, _
                                  Path.GetFileNameWithoutExtension(AppInf.AssemblyName))

        Me.Text = String.Format("About {0}", ApplicationTitle)


        Me.txt_Version.Text = AppInf.Version.ToString
        Me.txt_Description.Text = AppInf.Description

    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_OK.Click
        Me.Close()
    End Sub


End Class
