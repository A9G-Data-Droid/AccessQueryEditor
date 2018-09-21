Imports System.Windows.Forms

Public Class frm_UDL

#Region "Added Fields & Subs"
    Public DataBasePath As String = ""
    Public PassWord As String = ""

#End Region

#Region "Controls Events Handlers"

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        If validateValues() AndAlso TestDBCon() Then

            DataBasePath = txt_Path.Text
            PassWord = getPassWord()
        Else
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
        End If

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        DataBasePath = ""
        PassWord = ""
    End Sub

    Private Sub chk_ShowPW_Chars_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_ShowPW_Chars.CheckedChanged
        Me.txt_PW.PasswordChar = If(chk_ShowPW_Chars.Checked, "", "*")
    End Sub

    Private Sub btn_Browse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Browse.Click
        Using OpenDialog As New OpenFileDialog
            With OpenDialog

                .Filter = "MS Access Files (*.mdb)|*.mdb"
                .Title = "Select the Destenation DB to Execute the Queries"
                .Multiselect = False
                If OpenDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    txt_Path.Text = .FileName
                End If
            End With
        End Using
    End Sub

    Private Sub chk_BlankPW_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_BlankPW.CheckedChanged
        Me.txt_PW.Enabled = Not chk_BlankPW.Checked
    End Sub

    Private Sub btn_TestCon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_TestCon.Click

        If ValidateValues() Then
            Try
                If TestDBCon() = True Then
                    MsgBox("Test connection succeeded.", Title:="Success")
                Else
                    MsgBox("Test connection failed", MsgBoxStyle.Exclamation, "Caution")
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Caution")
            End Try
        End If

    End Sub

    Private Function TestDBCon() As Boolean
        Dim tmpdb As New DataBase(Me.txt_Path.Text, getPassWord)
        If tmpdb.OpenDB() = True Then
            tmpdb.CloseDB()
            Return True
        Else
            Return False
        End If

    End Function
#End Region


#Region "Form Events Handlers"
    Private Sub frm_UDL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txt_PW.Text = PassWord
        Me.txt_Path.Text = DataBasePath
    End Sub
#End Region

#Region "Added Subs & Functions"

    Private Function getPassWord() As String
        Return If(chk_BlankPW.Checked, "", Me.txt_PW.Text)
    End Function

    Private Function validateValues() As Boolean
        If Not IO.File.Exists(Me.txt_Path.Text) Then
            MsgBox("DataBase file is not existed", MsgBoxStyle.Exclamation, "Coution")
            Return False
        End If

        Return True
    End Function

#End Region

End Class
