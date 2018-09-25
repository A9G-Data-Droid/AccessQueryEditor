Imports System.IO

Public Class ConnectionDialog

#Region "Added Fields & Subs"

    Public Property DataBasePath as String
        get
            Return txt_Path.Text
        End Get
        Set
            txt_Path.Text = value
        End Set
    End Property

    Public Property PassWord as String
        get
            Return If(chk_BlankPW.Checked, "", txt_PW.Text)
        End Get
        set
            txt_PW.Text = value
        End Set
    End Property

#End Region

#Region "Controls Events Handlers"

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        If validateValues() Then
            'DataBasePath = txt_Path.Text
            'PassWord = getPassWord()
        Else
            DialogResult = DialogResult.Cancel
        End If
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        txt_Path.Text = ""
        txt_PW.Text = ""
    End Sub

    Private Sub Chk_ShowPW_Chars_CheckedChanged(sender As Object, e As EventArgs) _
        Handles chk_ShowPW_Chars.CheckedChanged
        txt_PW.PasswordChar = If(chk_ShowPW_Chars.Checked, "", "*")
    End Sub

    Private Sub Btn_Browse_Click(sender As Object, e As EventArgs) Handles btn_Browse.Click
        Using openDialog As New OpenFileDialog
            With openDialog
                .Filter = "MS Access Files (*.mdb;*.accdb)|*.mdb;*.accdb"
                .Title = "Select the target DataBase"
                .Multiselect = False

                Dim dResult = openDialog.ShowDialog()
                If dResult = DialogResult.OK Then
                    txt_Path.Text = .FileName
                End If
            End With
        End Using
    End Sub

    Private Sub Chk_BlankPW_CheckedChanged(sender As Object, e As EventArgs) Handles chk_BlankPW.CheckedChanged
        txt_PW.Enabled = Not chk_BlankPW.Checked
    End Sub

    Private Async Sub Btn_TestCon_Click(sender As Object, e As EventArgs) Handles btn_TestCon.Click
        If ValidateValues() Then
            dim testDb As New DataBase(DataBasePath, PassWord)
            Try
                If Await testDb.TestConnection Then
                    MsgBox("Test connection succeeded.", Title := "Success")
                Else
                    MsgBox("Test connection failed", MsgBoxStyle.Exclamation, "Caution")
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Caution")
            End Try
        End If
    End Sub

#End Region

#Region "Form Events Handlers"

    Private Sub Frm_UDL_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txt_PW.Text = PassWord
        txt_Path.Text = DataBasePath
    End Sub

#End Region

    Private Function ValidateValues() As Boolean
        If Not File.Exists(txt_Path.Text) Then
            MsgBox("DataBase file not found.", MsgBoxStyle.Exclamation, "Caution!")
            Return False
        End If

        Return True
    End Function
End Class
