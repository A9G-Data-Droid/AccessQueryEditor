Imports QueryEditor.ColoringWords

Public Class EditorOptions

#Region "Added Functions & Subs"

    Private Sub ColorizeButtons()
        btn_Literal_OperatorsColor.BackColor = ColoredWords.SqlLiteralOperatorsColor
        btn_FunctionsColor.BackColor = ColoredWords.FunctionsColor
        btn_OperatorsColor.BackColor = ColoredWords.SqlOperatorsColor
        btn_KeyWords_Color.BackColor = ColoredWords.KeyWordsColor
        btn_StringsColor.BackColor = ColoredWords.StringsColor
        btn_CommentsColor.BackColor = ColoredWords.CommentsColor
        btn_PunctuationColor.BackColor = ColoredWords.SqlPunctuationColor
    End Sub

#End Region

    Private Shared Sub Btn_Color_Click(sender As Object, e As EventArgs) _
        Handles btn_Literal_OperatorsColor.Click, btn_CommentsColor.Click,
                btn_FunctionsColor.Click, btn_OperatorsColor.Click,
                btn_KeyWords_Color.Click, btn_StringsColor.Click,
                btn_PunctuationColor.Click

        Using clrDialog As New ColorDialog
            With clrDialog
                .ShowDialog()
                CType(sender, Button).BackColor = .Color
            End With
        End Using
    End Sub

    Private Sub Frm_EditorOptions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ColorizeButtons()
    End Sub

    Private Sub Close_Click(sender As Object, e As EventArgs) _
        Handles btn_Cancel.Click, btn_OK.Click

        If CType(sender, Button) Is btn_OK Then
            ColoredWords.SqlLiteralOperatorsColor = btn_Literal_OperatorsColor.BackColor
            ColoredWords.FunctionsColor = btn_FunctionsColor.BackColor
            ColoredWords.SqlOperatorsColor = btn_OperatorsColor.BackColor
            ColoredWords.KeyWordsColor = btn_KeyWords_Color.BackColor
            ColoredWords.StringsColor = btn_StringsColor.BackColor
            ColoredWords.CommentsColor = btn_CommentsColor.BackColor
            ColoredWords.SqlPunctuationColor = btn_PunctuationColor.BackColor
        End If
        Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ColoredWords.ResetColors()
        ColorizeButtons()
    End Sub
End Class