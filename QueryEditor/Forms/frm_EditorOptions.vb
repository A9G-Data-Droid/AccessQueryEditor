Imports QueryEditor.ColoringWords

Public Class frm_EditorOptions

#Region "Added Functions & Subs"
    Sub GetColorsAndApplyBtns()
        btn_Literal_OperatorsColor.BackColor = ColoredWords.SqlLiteralOperatorsColor
        btn_FunctionsColor.BackColor = ColoredWords.FunctionsColor
        btn_OperatorsColor.BackColor = ColoredWords.SqlOperatorsColor
        btn_KeyWords_Color.BackColor = ColoredWords.KeyWordsColor
        btn_StringsColor.BackColor = ColoredWords.StringsColor
        btn_CommentsColor.BackColor = ColoredWords.CommentsColor
        btn_PunctuationColor.BackColor = ColoredWords.SqlPunctuationColor

    End Sub

#End Region

    Private Sub btn_Color_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
            Handles btn_Literal_OperatorsColor.Click, btn_CommentsColor.Click, _
                    btn_FunctionsColor.Click, btn_OperatorsColor.Click, _
                    btn_KeyWords_Color.Click, btn_StringsColor.Click, _
                    btn_PunctuationColor.Click

        Using ClrDialog As New ColorDialog
            With ClrDialog
                .ShowDialog()
                CType(sender, Button).BackColor = .Color
            End With
        End Using
    End Sub

    Private Sub frm_EditorOptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.GetColorsAndApplyBtns()

    End Sub

    Private Sub Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
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
        Me.Close()


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ColoredWords.ResetColors()
        GetColorsAndApplyBtns()
    End Sub
End Class