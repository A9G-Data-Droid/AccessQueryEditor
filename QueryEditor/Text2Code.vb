Imports System.Text

Public Class Text2Code
    Public Text2Convert As String = ""

#Region "Controls Events Handlers"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        comb_Lang.SelectedIndex = 0
        txt_PlainText.Text = Text2Convert
    End Sub

    Private Sub Btn_Convert_Click(sender As Object, e As EventArgs) Handles btn_Convert.Click
        If comb_Lang.SelectedIndex = 0 Then
            txt_CodeText.Text = ConvertToVBCode(txt_PlainText.Lines,
                                                   chk_StrBldr.Checked,
                                                   chk_AddNewLine.Checked)
        Else
            txt_CodeText.Text = ConvertToCSharpCode(txt_PlainText.Lines,
                                                       chk_StrBldr.Checked,
                                                       chk_AddNewLine.Checked)
        End If
    End Sub

    Private Sub Btn_CpyOutPut_Click(sender As Object, e As EventArgs) Handles btn_CpyOutPut.Click
        If txt_CodeText.Text <> "" Then
            My.Computer.Clipboard.Clear()
            My.Computer.Clipboard.SetText(txt_CodeText.Text)
        End If
    End Sub

#End Region

#Region "Added Funs & Subs"

    Private Shared Function ConvertToVbCode(lines As IEnumerable, useStringBuilder As Boolean,
                                            addNewLineSymbol As Boolean) As String
        Dim output As New StringBuilder(1024)

        If useStringBuilder Then
            output.Append("Dim str As New System.Text.StringBuilder(1024)" & Environment.NewLine & Environment.NewLine)
            For Each line As String In lines
                output.Append(String.Format("str.Append (""{0}""{1}) {2}",
                                            line.Replace("""", """"""),
                                            If(addNewLineSymbol, " & Environment.NewLine ", ""),
                                            ControlChars.CrLf))
            Next
        Else
            output.Append("Dim str As String = """"" & Environment.NewLine & Environment.NewLine)
            For Each line As String In lines
                output.Append(String.Format("str &= ""{0}""{1} {2}",
                                            line.Replace("""", """"""),
                                            If(addNewLineSymbol, " & Environment.NewLine ", ""),
                                            Environment.NewLine))
            Next
        End If

        Return output.ToString()
    End Function

    Private Shared Function ConvertToCSharpCode(lines As IEnumerable, useStringBuilder As Boolean,
                                                addNewLineSymbole As Boolean) As String
        Dim output As New StringBuilder(1024)

        If useStringBuilder Then
            output.Append("StringBuilder str = new StringBuilder(1024);" & Environment.NewLine & Environment.NewLine)
            For Each line As String In lines
                output.Append(String.Format("str.Append (@""{0}""{1}); {2}",
                                            line.Replace("""", """"""),
                                            If(addNewLineSymbole,
                                               " + Environment.NewLine ", ""),
                                            Environment.NewLine))
            Next
        Else
            output.Append("String str = """";" & Environment.NewLine & Environment.NewLine)
            For Each line As String In lines
                output.Append(String.Format("str += @""{0}""{1}; {2}",
                                            line.Replace("""", """"""),
                                            If(addNewLineSymbole,
                                               " + Environment.NewLine ", ""),
                                            Environment.NewLine))
            Next
        End If

        Return output.ToString()
    End Function

#End Region
End Class