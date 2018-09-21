Public Class frm_Text2Code

    Public Text2Convert As String = ""

#Region "Controls Events Handlers"


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.comb_Lang.SelectedIndex = 0
        Me.txt_PlainText.Text = Text2Convert
    End Sub

    Private Sub btn_Convert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Convert.Click
        If Me.comb_Lang.SelectedIndex = 0 Then
            Me.txt_CodeText.Text = Me.ConvertToVBCode(Me.txt_PlainText.Lines, _
                                                      Me.chk_StrBldr.Checked, _
                                                      Me.chk_AddNewLine.Checked)
        Else
            Me.txt_CodeText.Text = Me.ConvertToCSharpCode(Me.txt_PlainText.Lines, _
                                                          Me.chk_StrBldr.Checked, _
                                                          Me.chk_AddNewLine.Checked)
        End If
    End Sub

    Private Sub btn_CpyOutPut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_CpyOutPut.Click
        If Me.txt_CodeText.Text <> "" Then
            My.Computer.Clipboard.Clear()
            My.Computer.Clipboard.SetText(Me.txt_CodeText.Text)
        End If
    End Sub
#End Region

#Region "Added Funs & Subs"
    Private Function ConvertToVBCode(ByVal lines As String(), ByVal useStringBuilder As Boolean, ByVal addNewLineSymbole As Boolean) As String
        Dim Output As New System.Text.StringBuilder(1024)

        If useStringBuilder Then
            Output.Append("Dim str As New System.Text.StringBuilder(1024)" & Environment.NewLine & Environment.NewLine)
            For Each Line As String In lines
                Output.Append(String.Format("str.Append (""{0}""{1}) {2}", _
                                            Line.Replace("""", """"""), _
                                            If(addNewLineSymbole, " & Environment.NewLine ", ""), _
                                            ControlChars.CrLf))
            Next
        Else
            Output.Append("Dim str As String = """"" & Environment.NewLine & Environment.NewLine)
            For Each Line As String In lines
                Output.Append(String.Format("str &= ""{0}""{1} {2}", _
                                            Line.Replace("""", """"""), _
                                            If(addNewLineSymbole, " & Environment.NewLine ", ""), _
                                            Environment.NewLine))
            Next
        End If


        Return Output.ToString()
    End Function

    Private Function ConvertToCSharpCode(ByVal lines As String(), ByVal useStringBuilder As Boolean, ByVal addNewLineSymbole As Boolean) As String
        Dim output As New System.Text.StringBuilder(1024)

        If useStringBuilder Then
            output.Append("StringBuilder str = new StringBuilder(1024);" & Environment.NewLine & Environment.NewLine)
            For Each Line As String In lines
                output.Append(String.Format("str.Append (@""{0}""{1}); {2}", _
                                            Line.Replace("""", """"""), _
                                            If(addNewLineSymbole, _
                                               " + Environment.NewLine ", ""), _
                                               Environment.NewLine))
            Next
        Else
            output.Append("String str = """";" & Environment.NewLine & Environment.NewLine)
            For Each Line As String In lines
                output.Append(String.Format("str += @""{0}""{1}; {2}", _
                                            Line.Replace("""", """"""), _
                                            If(addNewLineSymbole, _
                                               " + Environment.NewLine ", ""), _
                                               Environment.NewLine))
            Next
        End If


        Return output.ToString()
    End Function

#End Region

End Class