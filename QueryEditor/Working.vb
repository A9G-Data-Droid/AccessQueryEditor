Public Class Working
    Private _message As String = ""

    Public Property Message As String
' ReSharper disable once UnusedMember.Global
        Get
            Return _message
        End Get

        Set
            If value <> _message Then
                _message = value
                lbl_Task.Text = value
            End If
        End Set
    End Property
End Class