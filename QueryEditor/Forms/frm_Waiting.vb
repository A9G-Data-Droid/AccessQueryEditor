Public Class frm_Working

    Private _Task As String = ""
    Public Property Task() As String
        Get
            Return _Task
        End Get
        Set(ByVal value As String)
            If value <> _Task Then
                _Task = value
                Me.lbl_Task.Text = value
            End If
        End Set
    End Property

    Public Overloads Sub Show(ByVal task As String)
        Me.Task = task
        Me.Show()
        Application.DoEvents()
    End Sub
 
    Private Sub frm_Working_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class