Public Class frm_AddParams

    Private Sub Closing_Clicks(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles btn_OK.Click, btn_Cancel.Click

        
        If DialogResult = Windows.Forms.DialogResult.Cancel Then
            Close()
            Exit Sub
        End If

        Try

            DGV_Params.AllowUserToAddRows = False 'remove the last Null row
            Dim NotWanted As Object() = New Object() {"", DBNull.Value, Nothing}

            Query.QueryParams.Clear()
            For Each Row As DataGridViewRow In _
                   From DGV_Row As DataGridViewRow In DGV_Params.Rows _
                   Where Not NotWanted.Contains(DGV_Row.Cells(0).Value) And _
                         Not NotWanted.Contains(DGV_Row.Cells(1).Value)


                
                Query.QueryParams.Add(CType(Row.Cells(0).Value, String), _
                                CType(Row.Cells(1).Value, String))
            Next
            Close()
        Catch ex As Exception

            MsgBox("Error Encountered" & Environment.NewLine & Environment.NewLine & _
                   ex.Message, _
                   MsgBoxStyle.Exclamation, "Caution")
            DialogResult = Windows.Forms.DialogResult.Cancel

        End Try

    End Sub

    Private Sub frm_AddParams_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.btn_Cancel.PerformClick()
        End If
    End Sub

    Private Sub frm_AddParams_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DGV_Params.Rows.Clear()

        If Query.QueryParams IsNot Nothing AndAlso Query.QueryParams.Count > 0 Then

            For Each P In Query.QueryParams
                DGV_Params.Rows.Add(New Object() {P.Key, P.Value})
            Next

        End If
    End Sub
End Class