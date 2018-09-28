Public Class FrmAddParams
    Private Sub Closing_Clicks(sender As Object, e As EventArgs) _
        Handles btn_OK.Click, btn_Cancel.Click

        If DialogResult = DialogResult.Cancel Then
            Close()
            Exit Sub
        End If

        Try
            DGV_Params.AllowUserToAddRows = False 'remove the last Null row
            Dim notWanted = New Object() {"", DBNull.Value, Nothing}

            Editor.CurrentDb.Query.QueryParams.Clear()
            For Each row As DataGridViewRow In _
                From dgvRow As DataGridViewRow In DGV_Params.Rows _
                    Where Not NotWanted.Contains(dgvRow.Cells(0).Value) And
                          Not NotWanted.Contains(dgvRow.Cells(1).Value)

                Editor.CurrentDb.Query.QueryParams.Add(CType(Row.Cells(0).Value, String),
                                                CType(Row.Cells(1).Value, String))
            Next
            Close()
        Catch ex As Exception
            MsgBox("Error Encountered" & Environment.NewLine & Environment.NewLine &
                   ex.Message,
                   MsgBoxStyle.Exclamation, "Caution")
            DialogResult = DialogResult.Cancel
        End Try
    End Sub

    Private Sub Frm_AddParams_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            btn_Cancel.PerformClick()
        End If
    End Sub

    Private Sub Frm_AddParams_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DGV_Params.Rows.Clear()

        If Editor.CurrentDb.Query.QueryParams IsNot Nothing AndAlso Editor.CurrentDb.Query.QueryParams.Count > 0 Then
            For Each P In Editor.CurrentDb.Query.QueryParams
                DGV_Params.Rows.Add(New Object() {P.Key, P.Value})
            Next
        End If
    End Sub
End Class