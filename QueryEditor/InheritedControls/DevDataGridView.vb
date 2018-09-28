Imports System.ComponentModel

<ToolboxBitmap(GetType(DataGridView))>
Public Class DevDataGridView
    Inherits DataGridView

#Region "IndexOnRowHeader"

    ''' <summary>
    '''     the internal value of the IndexOnRowHeader Property,
    '''     and it's set to true as an initialization
    ''' </summary>
    Private _indexOnRowHeader As Boolean = True

    ''' <summary>Indicates whether each row has a header of its index(starts by 1) or not</summary>
    <Category("Appearance")> _
    <Description("Indicates whether each row has a header of its index(starts by 1) or not")>
    Public Property IndexOnRowHeader As Boolean
        Get
            Return _IndexOnRowHeader
        End Get

        Set
            If _IndexOnRowHeader <> value Then
                _IndexOnRowHeader = value

                If RowHeadersVisible Then ' for Optimization , 
                    '     no need to set it if it's not Visible

                    For Each row As DataGridViewRow In Rows
                        row.HeaderCell.Value = GetRowHeaderText(row.Index)
                    Next
                End If
            End If
        End Set
    End Property

    ''' <summary>
    '''     The header of a row could be Set from different places,
    '''     so I made it in a Function to uniform the operations that are done to get it
    ''' </summary>
    ''' <param name="rowIndex">Index of the row whose header you're asking for</param>
    Private Function GetRowHeaderText(rowIndex As Integer) As String
        Return If(IndexOnRowHeader,
                  CType(RowIndex + 1, String),
                  String.Empty)
    End Function

#End Region

    ''' <summary>
    '''     Gives a little tip about why this control is made for
    ''' </summary>
    <CategoryAttribute("Misc")> _
    <Description("Gives a little tip about why this control is made for")>
    Public ReadOnly Property NewFeatures As String
        Get
            Return "1 - Inherits System.Windows.Forms.DataGridView :" & Environment.NewLine &
                   "2 - Enables you from putting the Number of each Row in its HeaderCell"
        End Get
    End Property

    ''' <summary>Occurs when the rows count of the calling instance is changed</summary>
    Public Event RowsCountChanged(sender As Object, e As EventArgs)

    ''' <summary>raises RowsCountChanged events in the needed time</summary>
    Private Sub OnRowsCountChanged() Handles Me.RowsAdded, Me.RowsRemoved
        RaiseEvent RowsCountChanged(Me, New EventArgs)
    End Sub

    Protected Overrides Sub OnCellFormatting(e As DataGridViewCellFormattingEventArgs)
        MyBase.OnCellFormatting(e)

        If RowHeadersVisible Then ' for Optimization , 
            '                       no need to set it if it's not Visible

            Rows(e.RowIndex).HeaderCell.Value = GetRowHeaderText(e.RowIndex)
        End If
    End Sub
End Class
