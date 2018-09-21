Imports System.ComponentModel
<ToolboxBitmap(GetType(System.Windows.Forms.DataGridView))> _
Public Class DevDataGridView
    Inherits System.Windows.Forms.DataGridView

#Region "Added Properties"

#Region "IndexOnRowHeader Dealing"

    ''' <summary>the internal value of the IndexOnRowHeader Property, 
    ''' and it's set to true as an initialization
    ''' </summary>
    Private _IndexOnRowHeader As Boolean = True

    ''' <summary>Indicates whether each row has a header of its index(starts by 1) or not</summary>
    <CategoryAttribute("Appearance")> _
    <Description("Indicates whether each row has a header of its index(starts by 1) or not")> _
    Public Property IndexOnRowHeader() As Boolean
        Get
            Return _IndexOnRowHeader
        End Get
        Set(ByVal value As Boolean)
            If _IndexOnRowHeader <> value Then
                _IndexOnRowHeader = value

                If RowHeadersVisible Then ' for Optimization , 
                    '     no need to set it if it's not Visible

                    For Each row As DataGridViewRow In Me.Rows
                        row.HeaderCell.Value = GetRowHeaderText(row.Index)
                    Next
                End If
            End If
        End Set
    End Property

    ''' <summary>
    ''' The header of a row could be Set from different places, 
    ''' so I made it in a Function to uniform the operations that are done to get it
    ''' </summary>
    ''' <param name="RowIndex">Index of the row whose header you're asking for</param>
    Private Function GetRowHeaderText(ByVal RowIndex As Integer) As String
        Return If(Me.IndexOnRowHeader, _
                  CType(RowIndex + 1, String), _
                  String.Empty)
    End Function

#End Region

#Region "Misc"

    ''' <summary>
    ''' Gives a little tip about why this control is made for
    ''' </summary>
    <CategoryAttribute("Misc")> _
    <Description("Gives a little tip about why this control is made for")> _
    Public ReadOnly Property NewFeatures() As String
        Get
            Return "1 - Inherits System.Windows.Forms.DataGridView :" & Environment.NewLine & _
                   "2 - Enables you from putting the Number of each Row in its HeaderCell"

        End Get
    End Property

#End Region
#End Region

#Region "Added Events"
    ''' <summary>Occurs when the rows count of the calling instance is changed</summary>
    Public Event RowsCountChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    ''' <summary>raises RowsCountChanged events in the needed time</summary>
    Private Sub OnRowsCountChanged() Handles Me.RowsAdded, Me.RowsRemoved
        RaiseEvent RowsCountChanged(Me, New System.EventArgs)
    End Sub

#End Region

#Region "Overrides Subs and Functions"
    Protected Overrides Sub OnCellFormatting(ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs)
        MyBase.OnCellFormatting(e)

        If RowHeadersVisible Then ' for Optimization , 
            '                       no need to set it if it's not Visible

            Rows(e.RowIndex).HeaderCell.Value = GetRowHeaderText(e.RowIndex)
        End If

    End Sub


#End Region

End Class
