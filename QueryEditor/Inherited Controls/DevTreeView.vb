Imports System.ComponentModel

<ToolboxBitmap(GetType(System.Windows.Forms.TreeView))> _
Public Class DevTreeView
    Inherits TreeView

#Region "LastUsedNode  Stuff"
    ''' <summary>
    ''' this will be used with context menu that deals with each nodes 
    ''' when it opens (the context menue I mean) which node will be affected ?......
    ''' this field will determain this node......
    ''' and its value will be set by seviral events of the TreeView
    ''' </summary>
    Private _LastUsedNode As TreeNode

    ''' <summary>
    ''' this will be used with context menu that deals with each nodes 
    ''' when it opens (the context menue I mean) which node will be affected ?......
    ''' this Property will determain this node......
    ''' </summary>
    <Browsable(False)> _
    Public ReadOnly Property LastUsedNode() As TreeNode
        Get
            Return _LastUsedNode
        End Get
    End Property

    Protected Overrides Sub OnAfterCheck(ByVal e As System.Windows.Forms.TreeViewEventArgs)
        MyBase.OnAfterCheck(e)
        _LastUsedNode = e.Node
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)

        Dim CursorPos = New Point(e.X, e.Y)
        _LastUsedNode = Me.HitTest(CursorPos).Node
    End Sub

#End Region

End Class
