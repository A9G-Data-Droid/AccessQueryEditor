Imports System.ComponentModel

<ToolboxBitmap(GetType(TreeView))>
Public Class DevTreeView
    Inherits TreeView

#Region "LastUsedNode  Stuff"

    ''' <summary>
    '''     this will be used with context menu that deals with each nodes
    '''     when it opens (the context menu I mean) which node will be affected ?......
    '''     this field will determine this node......
    '''     and its value will be set by several events of the TreeView
    ''' </summary>
    Private _lastUsedNode As TreeNode

    ''' <summary>
    '''     this will be used with context menu that deals with each nodes
    '''     when it opens (the context menu I mean) which node will be affected ?......
    '''     this Property will determine this node......
    ''' </summary>
    <Browsable(False)>
    Public ReadOnly Property LastUsedNode As TreeNode
        Get
            Return _lastUsedNode
        End Get
    End Property

    ''' <summary>
    '''     Sets the last node on check
    ''' </summary>
    ''' <param name="e">The event</param>
    Protected Overrides Sub OnAfterCheck(e As TreeViewEventArgs)
        MyBase.OnAfterCheck(e)
        _lastUsedNode = e.Node
    End Sub

    ''' <summary>
    '''     Sets the last node on click.
    ''' </summary>
    ''' <param name="e">Click event</param>
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        Dim cursorPos = New Point(e.X, e.Y)
        _lastUsedNode = HitTest(cursorPos).Node
    End Sub

#End Region
End Class
