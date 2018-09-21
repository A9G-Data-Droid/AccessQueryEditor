Namespace BindingDataToNode

    Public Class NodeBindingData

        Public NodeType As NodeTypes = NodeTypes.None


        Public Sub New(ByVal type As NodeTypes)
            MyBase.new()
            Me.NodeType = type
        End Sub
        Public Sub New()
            MyBase.new()
            Me.NodeType = NodeTypes.None
        End Sub
    End Class

    Public Enum NodeTypes As Byte

        None = 0
        FolderNode = 2
        FunctionNode = 4
        TableNode = 8
        QueryNode = 16
        FieldNode = 32
        DataBaseNode = 64
    End Enum

End Namespace