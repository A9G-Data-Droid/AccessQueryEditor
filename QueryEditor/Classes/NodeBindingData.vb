Namespace BindingDataToNode
    Public Class NodeBindingData
        Public ReadOnly NodeType As NodeTypes = NodeTypes.None

        Public Sub New(type As NodeTypes)
            MyBase.New()
            NodeType = type
        End Sub

' ReSharper disable once UnusedMember.Global
        Public Sub New()
            MyBase.New()
            NodeType = NodeTypes.None
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