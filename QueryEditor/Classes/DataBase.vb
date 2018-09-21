Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Imports QueryEditor.Queries
Public Class DataBase

#Region "Connection"

#Region "Private Fields"
    ''' <summary>the member that is responsible for connecting to the database </summary>
    Private Con As New OleDbConnection
#End Region


    Sub SetConnectionString()
        If Con IsNot Nothing AndAlso Con.State = ConnectionState.Open Then
            Con.Close()
        End If

        Con.ConnectionString = GetConStr()

    End Sub

    ''' <summary>
    ''' opens the database that calling instance is representing, 
    ''' but it doen't populate its members with the info from the source
    ''' </summary>
    Public Function OpenDB() As Boolean
        Try
            If Con.State <> ConnectionState.Open Then

                Con.Open()
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Caution")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' closes the database that calling instance is representing, 
    ''' if it's already closed then no procedure will be done
    ''' </summary>
    Public Function CloseDB() As Boolean
        Try
            If CBool(Con.State And Not ConnectionState.Closed) Then
                Con.Close()
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Caution")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' gets the connection string required to connect to the database ,
    ''' connection string will be gathered from other instance members
    ''' </summary>
    Private Function GetConStr() As String

        Dim pw = If(_PassWord = "", "", "Jet OLEDB:Database Password = " & _PassWord & ";")
        Dim rtn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & _
                _Path & """;" & pw

        Return rtn

    End Function

#End Region

#Region "Dealing With the DataBase Schema"

#Region "Internally used variables and functions"

    Private populated As Boolean = False

    ''' <summary>
    ''' fills the instace members with the info from the source
    ''' </summary>
    ''' <param name="sender">a value to determine the caller of this method</param>
    ''' <remarks></remarks>
    Private Sub populate(ByVal sender As Object)

        Try
            Con.Open()

            _Tables = GetTables()
            _Queries = GetQueries()

            Con.Close()

            _All_Names = From x In New String() {}

            Dim AllObjects = _Tables.Union(_Queries) 'not we can't use Tables/Queries properties cuz populated Field is still False ....they will call Populate again 
            Dim tmp0 As New List(Of String)
            For Each T In AllObjects
                Application.DoEvents()
                tmp0.Add(T.TableName)
                tmp0.AddRange(From C As DataColumn In T.Columns _
                              Select C.ColumnName)
            Next

            _All_Names = From N In tmp0 Select N Distinct

            populated = True

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "Properties"

    Private _PassWord As String = ""
    ''' <summary>Gets and Sets the password of the calling instance </summary>
    Public Property PassWord() As String
        Get
            Return _PassWord
        End Get
        Set(ByVal value As String)
            _PassWord = value
            Con.ConnectionString = GetConStr()
            populated = False
        End Set
    End Property

    Private _Path As String = ""
    ''' <summary>Gets and Sets the Path of the calling instance </summary>
    Public Property Path() As String
        Get
            Return _Path
        End Get
        Set(ByVal value As String)
            _Path = value
            Con.ConnectionString = GetConStr()
            populated = False
        End Set
    End Property

    Private _Tables As IEnumerable(Of DataTable)
    ''' <summary>a list of all Tables Names in the calling instance</summary>
    Public ReadOnly Property Tables() As IEnumerable(Of DataTable)
        Get
            If Not populated Then populate("Tables Property : if Not populated")

            If _Tables Is Nothing Then
                _Tables = From t In New DataTable() {}
            End If

            Return _Tables
        End Get
    End Property

    Private _Queries As IEnumerable(Of DataTable)
    ''' <summary>a list of all Queries Names in the calling instance</summary>
    Public ReadOnly Property Queries() As IEnumerable(Of DataTable)
        Get
            If Not populated Then populate("Queries Property : if Not populated")

            If _Queries Is Nothing Then
                _Queries = From t In New DataTable() {}
            End If

            Return _Queries
        End Get
    End Property

    Private _All_Names As IEnumerable(Of String)
    ''' <summary>
    ''' a list of all Objects Names (Tables/Queries/Columns) in the calling instance
    ''' this is gonna be used for faster DB Objects queries , 
    ''' so it doesn't iterate Tables or Queries to get Columns Names
    ''' </summary>
    Public ReadOnly Property All_Names() As IEnumerable(Of String)
        Get
            Return _All_Names
        End Get
    End Property

    ''' <summary>
    ''' gets a list of DB Objects Aoutocomplete words for a given word
    ''' </summary>
    ''' <param name="Word">the words to search for Autocompletes</param>
    Public ReadOnly Property GetAutoComplete(ByVal Word As String) As IEnumerable(Of String)
        Get

            If Path.Trim <> "" Then 'if DB is not Specified yet that means "DB Objects" are not available
                Return From N In All_Names _
                       Where N.StartsWith(Word, StringComparison.OrdinalIgnoreCase) _
                       Select "[" & N & "]"
            Else
                Return (From x In New String() {})

            End If

        End Get
    End Property
#End Region

#Region "Subs And Functions"

    Public Sub New(ByVal DBPath As String, ByVal DBPassword As String)
        MyBase.New()
        Me.Path = DBPath
        Me.PassWord = DBPassword
        SetConnectionString()

    End Sub

    ''' <summary>
    ''' a list of Queries in the calling instance (returned as DataTable Type)
    ''' </summary>
    Private Function GetQueries() As IEnumerable(Of DataTable)

        Dim Tbl As DataTable
        Tbl = Con.GetOleDbSchemaTable(OleDbSchemaGuid.Procedures, _
                                      New Object() {})


        Dim Rslt0 = From Row As DataRow In Tbl.Rows _
                   Where Not Row.Item("PROCEDURE_NAME").StartsWith("~") _
                   Select GetDataTableSchema("[" & CType(Row.Item("PROCEDURE_NAME"), String) & "]")


        Tbl = Con.GetOleDbSchemaTable(OleDbSchemaGuid.Views, _
                              New Object() {})

        Dim Rslt1 = From Row As DataRow In Tbl.Rows _
                    Select GetDataTableSchema("[" & CType(Row.Item("TABLE_NAME"), String) & "]")


        Return Rslt0.Union(Rslt1)

    End Function

    '''<summary > a list of Talbes in the calling instance (returned as DataTable Type)</summary>
    Private Function GetTables() As IEnumerable(Of DataTable)

        Dim Tbl As DataTable

        Tbl = Con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables_Info, New Object() {})

        Dim GoodTablesTypes = New String() {"TABLE", "LINK", "PASS-THROUGH"} 'I'm afraid there are more ,I didn't try all types , but I'll try to

        Dim Fun_GoodTableType = Function(Table_Type As String) _
                                   GoodTablesTypes.Contains(Table_Type)

        'gets the TablesNames only
        Dim TablesNames = From Row As DataRow In Tbl.Rows _
                          Where Fun_GoodTableType(CType(Row.Item("TABLE_TYPE"), String)) _
                          Select "[" & CType(Row.Item("TABLE_NAME"), String) & "]" Distinct


        'Now Fetch the Schema of tables
        Dim Rslt = From TblName As String In TablesNames _
                   Select GetDataTableSchema(TblName)

        Return Rslt
    End Function

    ''' <summary>
    ''' re-fill instance members with a Info from the source
    ''' </summary>
    Public Sub RefreshSchema()
        populate("RefreshSchema")
    End Sub

#End Region

#End Region

#Region "Dealing with the Data it self"

    ''' <summary>Executes a given sql statement and returns the count of the affected records</summary>
    ''' <param name="Query">the SqlQuery statement to execute</param>
    ''' <param name="Errors">this will hold the errors msgs in case things didn't go right</param>
    ''' <returns>count of the affected records</returns>
    Public Function ExecuteSQLCommand(ByVal Query As SqlQuery, _
                                      Optional ByRef Errors As String = "") As Long

        Dim count As Long

        Dim I_Opened_it As Boolean = False

        If Con.State <> ConnectionState.Open Then
            Con.Open()
            I_Opened_it = True
        End If

        Dim cmd = New OleDbCommand(Query.Sql, Con)

        Try
            If Query.QueryParams IsNot Nothing AndAlso Query.QueryParams.Count > 0 Then
                cmd.Parameters.Clear()
                For Each p In Query.QueryParams
                    cmd.Parameters.AddWithValue(p.Key, p.Value)
                Next
            End If
            count = cmd.ExecuteNonQuery()
        Catch ex As Exception
            Errors = ex.Message

        Finally
            If I_Opened_it Then
                Con.Close()
            End If

        End Try

        Return count

    End Function

    ''' <summary>Returns a DataTable instance that holds the returned data of a specified Sql statement</summary>
    ''' <param name="Query">SqlQuery statement to execute</param>
    ''' <param name="Errors">this will hold the errors msgs in case things didn't go right</param>
    ''' <param name="TableName">a value to be set to the name of the returned table</param>
    Public Function GetDataTable(ByVal Query As SqlQuery, _
                                 Optional ByRef Errors As String = "", _
                                 Optional ByVal TableName As String = "Table") As DataTable

        Dim cmd = New OleDbCommand(Query.Sql, Con)

        Dim da As New OleDb.OleDbDataAdapter
        Dim ds As New DataSet
        Dim dt As New DataTable(TableName)

        Try
            da.SelectCommand = cmd
            If Query.QueryParams IsNot Nothing AndAlso Query.QueryParams.Count > 0 Then
                cmd.Parameters.Clear()
                For Each p In Query.QueryParams
                    cmd.Parameters.AddWithValue(p.Key, p.Value)
                Next
            End If

            da.Fill(ds, TableName)

            If ds.Tables.Count > 0 Then
                dt = ds.Tables(0)
                dt.TableName = TableName
            End If

        Catch ex As Data.OleDb.OleDbException
            Errors = ex.Message
        End Try

        cmd.Dispose()
        cmd = Nothing
        Return dt
    End Function

    ''' <summary>Returns a DataTable instance that holds no data , it has only the schema of the specified exised Table Name</summary>
    ''' <param name="Errors">this will hold the errors msgs in case things didn't go right</param>
    ''' <param name="TableName">the name of the table to return the schema of</param>
    Public Function GetDataTableSchema(ByVal TableName As String, _
                                       Optional ByRef Errors As String = "" _
                                       ) As DataTable


        Dim cmd = New OleDbCommand("Select * From " & TableName, Con)

        Dim da As New OleDb.OleDbDataAdapter
        Dim ds As New DataSet
        Dim dt As New DataTable()

        Try

            da.SelectCommand = cmd
            da.FillSchema(ds, SchemaType.Source)

            If ds.Tables.Count > 0 Then
                dt = ds.Tables(0)
            End If

            dt.TableName = Regex.Replace(TableName, "\[|\]", "")
        Catch ex As Data.OleDb.OleDbException
            Errors = ex.Message
        End Try

        Return dt
    End Function

#End Region

#Region "Overrided Subs"

    Protected Overrides Sub Finalize()
        If Con IsNot Nothing Then
            If CBool(Con.State And ConnectionState.Open) Then
                Con.Close()
            End If

            Con = Nothing
        End If


        MyBase.Finalize()
    End Sub
#End Region

End Class
