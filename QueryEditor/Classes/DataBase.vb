Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports QueryEditor.Queries

Public Class DataBase
#Region "Properties"
    Public Property PassWord As String = ""

    Public Property Path As String = ""

    Private ReadOnly Property ConnectionString As String
    Get
        Dim connectionBuilder As New OleDbConnectionStringBuilder
        With connectionBuilder
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .DataSource = Path
            .PersistSecurityInfo = False
            '.Add("Connect Timeout", 15)
            If PassWord.Length > 0 Then .Add("Jet OLEDB:Database Password", PassWord)
        End With
        
        Return connectionBuilder.ToString()
    End Get
    End Property

    ''' <summary>Represents the Query we're working on</summary>
    Public Property Query As New SqlQuery("")

    Private _tables As IEnumerable(Of DataTable) = From t In New DataTable() {}
    ''' <summary>a list of all Tables Names in the calling instance</summary>
    Public ReadOnly Property Tables As IEnumerable(Of DataTable)
        Get
            'If Not _populated Then Populate()  '"Tables Property : if Not populated"
            Return _tables
        End Get
    End Property

    Private _queries As IEnumerable(Of DataTable) = From t In New DataTable() {}

    ''' <summary>a list of all Queries Names in the calling instance</summary>
    Public ReadOnly Property Queries As IEnumerable(Of DataTable)
        Get
            'If Not _populated Then populate()  '"Queries Property : if Not populated"
            Return _queries
        End Get
    End Property

    Private _tableCount As Integer
    Public ReadOnly Property TableCount as Integer
    Get
        Return _tableCount
    End Get
    End Property

    Private _queryCount As Integer
    Public ReadOnly Property QueryCount as Integer
        Get
            Return _queryCount
        End Get
    End Property

    Private _allNames As IEnumerable(Of String)
    ''' <summary>
    '''     a list of all Objects Names (Tables/Queries/Columns) in the calling instance
    '''     this is gonna be used for faster DB Objects queries ,
    '''     so it doesn't iterate Tables or Queries to get Columns Names
    ''' </summary>
    Private ReadOnly Property AllNames As IEnumerable(Of String)
        Get
            Return _allNames
        End Get
    End Property

    ''' <summary>
    '''     gets a list of DB Objects Autocomplete words for a given word
    ''' </summary>
    ''' <param name="word">the words to search for Autocomplete</param>
    Public ReadOnly Property GetAutoComplete(word As String) As IEnumerable(Of String)
        Get
            If Path.Trim <> "" Then 'if DB is not Specified yet that means "DB Objects" are not available
                Return From n In AllNames _
                    Where N.StartsWith(Word, StringComparison.OrdinalIgnoreCase) _
                    Select "[" & N & "]"
            Else
                Return (From x In New String() {})

            End If
        End Get
    End Property

#End Region
    ''' <summary>
    '''     Class Constructor.
    ''' </summary>
    ''' <param name="dbPath">Path to Database File.</param>
    ''' <param name="dbPassword">Password to that database.</param>
    Public Sub New(dbPath As String, dbPassword As String)
        MyBase.New()
        Path = dbPath
        PassWord = dbPassword
    End Sub

    Public Async Function TestConnection() As Task(of Boolean)
        using dbConnection = New OleDbConnection(ConnectionString())
            Try
                Await dbConnection.OpenAsync()
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
                Return False
            End Try
        End Using

        Return True
    End Function

#Region "DataBase Schema"

    'Private _populated As Boolean = False

    ''' <summary>
    '''     fills the instance members with the info from the source
    ''' </summary>
    ''' <remarks></remarks>
    Private Async Function Populate() As Task
        _Tables = Await GetTables()
        _Queries = Await GetQueries()

        Try ' to join the data
            _allNames = From x In New String() {}

            Dim allObjects = _Tables.Union(_Queries) _
            'not we can't use Tables/Queries properties cuz populated Field is still False ....they will call Populate again 
            Dim tmp0 As New List(Of String)
            For Each T In AllObjects
                Dim dataColumns = T.Columns
                tmp0.Add(T.TableName)
                tmp0.AddRange(From thisColumn As DataColumn In dataColumns Select thisColumn.ColumnName)
            Next

            _allNames = From n In tmp0 Select n=N Distinct
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            ' ???
        End Try
    End Function

    ''' <summary>
    '''     A list of Queries in the calling instance (returned as DataTable Type)
    ''' </summary>
    Private Async Function GetQueries() As Task(Of IEnumerable(Of DataTable))
        Using dbConnection = New OleDbConnection(ConnectionString())
            Await dbConnection.OpenAsync()

            Dim procedureTable = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Procedures, New Object() {})
            Dim procedures = From row As DataRow In procedureTable.Rows _
                    Where Not Row.Item("PROCEDURE_NAME").StartsWith("~") _
                    Select GetDataTableSchema("[" & CType(Row.Item("PROCEDURE_NAME"), String) & "]")

            Dim viewsTable = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Views, New Object() {})
            Dim queryNames = From row As DataRow In viewsTable.Rows _
                    Select GetDataTableSchema("[" & CType(Row.Item("TABLE_NAME"), String) & "]")
        
            Dim queryList = procedures.Union(queryNames).ToList()
            _queryCount = queryList.Count()

            Return queryList
        End Using
    End Function

    ''' <summary> a list of Tables in the calling instance (returned as DataTable Type)</summary>
    Private Async Function GetTables() As Task(Of IEnumerable(Of DataTable))
        Using dbConnection = New OleDbConnection(ConnectionString())
            Await dbConnection.OpenAsync()

            Dim tablesInfo = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables_Info, New Object() {})

            Dim goodTablesTypes = New String() {"TABLE", "LINK", "PASS-THROUGH"} _
            'I'm afraid there are more ,I didn't try all types , but I'll try to

            Dim funGoodTableType = Function(tableType As String) _
                    GoodTablesTypes.Contains(tableType)

            'gets the TablesNames only
            Dim tablesNames = From row As DataRow In tablesInfo.Rows _
                    Where funGoodTableType(CType(Row.Item("TABLE_TYPE"), String)) _
                    Select "[" & CType(Row.Item("TABLE_NAME"), String) & "]" Distinct

            'Now Fetch the Schema of tables
            Dim tableSchema = From tblName As String In TablesNames _
                    Select GetDataTableSchema(TblName)
        
            Dim tableList = tableSchema.ToList()
            _tableCount = tableList.Count()

            Return tableList
        End Using
    End Function

    ''' <summary>
    '''     re-fill instance members with a Info from the source
    ''' </summary>
    Public Async Function RefreshSchema() As Task
        Await Populate()  '"RefreshSchema"        
    End Function

#End Region

#Region "Dealing with the Data itself"

    ''' <summary>Executes a given sql statement and returns the count of the affected records</summary>
    ''' <param name="commandQuery">the SqlQuery statement to execute</param>
    ''' <returns>count of the affected records</returns>
    Public Async Function ExecuteSqlCommand(commandQuery As SqlQuery) As Task(of Long)
        Dim count As Long
        Using dbConnection = New OleDbConnection(ConnectionString())
            Using dbCommand = New OleDbCommand(commandQuery.Sql, dbConnection)
                If commandQuery.QueryParams IsNot Nothing AndAlso commandQuery.QueryParams.Count > 0 Then
                    dbCommand.Parameters.Clear()
                    For Each p In commandQuery.QueryParams
                        dbCommand.Parameters.AddWithValue(p.Key, p.Value)
                    Next
                End If

                Try 'to run the command
                    count = Await dbCommand.ExecuteNonQueryAsync()
                Catch ex As Exception
                    Editor.ShowResults2MsgTab(ex.Message & Environment.NewLine)
                Finally
                    ' ???
                End Try

                Return count
            End Using
        End Using
    End Function

    ''' <summary>Returns a DataTable instance that holds the returned data of a specified Sql statement</summary>
    ''' <param name="thisQuery">SqlQuery statement to execute</param>
    ''' <param name="tableName">a value to be set to the name of the returned table</param>
    Public Async Function GetDataTable(thisQuery As SqlQuery, Optional ByVal tableName As String = "Table") As Task(Of DataTable)
        Using dbConnection = New OleDbConnection(ConnectionString())
            Using dbCommand = New OleDbCommand(thisQuery.Sql, dbConnection)
                Using dataAdapter = New OleDbDataAdapter
                    With dataAdapter
                        .SelectCommand = dbCommand
                        .ContinueUpdateOnError = True
                        .AcceptChangesDuringFill = True
                        .AcceptChangesDuringUpdate = True
                        .ResetFillLoadOption()
                        '.GetFillParameters()
                    End With

                    dataAdapter.SelectCommand = dbCommand
                    Dim currentDataSet As New DataSet
                    Dim resultTable As New DataTable(TableName)

                    ' What is happening here? cmd is never used after it is modified?
                    If thisQuery.QueryParams IsNot Nothing AndAlso thisQuery.QueryParams.Count > 0 Then
                        dbCommand.Parameters.Clear()
                        For Each p In thisQuery.QueryParams
                            dbCommand.Parameters.AddWithValue(p.Key, p.Value)
                        Next
                    End If

                    Try ' to get the data
                        Await Task.Run(Sub() dataAdapter.Fill(currentDataSet, TableName))
                    Catch ex As OleDbException
                        Editor.ShowResults2MsgTab(ex.Message)
                    End Try

                    If currentDataSet.Tables.Count > 0 Then
                        resultTable = currentDataSet.Tables(0)
                        resultTable.TableName = TableName
                    End If

                    Return resultTable
                End Using
            End Using
        End Using
    End Function

    ''' <summary>Returns a DataTable instance that holds no data, it has only the schema of the specified existed Table Name</summary>
    ''' <param name="tableName">the name of the table to return the schema of</param>
    Private Function GetDataTableSchema(tableName As String) As DataTable
        Using dbConnection = New OleDbConnection(ConnectionString())
            Using dbCommand = New OleDbCommand("SELECT * FROM " & tableName, dbConnection)
                Using dataAdapter = New OleDbDataAdapter
                    With dataAdapter
                        .SelectCommand = dbCommand
                        .ContinueUpdateOnError = True
                        .AcceptChangesDuringFill = True
                        .AcceptChangesDuringUpdate = True
                        .ResetFillLoadOption()
                        '.GetFillParameters()
                    End With

                    Dim theDataSet As New DataSet                    
                    Try ' to get the data
                        'dataAdapter.Fill(theDataSet, tableName)
                        dataAdapter.FillSchema(theDataSet, SchemaType.Source)  ' Produces memory exception?
                    Catch ex As Exception
                        Debug.Print(ex.Message)
                    End Try

                    Dim firstTable As New DataTable()
                    If theDataSet.Tables.Count > 0 Then
                        firstTable = theDataSet.Tables(0)
                    End If

                    firstTable.TableName = Regex.Replace(tableName, "\[|\]", "")

                    Return firstTable
                End Using
            End Using
        End Using
    End Function

#End Region

#Region "Overrided Subs"

    'Protected Overrides Sub Finalize()
    '    If _connection IsNot Nothing Then
    '        If CBool(_connection.State And ConnectionState.Open) Then
    '            _connection.Close()
    '        End If

    '        _connection = Nothing
    '    End If

    '    MyBase.Finalize()
    'End Sub

#End Region
End Class
