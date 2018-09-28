Imports System.Text.RegularExpressions
Imports QueryEditor.ColoringWords

Namespace Queries
    <Flags>
    Public Enum SqlQueryType
        ParameterizedSqlQuery = 1
        SelectSqlQuery = 2
        NoneQuerySqlQuery = 4
        'we'll see the others later
    End Enum

    Public Class SqlQuery
        Private _allSql As String = ""
        Private _noCommentsSql As String = ""

        ''' <summary> Get or Sets the sql statement text </summary>
        ''' <param name="commented">determine if the returned value will be with comments or not</param>
        Public Property Sql(Optional ByVal commented As Boolean = False) As String
            Get
                If Commented Then
                    Return _AllSql
                Else
                    Return _NoCommentsSql
                End If
            End Get

            Private Set
                If _AllSql <> value Then
                    _AllSql = value.Trim

                    'next is to remove the commented statements
                    _NoCommentsSql = Regex.Replace(_AllSql,
                                                   ColoredWords.GetCommentsPattern,
                                                   " ",
                                                   RegexOptions.IgnoreCase Or RegexOptions.IgnorePatternWhitespace).Trim
                End If
            End Set
        End Property

        Public Sub New(sql As String)
            Me.Sql = sql
        End Sub

        ''' <summary>
        '''     Gets a value indicating the type of the SQL statement.
        '''     it returns one of the following values :
        '''     a ( SELECT ) statement or a ( NonQuery ) Statement
        '''     or a( Parameterized SELECT ) Statement or a ( Parameterized NonQuery ) Statement.
        ''' </summary>
        Public ReadOnly Property QueryType As SqlQueryType
            Get
                Const searchOption As RegexOptions = RegexOptions.IgnorePatternWhitespace Or RegexOptions.IgnoreCase

                If Regex.IsMatch(_NoCommentsSql, "^SELECT|TRANSFORM", searchOption) Then 'it's a SELECT
                    Return SqlQueryType.SelectSqlQuery


                ElseIf Regex.IsMatch(_NoCommentsSql, "^PARAMETERS", SearchOption) Then 'it's a Parameterized

                    'next is to see what is after the parameters 
                    'is it a SELECT or a NoneQuerySqlQuery
                    Dim withNoParameters = Regex.Replace(_NoCommentsSql,
                                                         "^PARAMETERS[^'""]*?;", "",
                                                         searchOption).Trim


                    If Regex.IsMatch(WithNoParameters, "^SELECT|TRANSFORM", searchOption) Then
                        Return SqlQueryType.ParameterizedSqlQuery Or SqlQueryType.SelectSqlQuery
                    Else
                        Return SqlQueryType.ParameterizedSqlQuery Or SqlQueryType.NoneQuerySqlQuery
                    End If
                Else 'it's not a Parameterized neither a SELECT then it's a NoneQuerySqlQuery
                    Return SqlQueryType.NoneQuerySqlQuery
                End If
            End Get
        End Property

#Region "Dealing with the Parameters"

        ''' <summary>the internal object of the property QueryParams</summary>
        Private _queryParams As New Dictionary(Of String, String)

        ''' <summary>Gets or Sets the query parameters that will be used for Parameterized queries</summary>
        Public Property QueryParams As Dictionary(Of String, String)
            Get
                Return _QueryParams
            End Get

            Set
                If _QueryParams IsNot value Then
                    _QueryParams.Clear()
                    _QueryParams = Nothing
                    _QueryParams = value
                End If
            End Set
        End Property

        ''' <summary>
        '''     gets a list of the available parameters that start with a given word
        ''' </summary>
        ''' <param name="word">the word to get a [Parameters Auto Complete list] for</param>
        Public Function GetParamsAutoCompleteList(word As String) As IEnumerable(Of String)
            Dim enumerableResult As IEnumerable(Of String)
            Try
                enumerableResult = From p In QueryParams _
                    Where P.Key.StartsWith(word, StringComparison.OrdinalIgnoreCase) _
                    Select "[" & P.Key & "]"

            Catch ex As Exception
                enumerableResult = From x In New String() {} ' to not return a Nothing
            End Try

            Return enumerableResult
        End Function

#End Region
    End Class
End Namespace
