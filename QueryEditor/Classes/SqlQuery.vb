
Imports System.Text.RegularExpressions
Imports QueryEditor.ColoringWords

Namespace Queries
    <Flags()> _
    Public Enum SqlQueryType
        ParameterizedSqlQuery = 1
        SelectSqlQuery = 2
        NoneQuerySqlQuery = 4
        'we'll see the others later
    End Enum

    Public Class SqlQuery
        Private _AllSql As String = ""
        Private _NoCommentsSql As String = ""

        ''' <summary> Get or Sets the sql statement text </summary>
        ''' <param name="Commented">determine if the returned value will be with comments or not</param> 
        Public Property Sql(Optional ByVal Commented As Boolean = False) As String
            Get
                If Commented Then
                    Return _AllSql
                Else
                    Return _NoCommentsSql
                End If
            End Get

            Set(ByVal value As String)
                If _AllSql <> value Then
                    _AllSql = value.Trim

                    'next is to remove the commented statements
                    _NoCommentsSql = Regex.Replace(_AllSql, _
                                       ColoredWords.GetCommentsPattern, _
                                        " ", _
                                        RegexOptions.IgnoreCase Or RegexOptions.IgnorePatternWhitespace).Trim


                End If
            End Set
        End Property

        Public Sub New(ByVal sql As String)
            Me.Sql = sql
        End Sub

        ''' <summary>Gets a value indicating the type of the SQL statement.
        ''' it returns one of the following values : 
        ''' a ( SELECT ) statement or a ( NonQuery ) Statement 
        ''' or a( Parameterized SELECT ) Statement or a ( Parameterized NonQuery ) Statement. 
        ''' </summary>
        Public ReadOnly Property QueryType() As SqlQueryType
            Get
                Dim SearchOption = RegexOptions.IgnorePatternWhitespace Or RegexOptions.IgnoreCase

                If Regex.IsMatch(_NoCommentsSql, "^SELECT|TRANSFORM", SearchOption) Then 'it's a SELECT
                    Return SqlQueryType.SelectSqlQuery


                ElseIf Regex.IsMatch(_NoCommentsSql, "^PARAMETERS", SearchOption) Then 'it's a Parameterized

                    'next is to see what is after the parameters 
                    'is it a SELECT or a NoneQuerySqlQuery
                    Dim WithNoParameters = Regex.Replace(_NoCommentsSql, _
                                                         "^PARAMETERS[^'""]*?;", "", _
                                                         SearchOption).Trim


                    If Regex.IsMatch(WithNoParameters, "^SELECT|TRANSFORM", SearchOption) Then
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
        Private _QueryParams As New Generic.Dictionary(Of String, String)

        ''' <summary>Gets or Sets the query parameters that will be used for Parameterized queries</summary>
        Public Property QueryParams() As Generic.Dictionary(Of String, String)
            Get
                Return _QueryParams
            End Get
            Set(ByVal value As Generic.Dictionary(Of String, String))
                If _QueryParams IsNot value Then
                    _QueryParams.Clear()
                    _QueryParams = Nothing
                    _QueryParams = value
                End If
            End Set
        End Property

        ''' <summary>
        ''' gets a list of the available parameters that start with a given word
        ''' </summary>
        ''' <param name="word">the word to get a [Parameters Auto Complete list] for</param>
        Public Function GetParamsAutoCompleteList(ByVal word As String) As IEnumerable(Of String)
            Dim Rslt As IEnumerable(Of String)
            Try

                Rslt = From P In QueryParams _
                       Where P.Key.StartsWith(word, StringComparison.OrdinalIgnoreCase) _
                       Select "[" & P.Key & "]"

            Catch ex As Exception
                Rslt = From x In New String() {} ' to not return a Nothing
            End Try

            Return Rslt

        End Function

#End Region

    End Class

End Namespace
