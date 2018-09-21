#Region "Notes"

#End Region

#Region "Imports"
Imports Microsoft.Win32
Imports QueryEditor.Queries
#End Region

Module Mod_General

    'Public Const RegKey = "Software\QueryEditor" 'this key wil be in the HKCU key , 
    '                                       could be used in case we wanted to make some kinda MRU list

    ''' <summary>
    ''' one of the variables that holds an Index in the "ImageList" of the "TrView_dbObjects" 
    ''' in the main form ....  Iput them as Public variables 
    ''' so I can change their values without affecting codes that use them
    ''' </summary>
    Public TextIcon, NumberIcon, DateIcon, BooleanIcon, BinarryArrayIcon, _
           DBIcon, TableIcon, QueryIcon, GroupIcon, FunctionIcon As Integer


#Region "Working with the database "
    ''' <summary>Represents the database we're conneting to and executing queries on</summary>
    Public DB As New DataBase("", "") ' this declaration won't open the database

    ''' <summary>
    '''to do some things that not related to the db layer 
    ''' (like I have a tight system analysation ........ db layer.... , lololololo
    ''' </summary>
    ''' <param name="_DataBaseName">the database name to connect to</param>
    Public Function SetConStr(ByVal _DataBaseName As String) As Boolean

        DB.Path = _DataBaseName
        Return True
    End Function


#End Region

    ''' <summary>Represents the Query we're working on</summary>
    Public Query As New SqlQuery("")

End Module