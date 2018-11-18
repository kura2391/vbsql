
Public Class Update
    Inherits SqlAbstract

    Private _where As New Where()

    'save set sentences 
    Private _set As String

    'save parameter
    Private _params As Parameters

    Public Sub New(connection As Connection)
        MyBase.New(connection)

        _set = ""
        _params = New Parameters("@set")
    End Sub
    Public Sub New(connectionString As String)
        MyBase.New(connectionString)

        _set = ""
        _params = New Parameters("@set")
    End Sub


    Public Function table(tableName As String) As Update
        _table = tableName
        Return Me
    End Function


    Public Function [set](ht As Hashtable) As Update
        If _set <> "" Then
            _set &= ","
        End If
        For Each col As String In ht.Keys
            If IsDBNull(ht(col)) Then
                _set &= col & "=NULL,"
            Else
                _params.add(ht(col))
                _set &= col & "=" & _params.getLatestParameterName & ","
            End If

        Next

        _set = _set.Substring(0, _set.Length - 1)
        Return Me
    End Function


    Public Function [set](columnName As String, value As Object) As Update
        If _set <> "" Then
            _set &= ","
        End If

        If IsDBNull(value) Then
            _set &= columnName & "=" & value
        Else
            _params.add(value)
            _set &= columnName & "=" & _params.getLatestParameterName
        End If

        Return Me
    End Function

    'set where
    Public Function where(conditions As String, Optional ByVal col() As String = Nothing) As Update
        _where.add(conditions, param:=col)
        Return Me
    End Function


    'execute sql and return integer 
    Public Function execute() As Integer
        Return _connection.execute(buildSql(), buildParameter())
    End Function

    '
    Function buildSql() As String
        Dim sql As String = "UPDATE " & _table & " SET "
        sql &= _set

        If Not _where.isEmpty() Then
            sql &= " WHERE " & _where.getSql()
        End If
        Return sql
    End Function

    Private Function buildParameter() As SqlClient.SqlParameter()
        Dim p As New List(Of SqlClient.SqlParameter)
        p.AddRange(_params.getParamsArray)
        p.AddRange(_where.getParamList)

        Return p.ToArray()
    End Function

End Class


