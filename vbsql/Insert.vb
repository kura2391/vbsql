
Public Class Insert
    Inherits SqlAbstract

    'save lastinsertid
    Private _lastInsertId As Boolean = False

    'save string of value and columns
    Private _columnstr As String = Nothing
    Private _valuestr As String = Nothing

    Private _params As Parameters

    Public Sub New(connection As Connection)
        MyBase.New(connection)
        _columnstr = ""
        _valuestr = ""

        _params = New Parameters("@insert")
    End Sub

    Public Sub New(connectionString As String)
        MyBase.New(connectionString)
        _columnstr = ""
        _valuestr = ""
        _params = New Parameters("@insert")
    End Sub


    'set table
    Public Function into(table As String) As Insert
        _table = table
        Return Me
    End Function

    ''set insertdata
    Public Function values(ht As Hashtable) As Insert
        If _columnstr <> "" Then
            _columnstr &= ","
            _valuestr &= ","
        End If
        For Each col As String In ht.Keys

            _columnstr &= col & ","

            If IsDBNull(ht(col)) Then
                _valuestr &= "NULL,"
            Else
                _params.add(ht(col))
                _valuestr &= _params.getLatestParameterName & ","
            End If

        Next

        _valuestr = _valuestr.Substring(0, _valuestr.Length - 1)
        _columnstr = _columnstr.Substring(0, _columnstr.Length - 1)
        Return Me
    End Function

    'set lastInsertId
    Public Function lastInsertId() As Insert
        _lastInsertId = True
        Return Me
    End Function


    'execute sql and return integer 
    Public Function execute() As Integer
        'check()

        Return _connection.execute(buildSql(), buildParameter())
    End Function

    'create sql
    Private Function buildSql()

        Dim sql As String = "INSERT INTO " & _table & " ("

        sql &= _columnstr & ") VALUES ("
        sql &= _valuestr & ")"

        If _lastInsertId Then
            sql &= "SELECT SCOPE_IDENTITY();"
        End If
        Return sql
    End Function

    Public Function buildParameter() As SqlClient.SqlParameter()
        Return _params.getParamsArray
    End Function



End Class
