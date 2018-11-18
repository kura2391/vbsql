
Public Class Insert
    Inherits SqlAbstract

    'save lastinsertid
    Private _lastInsertId As Boolean = False

    'save string of value and columns
    Private _columnstr As String = Nothing
    Private _valuestr As String = Nothing

    'if values(dt as datatable) is used, check flag to override the columnstr by value(ht as hashtable) 

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
    Public Function value(ht As Hashtable) As Insert
        _columnstr = ""
        _valuestr = ""
        _params.clear()

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

    Public Function value(columnName As String, data As Object) As Insert

        If _columnstr <> "" Then
            _columnstr &= ","
            _valuestr &= ","
        End If
        _columnstr &= columnName
        If IsDBNull(data) Then
            _valuestr &= "NULL"
        Else
            _params.add(data)
            _valuestr &= _params.getLatestParameterName
        End If

        Return Me
    End Function


    'values
    Public Function values(dt As DataTable) As Insert
        _columnstr = ""
        _valuestr = ""
        _params.clear()

        Dim cols(dt.Columns.Count) As String
        For i As Integer = 0 To dt.Columns.Count - 1
            _columnstr &= dt.Columns(i).ColumnName & ","
            cols(i) = dt.Columns(i).ColumnName
        Next

        For Each row As DataRow In dt.Rows
            For i As Integer = 0 To dt.Columns.Count - 1
                If IsDBNull(row(cols(i))) Then
                    _valuestr &= "NULL,"
                Else
                    _params.add(row(cols(i)))

                    _valuestr &= _params.getLatestParameterName & ","
                End If

            Next
            _valuestr = _valuestr.Substring(0, _valuestr.Length - 1)
            _valuestr &= "),("
        Next

        _valuestr = _valuestr.Substring(0, _valuestr.Length - 3)
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
