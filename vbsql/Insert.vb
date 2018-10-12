Public Class Insert
    Private _into As String = Nothing
    Private _columns As Hashtable = Nothing
    Private _lastInsertId As Boolean = False

    'set table
    Public Function into(table As String) As Insert
        _into = table
        Return Me
    End Function

    'set hashtable as insertdata
    Public Function values(hashtable As Hashtable) As Insert
        _columns = hashtable
        Return Me
    End Function

    'set lastInsertId
    Public Function lastInsertId() As Insert
        _lastInsertId = True
        Return Me
    End Function


    'execute sql and return integer 
    Public Function execute(connectionString As String) As Integer
        checkInto()
        checkColumns()

        Dim cn As New SqlClient.SqlConnection
        Dim sql As New SqlClient.SqlCommand
        cn.ConnectionString = connectionString
        sql = cn.CreateCommand
        sql.CommandText = getSql()
        cn.Open()

        execute = sql.ExecuteScalar()

        cn.Close()
        sql.Dispose()
        cn.Dispose()
    End Function

    'create sql
    Function getSql()
        Dim sql As String = " INSERT INTO "
        Dim values As String = ""

        sql &= _into & "("
        For Each key As String In _columns.Keys
            sql &= key & ","
            values &= "'" & _columns(key) & "',"
        Next
        sql = sql.Substring(0, sql.Length - 1)
        values = values.Substring(0, values.Length - 1)

        sql &= ") VALUES(" & values & ");"
        If _lastInsertId Then
            sql &= "SELECT SCOPE_IDENTITY();"
        End If
        Return sql
    End Function
    'insert into test(text,number,Date) values('number','99','2018-10-10')

    'if _from is not set, throw error
    Sub checkInto()
        If IsNothing(_into) OrElse _into.Trim() = "" Then
            Throw New Exception("no table selected")
        End If
    End Sub

    'check _columns
    Sub checkColumns()
        If IsNothing(_columns) OrElse _columns.Count = 0 Then
            Throw New Exception("no data set")
        End If
    End Sub


End Class
