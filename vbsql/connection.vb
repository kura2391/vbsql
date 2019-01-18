Public Class Connection

    'save connectionstring
    Private _cnStr As String

    'connection and Command class
    Private _cn As SqlClient.SqlConnection
    Private _cmd As SqlClient.SqlCommand

    'save if transaction
    Private _tran As SqlClient.SqlTransaction
    Private _transaction As Boolean = False


    Public Sub New(connectionString As String)
        _cnStr = connectionString
    End Sub

    Public Sub New(server As String, userId As String, password As String, initialCatalog As String)
        Dim _connectionString As String
        _connectionString = ""
        _connectionString &= "Server=" & server & ";"
        _connectionString &= "User ID=" & userId & ";"
        _connectionString &= "Password=" & password & ";"
        _connectionString &= "Initial Catalog=" & initialCatalog
        _cnStr = _connectionString
    End Sub

    'begin transaction
    Public Sub beginTransaction()
        If _transaction Then
            Throw New Exception("vbsql error: transaction already started or not closed")
            Exit Sub
        End If

        _transaction = True

        _cn = New SqlClient.SqlConnection
        _cn.ConnectionString = _cnStr
        _cn.Open()
        _tran = _cn.BeginTransaction

    End Sub



    'execute delete insert update 
    Public Function execute(sql As String, values() As SqlClient.SqlParameter) As Integer
        If Not _transaction Then
            Return execNonTransaction(sql, values)
        End If

        Try
            'transaction process
            _cmd = _cn.CreateCommand
            _cmd.CommandText = sql
            _cmd.Parameters.AddRange(values)
            _cmd.Transaction = _tran
            execute = _cmd.ExecuteScalar()
            _cmd.Dispose()
        Catch ex As Exception
            Throw New Exception("vbsql error:" & ex.Message)
        End Try

    End Function

    'commit transaction data
    Public Sub commitTransaction()
        If Not _transaction Then
            Throw New Exception("vbsql error: not transaction")
        End If

        _tran.Commit()


    End Sub

    'rollback 
    Public Sub rollbackTransaction()
        If Not _transaction Then
            Throw New Exception("vbsql error: not transaction")
        End If
        If Not _tran Is Nothing Then
            _tran.Rollback()
        End If
    End Sub



    'close 
    Public Sub closeTransaction()
        If Not _transaction Then
            Throw New Exception("vbsql error: not transaction")
        End If

        If Not _cn.State = ConnectionState.Closed Then
            _cn.Close()
        End If
        _cn.Dispose()
        _tran.Dispose()

        _transaction = False

    End Sub



    'execute with no transaction
    Private Function execNonTransaction(sql As String, values() As SqlClient.SqlParameter) As Integer
        _cn = New SqlClient.SqlConnection
        _cmd = New SqlClient.SqlCommand
        _cn.ConnectionString = _cnStr
        _cmd = _cn.CreateCommand
        _cmd.CommandText = sql
        _cmd.Parameters.AddRange(values)
        _cn.Open()

        execNonTransaction = _cmd.ExecuteScalar()

        _cn.Close()
        _cmd.Dispose()
        _cn.Dispose()
    End Function






    'execute select 
    Public Function executeSelect(sql As String, values() As SqlClient.SqlParameter) As DataTable
        Dim dt As New DataTable
        Dim da = New SqlClient.SqlDataAdapter(sql, _cnStr)
        da.SelectCommand.Parameters.AddRange(values)
        da.Fill(dt)
        Return dt
    End Function


    'select
    Public Function [select](table As String, columns As String()) As [Select]

        Return New [Select](Me).select(columns).from(table)
    End Function

    'insert
    Public Function insert(table As String) As Insert
        Return New Insert(Me).into(table)
    End Function

    'update
    Public Function Update(table As String) As Update
        Return New Update(Me).table(table)
    End Function

    'delete
    Public Function Delete(table As String) As Delete
        Return New Delete(Me).from(table)
    End Function
End Class
