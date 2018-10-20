﻿Public Class Connection

    'save connectionstring
    Private _cnStr As String
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

    'execute delete insert update 
    Public Function execute(sql As String, values() As SqlClient.SqlParameter) As Integer
        Dim cn As New SqlClient.SqlConnection
        Dim cmd As New SqlClient.SqlCommand
        cn.ConnectionString = _cnStr
        cmd = cn.CreateCommand
        cmd.CommandText = sql
        cmd.Parameters.AddRange(values)
        cn.Open()

        execute = cmd.ExecuteScalar()

        cn.Close()
        cmd.Dispose()
        cn.Dispose()

    End Function


    'execute select 
    Public Function executeSelect(sql As String, values() As SqlClient.SqlParameter) As DataTable
        Dim dt As New DataTable
        Dim da = New SqlClient.SqlDataAdapter(sql, _cnStr)
        da.SelectCommand.Parameters.AddRange(values)
        da.Fill(dt)
        Return dt
    End Function
End Class
