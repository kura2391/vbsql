Public Class SqlAbstract

    'save parameter
    'save table parameter
    Protected _table As SqlClient.SqlParameter

    'save connection
    Protected _connection As Connection = Nothing


    Public Sub New(connection As Connection)
        _connection = connection
        _table = New SqlClient.SqlParameter
        _table.ParameterName = "table"
    End Sub

    Public Sub New(connectionString As String)
        _connection = New Connection(connectionString)
        _table = New SqlClient.SqlParameter
        _table.ParameterName = "table"
    End Sub



    'check connection class
    Protected Sub checkConnection()
        If IsNothing(_connection) Then
            Throw New Exception("no Connection set")
        End If
    End Sub
End Class
