Namespace VbSql
    Public Class SqlAbstract

        'save parameter
        'save table parameter
        'Protected _table As SqlClient.SqlParameter
        Protected _table As String = Nothing
        'save connection
        Protected _connection As Connection = Nothing

        'save variables
        Protected _variables As List(Of Parameter) = Nothing

        'save prefix
        Protected _prefix As String = Nothing


        Public Sub New(connection As Connection)
            _connection = connection
            _variables = New List(Of Parameter)
        End Sub

        Public Sub New(connectionString As String)
            _connection = New Connection(connectionString)
            _variables = New List(Of Parameter)
        End Sub


        'set parameterName to List(Of Parameter)
        Protected Sub setParameterName()
            For i As Integer = 0 To _variables.Count - 1
                _variables(i).setParameterName(_prefix & i.ToString)
            Next
        End Sub

        'check connection class
        Protected Sub checkConnection()
            If IsNothing(_connection) Then
                Throw New Exception("no Connection set")
            End If
        End Sub
        Protected Sub checkVariables()
            If _variables.Count = 0 Then
                Throw New Exception("no data inserted")
            End If
        End Sub

        Protected Sub checkTable()
            If IsNothing(_table) Then
                Throw New Exception("no table set")
            End If
        End Sub
    End Class
End Namespace
