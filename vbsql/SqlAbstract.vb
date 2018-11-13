
Public Class SqlAbstract

    'save table parameter
    Protected _table As String = Nothing
    'save connection
    Protected _connection As Connection = Nothing



    Public Sub New(connection As Connection)
        _connection = connection
    End Sub

    Public Sub New(connectionString As String)
        _connection = New Connection(connectionString)
    End Sub


End Class
