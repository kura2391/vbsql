Public Class Connection

    Private Property _connectionString As String

    Public Sub New(server As String, userId As String, password As String, initialCatalog As String)
        _connectionString = ""
        _connectionString &= "Server=" & server & ";"
        _connectionString &= "User ID=" & userId & ";"
        _connectionString &= "Password=" & password & ";"
        _connectionString &= "Initial Catalog=" & initialCatalog



    End Sub

    Public Sub New(connectionString As String)
        _connectionString = connectionString
    End Sub

End Class
