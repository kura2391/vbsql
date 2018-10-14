Module ConnectionStringModule
    Public Function createConnectionString(server As String, userId As String, password As String, initialCatalog As String) As String
        Dim _connectionString
        _connectionString = ""
        _connectionString &= "Server=" & server & ";"
        _connectionString &= "User ID=" & userId & ";"
        _connectionString &= "Password=" & password & ";"
        _connectionString &= "Initial Catalog=" & initialCatalog
        Return _connectionString
    End Function
End Module
