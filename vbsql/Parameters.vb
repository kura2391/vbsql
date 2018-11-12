Public Class Parameters
    Private _params As List(Of SqlClient.SqlParameter)

    Private _prefix As String

    Public Sub New(prefix As String)
        _prefix = prefix
        _params = New List(Of SqlClient.SqlParameter)
    End Sub

    Public Sub add(value As String, Optional type As SqlDbType = SqlDbType.NVarChar)

        Dim p As New SqlClient.SqlParameter()
        p.Value = value
        p.ParameterName = _prefix & _params.Count.ToString()
        p.SqlDbType = type

        _params.Add(p)
    End Sub

    Public Function getLatestParameterName() As String
        Return _params(_params.Count - 1).ParameterName
    End Function


    Public Function getParamsArray() As SqlClient.SqlParameter()
        Return _params.ToArray()
    End Function

End Class
