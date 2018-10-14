Public Class Parameter
    Private _param As List(Of SqlClient.SqlParameter)
    Private _pre As String

    Public Sub New(prefix As String)
        _pre = prefix
        _param = New List(Of SqlClient.SqlParameter)

    End Sub
    Public Function appendParameter(value As String, Optional type As SqlDbType = SqlDbType.NVarChar) As SqlClient.SqlParameter
        Dim p As New SqlClient.SqlParameter(_pre & _param.Count.ToString(), dbType:=type)
        p.Value = value
        Return p
    End Function

    Public Function getParameterList() As List(Of SqlClient.SqlParameter)
        Return _param
    End Function

    Public Function getParameter(index As Integer) As SqlClient.SqlParameter
        If _param.Count < index Then
            Throw New Exception("index overflow")
        End If
        Return _param(index)
    End Function

    Public Sub clear()
        _param.Clear()
    End Sub

    Public Function count() As Integer
        Return _param.Count()
    End Function
End Class
