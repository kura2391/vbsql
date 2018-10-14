Public Class Where
    Private _param As Parameter = Nothing
    Private _string As String = Nothing

    Public Sub New(prefix As String)
        _param = New Parameter(prefix)
    End Sub



    Public Sub [set](conditions As String, col() As String)
        clear()
        Dim i As Integer = 0
        Dim j As Integer = 0
        While conditions.IndexOf("?", i) <> -1
            append(conditions.Substring(i, conditions.IndexOf("?", i)), col(j))
            i = conditions.IndexOf("?") + 1
            j += 1
        End While
        _string &= conditions.Substring(i)
    End Sub


    Public Function sql() As String
        Return _string
    End Function

    Public Function getParameterList() As List(Of SqlClient.SqlParameter)
        Return _param.getParameterList
    End Function


    Public Sub append(conditions As String, col As String)
        _param.appendParameter(col)
        _string &= Replace(conditions, "?", "@" & _param.getParameter(_param.count - 1).ParameterName)
    End Sub

    Public Sub clear()
        _param.clear()
        _string = ""
    End Sub


    Public Function isEmpty() As Boolean
        If _string = "" Then
            Return True
        End If
        Return False
    End Function
End Class
