
Public Class Where
    'save parameter in where sentence
    'Private _param As List(Of SqlClient.SqlParameter) = Nothing


    'save sql string
    Private _sql As String = Nothing
    'save parameters
    Private _params As Parameters

    Public Sub New()
        _sql = ""

        _params = New Parameters("@where")
    End Sub

    'append parameter and append sql
    Public Sub add(conditions As String, Optional ByVal param As String() = Nothing)
        For Each p As String In param
            _params.add(p)
        Next

        _sql &= conditions

    End Sub


    'build and get sql
    Friend Function getSql() As String
        Dim array As SqlClient.SqlParameter() = _params.getParamsArray
        Dim ans As String = ""
        Dim find As Integer
        Dim start As Integer = 0
        For i As Integer = 0 To array.Count - 1
            find = _sql.IndexOf("?", start)
            ans &= _sql.Substring(start, find - start) & array(i).ParameterName
            start = find + 1
        Next
        If start < _sql.Length Then
            ans &= _sql.Substring(start)
        End If

        Return ans
    End Function

    Friend Function getParamList() As SqlClient.SqlParameter()
        Return _params.getParamsArray
    End Function



    Public Function isEmpty() As Boolean
        If _sql = "" Then
            Return True
        End If
        Return False
    End Function
End Class


