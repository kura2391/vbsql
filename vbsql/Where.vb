
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
        Dim ans As String = _sql

        For i As Integer = 0 To array.Length - 1
            ans = Replace(ans, "?", array(i).ParameterName, Count:=1)
        Next

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


