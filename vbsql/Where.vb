
Public Class Where
    'save parameter in where sentence
    Private _param As List(Of SqlClient.SqlParameter) = Nothing
    'save sql string
    Private _sql As String = Nothing
    'save prefix of ParameterName 
    Private _prefix As String = "@where"



    Public Sub New()
        _param = New List(Of SqlClient.SqlParameter)
        _sql = ""
    End Sub

    'append parameter and append sql
    Public Sub add(conditions As String, Optional ByVal param As Parameter() = Nothing)

        addParameter(param)

        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim qmark As Integer = conditions.IndexOf("?", i)
        While qmark <> -1
            _sql &= conditions.Substring(i, qmark - i) & "" & param(j).getParameterName
            i = conditions.IndexOf("?", qmark) + 1
            j += 1
            qmark = conditions.IndexOf("?", i)
        End While
        _sql &= conditions.Substring(i)
    End Sub

    'more easier add function
    Public Sub add(conditions As String, Optional ByVal param As String() = Nothing)
        Dim setParam(param.Count - 1) As Parameter
        For i As Integer = 0 To param.Count - 1
            setParam(i) = New Parameter(param(i))
        Next
        add(conditions, setParam)
    End Sub


    Private Sub addParameter(param() As Parameter)
        Dim count As Integer = _param.Count
        For i As Integer = count To count + param.Count - 1
            param(i - count).setParameterName(_prefix & i.ToString())
            _param.Add(param(i - count).getSqlParameter)
        Next

    End Sub



    Friend Function sql() As String
        Return _sql
    End Function

    Friend Function getParamList() As List(Of SqlClient.SqlParameter)
        Return _param
    End Function


    Private Sub clear()
        _param.Clear()
        _sql = ""
    End Sub


    Public Function isEmpty() As Boolean
        If _sql = "" Then
            Return True
        End If
        Return False
    End Function
End Class


