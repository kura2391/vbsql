
Public Class Parameter
    'columnName
    Private _col As String = Nothing
    'value and type
    Private _sqlParameter As SqlClient.SqlParameter = Nothing
    'for where sentence
    Public Sub New(value As String, Optional type As SqlDbType = SqlDbType.NVarChar)
        [set](value, type)
    End Sub
    'for update,insert sentence
    Public Sub New(columnName As String, value As String, Optional type As SqlDbType = SqlDbType.NVarChar)
        _col = columnName
        [set](value, type)
    End Sub
    'for just instance
    Public Sub New()
    End Sub

    'set value and type at where sentence
    Public Sub [set](value As String, Optional type As SqlDbType = SqlDbType.NVarChar)
        _sqlParameter = New SqlClient.SqlParameter
        _sqlParameter.Value = value
        _sqlParameter.SqlDbType = type
    End Sub

    'set columnName , value and value type at update and select sentence
    Public Sub [set](columnName As String, value As String, Optional type As SqlDbType = SqlDbType.NVarChar)
        _col = columnName
        [set](value, type)
    End Sub

    'getter
    Friend Function getColumnName() As String
        Return _col
    End Function

    'setter
    Friend Sub setParameterName(paramname As String)
        _sqlParameter.ParameterName = paramname
    End Sub
    'getter
    Friend Function getParameterName() As String
        Return _sqlParameter.ParameterName
    End Function

    'getter
    Friend Function getSqlParameter()
        Return _sqlParameter
    End Function
End Class
