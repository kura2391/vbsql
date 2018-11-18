
'set for Join method
Public Enum Jointype As Integer
    INNER
    LEFT
    RIGHT
    FULL
End Enum

Public Class [Select]
    Inherits SqlAbstract
    Private _select As String = ""
    Private _where As New Where()
    Private _order As String = ""
    Private _join As New List(Of String)
    Private _groupBy As String = ""

    Public Sub New(connection As Connection)
        MyBase.New(connection)
    End Sub

    Public Sub New(connectionString As String)
        MyBase.New(connectionString)
    End Sub

    'set table
    Public Function from(table As String) As [Select]
        _table = table
        Return Me
    End Function
    'set Columns
    Public Function [select](Optional ByVal columns As String() = Nothing) As [Select]
        If _select <> "" Then
            _select &= ","
        End If

        For i As Integer = 0 To columns.Length - 1
            _select &= columns(i) & ","
        Next
        _select = _select.Substring(0, _select.Length - 1)
        Return Me
    End Function

    'set Where

    Public Function where(conditions As String, Optional ByVal col() As String = Nothing) As [Select]
        _where.add(conditions, param:=col)
        Return Me
    End Function

    'set OrderBy
    Public Function orderBy(order As String) As [Select]
        _order = order
        Return Me
    End Function


    'set join
    Public Function join(table As String, conditions As String, Optional type As Jointype = Jointype.INNER) As [Select]
        Dim str As String = " "
        Select Case type
            Case Jointype.INNER
                str = " JOIN "
            Case Jointype.LEFT
                str = " LEFT JOIN "
            Case Jointype.RIGHT
                str = " RIGHT JOIN "
            Case Jointype.FULL
                str = " FULL JOIN "
        End Select
        str &= table & " ON " & conditions
        Me._join.Add(str)
        Return Me
    End Function

    'set groupby
    Public Function groupBy(group As String) As [Select]
        _groupBy &= group
        Return Me
    End Function


    'execute sql and get data as DataTable
    Public Function execute() As DataTable
        Return _connection.executeSelect(buildSql, buildParameter())
    End Function

    'create sql sentence 
    Private Function buildSql() As String
        Dim sql As String = "SELECT "

        sql &= _select

        sql &= " FROM " & _table

        sql &= ""
        If _join.Count <> 0 Then
            For Each s As String In _join
                sql &= s
            Next
        End If
        If Not _where.isEmpty() Then
            sql &= " WHERE "
            sql &= _where.getSql()
        End If
        If _groupBy <> "" Then
            sql &= " GROUP BY "
            sql &= _groupBy
        End If
        If _order <> "" Then
            sql &= " ORDER BY "
            sql &= _order
        End If
        Return sql
    End Function

    Private Function buildParameter() As SqlClient.SqlParameter()
        Dim p As New List(Of SqlClient.SqlParameter)

        p.AddRange(_where.getParamList())

        Return p.ToArray()
    End Function

    'if _from is not set, throw error



End Class