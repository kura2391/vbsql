Public Class Insert
    Private _into As String = Nothing
    Private _columns As Hashtable = Nothing

    Public Function into(table As String) As Insert
        _into = table
        Return Me
    End Function

    Public Function [set](hashtable As Hashtable) As Insert
        _columns = hashtable
        Return Me
    End Function

    'Public Function execute(connectionString As String) As Integer
    '    checkInto()

    '    Return Nothing
    'End Function


    'Function getSql()
    '    Dim sql As String = " INSERT INTO "
    '    sql &= _into
    '    sql &= " VALUES("

    '    Return sql
    'End Function
    'insert into test(text,number,date) values('number','99','2018-10-10')

    'if _from is not set, throw error
    Sub checkInto()
        If IsNothing(_into) OrElse _into.Trim() = "" Then
            Throw New Exception("no table selected")
        End If
    End Sub
End Class
