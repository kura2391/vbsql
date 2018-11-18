Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Vbsql

<TestClass()> Public Class WhereTest

    <TestMethod()> Public Sub getSqlTest()
        Dim where As New Where()
        Dim po As New PrivateObject(where)
        where.add("id=? AND date=?", {990, "2018-10-10"})

        Dim ans As String = po.Invoke("getSql")

        Dim actual As String = "id=@where0 AND date=@where1"
        Assert.AreEqual(actual, ans)
    End Sub

End Class