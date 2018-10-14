Imports System.Windows.Forms
Public Class ControllerToDb
    'save the correspondings (key = controller.name, value = db.columnName)
    Private _corresponds As Dictionary(Of String, String)

    'get the corresponds by using tab area
    Public Sub New(form As Windows.Forms.Form)
        _corresponds = New Dictionary(Of String, String)

        For Each c As Control In form.Controls
            If isAvailableType(c) Then
                _corresponds.Add(c.Name, c.Tag)
            End If
        Next

    End Sub


    Public Sub New()
        _corresponds = New Dictionary(Of String, String)
    End Sub



    'add collesponds manually
    Public Sub add(control As Control, columnName As String)
        If _corresponds.ContainsKey(control.Name) Then
            Throw New Exception(control.Name & " is already added.")
        End If
        If Not isAvailableType(control) Then
            Throw New Exception(control.GetType().ToString() & " is not available type")
        End If
        _corresponds.Add(control.Name, columnName)
    End Sub

    'rename columnName manually
    Public Sub update(control As Control, columnName As String)
        If Not _corresponds.ContainsKey(control.Name) Then
            Throw New Exception(control.Name & " is not registered")
        End If
        _corresponds(control.Name) = columnName
    End Sub

    'remove collesponds manually
    Public Sub remove(control As Control)
        If Not _corresponds.ContainsValue(control.Name) Then
            Throw New Exception(control.Name & " is not registered")
        End If
        _corresponds.Remove(control.Name)
    End Sub

    'return if its Type is available 
    Private Function isAvailableType(c As Control)
        Dim flag As Boolean = False
        If TypeOf c Is TextBox Then
            flag = True
        ElseIf TypeOf c Is RichTextBox Then
            flag = True
        ElseIf TypeOf c Is ComboBox Then
            flag = True
        ElseIf TypeOf c Is DateTimePicker Then
            flag = True
        ElseIf TypeOf c Is CheckBox Then
            flag = True
        End If
        Return flag
    End Function


End Class
