Public Class DataValue
    ' Private field
    Private _value As String

    ' Public property
    Public Property Value() As String
        Get
            Return _value
        End Get
        Set(value As String)
            _value = value
        End Set
    End Property

    ' Constructor
    Public Sub New(value As String)
        _value = value
    End Sub
End Class
