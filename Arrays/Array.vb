Public Class Array
    ' Private fields
    Private ReadOnly maxIndex As Integer
    Private _items As DataValue()
    Private lastValIndex As Integer

    ' Public property
    Public ReadOnly Property Items As DataValue()
        Get
            ' Returns shallow copy - should probably return deep copy but that would be expensive
            Return _items.Clone
        End Get
    End Property

    ' Constructor
    Public Sub New(maxIndex As Integer)
        Me.maxIndex = maxIndex

        ' Initialise array
        _items = New DataValue(maxIndex) {}
        EmptyArray()
    End Sub

    Public Sub EmptyArray()
        For i As Integer = 0 To maxIndex
            If _items(i) IsNot Nothing Then
                _items(i).Value = String.Empty
            Else
                _items(i) = New DataValue(String.Empty)
            End If
        Next
        lastValIndex = -1
    End Sub

    Public Sub Insert(newItem As String, inputPos As Integer)
        Dim curMaxIndex As Integer
        If lastValIndex < maxIndex Then
            If Not String.IsNullOrWhiteSpace(newItem) Then
                curMaxIndex = Math.Min(maxIndex, lastValIndex + 1)
                If inputPos >= 0 And inputPos <= curMaxIndex Then
                    ' Insert anywhere between 0 and the space after the current highest index
                    For i As Integer = lastValIndex + 1 To inputPos + 1 Step -1
                        ' Don't assign the object - just the value
                        _items(i).Value = _items(i - 1).Value
                    Next
                    ' Insert value once existing values have shuffled up
                    _items(inputPos).Value = newItem
                    lastValIndex = lastValIndex + 1
                Else
                    Throw New ArgumentException(String.Format(My.Resources.InvalidIndexError, Str(curMaxIndex)))
                End If
            End If
        Else
            Throw New Exception(My.Resources.ArrayOverflowError)
        End If
    End Sub

    Public Sub Edit(searchItem As String, newItem As String)
        Dim pos As Integer
        If Not String.IsNullOrWhiteSpace(searchItem) And Not String.IsNullOrWhiteSpace(newItem) Then
            pos = Search(searchItem)
            If pos = -1 Then
                Throw New ArgumentException(My.Resources.DataNonExistentError)
            Else
                _items(pos).Value = newItem
            End If
        End If
    End Sub

    Public Sub Delete(searchitem As String)
        Dim pos As Integer
        If lastValIndex > -1 Then
            If Not String.IsNullOrWhiteSpace(searchitem) Then
                pos = Search(searchitem)
                If pos = -1 Then
                    Throw New ArgumentException(My.Resources.DataNonExistentError)
                Else
                    For i As Integer = pos To lastValIndex - 1
                        ' Don't assign the object - just the value
                        _items(i).Value = _items(i + 1).Value
                    Next
                    ' Set last item to blank - can't set this in the loop since
                    ' if array was full there would be no next element to copy
                    ' and this would cause an index out of bounds error
                    _items(lastValIndex).Value = String.Empty
                    lastValIndex = lastValIndex - 1
                End If
            End If
        Else
            Throw New Exception(My.Resources.ArrayUnderflowError)
        End If
    End Sub

    Public Sub Append(newItem As String)
        If lastValIndex < maxIndex Then
            If Not String.IsNullOrWhiteSpace(newItem) Then
                lastValIndex = lastValIndex + 1
                _items(lastValIndex).Value = newItem
            End If
        Else
            Throw New Exception(My.Resources.ArrayOverflowError)
        End If
    End Sub

    Public Function Search(searchItem As String) As Integer
        Dim found As Boolean = False
        Dim pos As Integer = 0
        Do While Not found And pos <= lastValIndex
            If _items(pos).Value = searchItem Then
                found = True
            Else
                pos = pos + 1
            End If
        Loop
        If Not found Then
            pos = -1
        End If
        Search = pos
    End Function
End Class
