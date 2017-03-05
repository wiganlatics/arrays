Public Class frmArrays
    Const maxIndex As Integer = 9
    Dim letters(maxIndex) As DataValue
    Dim lastValIndex As Integer

    Private Sub frmArrays_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialise objects
        For i As Integer = 0 To maxIndex
            letters(i) = New DataValue(String.Empty)
        Next

        Reset()
    End Sub

    Private Sub ResetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetToolStripMenuItem.Click
        Reset()
    End Sub

    Private Sub Reset()
        EmptyArray()
        SetupArray()
        DisplayArray()
    End Sub

    Private Sub EmptyArray()
        For i As Integer = 0 To maxIndex
            letters(i).Value = String.Empty
        Next
        lastValIndex = -1
    End Sub

    Private Sub SetupArray()
        letters(0).Value = "D"
        letters(1).Value = "K"
        letters(2).Value = "Q"
        letters(3).Value = "V"
        lastValIndex = 3
    End Sub

    Private Sub DisplayArray()
        dgvView.DataSource = letters
    End Sub

    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click
        Insert()
    End Sub

    Private Sub InsertToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InsertToolStripMenuItem.Click
        Insert()
    End Sub

    Private Sub Insert()
        Dim inputPos As String
        Dim pos As Integer
        Dim newItem As String
        Dim curMaxIndex As Integer
        If lastValIndex < maxIndex Then
            newItem = InputBox(My.Resources.EnterNewDataPrompt)
            If Not String.IsNullOrWhiteSpace(newItem) Then
                ' User pressed OK and string is not empty or whitespace
                inputPos = InputBox(My.Resources.InsertLocationPrompt)
                If Not String.IsNullOrWhiteSpace(inputPos) Then
                    If Integer.TryParse(inputPos, pos) Then
                        ' User pressed OK and entered a valid number
                        curMaxIndex = Math.Min(maxIndex, lastValIndex + 1)
                        If pos >= 0 And pos <= curMaxIndex Then
                            ' Insert anywhere between 0 and the space after the current highest index
                            For i As Integer = lastValIndex + 1 To pos + 1 Step -1
                                ' Don't assign the object - just the value
                                letters(i).Value = letters(i - 1).Value
                            Next
                            ' Insert value once existing values have shuffled up
                            letters(pos).Value = newItem
                            lastValIndex = lastValIndex + 1
                            DisplayArray()
                        Else
                            MsgBox(String.Format(My.Resources.InvalidIndexError, Str(curMaxIndex)))
                        End If
                    Else
                        MsgBox(My.Resources.InvalidNumberError)
                    End If
                End If
            End If
        Else
            MsgBox(My.Resources.ArrayOverflowError)
        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Edit()
    End Sub

    Private Sub EditToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditToolStripMenuItem.Click
        Edit()
    End Sub

    Private Sub Edit()
        Dim searchitem As String
        Dim strInput As String
        Dim pos As Integer
        searchitem = InputBox(My.Resources.ChooseDataToEditPrompt)
        If Not String.IsNullOrWhiteSpace(searchitem) Then
            'User pressed OK and string is not empty or white space
            pos = Search(searchitem)
            If pos = -1 Then
                MsgBox(My.Resources.DataNonExistentError)
            Else
                strInput = InputBox(My.Resources.EnterNewDataPrompt)
                If Not String.IsNullOrWhiteSpace(strInput) Then
                    'User pressed OK and string is not empty or whitespace
                    letters(pos).Value = strInput
                    DisplayArray()
                End If
            End If
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Delete()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        Delete()
    End Sub

    Private Sub Delete()
        Dim searchitem As String
        Dim pos As Integer
        If lastValIndex > -1 Then
            searchitem = InputBox(My.Resources.ChooseDateToDeletePrompt)
            If Not String.IsNullOrWhiteSpace(searchitem) Then
                ' User pressed OK and string is not empty or white space
                pos = Search(searchitem)
                If pos = -1 Then
                    MsgBox(My.Resources.DataNonExistentError)
                Else
                    For i As Integer = pos To lastValIndex - 1
                        ' Don't assign the object - just the value
                        letters(i).Value = letters(i + 1).Value
                    Next
                    ' Set last item to blank - can't set this in the loop since
                    ' if array was full there would be no next element to copy
                    ' and this would cause an index out of bounds error
                    letters(lastValIndex).Value = String.Empty
                    lastValIndex = lastValIndex - 1
                    DisplayArray()
                End If
            End If
        Else
            MsgBox(My.Resources.ArrayUnderflowError)
        End If
    End Sub

    Private Sub btnAppend_Click(sender As Object, e As EventArgs) Handles btnAppend.Click
        Append()
    End Sub

    Private Sub AppendToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AppendToolStripMenuItem.Click
        Append()
    End Sub

    Private Sub Append()
        Dim strInput As String
        If lastValIndex < maxIndex Then
            strInput = InputBox(My.Resources.EnterNewDataPrompt)
            If Not String.IsNullOrWhiteSpace(strInput) Then
                'User pressed OK and string is not empty or whitespace
                lastValIndex = lastValIndex + 1
                letters(lastValIndex).Value = strInput
            End If
            DisplayArray()
        Else
            MsgBox(My.Resources.ArrayOverflowError)
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Public Function Search(searchitem As String) As Integer
        Dim found As Boolean = False
        Dim pos As Integer = 0
        Do While Not found And pos <= lastValIndex
            If letters(pos).Value = searchitem Then
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
