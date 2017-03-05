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
        dvgView.DataSource = letters
    End Sub

    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click
        Insert()
    End Sub

    Private Sub InsertToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InsertToolStripMenuItem.Click
        Insert()
    End Sub

    Private Sub Insert()

    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Edit()
    End Sub

    Private Sub EditToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditToolStripMenuItem.Click
        Edit()
    End Sub

    Private Sub Edit()

    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Delete()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        Delete()
    End Sub

    Private Sub Delete()

    End Sub

    Private Sub btnAppend_Click(sender As Object, e As EventArgs) Handles btnAppend.Click
        Append()
    End Sub

    Private Sub AppendToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AppendToolStripMenuItem.Click
        Append()
    End Sub

    Private Sub Append()

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Public Function Search(searchitem As String) As Integer
        Search = 0
    End Function
End Class
