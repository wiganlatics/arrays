''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright (c) 2017 Matthew Wright.                                                       '
' Licensed under MIT License. See LICENSE.txt for further details.                         '
'                                                                                          '
' This software should be distributed with a LICENSE.TXT file in the solution root.        '
' Alternatively  you can find a copy of the license in the github repository:              '
' https://github.com/wiganlatics/arrays.                                                   '
' The MIT License text is also available at: https://choosealicense.com/licenses/mit/.     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Class frmArrays
    Const maxIndex As Integer = 9
    Dim array As Array = New Array(maxIndex)

    Private Sub frmArrays_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Reset()
    End Sub

    Private Sub ResetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetToolStripMenuItem.Click
        Reset()
    End Sub

    Private Sub Reset()
        array.EmptyArray()
        SetupArray()
        DisplayArray()
    End Sub

    Private Sub SetupArray()
        array.Append("D")
        array.Append("K")
        array.Append("Q")
        array.Append("V")
    End Sub

    Private Sub DisplayArray()
        dgvView.DataSource = array.Items
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
        Try
            newItem = InputBox(My.Resources.EnterNewDataPrompt)
            If Not String.IsNullOrWhiteSpace(newItem) Then
                ' User pressed OK and string is not empty or white space
                inputPos = InputBox(My.Resources.InsertLocationPrompt)
                If Not String.IsNullOrWhiteSpace(inputPos) Then
                    If Integer.TryParse(inputPos, pos) Then
                        ' User pressed OK and entered a valid number
                        array.Insert(newItem, pos)
                        DisplayArray()
                    Else
                        MsgBox(My.Resources.InvalidNumberError)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Edit()
    End Sub

    Private Sub EditToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditToolStripMenuItem.Click
        Edit()
    End Sub

    Private Sub Edit()
        Dim searchItem As String
        Dim newItem As String
        Try
            searchItem = InputBox(My.Resources.ChooseDataToEditPrompt)
            If Not String.IsNullOrWhiteSpace(searchItem) Then
                ' User pressed OK and string is not empty or white space
                newItem = InputBox(My.Resources.EnterNewDataPrompt)
                array.Edit(searchItem, newItem)
                DisplayArray()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Delete()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        Delete()
    End Sub

    Private Sub Delete()
        Try
            array.Delete(InputBox(My.Resources.ChooseDateToDeletePrompt))
            DisplayArray()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnAppend_Click(sender As Object, e As EventArgs) Handles btnAppend.Click
        Append()
    End Sub

    Private Sub AppendToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AppendToolStripMenuItem.Click
        Append()
    End Sub

    Private Sub Append()
        Try
            array.Append(InputBox(My.Resources.EnterNewDataPrompt))
            DisplayArray()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
End Class
