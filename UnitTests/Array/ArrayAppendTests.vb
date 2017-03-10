Imports Arrays
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Tests the Append method of the Array class.
''' </summary>
<TestClass()> Public Class ArrayAppendTests
    ''' <summary>
    ''' Test that appending two distinct values to a new array succeeds.
    ''' </summary>
    <TestMethod()> Public Sub AppendSucceeds()
        Dim appendVal As String = "A"
        Dim appendVal2 As String = "B"
        AppendTestHelper(appendVal, appendVal2)
    End Sub

    ''' <summary>
    ''' Test that appending duplicate values to an array succeeds.
    ''' </summary>
    <TestMethod()> Public Sub AppendDuplicateSucceeds()
        Dim appendVal As String = "A"
        AppendTestHelper(appendVal, appendVal)
    End Sub

    ''' <summary>
    ''' Test that appending a blank value to an array has no impact.
    ''' Append a blank value in first space then ensure that it is overwritten when real data is appended.
    ''' </summary>
    <TestMethod()> Public Sub AppendBlankDataIgnored()
        Dim appendVal As String = "A"
        AppendBlankTestHelper(String.Empty)
        AppendBlankTestHelper(appendVal)
    End Sub

    ''' <summary>
    ''' Tests that appending to a full array will throw an exception
    ''' </summary>
    <TestMethod()> Public Sub AppendToFullArrayThrowsException()
        Const maxIndex As Integer = 3
        Dim array As Array = New Array(maxIndex)
        array.Append("A")
        array.Append("B")
        array.Append("C")
        array.Append("D")

        Try
            array.Append("E")
            Assert.Fail("Expected an overflow exception")
        Catch ex As Exception
            Assert.AreEqual(ex.Message, "Insufficient space!", "Exception thrown witrh an unexpected message.")
        End Try
    End Sub

    ''' <summary>
    ''' Test Helper for Append tests.
    ''' </summary>
    ''' <param name="elementOne">First element to append.</param>
    ''' <param name="elementTwo">Second element to append.</param>
    Private Sub AppendTestHelper(elementOne As String, elementTwo As String)
        Const maxIndex As Integer = 9
        Dim array As Array = New Array(maxIndex)
        array.Append(elementOne)
        array.Append(elementTwo)

        Assert.AreEqual(elementOne, array.Items(0).Value, "First value is incorrect.")
        Assert.AreEqual(elementTwo, array.Items(1).Value, "Second value is incorrect.")
        For i As Integer = 2 To maxIndex
            Assert.AreEqual(String.Empty, array.Items(i).Value, String.Format("Value at Index:{0} is incorrect.", i.ToString()))
        Next
    End Sub

    ''' <summary>
    ''' Test Helper for Append blank values.
    ''' </summary>
    ''' <param name="elementOne">Element to append.</param>
    Private Sub AppendBlankTestHelper(elementOne As String)
        Const maxIndex As Integer = 9
        Dim array As Array = New Array(maxIndex)
        array.Append(elementOne)

        Assert.AreEqual(elementOne, array.Items(0).Value, "First value is incorrect.")
        For i As Integer = 1 To maxIndex
            Assert.AreEqual(String.Empty, array.Items(i).Value, String.Format("Value at Index:{0} is incorrect.", i.ToString()))
        Next
    End Sub
End Class