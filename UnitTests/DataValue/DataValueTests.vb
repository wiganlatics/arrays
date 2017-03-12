''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright (c) 2017 Matthew Wright.                                                       '
' Licensed under MIT License. See LICENSE.txt for further details.                         '
'                                                                                          '
' This software should be distributed with a LICENSE.TXT file in the solution root.        '
' Alternatively  you can find a copy of the license in the github repository:              '
' https://github.com/wiganlatics/arrays.                                                   '
' The MIT License text is also available at: https://choosealicense.com/licenses/mit/.     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports Arrays
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class DataValueTests
    ''' <summary>
    ''' Test that the DataValue constructor sets the value field correctly.
    ''' </summary>
    <TestMethod()> Public Sub InitialiseDataValueSucceeds()
        Dim initialValue As String = "A"
        Dim dataValue As DataValue = New DataValue(initialValue)
        Assert.AreEqual(initialValue, dataValue.Value, "The DataValue was not initialised correctly.")
    End Sub

    ''' <summary>
    ''' Test that the DataValue value field can be updated correctly.
    ''' </summary>
    <TestMethod()> Public Sub SetDataValueSucceeds()
        Dim initialValue As String = "A"
        Dim updateValue As String = "B"
        Dim dataValue As DataValue = New DataValue(initialValue)
        dataValue.Value = updateValue
        Assert.AreEqual(updateValue, dataValue.Value, "The DataValue was not updated correctly.")
    End Sub
End Class