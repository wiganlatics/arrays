''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright (c) 2017 Matthew Wright.                                                       '
' Licensed under MIT License. See LICENSE.txt for further details.                         '
'                                                                                          '
' This software should be distributed with a LICENSE.TXT file in the solution root.        '
' Alternatively  you can find a copy of the license in the github repository:              '
' https://github.com/wiganlatics/arrays.                                                   '
' The MIT License text is also available at: https://choosealicense.com/licenses/mit/.     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
