Public Class Camino

    Private _name As String
    Private _section As Int32
    Private _bolla As Integer

    Public Sub New(name As String, section As Int32, Optional bolla As Byte = 254)

        _name = name
        _section = section
        _bolla = bolla

    End Sub

    Public Function getName() As String

        Return _name

    End Function

    Public Function getSection() As Integer

        Return _section

    End Function


End Class
