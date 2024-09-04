Public Class Progress(Of T)
    Private ReadOnly _callback As Action(Of T)

    Public Sub New(callback As Action(Of T))
        If callback Is Nothing Then
            Throw New ArgumentNullException("callback")
        End If
        _callback = callback
    End Sub

    Public Sub Report(value As T)
        _callback.Invoke(value)
    End Sub
End Class
