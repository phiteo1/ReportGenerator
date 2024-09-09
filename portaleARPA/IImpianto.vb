Public Interface IImpianto


    ReadOnly Property getChimneyList As List(Of Camino)

    Function getChimneyFromName(nome As String) As Camino

    Sub mainThread()

    Sub AddChimneyToList()

End Interface
