Imports Microsoft.VisualBasic
Imports System.Web
Public Class MySharedMethod

  
    Public Shared Function KillAllExcels() As Integer
        Dim proc As System.Diagnostics.Process
        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next
        Return 0
    End Function
    Public Shared Function GetChimneyName(ByVal param As Integer) As String
        Select Case (param)
            Case 1
                Return "E1" '(CDU)"
            Case 2
                Return "E2" ' (TSTC)"
            Case 3
                Return "E4" ' (HOT OIL)"
            Case 4
                Return "E7" ' (TIP)"
            Case 5
                Return "E8" ' (RHU/HDC)"
            Case 6
                Return "E9" ' (H2)"
            Case 7
                Return "E10" ' (ZOLFO)"
            Case 8
                Return "MASSICI DI BOLLA"
            Case Else
                Return "Undef. Cemney"
        End Select
    End Function
    Public Shared Function GetMainChimneyName(ByVal param As Integer) As String
        Select Case (param)
            Case 1
                Return "E1 (CDU)"
            Case 2
                Return "E2 (TSTC)"
            Case 3
                Return "E4 (HOT OIL)"
            Case 4
                Return "E7 (TIP)"
            Case 5
                Return "E8 (RHU/HDC)"
            Case 6
                Return "E9 (H2)"
            Case 7
                Return "E10 (ZOLFO)"
            Case 8
                Return "MASSICI DI BOLLA"
            Case Else
                Return "Undef. Cemney"
        End Select
    End Function
    Public Shared Function GetChimneyCoordinate(ByVal param As Integer) As String
        Select Case (param)
            Case 1
                Return "40°29'28.0489'' N;17°11'41.9436'' E"
            Case 2
                Return "40°29'32.9101'' N;17°11'50.4427'' E"
            Case 3
                Return "40°29'31.1051'' N;17°11'42.6837'' E"
            Case 4
                Return "40°29'32.8923'' N;17°11'43.6378'' E"
            Case 5
                Return "40°29'36.5911'' N;17°11'39.5699'' E"
            Case 6
                Return "40°29'10.0527'' N;17°11'48.2331'' E"
            Case 7
                Return "40°29'09.4365'' N;17°11'47.4353'' E"
            Case Else
                Return "Undef. Cemney"
        End Select
    End Function
    Public Shared Function GetChimneyComesFrom(ByVal param As Integer) As String
        Select Case (param)
            Case 1
                Return "CDU,PLAT,HDS1,HDT"
            Case 2
                Return "TSTC,HDS2,Claus 2-3-4,SCOT,Imp.H2,(U2200,U2500),EST,H2 EST (U9400) "
            Case 3
                Return "HOT OIL"
            Case 4
                Return "TIP"
            Case 5
                Return "RHU/HDC"
            Case 6
                Return "H2 (U4400)"
            Case 7
                Return "CLAUS"
            Case Else
                Return "Undef. Cemney"
        End Select
    End Function
    Public Shared Function GetChimneyHeight(ByVal param As Integer) As String
        Select Case (param)
            Case 1
                Return "H: 100 m Area Sezione : 11.52 m2 "
            Case 2
                Return "H: 120 m Area Sezione : 19.63 m2 "
            Case 3
                Return "H: 54.7 m Area Sezione : 1.98 m2 "
            Case 4
                Return "H: 20.1 m Area Sezione : 0.11 m2 "
            Case 5
                Return "H: 95 m Area Sezione : 2.01 m2 "
            Case 6
                Return "H: 40 m Area Sezione : 3.14 m2 "
            Case 7
                Return "H: 80 m Area Sezione : 3.14 m2 "
            Case Else
                Return "Undef. Cemney"
        End Select
    End Function

    Public Shared Function GetMonthName(ByVal param As Integer) As String
        Select Case (param)
            Case 1
                Return "Gennaio"
            Case 2
                Return "Febbraio"
            Case 3
                Return "Marzo"
            Case 4
                Return "Aprile"
            Case 5
                Return "Maggio"
            Case 6
                Return "Giugno"
            Case 7
                Return "Luglio"
            Case 8
                Return "Agosto"
            Case 9
                Return "Settembre"
            Case 10
                Return "Ottobre"
            Case 11
                Return "Novembre"
            Case 12
                Return "Dicembre"
        End Select
        Return ""
    End Function



End Class
