Imports System.IO

Public Class Logger
    Shared exePath As String = Application.StartupPath
    ' Ottiene la directory tre livelli sopra l'eseguibile
    Shared grandParentPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
    Private Shared logFile As String
    Private Shared logFilePath As String


    Public Shared Sub CreateLogDir()

        Dim logDir = Path.Combine(grandParentPath, "logger")
        If Not Directory.Exists(logDir) Then
            Try
                Directory.CreateDirectory(logDir)
            Catch ex As Exception
                Console.WriteLine("Errore nella scrittura del log: " & ex.Message)
            End Try

        End If

    End Sub


    ' Metodo per scrivere un messaggio di log
    Public Shared Sub Log(message As String)

        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HH")
        logFile = "log_£" & timestamp & ".txt"
        logFilePath = Path.Combine(grandParentPath, "logger", logFile)
        Try
            Using writer As StreamWriter = New StreamWriter(logFilePath, True)
                writer.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " - " & message)
            End Using
        Catch ex As Exception
            Console.WriteLine("Errore nella scrittura del log: " & ex.Message)
        End Try
    End Sub

    ' Metodo per scrivere un messaggio di log per gli errori
    Public Shared Sub LogError(message As String, ex As Exception)
        Try
            Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HH")
            logFile = "log_" & timestamp & ".txt"
            logFilePath = Path.Combine(grandParentPath, "logger", logFile)
            Using writer As StreamWriter = New StreamWriter(logFilePath, True)
                writer.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " - ERRORE: " & message)
                writer.WriteLine("Eccezione: " & ex.Message)
                writer.WriteLine("StackTrace: " & ex.StackTrace)
            End Using
        Catch e As Exception
            Console.WriteLine("Errore nella scrittura del log: " & e.Message)
        End Try
    End Sub

    Public Shared Sub LogWarning(message As String, ex As Exception)
        Try
            Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HH")
            logFile = "log_" & timestamp & ".txt"
            logFilePath = Path.Combine(grandParentPath, "logger", logFile)
            Using writer As StreamWriter = New StreamWriter(logFilePath, True)
                writer.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " - WARNING: " & message)
                writer.WriteLine("Eccezione: " & ex.Message)
                writer.WriteLine("StackTrace: " & ex.StackTrace)
            End Using
        Catch e As Exception
            Console.WriteLine("Errore nella scrittura del log: " & e.Message)
        End Try
    End Sub
End Class
