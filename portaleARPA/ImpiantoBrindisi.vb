Imports System.Linq
Imports System.Threading
Imports System.Configuration
Imports System.Globalization
Imports System.IO
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.Reflection
Imports System.ComponentModel
Imports GemBox.Spreadsheet

Public Class ImpiantoBrindisi
    Implements IImpianto

    Private formInstance As Form1 = DirectCast(Application.OpenForms("Form1"), Form1)
    Dim connectionString As String
    Dim connectionStringCTE As String
    Private _chimneyList As New List(Of Camino)
    Private actualState As Byte
    Private hiddenColumns As New List(Of String)()
    Private culture As System.Globalization.CultureInfo
    Dim hnf, htran, vleCo, vleNox As String
    Dim colToJ As Dictionary(Of Integer, Integer)
    Dim QAL2 As DataTable
        

    Enum State                  'State Machine of the downloading process
        DataLoading = 1
        TableLoading = 2
        SheetLoading = 3
        FinishedReport = 4
        Finished = 5
    End Enum

    Public Sub New()

        AddChimneyToList()
        InitCorrectionDict()
        culture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        culture.NumberFormat.NumberGroupSeparator = ""
        connectionString = ConfigurationManager.ConnectionStrings("AQMSDBCONN").ConnectionString
        connectionStringCTE = ConfigurationManager.ConnectionStrings("AQMSDBCONNCTE").ConnectionString
        QAL2 = New DataTable()


    End Sub


    Private Sub AddChimneyToList() Implements IImpianto.AddChimneyToList

        _chimneyList.Add(New Camino("CC1", 1))
        _chimneyList.Add(New Camino("CC2", 2))
        _chimneyList.Add(New Camino("CC3", 3))

    End Sub


    Public ReadOnly Property getChimenyList As List(Of Camino) Implements IImpianto.getChimneyList

        Get

            Return _chimneyList
        End Get

    End Property

    Public Function getChimneyFromName(name As String) As Camino Implements IImpianto.getChimneyFromName

        Dim chimney As Camino = _chimneyList.FirstOrDefault(Function(c) c.getName = name)

        Return chimney

    End Function


    Public Sub mainThread() Implements IImpianto.mainThread

        Dim exePath As String = Application.StartupPath                                                                                                                     ' Get the 2 layer up directory
        Dim grandParentPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim chimneyName As String = GetChimneyName(Form1.section)
        Dim reportPath As String = Path.Combine(grandParentPath, "report", chimneyName)
        Dim startDate As Date = Form1.startDate
        Dim endDate As Date = Form1.endDate


        UpdateProgressBarStatus(formInstance, True)

        UpdateTextBoxStatus(formInstance, formInstance.TextBox1, True)


        Dim barProgress As New Progress(Of Integer)(Sub(v)
                                                        UpdateProgressBarValue(formInstance, v)
                                                    End Sub)                                                                                                                    'Refresh the GUI when a change in the progress bar occours


        Dim StatusProgress As New Progress(Of Integer)(Sub(index)
                                                           Select Case index
                                                               Case 1
                                                                   UpdateTextBoxText(formInstance, formInstance.TextBox1, "Data Loading...")
                                                                   actualState = State.DataLoading
                                                               Case 2
                                                                   UpdateTextBoxText(formInstance, formInstance.TextBox1, "Table creation...")
                                                                   UpdateProgressBarStatus(formInstance, False)
                                                                   actualState = State.TableLoading
                                                               Case 3
                                                                   UpdateTextBoxText(formInstance, formInstance.TextBox1, "Sheet creation...")
                                                                   actualState = State.SheetLoading
                                                               Case 4
                                                                   If (Form1.reportType = 0) Then

                                                                       UpdateTextBoxText(formInstance, formInstance.TextBox1, "Year " & startDate.Year.ToString & " downloaded succesfully")

                                                                   ElseIf (Form1.reportType = 1) Then

                                                                       UpdateTextBoxText(formInstance, formInstance.TextBox1, "Month " & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:MMMM yyyy}", Date.Parse(startDate)) & " downloaded succesfully")

                                                                   End If
                                                                   actualState = State.FinishedReport
                                                               Case 5
                                                                   UpdateTextBoxText(formInstance, formInstance.TextBox1, "Report generation finished!")
                                                                   actualState = State.Finished
                                                                   EnableFormSafe(formInstance)
                                                                   HideFormSafe(formInstance)
                                                           End Select
                                                       End Sub)                                                                                                                         'Refresh the GUI when a change in the state occours

        Dim dataTable1 As DataTable = Nothing
        Dim dataTable2 As DataTable = Nothing

        If Not Directory.Exists(reportPath) Then
            Try
                Directory.CreateDirectory(reportPath)
            Catch ex As Exception
                MessageBox.Show("Errore nella creazione della directory.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                EnableFormSafe(formInstance)
                Return
            End Try
        End If

        While (startDate <= endDate)

            UpdateProgressBarValue(formInstance, 0)
            If (Not formInstance.ProgressBar1.Visible) Then
                UpdateProgressBarStatus(formInstance, True)
            End If

            dataTable1 = GetFirstCaminiTable(barProgress, startDate, endDate, Form1.section, Form1.reportType)
            If dataTable1 Is Nothing Then
                MessageBox.Show("Errore nell'acquisizione dei dati, consultare il file di log per i dettagli.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                EnableFormSafe(formInstance)
                Return
            Else
                preRenderCaminiTable(dataTable1)
                UpdateDgvDataSource(dataTable1, Form1.dgv)
            End If

            If Form1.reportType = 2 Then
                dataTable2 = GetDailySecondCaminiTable(barProgress, startDate, endDate, Form1.section, Form1.reportType)
            Else
                dataTable2 = GetSecondCaminiTable(barProgress, startDate, endDate, Form1.section, Form1.reportType)
            End If


            

            If dataTable2 Is Nothing Then
                MessageBox.Show("Errore nell'acquisizione dei dati, consultare il file di log per i dettagli.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                EnableFormSafe(formInstance)
                Return
            Else
                UpdateDgvDataSource(dataTable2, Form1.dgv2)
            End If


            If Form1.reportType = 0 Then

                downloadYearlyReportCamini(StatusProgress, startDate, endDate, reportPath)

                '    downloadYearlyReportCTE(StatusProgress, startDate, endDate, reportPath)

            ElseIf Form1.reportType = 1 Then


                downloadMonthlyReportCamini(StatusProgress, startDate, endDate, reportPath)

                '    downloadMonthlyReportCTE(StatusProgress, startDate, endDate, reportPath)
            Else

                downloadDailyReportCamini(StatusProgress, startDate, endDate, reportPath)


            End If


            Dim deltaTime As String
            If (Form1.reportType = 0) Then
                deltaTime = "yyyy"                                                                                                                                                      'Add one year or one month according to the report type choosed
            ElseIf (Form1.reportType = 1) Then
                deltaTime = "m"
            Else
                deltaTime = "d"
            End If

            startDate = DateAdd(deltaTime, 1, startDate)

        End While


    End Sub
   

    Private Function GetFirstCaminiTable(progress As Progress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, ByVal type As Int32) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim command As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim connection As New SqlConnection(connectionString)
        Dim queryNumber As Integer = 0
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount
        Dim methodName As String = GetCurrentMethod()
        Dim dataType As String = " ORDER BY N_RIGA"
        Dim ret As Integer

        If Form1.reportType = 0 Then                                                      ' It was needed thanks to the genius who wrote the logics in the portal :))
            type = 3
        ElseIf Form1.reportType = 1 Then
            type = 2
        Else
            type = 1
        End If

        Try
            ' Tenta di aprire la connessione
            connection.Open()
        Catch ex As Exception
            ' Gestione degli errori
            MessageBox.Show("Errore durante la connessione al database: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dt = Nothing
            Return dt
        End Try

        
        Dim q2Result = FillQAL2()

        If q2Result = 0 Then
            dt = Nothing
            Return dt
        End If

        dt.Columns.Add(New Data.DataColumn("INTESTAZIONE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("TFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("H2O", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2RIF", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("MWE", GetType(String)))

        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        Dim testCMD As Data.SqlClient.SqlCommand = New Data.SqlClient.SqlCommand("sp_AQMSNT_FILL_ARPA_REPORT_WEB", connection)
        testCMD.CommandType = Data.CommandType.StoredProcedure
        testCMD.Parameters.Add("@idsez", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@idsez").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@idsez").Value = Form1.section

        testCMD.Parameters.Add("@data", Data.SqlDbType.DateTime, 11)
        testCMD.Parameters("@data").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@data").Value = Format("{0:dd/MM/yyyy}", startTime.ToString()) 'RepggCal.SelectedDate.ToString("dd/MM/yyyy HH:mm:ss")

        testCMD.Parameters.Add("@tipoestrazione", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@tipoestrazione").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@tipoestrazione").Value = type

        testCMD.Parameters.Add("@retval", Data.SqlDbType.Int)
        testCMD.Parameters("@retval").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@HNF", Data.SqlDbType.Int)
        testCMD.Parameters("@HNF").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@H_TRANS", Data.SqlDbType.Int)
        testCMD.Parameters("@H_TRANS").Direction = Data.ParameterDirection.Output

        Try
            testCMD.ExecuteScalar()
        Catch ex As Exception
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della stored procedure: ", ex)
            dt = Nothing
            Return dt
        End Try

        ret = testCMD.Parameters("@retval").Value
        hnf = testCMD.Parameters("@HNF").Value.ToString()
        htran = testCMD.Parameters("@H_TRANS").Value.ToString()

        Dim log_statement As String = "SELECT * FROM [ARPA_REPORT_WEB] WHERE IDX_REPORT = " & ret.ToString() & dataType
        command = New System.Data.SqlClient.SqlCommand(log_statement, connection)

        Try
            reader = command.ExecuteReader()
        Catch ex As SqlException
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della query: ", ex)
            dt = Nothing
            Return dt
        End Try

        Dim dr As Data.DataRow = dt.NewRow()
        While reader.Read()
            Try
                dr("INTESTAZIONE") = reader("INTESTAZIONE")
                dr("CO") = String.Format("{0:n2}", reader("CO"))
                dr("NOX") = String.Format("{0:n2}", reader("NOX"))
                dr("O2") = String.Format("{0:n2}", reader("O2"))
                dr("TFUMI") = String.Format("{0:n2}", reader("TFUMI"))
                dr("PFUMI") = String.Format("{0:n2}", reader("PFUMI"))
                dr("H2O") = String.Format("{0:n2}", reader("H2O"))
                dr("O2RIF") = String.Format("{0:n2}", reader("O2RIF"))
                dr("MWE") = String.Format("{0:n2}", reader("MWE"))

                dt.Rows.Add(dr)
                dr = dt.NewRow()
            Catch ex As Exception
                Logger.LogWarning("[" & methodName & "]" & " Errore nella lettura dei dati: ", ex)
                Continue While
            End Try

        End While

        Return dt

    End Function

    Private Function GetSecondCaminiTable(progress As Progress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, ByVal type As Int32) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim command As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim connection As New SqlConnection(connectionString)
        Dim queryNumber As Integer = 3
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount
        Dim methodName As String = GetCurrentMethod()
        Dim dataType As String = " ORDER BY INS_ORDER"
        Dim ret As Integer
        Dim isMese As Integer

        If Form1.reportType = 0 Then                                                      ' It was needed thanks to the genius who wrote the logics in the portal :))
            type = 3
            isMese = 0
        ElseIf Form1.reportType = 1 Then
            type = 2
            isMese = 1
        End If

        Try
            ' Tenta di aprire la connessione
            connection.Open()
        Catch ex As Exception
            ' Gestione degli errori
            MessageBox.Show("Errore durante la connessione al database: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dt = Nothing
            Return dt
        End Try



        dt.Columns.Add(New Data.DataColumn("IDX_REPORT", GetType(Double)))
        dt.Columns.Add(New Data.DataColumn("INS_ORDER", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("ORA", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOX_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOTE_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_NOX", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("CO_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOTE_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_CO", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("O2_MIS", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2RIF", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_O2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOTE_O2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("TFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_TFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOTE_TFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_PFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOTE_PFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("ORE_NF", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_QFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOTE_QFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("UFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_UFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOTE_UFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("MWE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOTE_MW", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QGAS", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOTE_QGAS", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QFUELGAS", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOTE_QFGAS", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOTE_QCO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOTE_QNOX", GetType(String)))


        Dim testCMD As Data.SqlClient.SqlCommand = New Data.SqlClient.SqlCommand("sp_AQMSNT_FILL_ARPA_MESE_ANNO_REPORT", connection)
        testCMD.CommandType = Data.CommandType.StoredProcedure
        testCMD.Parameters.Add("@idsez", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@idsez").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@idsez").Value = Form1.section

        testCMD.Parameters.Add("@data", Data.SqlDbType.DateTime, 11)
        testCMD.Parameters("@data").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@data").Value = Format("{0:dd/MM/yyyy}", startTime.ToString()) 'RepggCal.SelectedDate.ToString("dd/MM/yyyy HH:mm:ss")

        testCMD.Parameters.Add("@IS_MESE", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@IS_MESE").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@IS_MESE").Value = isMese

        testCMD.Parameters.Add("@retval", Data.SqlDbType.Int)
        testCMD.Parameters("@retval").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@LL_GG_NOX", Data.SqlDbType.Float)
        testCMD.Parameters("@LL_GG_NOX").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@LL_GG_CO", Data.SqlDbType.Float)
        testCMD.Parameters("@LL_GG_CO").Direction = Data.ParameterDirection.Output

        Try
            testCMD.ExecuteScalar()
        Catch ex As Exception
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della stored procedure: ", ex)
            dt = Nothing
            Return dt
        End Try


        vleCo = testCMD.Parameters("@LL_GG_CO").Value.ToString()
        vleNox = testCMD.Parameters("@LL_GG_NOX").Value.ToString()
        ret = testCMD.Parameters("@retval").Value

        Dim log_statement As String = "SELECT * FROM [ARPA_WEB_MESE_ANNO_REPORT] WHERE IDX_REPORT = " & ret.ToString() & dataType
        command = New System.Data.SqlClient.SqlCommand(log_statement, connection)

        Try
            reader = command.ExecuteReader()
        Catch ex As SqlException
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della query: ", ex)
            dt = Nothing
            Return dt
        End Try

        Dim dr As Data.DataRow = dt.NewRow()
        While reader.Read()
            Try
                dr("IDX_REPORT") = reader("IDX_REPORT")
                dr("INS_ORDER") = String.Format("{0:n0}", reader("INS_ORDER"))
                dr("ORA") = reader("ORA") 'String.Format("{0:n2}", reader("NOX"))
                dr("NOX_IC") = reader("NOX_IC") 'String.Format("{0:n2}", Convert.ToDouble(reader("NOX_IC")))
                dr("DISP_NOX") = reader("DISP_NOX") 'String.Format("{0:n2}", reader("DISP_NOX"))
                dr("NOTE_NOX") = String.Format("{0:n2}", reader("NOTE_NOX"))
                dr("IS_BOLD_NOX") = reader("IS_BOLD_NOX")
                dr("CO_IC") = reader("CO_IC") 'String.Format("{0:n2}", Convert.ToDouble(reader("CO_IC")))
                dr("DISP_CO") = reader("DISP_CO") 'String.Format("{0:n2}", reader("DISP_CO"))
                dr("NOTE_CO") = String.Format("{0:n2}", reader("NOTE_CO"))
                dr("IS_BOLD_CO") = reader("IS_BOLD_CO")
                dr("O2_MIS") = String.Format("{0:n2}", reader("O2_MIS"))
                dr("O2RIF") = String.Format("{0:n2}", reader("O2_RIF"))
                dr("DISP_O2") = String.Format("{0:n2}", reader("DISP_O2"))
                dr("NOTE_O2") = String.Format("{0:n2}", reader("NOTE_O2"))
                dr("TFUMI") = String.Format("{0:n2}", reader("TFUMI"))
                dr("DISP_TFUMI") = String.Format("{0:n2}", reader("DISP_TFUMI"))
                dr("NOTE_TFUMI") = String.Format("{0:n2}", reader("NOTE_TFUMI"))
                dr("PFUMI") = String.Format("{0:n2}", reader("PFUMI"))
                dr("DISP_PFUMI") = String.Format("{0:n2}", reader("DISP_PFUMI"))
                dr("NOTE_PFUMI") = String.Format("{0:n2}", reader("NOTE_PFUMI"))
                dr("ORE_NF") = String.Format("{0:n2}", reader("ORE_NF"))
                dr("QFUMI") = String.Format("{0:n2}", reader("QFUMI"))
                dr("DISP_QFUMI") = String.Format("{0:n2}", reader("DISP_QFUMI"))
                dr("NOTE_QFUMI") = String.Format("{0:n2}", reader("NOTE_QFUMI"))
                dr("UFUMI") = String.Format("{0:n2}", reader("UFUMI"))
                dr("DISP_UFUMI") = String.Format("{0:n2}", reader("DISP_UFUMI"))
                dr("NOTE_UFUMI") = String.Format("{0:n2}", reader("NOTE_UFUMI"))
                dr("MWE") = String.Format("{0:n2}", reader("MWE"))
                dr("NOTE_MW") = String.Format("{0:n2}", reader("NOTE_MW"))
                dr("QGAS") = String.Format("{0:n2}", reader("QGAS"))
                dr("NOTE_QGAS") = String.Format("{0:n2}", reader("NOTE_QGAS"))
                dr("QFUELGAS") = String.Format("{0:n2}", reader("QFUELGAS"))
                dr("NOTE_QFGAS") = String.Format("{0:n2}", reader("NOTE_QFGAS"))
                dr("PORTATA_CO") = String.Format("{0:n2}", reader("PORTATA_CO"))
                dr("NOTE_QCO") = String.Format("{0:n2}", reader("NOTE_QCO"))
                dr("PORTATA_NOX") = String.Format("{0:n2}", reader("PORTATA_NOX"))
                dr("NOTE_QNOX") = String.Format("{0:n2}", reader("NOTE_QNOX"))

                dt.Rows.Add(dr)
                dr = dt.NewRow()
            Catch ex As Exception
                Logger.LogWarning("[" & methodName & "]" & " Errore nella lettura dei dati: ", ex)
                Continue While
            End Try

        End While

        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        Return dt

    End Function

    Private Function GetDailySecondCaminiTable(progress As Progress(Of Integer), startTime As DateTime, endTime As DateTime, section As Int32, ByVal type As Int32) As Data.DataTable

        Dim dt As New Data.DataTable()
        Dim command As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim connection As New SqlConnection(connectionString)
        Dim queryNumber As Integer = 3
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount
        Dim methodName As String = GetCurrentMethod()
        Dim dataType As String = " ORDER BY INS_ORDER"
        Dim ret As Integer
        

        Try
            ' Tenta di aprire la connessione
            connection.Open()
        Catch ex As Exception
            ' Gestione degli errori
            MessageBox.Show("Errore durante la connessione al database: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dt = Nothing
            Return dt
        End Try



        dt.Columns.Add(New Data.DataColumn("IDX_REPORT", GetType(Double)))
        dt.Columns.Add(New Data.DataColumn("INS_ORDER", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("ORA", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOX_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOX_TQ", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOX_COD", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_NOX", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("CO_IC", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("CO_TQ", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("CO_COD", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("IS_BOLD_CO", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("O2_MIS", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2_RIF", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("O2_COD", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("DISP_O2", GetType(String))) 'NUOVO
        dt.Columns.Add(New Data.DataColumn("TFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("TFUMI_COD", GetType(String)))   'NUOVO
        dt.Columns.Add(New Data.DataColumn("DISP_TFUMI", GetType(String)))  'NUOVO
        dt.Columns.Add(New Data.DataColumn("PFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PFUMI_COD", GetType(String)))   'NUOVO
        dt.Columns.Add(New Data.DataColumn("DISP_PFUMI", GetType(String)))  'NUOVO
        dt.Columns.Add(New Data.DataColumn("STATO_IMP", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QFUMI_COD", GetType(String)))   'NUOVO
        dt.Columns.Add(New Data.DataColumn("DISP_QFUMI", GetType(String)))  'NUOVO
        dt.Columns.Add(New Data.DataColumn("UFUMI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("UFUMI_COD", GetType(String)))   'NUOVO
        dt.Columns.Add(New Data.DataColumn("DISP_UFUMI", GetType(String)))  'NUOVO
        dt.Columns.Add(New Data.DataColumn("MWE", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QGAS", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("QFUELGAS", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("PORTATA_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("MULETTO_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("MULETTO_NO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("MULETTO_O2", GetType(String)))


        Dim testCMD As Data.SqlClient.SqlCommand = New Data.SqlClient.SqlCommand("sp_AQMSNT_FILL_ARPA_REALTIME", connection)
        testCMD.CommandType = Data.CommandType.StoredProcedure
        testCMD.Parameters.Add("@idsez", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@idsez").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@idsez").Value = Form1.section

        testCMD.Parameters.Add("@data", Data.SqlDbType.DateTime, 11)
        testCMD.Parameters("@data").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@data").Value = Format("{0:dd/MM/yyyy}", startTime.ToString()) 'RepggCal.SelectedDate.ToString("dd/MM/yyyy HH:mm:ss")

        testCMD.Parameters.Add("@retval", Data.SqlDbType.Int)
        testCMD.Parameters("@retval").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@percval_co", Data.SqlDbType.Int)
        testCMD.Parameters("@percval_co").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@percval_nox", Data.SqlDbType.Int)
        testCMD.Parameters("@percval_nox").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@totvlehhco", Data.SqlDbType.Int)
        testCMD.Parameters("@totvlehhco").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@totvlehhnox", Data.SqlDbType.Int)
        testCMD.Parameters("@totvlehhnox").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@totvleggco", Data.SqlDbType.Int)
        testCMD.Parameters("@totvleggco").Direction = Data.ParameterDirection.Output
        testCMD.Parameters.Add("@totvleggnox", Data.SqlDbType.Int)
        testCMD.Parameters("@totvleggnox").Direction = Data.ParameterDirection.Output

        Try
            testCMD.ExecuteScalar()
        Catch ex As Exception
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della stored procedure: ", ex)
            dt = Nothing
            Return dt
        End Try


        vleCo = testCMD.Parameters("@percval_co").Value
        vleNox = testCMD.Parameters("@percval_nox").Value
        ret = testCMD.Parameters("@retval").Value

        Dim log_statement As String = "SELECT * FROM [ARPA_WEB_REAL_TIME] WHERE IDX_REPORT = " & ret.ToString() & dataType
        command = New System.Data.SqlClient.SqlCommand(log_statement, connection)

        Try
            reader = command.ExecuteReader()
        Catch ex As SqlException
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della query: ", ex)
            dt = Nothing
            Return dt
        End Try

        Dim dr As Data.DataRow = dt.NewRow()

        While reader.Read()
            Try
                dr("IDX_REPORT") = reader("IDX_REPORT")
                dr("INS_ORDER") = String.Format("{0:n0}", reader("INS_ORDER"))
                dr("ORA") = reader("ORA") 'String.Format("{0:n2}", reader("NOX"))
                dr("NOX_IC") = String.Format("{0:n2}", reader("NOX_IC"))
                dr("NOX_TQ") = String.Format("{0:n2}", reader("NOX_TQ"))
                dr("NOX_COD") = String.Format("{0:n2}", reader("NOX_COD"))
                dr("DISP_NOX") = String.Format("{0:n2}", reader("DISP_NOX"))
                dr("IS_BOLD_NOX") = reader("IS_BOLD_NOX")
                dr("CO_IC") = String.Format("{0:n2}", reader("CO_IC"))
                dr("CO_TQ") = String.Format("{0:n2}", reader("CO_TQ"))
                dr("CO_COD") = String.Format("{0:n2}", reader("CO_COD"))
                dr("DISP_CO") = String.Format("{0:n2}", reader("DISP_CO"))
                dr("IS_BOLD_CO") = reader("IS_BOLD_CO")
                dr("O2_MIS") = String.Format("{0:n2}", reader("O2_MIS"))
                dr("O2_RIF") = String.Format("{0:n2}", reader("O2_RIF"))
                dr("O2_COD") = String.Format("{0:n2}", reader("O2_COD"))
                dr("DISP_O2") = String.Format("{0:n2}", reader("DISP_O2")) 'NUOVO
                dr("TFUMI") = String.Format("{0:n2}", reader("TFUMI"))
                dr("TFUMI_COD") = String.Format("{0:n2}", reader("TFUMI_COD")) 'NUOVO
                dr("DISP_TFUMI") = String.Format("{0:n2}", reader("DISP_TFUMI")) 'NUOVO
                dr("PFUMI") = String.Format("{0:n2}", reader("PFUMI"))
                dr("PFUMI_COD") = String.Format("{0:n2}", reader("PFUMI_COD")) 'NUOVO
                dr("DISP_PFUMI") = String.Format("{0:n2}", reader("DISP_PFUMI")) 'NUOVO
                dr("STATO_IMP") = String.Format("{0:n2}", reader("STATO_IMP"))
                dr("QFUMI") = String.Format("{0:n2}", reader("QFUMI"))
                dr("QFUMI_COD") = String.Format("{0:n2}", reader("QFUMI_COD")) 'NUOVO
                dr("DISP_QFUMI") = String.Format("{0:n2}", reader("DISP_QFUMI")) 'NUOVO
                dr("UFUMI") = String.Format("{0:n2}", reader("UFUMI"))
                dr("UFUMI_COD") = String.Format("{0:n2}", reader("UFUMI_COD")) 'NUOVO
                dr("DISP_UFUMI") = String.Format("{0:n2}", reader("DISP_UFUMI")) 'NUOVO
                dr("MWE") = String.Format("{0:n2}", reader("MWE"))

                dr("QGAS") = String.Format("{0:n2}", reader("QGAS"))

                dr("QFUELGAS") = String.Format("{0:n2}", reader("QFUELGAS"))

                dr("PORTATA_CO") = String.Format("{0:n2}", reader("PORTATA_CO"))

                dr("PORTATA_NOX") = String.Format("{0:n2}", reader("PORTATA_NOX"))
                dr("MULETTO_CO") = String.Format("{0:n2}", reader("MULETTO_CO"))
                dr("MULETTO_NO") = String.Format("{0:n2}", reader("MULETTO_NO"))
                dr("MULETTO_O2") = String.Format("{0:n2}", reader("MULETTO_O2"))

                dt.Rows.Add(dr)
                dr = dt.NewRow()
            Catch ex As Exception
                Logger.LogWarning("[" & methodName & "]" & " Errore nella lettura dei dati: ", ex)
                Continue While
            End Try

        End While

        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        Return dt

    End Function





    Private Sub downloadYearlyReportCamini(ComboStatus As Progress(Of Integer), startDate As Date, endDate As Date, reportDir As String)

        Dim exePath As String = Application.StartupPath
        Dim rootPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim templateDir As String = ConfigurationManager.AppSettings("TemplateDirectory")
        Dim reportTitle As String = ""
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")
        Dim wBookGemBox As ExcelFile = ExcelFile.Load(Path.Combine(rootPath, templateDir, "152_CONC_ANNOTG.xls")) 'Viene caricato il template Excel dal percorso specificato.
        Dim wSheetGemBox As ExcelWorksheet = wBookGemBox.Worksheets(0) 'Viene ottenuto il riferimento al primo foglio di lavoro del file Excel caricato.

        ComboStatus.Report(State.TableLoading)

        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("G1"), "152_CONC_ANNO", True)
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("G2"), "ENIPOWER - Centrale di Brindisi - Sezione Termoelettrica n° " + Form1.section.ToString(), True)
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("G3"), "Sistema di Misura delle Emissioni", True)
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("G4"), "Report Mensile concentrazioni medie  giornaliere (Nox,Co) - Dlgs 152 (70%)", True)
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("G5"), "Report Annuale Anno " & Date.Parse(startDate.ToString(), New System.Globalization.CultureInfo("it-IT")).Year, True)
        SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange("G1:R1"))
        SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange("G2:R2"))
        SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange("G3:R3"))
        SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange("G4:R4"))
        SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange("G5:R5"))
        SetRangeValueAndStyle(wSheetGemBox.NamedRanges.Item("HNF").Range, hnf, True) 'Numero ore di normale funzionamento di impianto 
        SetRangeValueAndStyle(wSheetGemBox.NamedRanges.Item("HTRANS").Range, htran, True)  'Numero ore di transitorio di impianto
        reportTitle = "152_CONC_ANNO_" & Form1.section.ToString() & "_" & startDate.Year



        Dim i As Integer
        Dim j As Integer
        Dim cc As Integer
        Dim app As String
        Dim col As Integer
        Dim insert_tab As Integer
        Dim colgv As Integer
        cc = 11


        ComboStatus.Report(State.SheetLoading)

        For i = 0 To Form1.dgv2.Rows.Count - 2
            wSheetGemBox.Rows.InsertEmpty(cc + i)
        Next
        For i = 0 To Form1.dgv2.Rows.Count - 1

            app = "B" & i + 11 & ":R" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "C" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "D" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "E" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "F" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "G" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "H" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "I" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "J" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "K" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "L" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "M" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "N" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "O" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "P" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "Q" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "R" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "S" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "T" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "U" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "V" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "W" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "X" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "Y" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "Z" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "AA" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "AB" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            For col = 1 To 28 - 1


                If colToJ.ContainsKey(col) Then
                    j = colToJ(col)
                End If

                If col = 6 Then
                    'j = j + 1
                    If Form1.dgv2.Rows(i).Cells(j).Value.ToString() = "&nbsp;" Then
                        wSheetGemBox.Columns(col).Cells(i + 10).Value = ""
                    Else
                        app = "C" + Convert.ToString(i + 10) + ":C" + Convert.ToString(i + 10)
                        wSheetGemBox.Columns(col).Cells(i + 10).Value = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                        wSheetGemBox.Cells(i + 10, col).Style.HorizontalAlignment = HorizontalAlignmentStyle.Center

                    End If
                ElseIf col = 9 Then
                    'j = j + 1
                    If Form1.dgv2.Rows(i).Cells(j).Value.ToString() = "&nbsp;" Then
                        wSheetGemBox.Columns(col).Cells(i + 10).Value = ""
                        wSheetGemBox.Cells(i + 10, col).Style.HorizontalAlignment = HorizontalAlignmentStyle.Center

                    Else
                        app = "E" + Convert.ToString(i + 10) + ":E" + Convert.ToString(i + 10)
                        wSheetGemBox.Columns(col).Cells(i + 10).Value = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                        wSheetGemBox.Cells(i + 10, col).Style.HorizontalAlignment = HorizontalAlignmentStyle.Center

                    End If
                    'col = col + 1
                Else
                    If Form1.dgv2.Rows(i).Cells(j).Value.ToString() = "&nbsp;" Then
                        wSheetGemBox.Columns(col).Cells(i + 10).Value = ""

                    Else
                        app = "C" + Convert.ToString(i + 10) + ":C" + Convert.ToString(i + 10)
                        wSheetGemBox.Columns(col).Cells(i + 10).Value = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                        wSheetGemBox.Cells(i + 10, col).Style.HorizontalAlignment = HorizontalAlignmentStyle.Center
                    End If
                End If
            Next
        Next

        col = 2
        insert_tab = i + 35
        For z = 0 To Form1.dgv.Rows.Count - 1


            colgv = 5
            For tabcounter = 2 To Form1.dgv.Columns.Count
                If Form1.dgv.Rows(z).Cells(tabcounter - 1).Value.ToString() = "&nbsp;" Then
                    wSheetGemBox.Columns(colgv).Cells(insert_tab).Value = ""
                Else
                    wSheetGemBox.Columns(colgv - 1).Cells(insert_tab - 1).Value = Form1.dgv.Rows(z).Cells(tabcounter - 1).Value.ToString() ' tabella media annuale.
                End If
                colgv += 1
            Next
            insert_tab = insert_tab + 1
            col = col + 1
        Next

        ComboStatus.Report(State.FinishedReport)

        Dim reportFileXls = reportTitle & ".xls"
        Dim reportFilePdf = reportTitle & ".pdf"
        Dim reportPath = Path.Combine(reportDir, reportFileXls)
        Dim reportPathPdf = Path.Combine(reportDir, reportFilePdf)
        wBookGemBox.Save(reportPath)
        wBookGemBox.Save(reportPathPdf)
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("it-IT")

        If (startDate = endDate) Then

            ComboStatus.Report(State.Finished)
            ShowCompletionDialog()
        End If
    End Sub

    Private Sub downloadMonthlyReportCamini(ComboStatus As Progress(Of Integer), startDate As Date, endDate As Date, reportDir As String)

        Dim exePath As String = Application.StartupPath
        Dim rootPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim templateDir As String = ConfigurationManager.AppSettings("TemplateDirectory")
        Dim reportTitle As String = ""
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")
        Dim wBookGemBox As ExcelFile = ExcelFile.Load(Path.Combine(rootPath, templateDir, "152_CONC_MESETG.xls")) 'Viene caricato il template Excel dal percorso specificato.
        Dim wSheetGemBox As ExcelWorksheet = wBookGemBox.Worksheets(0) 'Viene ottenuto il riferimento al primo foglio di lavoro del file Excel caricato.

        ComboStatus.Report(State.TableLoading)

        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("G1"), "152_CONC_MESE", True)
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("G2"), "ENIPOWER - Centrale di Brindisi - Sezione Termoelettrica n° " + Form1.section.ToString(), True)
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("G3"), "Sistema di Misura delle Emissioni", True)
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("G4"), "Report Mensile concentrazioni medie  giornaliere (Nox,Co) - Dlgs 152 (70%)", True)
        Dim startDateFormatted As DateTime = DateTime.Parse(startDate).Date
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("G5"), "Report Mensile del Mese di " & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:MMMM yyyy}", startDateFormatted), True)

        'MOD 3
        SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange("G1:R1"))
        SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange("G2:R2"))
        SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange("G3:R3"))
        SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange("G4:R4"))
        SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange("G5:R5"))
        SetRangeValueAndStyle(wSheetGemBox.NamedRanges.Item("HNF").Range, hnf, True)
        SetRangeValueAndStyle(wSheetGemBox.NamedRanges.Item("HTRANS").Range, htran, True)
        reportTitle = "152_CONC_MESECC_" & Form1.section.ToString() & "_" & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:MMMM_yyyy}", Date.Parse(startDate))

        Dim i As Integer
        Dim j As Integer
        Dim cc As Integer
        Dim app As String
        Dim col As Integer
        Dim insert_tab As Integer
        Dim colgv As Integer
        cc = 11

        ComboStatus.Report(State.SheetLoading)

        For i = 0 To Form1.dgv2.Rows.Count - 2
            wSheetGemBox.Rows.InsertEmpty(cc + i)
        Next
        For i = 0 To Form1.dgv2.Rows.Count - 1

            app = "B" & i + 11 & ":R" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "C" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "D" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "E" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "F" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "G" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "H" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "I" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "J" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "K" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "L" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "M" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "N" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "O" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "P" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "Q" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            app = "R" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            'MOD 4
            app = "S" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "T" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "U" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "V" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "W" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "X" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "Y" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "Z" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "AA" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))

            app = "AB" & i + 11
            SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange(app))
            col = 2
            'For j = 2 To 21 - 1
            'For j = 2 To 38 - 1
            'For j = 2 To 38 - 1
            '    System.Diagnostics.Debug.WriteLine("j:" & j & " - Valore: " & gv_monthlyrep.SelectedRow.Cells(j).Text.ToString())
            'Next
            For col = 1 To 28 - 1

                If colToJ.ContainsKey(col) Then
                    j = colToJ(col)
                End If
                'If j = 6 Then
                If col = 6 Then
                    'j = j + 1 
                    If Form1.dgv2.Rows(i).Cells(j).Value.ToString() = "&nbsp;" Then
                        wSheetGemBox.Columns(col).Cells(i + 10).Value = ""
                    Else
                        app = "C" + Convert.ToString(i + 11) + ":C" + Convert.ToString(i + 11)
                        wSheetGemBox.Columns(col).Cells(i + 10).Value = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                        wSheetGemBox.Cells(i + 10, col).Style.HorizontalAlignment = HorizontalAlignmentStyle.Center
                        Dim verifica As Integer
                        Dim resultVerifica As Boolean
                        resultVerifica = Int16.TryParse(Form1.dgv2.Rows(i).Cells(j - 1).Value.ToString(), verifica)

                        'If Convert.ToInt16(gv_monthlyrep.SelectedRow.Cells(j - 1).Text) = 1 Then
                        If resultVerifica And verifica = 1 Then
                            wSheetGemBox.Cells.GetSubrange(app).Style.Font.Weight = Convert.ToInt16(Form1.dgv2.Rows(i).Cells(j - 1).Value.ToString())
                        End If
                    End If
                    'col = col + 1
                ElseIf col = 10 Then
                    'j = j + 1 
                    If Form1.dgv2.Rows(i).Cells(j).Value.ToString() = "&nbsp;" Then
                        wSheetGemBox.Columns(col).Cells(i + 10).Value = ""
                    Else
                        app = "E" + Convert.ToString(i + 11) + ":E" + Convert.ToString(i + 11)
                        wSheetGemBox.Columns(col).Cells(i + 10).Value = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                        wSheetGemBox.Cells(i + 10, col).Style.HorizontalAlignment = HorizontalAlignmentStyle.Center
                        'wSheetGemBox.Cells.GetSubrange(app).Style.Font.Weight = Convert.ToInt16(gv_monthlyrep.SelectedRow.Cells(j - 1).Text)

                    End If
                    'col = col + 1
                Else
                    If Form1.dgv2.Rows(i).Cells(j).Value.ToString() = "&nbsp;" Then
                        wSheetGemBox.Columns(col).Cells(i + 10).Value = ""
                    Else
                        wSheetGemBox.Columns(col).Cells(i + 10).Value = Form1.dgv2.Rows(i).Cells(j).Value.ToString()
                        wSheetGemBox.Cells(i + 10, col).Style.HorizontalAlignment = HorizontalAlignmentStyle.Center
                    End If
                    'col = col + 1
                End If
            Next
        Next

        col = 2
        insert_tab = i + 35
        For z = 0 To Form1.dgv.Rows.Count - 1


            colgv = 5
            For tabcounter = 2 To Form1.dgv.Columns.Count
                If Form1.dgv.Rows(z).Cells(tabcounter - 1).Value.ToString() = "&nbsp;" Then
                    wSheetGemBox.Columns(insert_tab).Cells(colgv).Value = ""
                Else
                    wSheetGemBox.Columns(colgv - 1).Cells(insert_tab - 1).Value = Form1.dgv.Rows(z).Cells(tabcounter - 1).Value.ToString()

                End If
                colgv += 1
            Next
            insert_tab = insert_tab + 1
            col = col + 1
        Next

        ComboStatus.Report(State.FinishedReport)
        Dim reportFileXls = reportTitle & ".xls"
        Dim reportFilePdf = reportTitle & ".pdf"
        Dim reportPath = Path.Combine(reportDir, reportFileXls)
        Dim reportPathPdf = Path.Combine(reportDir, reportFilePdf)
        wBookGemBox.Save(reportPath)
        wBookGemBox.Save(reportPathPdf)
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("it-IT")

        If (startDate = endDate) Then

            ComboStatus.Report(State.Finished)
            ShowCompletionDialog()
        End If

    End Sub

    Private Sub downloadDailyReportCamini(ComboStatus As Progress(Of Integer), startDate As Date, endDate As Date, reportDir As String)

        Dim exePath As String = Application.StartupPath
        Dim rootPath As String = Directory.GetParent(Directory.GetParent(exePath).FullName).FullName
        Dim templateDir As String = ConfigurationManager.AppSettings("TemplateDirectory")
        Dim reportTitle As String = ""
        Dim methodName As String = GetCurrentMethod()
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")
        Dim wBookGemBox As ExcelFile = ExcelFile.Load(Path.Combine(rootPath, templateDir, "152_CONC_GIORNOTG.xls")) 'Viene caricato il template Excel dal percorso specificato.
        Dim wSheetGemBox As ExcelWorksheet = wBookGemBox.Worksheets(0) 'Viene ottenuto il riferimento al primo foglio di lavoro del file Excel caricato.

        ComboStatus.Report(State.TableLoading)

        Dim first_row As Integer = 10
        Dim first_col As Integer = 1

        'test: Dim GV_DAILYREP_ROWS_COUNT = 10


        ' Set values and styles for named ranges
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("F1"), "152_CONC_GIORNO", True)
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("F2"), "ENIPOWER - Centrale di Brindisi - Sezione Termoelettrica n� " & Form1.section.ToString(), True)
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("F3"), "Sistema di Misura delle Emissioni", True)
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("F4"), "Report Giornaliero concentrazioni medie orarie e giornaliera (Nox,Co) - Dlgs 152 (70%)", True)
        Dim startDateFormatted As DateTime = DateTime.Parse(startDate).Date
        SetRangeValueAndStyle(wSheetGemBox.Cells.GetSubrange("F5"), "Report Giornaliero del " & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:dd/MMMM/yyyy}", startDateFormatted), True)
        SetRangeBorderStyle(wSheetGemBox.Cells.GetSubrange("F1:P5"))
        SetRangeValueAndStyle(wSheetGemBox.NamedRanges.Item("HNF").Range, hnf, True)
        SetRangeValueAndStyle(wSheetGemBox.NamedRanges.Item("HTRANS").Range, htran, True)
        reportTitle = "152_CONC_GIORNOCC_" & Form1.section.ToString() & "_" & String.Format(New System.Globalization.CultureInfo("it-IT"), "{0:dd_MMMM_yyyy}", Date.Parse(startDate))
        Dim GV_DAILYREP_ROWS_COUNT = Form1.dgv2.Rows.Count

        Dim howmanyrows = If(GV_DAILYREP_ROWS_COUNT - 2 >= 0, GV_DAILYREP_ROWS_COUNT - 2, 0)

        wSheetGemBox.Rows.InsertCopy(first_row, howmanyrows, wSheetGemBox.Rows(first_row))

        ComboStatus.Report(State.SheetLoading)

        Dim row_offset As Integer = 0
        Dim EMPTY_STRING As String = "&nbsp;"

        ' warning: for loops in vbnet are inclusive that's why we have to subtract 2 from the column number
        For row_offset = 0 To (GV_DAILYREP_ROWS_COUNT - 1)

            For this_col = first_col To GetColumnNumber("AI")  ' Loop through columns B to AI 
                Dim current_cell = wSheetGemBox.Cells(first_row + row_offset, this_col)
                current_cell.Style.Borders.SetBorders(MultipleBorders.Outside, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin)
                current_cell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center
                current_cell.Style.VerticalAlignment = VerticalAlignmentStyle.Center
            Next

            Dim grid_offset = 2
            Dim col_offset = 1
            Dim gd_col = 0
            Dim gv_col = 0
            Dim skipped_cols = 0


            While gd_col < Form1.dgv2.Rows(row_offset).Cells.Count - 1
                Dim current_cell_col = gv_col + col_offset - skipped_cols
                Dim current_cell_row = first_row + row_offset
                gd_col = gv_col + grid_offset
                Dim current_cell = wSheetGemBox.Cells(current_cell_row, current_cell_col)
                Dim current_data = Form1.dgv2.Rows(row_offset).Cells(gd_col).Value.ToString()
                System.Diagnostics.Debug.WriteLine("current data:" & current_data)
                'handle special cases:
                If (gd_col = 7) Then
                    skipped_cols += 1
                    Dim cell_to_style = String.Format("C{0}", current_cell_row)
                    Try
                        If Convert.ToInt16(Form1.dgv2.Rows(row_offset).Cells(gd_col).Value.ToString()) = 1 Then
                            wSheetGemBox.Cells.GetSubrange(cell_to_style).Style.Font.Weight = ExcelFont.BoldWeight
                        End If
                    Catch ex As Exception
                        Logger.LogWarning("[" & methodName & "]" & " Errore nella lettura dei dati: ", ex)
                    End Try



                ElseIf (gd_col = 12) Then
                    skipped_cols += 1
                    Dim cell_to_style = String.Format("F{0}", current_cell_row)
                    current_cell.Value = current_data
                    Try
                        If Convert.ToInt16(Form1.dgv2.Rows(row_offset).Cells(gd_col).Value.ToString()) = 1 Then
                            wSheetGemBox.Cells.GetSubrange(cell_to_style).Style.Font.Weight = ExcelFont.BoldWeight
                        End If
                    Catch ex As Exception
                        Logger.LogWarning("[" & methodName & "]" & " Errore nella lettura dei dati: ", ex)
                    End Try


                ElseIf current_data = EMPTY_STRING Then
                    current_cell.Value = String.Empty
                Else
                    current_cell.Value = current_data
                End If

                gv_col += 1

            End While
        Next row_offset
        For current_row = 0 To GV_DAILYREP_ROWS_COUNT - 2
            Dim rowIndex = current_row + first_row
            Dim cellToTest = wSheetGemBox.Cells(rowIndex, GetColumnNumber("U"))

            'wSheet.Rows(rowIndex).Height = 600
            Dim testIsPassed As Boolean = cellToTest.Value <> Nothing AndAlso (cellToTest.Value = 31 Or cellToTest.Value = 32)

            If (testIsPassed) Then
                For this_col = GetColumnNumber("B") To GetColumnNumber("AI")
                    wSheetGemBox.Cells(rowIndex, this_col).Style.Font.Weight = ExcelFont.BoldWeight
                Next this_col
            End If
        Next current_row



        'populating the report
        Dim report_start_row = 38 + GV_DAILYREP_ROWS_COUNT - 1  'skip header and empty rows
        Dim report_start_col = 2 'skip the first column


        For this_row = report_start_row To report_start_row + Form1.dgv.Rows.Count - 1

            For this_col = report_start_col + 1 To Form1.dgv.Columns.Count - 1
                Dim current_data = Form1.dgv.Rows(this_row - report_start_row).Cells(this_col - report_start_col)
                Dim current_cell = wSheetGemBox.Cells(this_row, this_col + 1)

                If current_data.Value.ToString() = EMPTY_STRING Then

                    current_cell.Value = String.Empty
                Else
                    current_cell.Value = current_data.Value.ToString()

                End If
            Next
        Next

        SetRangeValue(wSheetGemBox, "DATA_QAL2_CO", Form1.section.ToString(), "CO")
        SetRangeValue(wSheetGemBox, "DATA_QAL2_NO", Form1.section.ToString(), "NOX")
        SetRangeValue(wSheetGemBox, "DATA_QAL2_O2", Form1.section.ToString(), "O2")



        SetCoefficientValue(wSheetGemBox, "A_QAL2_CO", Form1.section.ToString(), "COEFF_A")
        SetCoefficientValue(wSheetGemBox, "A_QAL2_NO", Form1.section.ToString(), "COEFF_A")
        SetCoefficientValue(wSheetGemBox, "A_QAL2_O2", Form1.section.ToString(), "COEFF_A")



        SetCoefficientValue(wSheetGemBox, "B_QAL2_CO", Form1.section.ToString(), "COEFF_B")
        SetCoefficientValue(wSheetGemBox, "B_QAL2_NO", Form1.section.ToString(), "COEFF_B")
        SetCoefficientValue(wSheetGemBox, "B_QAL2_O2", Form1.section.ToString(), "COEFF_B")


        SetRangeValueWithMultiplier(wSheetGemBox, "RANGE_QAL2_CO", Form1.section.ToString(), "Y_MAX", 1.1)
        SetRangeValueWithMultiplier(wSheetGemBox, "RANGE_QAL2_NO", Form1.section.ToString(), "Y_MAX", 1.1)

        ComboStatus.Report(State.FinishedReport)
        QAL2.Clear()
        QAL2.Columns.Clear()
        Dim reportFileXls = reportTitle & ".xls"
        Dim reportFilePdf = reportTitle & ".pdf"
        Dim reportPath = Path.Combine(reportDir, reportFileXls)
        Dim reportPathPdf = Path.Combine(reportDir, reportFilePdf)
        wBookGemBox.Save(reportPath)
        wBookGemBox.Save(reportPathPdf)

        If (startDate = endDate) Then
            ComboStatus.Report(State.Finished)
            ShowCompletionDialog()
        End If

    End Sub

    Sub SetRangeValue(wSheet As ExcelWorksheet, rangeName As String, Sezione As String, suffix As String)

        Dim qal_select = QAL2.Select(String.Format("CHINAM LIKE 'SME{0}_{1}'", Sezione, suffix)).FirstOrDefault()

        If qal_select IsNot Nothing Then
            Dim result = String.Format("{0:dd/MM/yyyy hh:mm}", qal_select.Item("Data_INI"))
            wSheet.NamedRanges.Item(rangeName).Range.Value = result
        Else
            wSheet.NamedRanges.Item(rangeName).Range.Value = String.Empty
        End If
    End Sub

    Sub SetCoefficientValue(wSheet As ExcelWorksheet, rangeName As String, Sezione As String, columnName As String)
        Dim result = QAL2.Select(String.Format("CHINAM LIKE 'SME{0}_{1}'", Sezione, columnName)).FirstOrDefault()

        If result IsNot Nothing Then
            wSheet.NamedRanges.Item(rangeName).Range.Value = result.Item(columnName)

        Else
            wSheet.NamedRanges.Item(rangeName).Range.Value = String.Empty
        End If

    End Sub

    Sub SetRangeValueWithMultiplier(wSheet As ExcelWorksheet, rangeName As String, Sezione As String, columnName As String, multiplier As Double)
        Dim result = QAL2.Select(String.Format("CHINAM LIKE 'SME{0}_{1}'", Sezione, columnName)).FirstOrDefault()
        If result IsNot Nothing Then
            wSheet.NamedRanges.Item(rangeName).Range.Value = DirectCast(result.Item(columnName), Double) * multiplier
        Else
            wSheet.NamedRanges.Item(rangeName).Range.Value = String.Empty
        End If
    End Sub

    Private Function GetColumnNumber(columnLetter As String) As Integer

        Dim base As Integer = Asc("A") - 1
        Dim result As Integer = 0

        For Each c As Char In columnLetter.ToUpper()
            result = result * 26 + Asc(c) - base
        Next

        Return result - 1

    End Function

    Private Function GetChimneyName(ByVal param As Integer) As String
        Select Case (param)
            Case 1
                Return "CC1"
            Case 2
                Return "CC2"
            Case 3
                Return "CC3"
        End Select
        Return "CC"
    End Function

    Function FillQAL2() As Byte

        Dim connection As New SqlConnection(connectionString)
        Dim command As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim methodName As String = GetCurrentMethod()
        Dim ret As Byte = 1
        QAL2.Columns.Add(New Data.DataColumn("Data_INI", GetType(DateTime)))
        QAL2.Columns.Add(New Data.DataColumn("CHINAM", GetType(String)))
        QAL2.Columns.Add(New Data.DataColumn("COEFF_A", GetType(Double)))
        QAL2.Columns.Add(New Data.DataColumn("COEFF_B", GetType(Double)))
        QAL2.Columns.Add(New Data.DataColumn("Y_MAX", GetType(Double)))


        Dim log_statement As String = "SELECT * FROM [Coeff_X_Validita] WHERE DATA_FIN IS NULL"

        connection.Open()
        Try
            command = New System.Data.SqlClient.SqlCommand(log_statement, connection)
        Catch ex As SqlException
            Logger.LogError("[" & methodName & "]" & " Errore durante l'esecuzione della query: ", ex)
            ret = 0
            Return ret
        End Try

        reader = command.ExecuteReader()

        Dim dr As Data.DataRow = QAL2.NewRow()
        While reader.Read()
            Try
                dr("Data_INI") = reader("Data_INI")
                dr("CHINAM") = reader("CHINAM")
                dr("COEFF_A") = reader("COEFF_A")
                dr("COEFF_B") = reader("COEFF_B")
                dr("Y_MAX") = reader("Y_MAX")
                QAL2.Rows.Add(dr)
                dr = QAL2.NewRow()
            Catch ex As Exception
                Logger.LogWarning("[" & methodName & "]" & " Errore nella lettura dei dati: ", ex)
                Continue While
            End Try
            
        End While

        connection.Close()

        Return ret

    End Function

    Protected Sub SetRangeValueAndStyle(ByRef range As Object, ByVal value As String, ByVal bold As Boolean)
        range.Value = value
        range.Style.Font.Weight = ExcelFont.MaxWeight
        If (bold) Then
            range.Style.Font.Weight = ExcelFont.BoldWeight
        End If
    End Sub

    Protected Sub SetRangeBorderStyle(ByRef range As Object)
        range.Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin)
    End Sub

    Private Sub ShowCompletionDialog()                                                                                  ' Crea un'istanza del form modale e la mostra in modalità                                       

        Dim completedDownloadForm As New Form2()
        completedDownloadForm.ShowDialog()

    End Sub


    Private Function GetCurrentMethod() As String
        Dim stackTrace As New StackTrace()
        Dim method As MethodBase = stackTrace.GetFrame(1).GetMethod()
        Return method.Name
    End Function

    Private Sub preRenderCaminiTable(dt As DataTable)

        For Each row As DataRow In dt.Rows
            ' Controlla il valore della colonna 5 (indice 4, dato che è zero-based)
            If row(5).ToString() = "1" Then
                ' Applica una formattazione simulata, ad esempio puoi impostare un flag o modificare i dati
                ' Poiché non hai una DataGridView visibile, potresti aggiungere una colonna per un flag di formattazione
                row(3) = "Bold Text"  ' Simulazione di formattazione, qui puoi anche solo gestire il dato
            End If

            ' Controlla il valore della colonna 8 (indice 7)
            If row(8).ToString() = "1" Then
                ' Anche qui puoi applicare una logica simile o un flag di formattazione
                row(6) = "Bold Text"
            End If
        Next

    End Sub

    Private Sub InitCorrectionDict()

        If colToJ Is Nothing Then
            colToJ = New Dictionary(Of Integer, Integer)
            colToJ.Add(1, 2)
            colToJ.Add(2, 5)
            colToJ.Add(3, 3)
            colToJ.Add(4, 4)
            colToJ.Add(5, 9)
            colToJ.Add(6, 7)
            colToJ.Add(7, 8)
            colToJ.Add(8, 14)
            colToJ.Add(9, 11)
            colToJ.Add(10, 12)
            colToJ.Add(11, 17)
            colToJ.Add(12, 15)
            colToJ.Add(13, 20)
            colToJ.Add(14, 18)
            colToJ.Add(15, 21)
            colToJ.Add(16, 24)
            colToJ.Add(17, 22)
            colToJ.Add(18, 27)
            colToJ.Add(19, 25)
            colToJ.Add(20, 29)
            colToJ.Add(21, 28)
            colToJ.Add(22, 31)
            colToJ.Add(23, 30)
            colToJ.Add(24, 33)
            colToJ.Add(25, 32)
            colToJ.Add(26, 34)
            colToJ.Add(27, 36)
        End If


    End Sub

    Private Function GetComboBoxSelectedIndex(frm As Form1, comboBox As ComboBox) As Integer
        If comboBox.InvokeRequired Then
            Return CInt(frm.Invoke(New Func(Of Form1, ComboBox, Integer)(AddressOf GetComboBoxSelectedIndex), frm, comboBox))
        Else
            Return comboBox.SelectedIndex
        End If
    End Function

    Private Function GetComboBoxSelectedItem(frm As Form1, comboBox As ComboBox) As String
        If comboBox.InvokeRequired Then
            Return CStr(frm.Invoke(New Func(Of Form1, ComboBox, String)(AddressOf GetComboBoxSelectedItem), frm, comboBox))
        Else
            Return CStr(comboBox.SelectedItem)
        End If
    End Function

    Private Function GetComboBoxSelectedItemFromIndex(frm As Form1, comboBox As ComboBox, indx As Integer) As String
        If comboBox.InvokeRequired Then
            Return CStr(frm.Invoke(New Func(Of Form1, ComboBox, Integer, String)(AddressOf GetComboBoxSelectedItemFromIndex), frm, comboBox, indx))
        Else
            Return CStr(comboBox.Items(indx).ToString())
        End If
    End Function

    Private Sub UpdateProgressBarStatus(frm As Form1, visibility As Boolean)
        If frm.ProgressBar1.InvokeRequired Then
            frm.Invoke(New Action(Of Form1, Boolean)(AddressOf UpdateProgressBarStatus), frm, visibility)
        Else
            frm.ProgressBar1.Visible = visibility
        End If
    End Sub

    Private Sub UpdateTextBoxStatus(frm As Form1, txtbox As TextBox, visibility As Boolean)
        If txtbox.InvokeRequired Then
            frm.Invoke(New Action(Of Form1, TextBox, Boolean)(AddressOf UpdateTextBoxStatus), frm, txtbox, visibility)
        Else
            txtbox.Visible = visibility
        End If
    End Sub



    Private Sub UpdateProgressBarValue(frm As Form1, value As Integer)
        If frm.ProgressBar1.InvokeRequired Then
            frm.Invoke(New Action(Of Form1, Integer)(AddressOf UpdateProgressBarValue), frm, value)
        Else
            frm.ProgressBar1.Value = value
        End If
    End Sub

    Private Sub UpdateTextBoxText(frm As Form1, txb As TextBox, text As String)
        If txb.InvokeRequired Then
            frm.Invoke(New Action(Of Form1, TextBox, String)(AddressOf UpdateTextBoxText), frm, txb, text)
        Else
            txb.Text = text
        End If
    End Sub

    Private Sub UpdateDgvDataSource(ds As DataTable, dataTable As DataGridView)
        If dataTable.InvokeRequired Then
            dataTable.Invoke(New Action(Of DataTable, DataGridView)(AddressOf UpdateDgvDataSource), ds, dataTable)
        Else
            dataTable.DataSource = ds
        End If
    End Sub

    Private Sub UpdateDgvColumnVisibility(visibility As Boolean, dgv As DataGridView, col As String)
        If dgv.InvokeRequired Then
            dgv.Invoke(New Action(Of Boolean, DataGridView, String)(AddressOf UpdateDgvColumnVisibility), visibility, dgv, col)
        Else
            dgv.Columns(col).Visible = visibility
        End If
    End Sub

    Private Sub EnableFormSafe(container As Control)
        If container.InvokeRequired Then
            container.Invoke(New Action(Of Control)(AddressOf EnableFormSafe), container)
        Else
            EnableControls(container)
            ResetForm()
        End If
    End Sub

    Private Sub EnableControls(container As Control)
        For Each ctrl As Control In container.Controls
            If Not ctrl.Equals(Form1.dgv) And Not ctrl.Equals(Form1.dgv2) And Not ctrl.Enabled Then
                ctrl.Enabled = True
            End If

            ' Ricorsione: se il controllo corrente contiene altri controlli, chiama EnableControls su di essi
            If ctrl.HasChildren Then
                EnableControls(ctrl)
            End If
        Next
    End Sub

    Private Sub HideFormSafe(control As Control)
        If control.InvokeRequired Then
            control.Invoke(New Action(Of Control)(AddressOf HideFormSafe), control)
        Else
            control.Hide()
        End If
    End Sub

    Private Sub UpdateComboBoxIndex(frm As Form1, cmb As ComboBox, index As Integer)
        If cmb.InvokeRequired Then
            frm.Invoke(New Action(Of Form1, ComboBox, Integer)(AddressOf UpdateComboBoxIndex), frm, cmb, index)
        Else
            cmb.SelectedIndex = index
        End If
    End Sub

    Private Sub UpdateTimePicker(frm As Form1, dtp As DateTimePicker, data As Date)
        If dtp.InvokeRequired Then
            frm.Invoke(New Action(Of Form1, DateTimePicker, Date)(AddressOf UpdateTimePicker), frm, dtp, data)
        Else
            dtp.Value = data
        End If
    End Sub

    Private Sub UpdateButtonStatus(frm As Form1, btn As Button, visibility As Boolean)
        If btn.InvokeRequired Then
            frm.Invoke(New Action(Of Form1, Button, Boolean)(AddressOf UpdateButtonStatus), frm, btn, visibility)
        Else
            btn.Visible = visibility
        End If
    End Sub

    Private Sub ResetForm()

        Dim formInstance As Form1 = DirectCast(Application.OpenForms("Form1"), Form1)
        formInstance.ProgressBar1.Visible = False                                                                                      'RESETTARE BOLLA A 254
        UpdateComboBoxIndex(formInstance, formInstance.ComboBox1, 0)
        UpdateComboBoxIndex(formInstance, formInstance.ComboBox2, 0)
        UpdateTimePicker(formInstance, formInstance.DateTimePicker2, Date.Now)
        UpdateTimePicker(formInstance, formInstance.DateTimePicker1, Date.Now.AddYears(-1))
        UpdateButtonStatus(formInstance, formInstance.Button1, True)
        formInstance.TextBox1.Text = "Data Loading..."
        UpdateTextBoxStatus(formInstance, formInstance.TextBox1, False)

    End Sub

End Class
